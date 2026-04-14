#Requires -RunAsAdministrator
<#
.SYNOPSIS
    VDI Golden Image Health Check Report — Generic / Interactive

.DESCRIPTION
    Generates a timestamped HTML health report for any Horizon VDI environment.
    Prompts the user for all environment-specific values at runtime.
   
    Checks covered:
      - System baseline (OS, domain, activation, disk, sysprep)
      - AppX provisioned vs per-user drift + sysprep blockers
      - DCOM 10016 audit
      - NVIDIA vGPU driver version + Control Panel Client
      - Horizon Agent services
      - App Volumes Agent + Manager connectivity
      - VMware Tools
      - Windows Store state
      - Shell folder integrity
      - Guest Customization log tail
      - Event log critical errors since last boot
      - TPM / Secure Boot / AV
      - Network (NIC, DNS, DC, optional servers)

.PARAMETER Silent
    Skip interactive prompts and use only auto-detected values.
    Connectivity checks that require FQDNs will be skipped.

.PARAMETER OutputPath
    Override output folder. If not provided, user is prompted.

.EXAMPLE
    .\Invoke-VDIHealthReport.ps1
    .\Invoke-VDIHealthReport.ps1 -Silent
    .\Invoke-VDIHealthReport.ps1 -OutputPath "D:\Reports"

.NOTES
    Version : 1.1
    Generic  : No environment-specific values hardcoded
#>

[CmdletBinding()]
param(
    [switch]$Silent,
    [string]$OutputPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "SilentlyContinue"

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║          VDI GOLDEN IMAGE HEALTH CHECK REPORT           ║" -ForegroundColor Cyan
    Write-Host "  ║          Omnissa Horizon  |  Generic Edition            ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Prompt-Input {
    param(
        [string]$Prompt,
        [string]$Default = "",
        [switch]$Optional
    )
    $label = if ($Default) { "$Prompt [default: $Default]" } `
             elseif ($Optional) { "$Prompt [press Enter to skip]" } `
             else { $Prompt }
    Write-Host "  ► " -NoNewline -ForegroundColor Cyan
    Write-Host $label -NoNewline -ForegroundColor White
    Write-Host ": " -NoNewline
    $val = Read-Host
    if ([string]::IsNullOrWhiteSpace($val)) { return $Default }
    return $val.Trim()
}

function Write-Section {
    param([string]$Text)
    Write-Host ""
    Write-Host "  ─── $Text " -ForegroundColor Yellow
}

# ─────────────────────────────────────────────────────────────────────────────
# INTERACTIVE SETUP
# ─────────────────────────────────────────────────────────────────────────────
Write-Banner

if (-not $Silent) {
    Write-Host "  This script prompts for environment details before running." -ForegroundColor Gray
    Write-Host "  Fields with defaults: press Enter to accept." -ForegroundColor Gray
    Write-Host "  Optional fields: press Enter to skip." -ForegroundColor Gray

    Write-Section "ENVIRONMENT IDENTITY"
    $EnvName           = Prompt-Input "Environment / Pool Name (e.g. Production, Dev, QA)" -Default "VDI"
    $PoolType          = Prompt-Input "Pool Type (FullClone / InstantClone / LinkedClone)"  -Default "FullClone"

    Write-Section "OUTPUT"
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath    = Prompt-Input "Report output folder" -Default "C:\Temp\VDIHealthReports"
    }

    Write-Section "NVIDIA vGPU"
    $ExpectedDriverVer = Prompt-Input "Expected vGPU guest driver version" -Default "582.16"

    Write-Section "APP VOLUMES MANAGER"
    $AppVolManagerFQDN = Prompt-Input "App Volumes Manager FQDN or IP" -Optional
    $AppVolManagerPort = if ($AppVolManagerFQDN) {
                             Prompt-Input "App Volumes Manager port" -Default "443"
                         } else { "443" }

    Write-Section "HORIZON CONNECTION SERVER (optional)"
    $HorizonCS         = Prompt-Input "Horizon Connection Server FQDN or IP" -Optional

    Write-Section "ADDITIONAL CONNECTIVITY CHECKS (optional)"
    $ExtraHost1        = Prompt-Input "Additional host to test (e.g. file server, DNS)" -Optional
    $ExtraHost2        = Prompt-Input "Additional host #2" -Optional

    Write-Host ""
    Write-Host "  ── Configuration confirmed. Starting health checks..." -ForegroundColor Green
    Write-Host ""

} else {
    $EnvName           = "VDI"
    $PoolType          = "FullClone"
    $ExpectedDriverVer = "582.16"
    $AppVolManagerFQDN = ""
    $AppVolManagerPort = "443"
    $HorizonCS         = ""
    $ExtraHost1        = ""
    $ExtraHost2        = ""
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = "C:\Temp\VDIHealthReports" }
    Write-Host "  [Silent mode] Running with auto-detected values..." -ForegroundColor Yellow
}

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT SETUP
# ─────────────────────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = Join-Path $OutputPath "VDIHealth_$($env:COMPUTERNAME)_$Timestamp.html"
$RunTime    = Get-Date -Format "dddd, MMMM dd yyyy  HH:mm:ss"

# ─────────────────────────────────────────────────────────────────────────────
# REPORT HELPERS
# ─────────────────────────────────────────────────────────────────────────────
function Get-StatusBadge {
    param([string]$Status)
    switch ($Status) {
        "PASS" { return '<span class="badge pass">PASS</span>' }
        "FAIL" { return '<span class="badge fail">FAIL</span>' }
        "WARN" { return '<span class="badge warn">WARN</span>' }
        "INFO" { return '<span class="badge info">INFO</span>' }
        "SKIP" { return '<span class="badge skip">SKIP</span>' }
        default { return '<span class="badge info">UNKN</span>' }
    }
}

function New-CheckRow {
    param(
        [string]$Check,
        [string]$Status,
        [string]$Detail,
        [string]$Recommendation = ""
    )
    $badge = Get-StatusBadge -Status $Status
    $rec   = if ($Recommendation) { "<div class='rec'>$Recommendation</div>" } else { "" }
    return "<tr class='row-$($Status.ToLower())'><td class='check-name'>$Check</td><td>$badge</td><td>$Detail$rec</td></tr>"
}

function Test-ServiceStatus {
    param([string]$ServiceName, [string]$ExpectedState = "Running")
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) { return @{ Status="FAIL"; Detail="Service not found: $ServiceName" } }
    if ($svc.Status -eq $ExpectedState) {
        return @{ Status="PASS"; Detail="$ServiceName — $($svc.Status)" }
    }
    return @{ Status="FAIL"; Detail="$ServiceName — Expected: $ExpectedState | Actual: $($svc.Status)" }
}

function Test-PortConnection {
    param([string]$HostName, [int]$Port, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($HostName)) {
        return @{ Status="SKIP"; Detail="$Label — no host provided"; Recommendation="" }
    }
    $result = Test-NetConnection -ComputerName $HostName -Port $Port -InformationLevel Quiet -WarningAction SilentlyContinue
    return @{
        Status         = if ($result) {"PASS"} else {"FAIL"}
        Detail         = "$Label — ${HostName}:${Port} — $(if ($result) {'Reachable'} else {'UNREACHABLE'})"
        Recommendation = if (-not $result) {"Check DNS resolution and firewall rules for ${HostName}:${Port}"} else {""}
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — SYSTEM BASELINE
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [1/10] System Baseline..." -ForegroundColor Cyan
$rows_system = @()

$os        = Get-CimInstance Win32_OperatingSystem
$cs        = Get-CimInstance Win32_ComputerSystem
$buildNum  = [int]$os.BuildNumber
$uptime    = (Get-Date) - $os.LastBootUpTime
$ramGB     = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
$disk      = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$diskFree  = [math]::Round($disk.FreeSpace / 1GB, 1)
$diskTotal = [math]::Round($disk.Size / 1GB, 1)
$diskPct   = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 0)

$rows_system += New-CheckRow "OS Version"        "INFO" "$($os.Caption) — Build $buildNum | $($os.OSArchitecture)"
$rows_system += New-CheckRow "Hostname"          "INFO" $env:COMPUTERNAME
$rows_system += New-CheckRow "Environment"       "INFO" "Name: $EnvName | Pool Type: $PoolType"

$bStatus = if ($buildNum -ge 26100) {"PASS"} else {"WARN"}
$rows_system += New-CheckRow "Windows 11 24H2 Build" $bStatus "Build $buildNum  (24H2 minimum = 26100)" `
    $(if ($bStatus -eq "WARN") {"Build is below 24H2 baseline. Verify ISO source."} else {""})

$dStatus = if ($cs.PartOfDomain) {"PASS"} else {"WARN"}
$rows_system += New-CheckRow "Domain Join" $dStatus `
    $(if ($cs.PartOfDomain) {"Joined: $($cs.Domain)"} else {"NOT domain joined"}) `
    $(if (-not $cs.PartOfDomain) {"Verify Guest Customization completed successfully."} else {""})

$lic = Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%' AND LicenseStatus=1" -ErrorAction SilentlyContinue | Select-Object -First 1
$aStatus = if ($lic) {"PASS"} else {"WARN"}
$rows_system += New-CheckRow "Windows Activation" $aStatus `
    $(if ($lic) {"Activated — $($lic.Name)"} else {"Not activated or unable to verify"}) `
    $(if (-not $lic) {"Check KMS connectivity or MAK activation."} else {""})

$rows_system += New-CheckRow "Last Boot / Uptime" "INFO" `
    "Booted: $($os.LastBootUpTime.ToString('MM/dd/yyyy HH:mm:ss'))  |  Uptime: $([math]::Floor($uptime.TotalHours))h $($uptime.Minutes)m"

$ramStatus = if ($ramGB -ge 4) {"PASS"} else {"WARN"}
$rows_system += New-CheckRow "Physical RAM" $ramStatus "$ramGB GB" `
    $(if ($ramGB -lt 4) {"Minimum 4GB recommended for VDI desktop."} else {""})

$diskStatus = if ($diskPct -ge 20) {"PASS"} elseif ($diskPct -ge 10) {"WARN"} else {"FAIL"}
$rows_system += New-CheckRow "Disk C: Free Space" $diskStatus "$diskFree GB free of $diskTotal GB ($diskPct% free)" `
    $(if ($diskStatus -eq "FAIL") {"Critical — less than 10% free. Image may be oversized."} `
      elseif ($diskStatus -eq "WARN") {"Consider cleanup or expanding template disk."} else {""})

$setupKey  = Get-ItemProperty "HKLM:\SYSTEM\Setup" -ErrorAction SilentlyContinue
$oobe      = $setupKey.OOBEInProgress
$setupType = $setupKey.SetupType
$spStatus  = if ($oobe -eq 0 -and $setupType -eq 0) {"PASS"} else {"WARN"}
$rows_system += New-CheckRow "Sysprep / OOBE State" $spStatus `
    "OOBEInProgress=$oobe | SetupType=$setupType  (both must be 0 on deployed clone)" `
    $(if ($spStatus -eq "WARN") {"VM may still be in generalize/OOBE state. Verify Guest Customization completed."} else {""})

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — VMWARE TOOLS + HORIZON AGENT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [2/10] VMware Tools + Horizon Agent..." -ForegroundColor Cyan
$rows_horizon = @()

$toolsResult = Test-ServiceStatus -ServiceName "VMTools"
$rows_horizon += New-CheckRow "VMware Tools Service" $toolsResult.Status $toolsResult.Detail `
    $(if ($toolsResult.Status -eq "FAIL") {"VMware Tools required for Guest Customization and Horizon."} else {""})

$toolsReg = Get-ItemProperty "HKLM:\SOFTWARE\VMware, Inc.\VMware Tools" -ErrorAction SilentlyContinue
$toolsVer  = if ($toolsReg) {$toolsReg.Version} else {"Not found"}
$rows_horizon += New-CheckRow "VMware Tools Version" $(if ($toolsReg) {"INFO"} else {"FAIL"}) $toolsVer

$horizonSvcs = @(
    @{ Name="vmware-viewagent";  Label="Horizon View Agent";         Required=$true  },
    @{ Name="wsnm";              Label="Horizon WSNM (Blast/PCoIP)"; Required=$true  },
    @{ Name="CSVD";              Label="Horizon CSVD (USB/Scanner)"; Required=$false },
    @{ Name="vmware-view-usbd"; Label="Horizon USB Arbitrator";     Required=$false }
)
foreach ($svc in $horizonSvcs) {
    $r = Test-ServiceStatus -ServiceName $svc.Name
    $finalStatus = if ($r.Status -eq "FAIL" -and -not $svc.Required) {"WARN"} else {$r.Status}
    $rows_horizon += New-CheckRow $svc.Label $finalStatus $r.Detail
}

$hReg = Get-ItemProperty "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\Agent" -ErrorAction SilentlyContinue
$hVer = if ($hReg -and $hReg.ProductVersion) {$hReg.ProductVersion} else {"Not found in registry"}
$rows_horizon += New-CheckRow "Horizon Agent Version" $(if ($hReg) {"INFO"} else {"WARN"}) $hVer `
    $(if (-not $hReg) {"Verify Horizon Agent installed correctly."} else {""})

$gcLog = "C:\Windows\Temp\vmware-imc\guestcust.log"
if (Test-Path $gcLog) {
    $gcTail   = (Get-Content $gcLog -Tail 5) -join " | "
    $gcStatus = if ($gcTail -match "success|Customization finished") {"PASS"} `
                elseif ($gcTail -match "error|fail") {"FAIL"} else {"INFO"}
    $rows_horizon += New-CheckRow "Guest Customization Log (last 5 lines)" $gcStatus $gcTail `
        $(if ($gcStatus -eq "FAIL") {"Review full log: C:\Windows\Temp\vmware-imc\guestcust.log"} else {""})
} else {
    $rows_horizon += New-CheckRow "Guest Customization Log" "WARN" "Not found: $gcLog" `
        "Log absent may indicate customization never ran or VM was not deployed via Horizon pool."
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — APP VOLUMES AGENT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [3/10] App Volumes Agent..." -ForegroundColor Cyan
$rows_appvol = @()

$svResult = Test-ServiceStatus -ServiceName "svservice"
$rows_appvol += New-CheckRow "App Volumes Agent (svservice)" $svResult.Status $svResult.Detail `
    $(if ($svResult.Status -eq "FAIL") {"svservice not running. Reinstall App Volumes Agent."} else {""})

$avReg = Get-ItemProperty "HKLM:\SOFTWARE\CloudVolumes\Agent" -ErrorAction SilentlyContinue
if (-not $avReg) { $avReg = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\CloudVolumes\Agent" -ErrorAction SilentlyContinue }
$avVer = if ($avReg -and $avReg.Version) {$avReg.Version} else {"Not found"}
$rows_appvol += New-CheckRow "App Volumes Agent Version" $(if ($avReg) {"INFO"} else {"WARN"}) $avVer

if (-not [string]::IsNullOrWhiteSpace($AppVolManagerFQDN)) {
    $mgrResult = Test-PortConnection -HostName $AppVolManagerFQDN -Port ([int]$AppVolManagerPort) -Label "App Volumes Manager"
    $rows_appvol += New-CheckRow "App Volumes Manager Connectivity" $mgrResult.Status $mgrResult.Detail $mgrResult.Recommendation
} else {
    $rows_appvol += New-CheckRow "App Volumes Manager Connectivity" "SKIP" "No manager host provided."
}

foreach ($p in @("C:\Program Files (x86)\CloudVolumes\Agent\log", "C:\ProgramData\CloudVolumes\Logs")) {
    $pStatus = if (Test-Path $p) {"PASS"} else {"WARN"}
    $rows_appvol += New-CheckRow "App Volumes Log Path" $pStatus $p `
        $(if ($pStatus -eq "WARN") {"Log directory missing — agent may not be installed correctly."} else {""})
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — NVIDIA vGPU
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [4/10] NVIDIA vGPU..." -ForegroundColor Cyan
$rows_nvidia = @()

$nvidiaSmi = "C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"

if (Test-Path $nvidiaSmi) {
    $driverVer    = (& $nvidiaSmi --query-gpu=driver_version --format=csv,noheader 2>$null).Trim()
    $driverStatus = if ($driverVer -eq $ExpectedDriverVer) {"PASS"} elseif ($driverVer) {"WARN"} else {"FAIL"}
    $rows_nvidia += New-CheckRow "NVIDIA Driver Version" $driverStatus `
        "Installed: $driverVer  |  Expected: $ExpectedDriverVer" `
        $(if ($driverStatus -eq "WARN") {"Driver version mismatch. Source correct driver from NVIDIA licensing portal."} `
          elseif ($driverStatus -eq "FAIL") {"Driver not detected. Install vGPU guest driver from NVIDIA licensing portal."} else {""})

    $gpuName  = (& $nvidiaSmi --query-gpu=name --format=csv,noheader 2>$null).Trim()
    $rows_nvidia += New-CheckRow "GPU Name / Profile" "INFO" $gpuName

    $gpuState = (& $nvidiaSmi --query-gpu=pstate,temperature.gpu,utilization.gpu --format=csv,noheader 2>$null).Trim()
    $rows_nvidia += New-CheckRow "GPU P-State / Temp / Utilization" "INFO" $gpuState

    $gpuErrors = (& $nvidiaSmi --query-gpu=ecc.errors.corrected.volatile.total,ecc.errors.uncorrected.volatile.total --format=csv,noheader 2>$null).Trim()
    if ($gpuErrors -and $gpuErrors -notmatch "N/A") {
        $errVals   = $gpuErrors -split ","
        $eccStatus = if ([int]$errVals[1].Trim() -gt 0) {"WARN"} else {"PASS"}
        $rows_nvidia += New-CheckRow "GPU ECC Errors (Uncorrected)" $eccStatus `
            "Corrected: $($errVals[0].Trim())  |  Uncorrected: $($errVals[1].Trim())"
    }
} else {
    $rows_nvidia += New-CheckRow "nvidia-smi.exe" "FAIL" "Not found at: $nvidiaSmi" `
        "Install vGPU guest driver from NVIDIA licensing portal (nvid.nvidia.com). Do NOT use consumer/GeForce driver."
}

$nvCpPath = "C:\Program Files\NVIDIA Corporation\Control Panel Client\nvcplui.exe"
$cpStatus = if (Test-Path $nvCpPath) {"PASS"} else {"FAIL"}
$rows_nvidia += New-CheckRow "NVIDIA Control Panel Client" $cpStatus `
    $(if ($cpStatus -eq "PASS") {$nvCpPath} else {"NOT FOUND: $nvCpPath"}) `
    $(if ($cpStatus -eq "FAIL") {"Control Panel Client missing. Source full vGPU guest driver package from NVIDIA licensing portal. Do not use GeForce/consumer driver."} else {""})

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — APPX PROVISIONED STATE
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [5/10] AppX Provisioned State..." -ForegroundColor Cyan
$rows_appx = @()

$provisionedPkgs  = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
$provisionedNames = $provisionedPkgs | Select-Object -ExpandProperty PackageName
$userPkgs         = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue

$perUserOnly = $userPkgs | Where-Object {
    $n = $_.Name
    -not ($provisionedNames | Where-Object { $_ -like "*$n*" })
} | Select-Object Name, PackageFullName | Sort-Object Name

$rows_appx += New-CheckRow "Total Provisioned Packages" "INFO" "$($provisionedPkgs.Count) provisioned system-wide"

$puStatus = if ($perUserOnly.Count -eq 0) {"PASS"} else {"WARN"}
$rows_appx += New-CheckRow "Per-User Only Packages (Sysprep Risk)" $puStatus `
    "$($perUserOnly.Count) packages installed per-user but NOT provisioned" `
    $(if ($perUserOnly.Count -gt 0) {"Run guyrleech Fix-SysprepAppxErrors.ps1 to remediate before sealing image."} else {""})

$inkPkg    = $userPkgs | Where-Object {$_.Name -like "*Ink.Handwriting*"} | Select-Object -First 1
$inkStatus = if ($inkPkg) {"WARN"} else {"PASS"}
$rows_appx += New-CheckRow "Microsoft.Ink.Handwriting (Known Sysprep Blocker)" $inkStatus `
    $(if ($inkPkg) {"PRESENT: $($inkPkg.PackageFullName)"} else {"Not found — clean"}) `
    $(if ($inkPkg) {"Remove: Get-AppxPackage -AllUsers *Ink.Handwriting* | Remove-AppxPackage -AllUsers"} else {""})

if ($perUserOnly.Count -gt 0) {
    $pkgList = ($perUserOnly | Select-Object -First 10 | ForEach-Object {$_.Name}) -join "<br>"
    $rows_appx += New-CheckRow "Per-User Package List (top 10)" "WARN" $pkgList
}

$store       = Get-AppxPackage -AllUsers -Name "Microsoft.WindowsStore" -ErrorAction SilentlyContinue
$storeStatus = if ($store -and $store.Status -eq "Ok") {"PASS"} elseif ($store) {"WARN"} else {"FAIL"}
$rows_appx += New-CheckRow "Windows Store" $storeStatus `
    $(if ($store) {"Status: $($store.Status)  |  Version: $($store.Version)"} else {"Not installed"}) `
    $(if ($storeStatus -ne "PASS") {"Re-register Store AppX or re-provision via DISM /Add-Capability."} else {""})

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 — DCOM + EVENT LOG
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [6/10] DCOM + Event Log Audit..." -ForegroundColor Cyan
$rows_events = @()

$dcom10016  = Get-WinEvent -LogName System -FilterHashtable @{Id=10016; StartTime=(Get-Date).AddHours(-24)} -ErrorAction SilentlyContinue
$dcomCount  = if ($dcom10016) {$dcom10016.Count} else {0}
$dcomStatus = if ($dcomCount -eq 0) {"PASS"} elseif ($dcomCount -lt 10) {"WARN"} else {"FAIL"}
$rows_events += New-CheckRow "DCOM 10016 Errors (Last 24h)" $dcomStatus "$dcomCount occurrences" `
    $(if ($dcomStatus -ne "PASS") {"Fix DCOM permissions via dcomcnfg, or disable wscsvc if Security Center not required."} else {""})

$wscDcom = $dcom10016 | Where-Object {$_.Message -like "*WscDataProtection*"} | Select-Object -First 1
if ($wscDcom) {
    $rows_events += New-CheckRow "DCOM 10016 — WscDataProtection" "WARN" `
        "WscDataProtection LocalLaunch permission denied to SYSTEM" `
        "Grant NT AUTHORITY\SYSTEM Local Launch + Local Activation via dcomcnfg, OR disable wscsvc."
}

$bootTime  = $os.LastBootUpTime
$appErrors = Get-WinEvent -LogName Application -FilterHashtable @{Level=2; StartTime=$bootTime} -ErrorAction SilentlyContinue | Select-Object -First 15
$appErrCnt = if ($appErrors) {$appErrors.Count} else {0}
$aeStatus  = if ($appErrCnt -eq 0) {"PASS"} elseif ($appErrCnt -lt 5) {"WARN"} else {"FAIL"}
$rows_events += New-CheckRow "Application Errors (Since Last Boot)" $aeStatus `
    "$appErrCnt errors since boot ($($bootTime.ToString('MM/dd HH:mm')))"

if ($appErrors) {
    $errList = ($appErrors | Select-Object -First 5 | ForEach-Object {
        "$($_.TimeCreated.ToString('HH:mm:ss')) | $($_.ProviderName) | $($_.Message.Substring(0,[Math]::Min(120,$_.Message.Length)))"
    }) -join "<br>"
    $rows_events += New-CheckRow "Top Application Errors" "INFO" $errList
}

$sysErrors = Get-WinEvent -LogName System -FilterHashtable @{Level=2; StartTime=$bootTime} -ErrorAction SilentlyContinue | Select-Object -First 10
$sysErrCnt = if ($sysErrors) {$sysErrors.Count} else {0}
$seStatus  = if ($sysErrCnt -eq 0) {"PASS"} elseif ($sysErrCnt -lt 5) {"WARN"} else {"FAIL"}
$rows_events += New-CheckRow "System Errors (Since Last Boot)" $seStatus "$sysErrCnt errors since boot"

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 — SHELL + PROFILE INTEGRITY
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [7/10] Shell + Profile Integrity..." -ForegroundColor Cyan
$rows_shell = @()

$shellFolders = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -ErrorAction SilentlyContinue
$sfStatus     = if ($shellFolders) {"PASS"} else {"FAIL"}
$rows_shell += New-CheckRow "HKCU Shell Folders Key" $sfStatus `
    $(if ($shellFolders) {"Present"} else {"MISSING — may cause Explorer errors at login"}) `
    $(if (-not $shellFolders) {"Profile corruption likely. Rebuild golden image profile."} else {""})

$odRun = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue).OneDrive
$rows_shell += New-CheckRow "OneDrive Autorun (HKLM Run)" "INFO" `
    $(if ($odRun) {"Present: $odRun"} else {"Not in HKLM Run (may be in HKCU — normal)"})

$explorerProc = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
$rows_shell += New-CheckRow "Explorer.exe Process" `
    $(if ($explorerProc) {"PASS"} else {"FAIL"}) `
    $(if ($explorerProc) {"Running (PID: $($explorerProc.Id -join ', '))"} else {"NOT running"}) `
    $(if (-not $explorerProc) {"Shell crash or not yet started."} else {""})

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 8 — TPM + SECURITY
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [8/10] TPM + Security..." -ForegroundColor Cyan
$rows_security = @()

$tpm = Get-CimInstance -Namespace "Root\CIMv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction SilentlyContinue
if ($tpm) {
    $tpmStatus = if ($tpm.IsEnabled_InitialValue -and $tpm.IsActivated_InitialValue) {"PASS"} else {"WARN"}
    $rows_security += New-CheckRow "TPM / vTPM State" $tpmStatus `
        "Enabled: $($tpm.IsEnabled_InitialValue)  |  Activated: $($tpm.IsActivated_InitialValue)  |  Spec: $($tpm.SpecVersion)"
} else {
    $rows_security += New-CheckRow "TPM / vTPM State" "WARN" "Not detected or WMI query failed" `
        "Horizon full clone pools can provision vTPM per-clone at creation — template itself does not require vTPM."
}

$secureBoot = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue
$rows_security += New-CheckRow "Secure Boot" `
    $(if ($secureBoot) {"PASS"} else {"INFO"}) `
    $(if ($secureBoot -eq $true) {"Enabled"} elseif ($secureBoot -eq $false) {"Disabled"} else {"Not supported / BIOS mode"})

$av       = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
$avStatus = if ($av) {"INFO"} else {"WARN"}
$rows_security += New-CheckRow "Antivirus Product" $avStatus `
    $(if ($av) {($av | ForEach-Object {$_.displayName}) -join ", "} else {"No AV detected via SecurityCenter2"})

$wdResult  = Test-ServiceStatus -ServiceName "WinDefend"
$rows_security += New-CheckRow "Windows Defender Service" $wdResult.Status $wdResult.Detail

$wscResult = Test-ServiceStatus -ServiceName "wscsvc"
$rows_security += New-CheckRow "Windows Security Center (wscsvc)" $wscResult.Status $wscResult.Detail `
    $(if ($wscResult.Status -eq "PASS") {"wscsvc generates DCOM 10016 errors — disable only if confirmed safe in this environment."} else {""})

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 9 — NETWORK
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [9/10] Network..." -ForegroundColor Cyan
$rows_network = @()

$nics = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.InterfaceAlias -notlike "*Loopback*"}
foreach ($nic in $nics) {
    $rows_network += New-CheckRow "NIC: $($nic.InterfaceAlias)" "INFO" "$($nic.IPAddress) / $($nic.PrefixLength)"
}

$dns = Get-DnsClientServerAddress -AddressFamily IPv4 |
    Where-Object {$_.InterfaceAlias -notlike "*Loopback*" -and $_.ServerAddresses.Count -gt 0}
foreach ($d in $dns) {
    $rows_network += New-CheckRow "DNS: $($d.InterfaceAlias)" "INFO" ($d.ServerAddresses -join ", ")
}

if ($cs.PartOfDomain) {
    $dcResult = Test-PortConnection -HostName $cs.Domain -Port 389 -Label "Domain Controller LDAP"
    $rows_network += New-CheckRow "Domain Controller LDAP (389)" $dcResult.Status $dcResult.Detail $dcResult.Recommendation
}

if (-not [string]::IsNullOrWhiteSpace($HorizonCS)) {
    $hcsResult = Test-PortConnection -HostName $HorizonCS -Port 443 -Label "Horizon Connection Server"
    $rows_network += New-CheckRow "Horizon Connection Server (443)" $hcsResult.Status $hcsResult.Detail $hcsResult.Recommendation
} else {
    $rows_network += New-CheckRow "Horizon Connection Server" "SKIP" "No host provided."
}

foreach ($eh in @($ExtraHost1, $ExtraHost2)) {
    if (-not [string]::IsNullOrWhiteSpace($eh)) {
        $r = Test-NetConnection -ComputerName $eh -InformationLevel Quiet -WarningAction SilentlyContinue
        $rows_network += New-CheckRow "Additional Host: $eh" $(if ($r) {"PASS"} else {"FAIL"}) `
            $(if ($r) {"Reachable"} else {"UNREACHABLE"})
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY SCORE
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  [10/10] Building Report..." -ForegroundColor Cyan

$allRows       = @($rows_system + $rows_horizon + $rows_appvol + $rows_nvidia + $rows_appx + $rows_events + $rows_shell + $rows_security + $rows_network)
$totalFail     = ([regex]::Matches($allRows -join "", "badge fail")).Count
$totalWarn     = ([regex]::Matches($allRows -join "", "badge warn")).Count
$totalPass     = ([regex]::Matches($allRows -join "", "badge pass")).Count
$overallStatus = if ($totalFail -gt 0) {"CRITICAL"} elseif ($totalWarn -gt 3) {"DEGRADED"} elseif ($totalWarn -gt 0) {"WARNING"} else {"HEALTHY"}
$overallColor  = switch ($overallStatus) {
    "CRITICAL" {"#ff3b3b"} "DEGRADED" {"#ff8c00"} "WARNING" {"#f5c518"} default {"#00e676"}
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML
# ─────────────────────────────────────────────────────────────────────────────
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>VDI Health Report — $env:COMPUTERNAME</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Share+Tech+Mono&family=Exo+2:wght@300;400;600;800&display=swap');
:root{--bg:#0a0e17;--surface:#111827;--surface2:#1a2233;--border:#1e3a5f;--accent:#00b4ff;--text:#c9d8e8;--text-dim:#5a7a9a;--pass:#00e676;--fail:#ff3b3b;--warn:#f5c518;--info:#00b4ff;--skip:#5a7a9a;--mono:'Share Tech Mono',monospace;--sans:'Exo 2',sans-serif}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:13px;line-height:1.6}
.header{background:linear-gradient(135deg,#0a1628,#0d1f3c 50%,#091420);border-bottom:1px solid var(--border);padding:32px 40px 24px;position:relative;overflow:hidden}
.header::before{content:'';position:absolute;inset:0;background:repeating-linear-gradient(0deg,transparent,transparent 39px,rgba(0,180,255,.03) 39px,rgba(0,180,255,.03) 40px),repeating-linear-gradient(90deg,transparent,transparent 39px,rgba(0,180,255,.03) 39px,rgba(0,180,255,.03) 40px);pointer-events:none}
.header-top{display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:16px}
.header h1{font-weight:800;font-size:22px;letter-spacing:.08em;text-transform:uppercase;color:var(--accent);text-shadow:0 0 20px rgba(0,180,255,.4)}
.header h2{font-family:var(--mono);font-size:12px;color:var(--text-dim);margin-top:4px;letter-spacing:.05em}
.overall-badge{font-weight:800;font-size:15px;letter-spacing:.12em;padding:10px 24px;border-radius:4px;border:2px solid;text-transform:uppercase;color:$overallColor;border-color:$overallColor;box-shadow:0 0 20px ${overallColor}40;white-space:nowrap}
.meta-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-top:20px}
.meta-item{background:rgba(0,180,255,.05);border:1px solid var(--border);border-radius:4px;padding:8px 12px}
.meta-label{font-size:10px;text-transform:uppercase;letter-spacing:.1em;color:var(--text-dim)}
.meta-value{font-family:var(--mono);font-size:12px;color:var(--accent);margin-top:2px;word-break:break-all}
.score-bar{background:var(--surface);border-bottom:1px solid var(--border);padding:16px 40px;display:flex;align-items:center;gap:32px;flex-wrap:wrap}
.score-num{font-family:var(--mono);font-size:26px;font-weight:bold}
.score-lbl{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--text-dim)}
.score-num.pass{color:var(--pass)}.score-num.warn{color:var(--warn)}.score-num.fail{color:var(--fail)}
.score-divider{width:1px;height:40px;background:var(--border)}
.content{padding:24px 40px}
.section{background:var(--surface);border:1px solid var(--border);border-radius:6px;margin-bottom:20px;overflow:hidden}
.section-header{background:linear-gradient(90deg,var(--surface2),var(--surface));border-bottom:1px solid var(--border);padding:12px 16px;display:flex;align-items:center;gap:10px}
.section-icon{font-size:16px}
.section-title{font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.1em;color:var(--accent)}
.section-count{margin-left:auto;font-family:var(--mono);font-size:11px;color:var(--text-dim)}
table{width:100%;border-collapse:collapse}
th{background:var(--surface2);color:var(--text-dim);font-size:10px;text-transform:uppercase;letter-spacing:.1em;padding:8px 14px;text-align:left;border-bottom:1px solid var(--border)}
td{padding:9px 14px;border-bottom:1px solid rgba(30,58,95,.5);vertical-align:top}
tr:last-child td{border-bottom:none}
.check-name{font-weight:600;font-size:12px;color:var(--text);white-space:nowrap;width:290px}
td:nth-child(3){font-family:var(--mono);font-size:11px;color:var(--text-dim);word-break:break-word}
.rec{margin-top:5px;font-family:var(--sans);font-size:11px;color:var(--warn);opacity:.85}
tr.row-fail{background:rgba(255,59,59,.04)}tr.row-warn{background:rgba(245,197,24,.04)}tr.row-pass{background:rgba(0,230,118,.02)}
tr:hover{background:rgba(0,180,255,.05)!important}
.badge{display:inline-block;font-family:var(--mono);font-size:10px;font-weight:bold;padding:3px 9px;border-radius:3px;letter-spacing:.08em;white-space:nowrap}
.badge.pass{background:rgba(0,230,118,.12);color:var(--pass);border:1px solid rgba(0,230,118,.3)}
.badge.fail{background:rgba(255,59,59,.12);color:var(--fail);border:1px solid rgba(255,59,59,.3)}
.badge.warn{background:rgba(245,197,24,.12);color:var(--warn);border:1px solid rgba(245,197,24,.3)}
.badge.info{background:rgba(0,180,255,.12);color:var(--info);border:1px solid rgba(0,180,255,.3)}
.badge.skip{background:rgba(90,122,154,.12);color:var(--skip);border:1px solid rgba(90,122,154,.3)}
.footer{border-top:1px solid var(--border);padding:16px 40px;display:flex;justify-content:space-between;align-items:center;font-family:var(--mono);font-size:11px;color:var(--text-dim);background:var(--surface)}
@media print{body{background:#fff;color:#000}.section,td,th{border-color:#ccc!important}}
</style>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div>
      <div style="font-family:var(--mono,monospace);font-size:11px;color:#5a7a9a;margin-bottom:4px;letter-spacing:.15em;">
        OMNISSA HORIZON VDI // $($PoolType.ToUpper()) // WINDOWS 11
      </div>
      <h1>VDI Health Report</h1>
      <h2>$env:COMPUTERNAME &nbsp;|&nbsp; $EnvName &nbsp;|&nbsp; $RunTime</h2>
    </div>
    <div class="overall-badge">$overallStatus</div>
  </div>
  <div class="meta-grid">
    <div class="meta-item"><div class="meta-label">Computer</div><div class="meta-value">$env:COMPUTERNAME</div></div>
    <div class="meta-item"><div class="meta-label">Environment</div><div class="meta-value">$EnvName</div></div>
    <div class="meta-item"><div class="meta-label">Pool Type</div><div class="meta-value">$PoolType</div></div>
    <div class="meta-item"><div class="meta-label">Run As</div><div class="meta-value">$env:USERNAME</div></div>
    <div class="meta-item"><div class="meta-label">Expected GPU Driver</div><div class="meta-value">$ExpectedDriverVer</div></div>
    <div class="meta-item"><div class="meta-label">App Vol Manager</div><div class="meta-value">$(if ($AppVolManagerFQDN) {$AppVolManagerFQDN} else {'Not specified'})</div></div>
  </div>
</div>

<div class="score-bar">
  <div style="display:flex;align-items:center;gap:8px"><div class="score-num pass">$totalPass</div><div class="score-lbl">Passed</div></div>
  <div class="score-divider"></div>
  <div style="display:flex;align-items:center;gap:8px"><div class="score-num warn">$totalWarn</div><div class="score-lbl">Warnings</div></div>
  <div class="score-divider"></div>
  <div style="display:flex;align-items:center;gap:8px"><div class="score-num fail">$totalFail</div><div class="score-lbl">Failed</div></div>
  <div class="score-divider"></div>
  <div style="font-family:monospace;font-size:11px;color:#5a7a9a">Overall: <span style="color:$overallColor;font-weight:bold">$overallStatus</span></div>
</div>

<div class="content">
  <div class="section">
    <div class="section-header"><span class="section-icon">🖥</span><span class="section-title">System Baseline</span><span class="section-count">OS · Domain · Activation · Disk · Sysprep</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_system -join "")</table>
  </div>
  <div class="section">
    <div class="section-header"><span class="section-icon">🔷</span><span class="section-title">VMware Tools + Horizon Agent</span><span class="section-count">Services · Version · Guest Customization Log</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_horizon -join "")</table>
  </div>
  <div class="section">
    <div class="section-header"><span class="section-icon">📦</span><span class="section-title">App Volumes Agent</span><span class="section-count">svservice · Version · Manager Connectivity · Log Paths</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_appvol -join "")</table>
  </div>
  <div class="section">
    <div class="section-header"><span class="section-icon">🎮</span><span class="section-title">NVIDIA vGPU</span><span class="section-count">Driver Version · GPU Profile · Control Panel Client · ECC</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_nvidia -join "")</table>
  </div>
  <div class="section">
    <div class="section-header"><span class="section-icon">📱</span><span class="section-title">AppX Provisioned State</span><span class="section-count">Sysprep Blockers · Per-User Drift · Windows Store</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_appx -join "")</table>
  </div>
  <div class="section">
    <div class="section-header"><span class="section-icon">📋</span><span class="section-title">DCOM + Event Log</span><span class="section-count">10016 Errors · WscDataProtection · App + System Errors</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_events -join "")</table>
  </div>
  <div class="section">
    <div class="section-header"><span class="section-icon">🪟</span><span class="section-title">Shell + Profile Integrity</span><span class="section-count">Shell Folders · OneDrive · Explorer</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_shell -join "")</table>
  </div>
  <div class="section">
    <div class="section-header"><span class="section-icon">🔒</span><span class="section-title">TPM + Security</span><span class="section-count">vTPM · Secure Boot · Defender · AV · wscsvc</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_security -join "")</table>
  </div>
  <div class="section">
    <div class="section-header"><span class="section-icon">🌐</span><span class="section-title">Network</span><span class="section-count">NIC · DNS · Domain Controller · Horizon CS · App Vol Manager</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_network -join "")</table>
  </div>
</div>

<div class="footer">
  <div>$ReportFile</div>
  <div>VDI Health Report v1.1 — Generic Edition</div>
</div>

</body>
</html>
"@

$html | Out-File -FilePath $ReportFile -Encoding UTF8 -Force

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
$statusColor = if ($overallStatus -eq "HEALTHY") {"Green"} elseif ($overallStatus -eq "WARNING") {"Yellow"} else {"Red"}
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor $statusColor
Write-Host "  ║  COMPLETE  —  $overallStatus" -ForegroundColor $statusColor
Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor DarkGray
Write-Host "  ║  PASS : $totalPass" -ForegroundColor Green
Write-Host "  ║  WARN : $totalWarn" -ForegroundColor Yellow
Write-Host "  ║  FAIL : $totalFail" -ForegroundColor Red
Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor DarkGray
Write-Host "  ║  $ReportFile" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor DarkGray
Write-Host ""

Start-Process $ReportFile
