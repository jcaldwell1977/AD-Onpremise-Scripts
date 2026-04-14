#Requires -RunAsAdministrator
<#
.SYNOPSIS
    VDI Golden Image Health Check Report
    Nova Pool | Omnissa Horizon 2512 | Windows 11 24H2 | NVIDIA vGPU 582.x

.DESCRIPTION
    Generates a timestamped HTML health report covering:
      - System baseline (OS, domain, activation)
      - AppX provisioned vs per-user drift
      - DCOM 10016 audit
      - NVIDIA vGPU driver + Control Panel Client
      - Horizon Agent services
      - App Volumes Agent + Manager connectivity
      - VMware Tools
      - Windows Store state
      - Shell folder integrity
      - Guest Customization log tail
      - Event log critical errors (last boot)
      - Sysprep state

.PARAMETER AppVolManagerFQDN
    FQDN or IP of your App Volumes Manager. Default: appvolmgr.corp.aechelon.com

.PARAMETER OutputPath
    Folder to save HTML report. Default: C:\Temp\VDIHealthReports

.PARAMETER ExpectedDriverVersion
    Expected NVIDIA vGPU guest driver version. Default: 582.16

.PARAMETER HorizonConnectionServer
    Optional: Horizon Connection Server to test connectivity. Default: skip

.EXAMPLE
    .\Invoke-VDIHealthReport.ps1
    .\Invoke-VDIHealthReport.ps1 -AppVolManagerFQDN "appvol.corp.aechelon.com" -OutputPath "D:\Reports"

.NOTES
    Author  : VDI Engineering
    Version : 1.0
    Target  : Windows 11 24H2 | Horizon 2512 Full Clone | Nova Pool
#>

[CmdletBinding()]
param(
    [string]$AppVolManagerFQDN     = "appvolmgr.corp.aechelon.com",
    [string]$OutputPath            = "C:\Temp\VDIHealthReports",
    [string]$ExpectedDriverVersion = "582.16",
    [string]$HorizonConnectionServer = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "SilentlyContinue"

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
function Get-StatusBadge {
    param([string]$Status)
    switch ($Status) {
        "PASS"    { return '<span class="badge pass">✔ PASS</span>' }
        "FAIL"    { return '<span class="badge fail">✖ FAIL</span>' }
        "WARN"    { return '<span class="badge warn">⚠ WARN</span>' }
        "INFO"    { return '<span class="badge info">ℹ INFO</span>' }
        "SKIP"    { return '<span class="badge skip">— SKIP</span>' }
        default   { return '<span class="badge info">? UNKN</span>' }
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
    $rec   = if ($Recommendation) { "<div class='rec'>↳ $Recommendation</div>" } else { "" }
    return "<tr class='row-$($Status.ToLower())'><td class='check-name'>$Check</td><td>$badge</td><td>$Detail$rec</td></tr>"
}

function Test-ServiceStatus {
    param([string]$ServiceName, [string]$ExpectedState = "Running")
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) { return @{ Status="FAIL"; Detail="Service not found: $ServiceName" } }
    if ($svc.Status -eq $ExpectedState) {
        return @{ Status="PASS"; Detail="$ServiceName — $($svc.Status)" }
    } else {
        return @{ Status="FAIL"; Detail="$ServiceName — Expected: $ExpectedState | Actual: $($svc.Status)" }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT SETUP
# ─────────────────────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$Timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile  = Join-Path $OutputPath "VDI_HealthReport_$($env:COMPUTERNAME)_$Timestamp.html"
$RunTime     = Get-Date -Format "dddd, MMMM dd yyyy  HH:mm:ss"

Write-Host "`n[VDI Health Check] Starting... $RunTime" -ForegroundColor Cyan
Write-Host "[VDI Health Check] Computer: $env:COMPUTERNAME" -ForegroundColor Cyan

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — SYSTEM BASELINE
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`n[1/10] System Baseline..." -ForegroundColor Yellow
$rows_system = @()

# OS Info
$os = Get-CimInstance Win32_OperatingSystem
$rows_system += New-CheckRow -Check "OS Version" -Status "INFO" `
    -Detail "$($os.Caption) — Build $($os.BuildNumber) | $($os.OSArchitecture)"

# 24H2 Build Check (Build 26100+)
$buildNum = [int]$os.BuildNumber
$buildStatus = if ($buildNum -ge 26100) { "PASS" } else { "WARN" }
$rows_system += New-CheckRow -Check "24H2 Build Verified" -Status $buildStatus `
    -Detail "Build $buildNum (24H2 = 26100+)" `
    -Recommendation $(if ($buildStatus -eq "WARN") { "Image may not be 24H2. Verify ISO source from M365 Admin Center." } else { "" })

# Hostname
$rows_system += New-CheckRow -Check "Hostname" -Status "INFO" -Detail $env:COMPUTERNAME

# Domain Join
$cs = Get-CimInstance Win32_ComputerSystem
$domainStatus = if ($cs.PartOfDomain) { "PASS" } else { "WARN" }
$rows_system += New-CheckRow -Check "Domain Join" -Status $domainStatus `
    -Detail $(if ($cs.PartOfDomain) { "Joined: $($cs.Domain)" } else { "NOT domain joined" }) `
    -Recommendation $(if (-not $cs.PartOfDomain) { "Verify Guest Customization completed successfully." } else { "" })

# Windows Activation
$licStatus = (Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%' AND LicenseStatus=1" -ErrorAction SilentlyContinue | Select-Object -First 1)
$actStatus = if ($licStatus) { "PASS" } else { "WARN" }
$rows_system += New-CheckRow -Check "Windows Activation" -Status $actStatus `
    -Detail $(if ($licStatus) { "Activated — $($licStatus.Name)" } else { "Not activated or unable to verify" }) `
    -Recommendation $(if (-not $licStatus) { "Check KMS connectivity or MAK activation." } else { "" })

# Uptime
$uptime = (Get-Date) - $os.LastBootUpTime
$rows_system += New-CheckRow -Check "Last Boot / Uptime" -Status "INFO" `
    -Detail "Booted: $($os.LastBootUpTime.ToString('MM/dd/yyyy HH:mm:ss'))  |  Uptime: $([math]::Floor($uptime.TotalHours))h $($uptime.Minutes)m"

# RAM
$ramGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
$ramStatus = if ($ramGB -ge 4) { "PASS" } else { "WARN" }
$rows_system += New-CheckRow -Check "Physical RAM" -Status $ramStatus `
    -Detail "$ramGB GB" `
    -Recommendation $(if ($ramGB -lt 4) { "Minimum 4GB recommended for VDI desktop." } else { "" })

# Disk (C:)
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$diskFreeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
$diskTotalGB = [math]::Round($disk.Size / 1GB, 1)
$diskPct = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 0)
$diskStatus = if ($diskPct -ge 20) { "PASS" } elseif ($diskPct -ge 10) { "WARN" } else { "FAIL" }
$rows_system += New-CheckRow -Check "Disk C: Free Space" -Status $diskStatus `
    -Detail "$diskFreeGB GB free of $diskTotalGB GB ($diskPct% free)" `
    -Recommendation $(if ($diskStatus -eq "FAIL") { "Critical: less than 10% free. Golden image may be too large." } elseif ($diskStatus -eq "WARN") { "Consider cleanup or expanding template disk." } else { "" })

# Sysprep State
$setupKey = Get-ItemProperty "HKLM:\SYSTEM\Setup" -ErrorAction SilentlyContinue
$oobe     = $setupKey.OOBEInProgress
$setupType = $setupKey.SetupType
$sysprepStatus = if ($oobe -eq 0 -and $setupType -eq 0) { "PASS" } else { "WARN" }
$rows_system += New-CheckRow -Check "Sysprep / OOBE State" -Status $sysprepStatus `
    -Detail "OOBEInProgress=$oobe | SetupType=$setupType  (both must be 0 on deployed clone)" `
    -Recommendation $(if ($sysprepStatus -eq "WARN") { "VM may still be in generalize/OOBE state. Verify Guest Customization completed." } else { "" })

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — VMWARE / HORIZON AGENT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[2/10] VMware Tools + Horizon Agent..." -ForegroundColor Yellow
$rows_horizon = @()

# VMware Tools
$toolsSvc = Test-ServiceStatus -ServiceName "VMTools"
$rows_horizon += New-CheckRow -Check "VMware Tools Service" -Status $toolsSvc.Status -Detail $toolsSvc.Detail `
    -Recommendation $(if ($toolsSvc.Status -eq "FAIL") { "VMware Tools not running — required for Guest Customization and Horizon." } else { "" })

# VMware Tools Version
$toolsReg = Get-ItemProperty "HKLM:\SOFTWARE\VMware, Inc.\VMware Tools" -ErrorAction SilentlyContinue
$toolsVer  = if ($toolsReg) { $toolsReg.Version } else { "Not found" }
$rows_horizon += New-CheckRow -Check "VMware Tools Version" -Status $(if ($toolsReg) {"INFO"} else {"FAIL"}) -Detail $toolsVer

# Horizon Agent Services
$horizonServices = @(
    @{ Name="vmware-viewagent";     Label="Horizon View Agent" },
    @{ Name="wsnm";                 Label="Horizon WSNM (Blast/PCoIP)" },
    @{ Name="CSVD";                 Label="Horizon CSVD (USB/Scanner)" },
    @{ Name="vmware-view-usbd";     Label="Horizon USB Arbitrator" }
)
foreach ($svc in $horizonServices) {
    $result = Test-ServiceStatus -ServiceName $svc.Name
    # CSVD and USB are optional, downgrade to WARN if missing
    $finalStatus = if ($result.Status -eq "FAIL" -and $svc.Name -in @("CSVD","vmware-view-usbd")) { "WARN" } else { $result.Status }
    $rows_horizon += New-CheckRow -Check $svc.Label -Status $finalStatus -Detail $result.Detail
}

# Horizon Agent Version from Registry
$horizonReg = Get-ItemProperty "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\Agent" -ErrorAction SilentlyContinue
$horizonVer  = if ($horizonReg -and $horizonReg.ProductVersion) { $horizonReg.ProductVersion } else { "Not found in registry" }
$rows_horizon += New-CheckRow -Check "Horizon Agent Version" -Status $(if ($horizonReg) {"INFO"} else {"WARN"}) -Detail $horizonVer `
    -Recommendation $(if (-not $horizonReg) { "Verify Horizon Agent 2512 installed correctly." } else { "" })

# Guest Customization Log
$gcLog = "C:\Windows\Temp\vmware-imc\guestcust.log"
if (Test-Path $gcLog) {
    $gcTail = (Get-Content $gcLog -Tail 5) -join " | "
    $gcStatus = if ($gcTail -match "success|Customization finished") { "PASS" } elseif ($gcTail -match "error|fail") { "FAIL" } else { "INFO" }
    $rows_horizon += New-CheckRow -Check "Guest Customization Log (tail)" -Status $gcStatus -Detail $gcTail `
        -Recommendation $(if ($gcStatus -eq "FAIL") { "Review full log: C:\Windows\Temp\vmware-imc\guestcust.log" } else { "" })
} else {
    $rows_horizon += New-CheckRow -Check "Guest Customization Log" -Status "WARN" `
        -Detail "Log not found: $gcLog" `
        -Recommendation "Log absent may indicate customization never ran or VM was not deployed via Horizon pool."
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — APP VOLUMES AGENT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[3/10] App Volumes Agent..." -ForegroundColor Yellow
$rows_appvol = @()

# svservice
$svResult = Test-ServiceStatus -ServiceName "svservice"
$rows_appvol += New-CheckRow -Check "App Volumes Agent (svservice)" -Status $svResult.Status -Detail $svResult.Detail `
    -Recommendation $(if ($svResult.Status -eq "FAIL") { "svservice not running. Reinstall App Volumes Agent or check installer log." } else { "" })

# Agent Version from Registry
$avReg = Get-ItemProperty "HKLM:\SOFTWARE\CloudVolumes\Agent" -ErrorAction SilentlyContinue
if (-not $avReg) { $avReg = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\CloudVolumes\Agent" -ErrorAction SilentlyContinue }
$avVer = if ($avReg -and $avReg.Version) { $avReg.Version } else { "Not found" }
$rows_appvol += New-CheckRow -Check "App Volumes Agent Version" -Status $(if ($avReg) {"INFO"} else {"WARN"}) -Detail $avVer

# Manager Connectivity
if ($AppVolManagerFQDN -and $AppVolManagerFQDN -ne "") {
    $mgr443 = Test-NetConnection -ComputerName $AppVolManagerFQDN -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue
    $mgrStatus = if ($mgr443) { "PASS" } else { "FAIL" }
    $rows_appvol += New-CheckRow -Check "App Volumes Manager 443 Connectivity" -Status $mgrStatus `
        -Detail "$AppVolManagerFQDN : 443 — $(if ($mgr443) {'Reachable'} else {'UNREACHABLE'})" `
        -Recommendation $(if (-not $mgr443) { "Check firewall rules and DNS resolution for $AppVolManagerFQDN" } else { "" })
} else {
    $rows_appvol += New-CheckRow -Check "App Volumes Manager Connectivity" -Status "SKIP" `
        -Detail "No FQDN provided. Use -AppVolManagerFQDN parameter."
}

# Agent Log Path
$avAgentLog = "C:\Program Files (x86)\CloudVolumes\Agent\log"
$avLogStatus = if (Test-Path $avAgentLog) { "PASS" } else { "WARN" }
$rows_appvol += New-CheckRow -Check "App Volumes Agent Log Path" -Status $avLogStatus `
    -Detail $avAgentLog `
    -Recommendation $(if ($avLogStatus -eq "WARN") { "Agent log directory missing — agent may not be installed correctly." } else { "" })

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — NVIDIA vGPU
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[4/10] NVIDIA vGPU..." -ForegroundColor Yellow
$rows_nvidia = @()

$nvidiaSmi = "C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"

if (Test-Path $nvidiaSmi) {
    # Driver Version
    $driverVer = & $nvidiaSmi --query-gpu=driver_version --format=csv,noheader 2>$null
    $driverVer = $driverVer.Trim()
    $driverStatus = if ($driverVer -eq $ExpectedDriverVersion) { "PASS" } elseif ($driverVer) { "WARN" } else { "FAIL" }
    $rows_nvidia += New-CheckRow -Check "NVIDIA Driver Version" -Status $driverStatus `
        -Detail "Installed: $driverVer  |  Expected: $ExpectedDriverVersion" `
        -Recommendation $(if ($driverStatus -eq "WARN") { "Driver version mismatch. Source correct guest driver from nvid.nvidia.com." } elseif ($driverStatus -eq "FAIL") { "Driver not detected. Reinstall from NVIDIA vGPU licensing portal." } else { "" })

    # GPU Name / vGPU Profile
    $gpuName = & $nvidiaSmi --query-gpu=name --format=csv,noheader 2>$null
    $rows_nvidia += New-CheckRow -Check "GPU Name / Profile" -Status "INFO" -Detail $gpuName.Trim()

    # GPU Status
    $gpuStatus = & $nvidiaSmi --query-gpu=pstate,temperature.gpu,utilization.gpu --format=csv,noheader 2>$null
    $rows_nvidia += New-CheckRow -Check "GPU P-State / Temp / Utilization" -Status "INFO" -Detail $gpuStatus.Trim()

    # ECC / Error Check
    $gpuErrors = & $nvidiaSmi --query-gpu=ecc.errors.corrected.volatile.total,ecc.errors.uncorrected.volatile.total --format=csv,noheader 2>$null
    if ($gpuErrors -and $gpuErrors -notmatch "N/A") {
        $errVals = $gpuErrors.Trim() -split ","
        $eccStatus = if ([int]$errVals[1].Trim() -gt 0) { "WARN" } else { "PASS" }
        $rows_nvidia += New-CheckRow -Check "GPU ECC Errors (Uncorrected)" -Status $eccStatus `
            -Detail "Corrected: $($errVals[0].Trim())  |  Uncorrected: $($errVals[1].Trim())"
    }
} else {
    $rows_nvidia += New-CheckRow -Check "nvidia-smi.exe" -Status "FAIL" `
        -Detail "Not found at: $nvidiaSmi" `
        -Recommendation "Install vGPU guest driver 582.16 from nvid.nvidia.com. Do NOT use GeForce/consumer driver."
}

# Control Panel Client
$nvCpPath = "C:\Program Files\NVIDIA Corporation\Control Panel Client\nvcplui.exe"
$cpStatus = if (Test-Path $nvCpPath) { "PASS" } else { "FAIL" }
$rows_nvidia += New-CheckRow -Check "NVIDIA Control Panel Client" -Status $cpStatus `
    -Detail $(if ($cpStatus -eq "PASS") { $nvCpPath } else { "NOT FOUND: $nvCpPath" }) `
    -Recommendation $(if ($cpStatus -eq "FAIL") { "Control Panel Client folder missing. Source full vGPU guest driver package from nvid.nvidia.com — consumer driver packages exclude this." } else { "" })

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — APPX PROVISIONED STATE
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[5/10] AppX Provisioned State Audit..." -ForegroundColor Yellow
$rows_appx = @()

# Get provisioned packages
$provisionedPkgs = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
$provisionedNames = $provisionedPkgs | Select-Object -ExpandProperty PackageName

# Get all user-installed packages
$userPkgs = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue

# Find per-user only (installed but NOT provisioned)
$perUserOnly = $userPkgs | Where-Object {
    $pkgName = $_.Name
    -not ($provisionedNames | Where-Object { $_ -like "*$pkgName*" })
} | Select-Object Name, PackageFullName | Sort-Object Name

# Known high-risk sysprep blockers
$highRiskPackages = @(
    "Microsoft.Ink.Handwriting",
    "MicrosoftWindows.Client.WebExperience",
    "Microsoft.Windows.Cortana",
    "Microsoft.BingWeather",
    "Microsoft.GamingApp",
    "Microsoft.XboxGameOverlay",
    "Microsoft.XboxGamingOverlay"
)

$rows_appx += New-CheckRow -Check "Total Provisioned Packages" -Status "INFO" `
    -Detail "$($provisionedPkgs.Count) provisioned system-wide"

$rows_appx += New-CheckRow -Check "Per-User Only Packages (Sysprep Risk)" `
    -Status $(if ($perUserOnly.Count -eq 0) { "PASS" } else { "WARN" }) `
    -Detail "$($perUserOnly.Count) packages installed per-user but NOT provisioned" `
    -Recommendation $(if ($perUserOnly.Count -gt 0) { "Run guyrleech Fix-SysprepAppxErrors.ps1 to remediate before sealing image." } else { "" })

# Check specifically for the known blocker
$inkHandwriting = $userPkgs | Where-Object { $_.Name -like "*Ink.Handwriting*" } | Select-Object -First 1
$inkStatus = if ($inkHandwriting) { "WARN" } else { "PASS" }
$rows_appx += New-CheckRow -Check "Microsoft.Ink.Handwriting (Known Blocker)" -Status $inkStatus `
    -Detail $(if ($inkHandwriting) { "PRESENT: $($inkHandwriting.PackageFullName)" } else { "Not found — clean" }) `
    -Recommendation $(if ($inkHandwriting) { "Remove with: Get-AppxPackage -AllUsers *Ink.Handwriting* | Remove-AppxPackage -AllUsers" } else { "" })

# List top per-user-only packages (up to 10)
if ($perUserOnly.Count -gt 0) {
    $pkgList = ($perUserOnly | Select-Object -First 10 | ForEach-Object { $_.Name }) -join "<br>"
    $rows_appx += New-CheckRow -Check "Per-User Package List (top 10)" -Status "WARN" -Detail $pkgList
}

# Windows Store
$store = Get-AppxPackage -AllUsers -Name "Microsoft.WindowsStore" -ErrorAction SilentlyContinue
$storeStatus = if ($store -and $store.Status -eq "Ok") { "PASS" } elseif ($store) { "WARN" } else { "FAIL" }
$rows_appx += New-CheckRow -Check "Windows Store (Microsoft.WindowsStore)" -Status $storeStatus `
    -Detail $(if ($store) { "Status: $($store.Status)  |  Version: $($store.Version)" } else { "Not installed / not provisioned" }) `
    -Recommendation $(if ($storeStatus -ne "PASS") { "Re-register: Add-AppxPackage -DisableDevelopmentMode -Register `"`$(`$_.InstallLocation)\AppxManifest.xml`"" } else { "" })

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 — DCOM / EVENT LOG AUDIT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[6/10] DCOM + Event Log Audit..." -ForegroundColor Yellow
$rows_events = @()

# DCOM 10016 count
$dcom10016 = Get-WinEvent -LogName System -FilterHashtable @{Id=10016; StartTime=(Get-Date).AddHours(-24)} -ErrorAction SilentlyContinue
$dcomCount  = if ($dcom10016) { $dcom10016.Count } else { 0 }
$dcomStatus = if ($dcomCount -eq 0) { "PASS" } elseif ($dcomCount -lt 10) { "WARN" } else { "FAIL" }
$rows_events += New-CheckRow -Check "DCOM 10016 Errors (Last 24h)" -Status $dcomStatus `
    -Detail "$dcomCount occurrences" `
    -Recommendation $(if ($dcomStatus -ne "PASS") { "Fix WscDataProtection DCOM permissions via dcomcnfg or disable wscsvc if Security Center not required." } else { "" })

# WscDataProtection specifically
$wscDcom = $dcom10016 | Where-Object { $_.Message -like "*WscDataProtection*" } | Select-Object -First 1
if ($wscDcom) {
    $rows_events += New-CheckRow -Check "DCOM 10016 — WscDataProtection" -Status "WARN" `
        -Detail "Confirmed: WscDataProtection LocalLaunch permission denied to SYSTEM" `
        -Recommendation "Grant NT AUTHORITY\SYSTEM Local Launch + Local Activation via dcomcnfg, OR disable wscsvc."
}

# Critical Application Errors (last boot)
$bootTime = $os.LastBootUpTime
$appErrors = Get-WinEvent -LogName Application -FilterHashtable @{Level=2; StartTime=$bootTime} -ErrorAction SilentlyContinue | Select-Object -First 15
$appErrCount = if ($appErrors) { $appErrors.Count } else { 0 }
$appErrStatus = if ($appErrCount -eq 0) { "PASS" } elseif ($appErrCount -lt 5) { "WARN" } else { "FAIL" }
$rows_events += New-CheckRow -Check "Application Critical Errors (Since Boot)" -Status $appErrStatus `
    -Detail "$appErrCount errors since last boot ($($bootTime.ToString('MM/dd HH:mm')))"

# List top app errors
if ($appErrors) {
    $errList = ($appErrors | Select-Object -First 5 | ForEach-Object {
        "$($_.TimeCreated.ToString('HH:mm:ss')) | $($_.ProviderName) | $($_.Message.Substring(0, [Math]::Min(120, $_.Message.Length)))"
    }) -join "<br>"
    $rows_events += New-CheckRow -Check "Top Application Errors" -Status "INFO" -Detail $errList
}

# System Critical Errors (last boot)
$sysErrors = Get-WinEvent -LogName System -FilterHashtable @{Level=2; StartTime=$bootTime} -ErrorAction SilentlyContinue | Select-Object -First 10
$sysErrCount = if ($sysErrors) { $sysErrors.Count } else { 0 }
$sysErrStatus = if ($sysErrCount -eq 0) { "PASS" } elseif ($sysErrCount -lt 5) { "WARN" } else { "FAIL" }
$rows_events += New-CheckRow -Check "System Critical Errors (Since Boot)" -Status $sysErrStatus `
    -Detail "$sysErrCount errors since last boot"

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 — SHELL + PROFILE INTEGRITY
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[7/10] Shell + Profile Integrity..." -ForegroundColor Yellow
$rows_shell = @()

# Shell folder paths
$shellFolders = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -ErrorAction SilentlyContinue
$shellStatus = if ($shellFolders) { "PASS" } else { "FAIL" }
$rows_shell += New-CheckRow -Check "HKCU Shell Folders Registry Key" -Status $shellStatus `
    -Detail $(if ($shellFolders) { "Present" } else { "MISSING — may cause Explorer errors at login" }) `
    -Recommendation $(if (-not $shellFolders) { "Shell folders key absent. Profile corruption likely. Rebuild golden image profile." } else { "" })

# OneDrive autorun
$odRun = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty "OneDrive" -ErrorAction SilentlyContinue
$rows_shell += New-CheckRow -Check "OneDrive Autorun (HKLM Run)" -Status "INFO" `
    -Detail $(if ($odRun) { "Present: $odRun" } else { "Not set in HKLM Run (may be in HKCU)" })

# Explorer running
$explorerProc = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
$rows_shell += New-CheckRow -Check "Explorer.exe Process" `
    -Status $(if ($explorerProc) { "PASS" } else { "FAIL" }) `
    -Detail $(if ($explorerProc) { "Running (PID: $($explorerProc.Id -join ', '))" } else { "NOT running" }) `
    -Recommendation $(if (-not $explorerProc) { "Explorer not running. Shell crash or not yet started." } else { "" })

# Taskbar / StartMenu registry
$taskbarReg = Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Taskband" -ErrorAction SilentlyContinue
$rows_shell += New-CheckRow -Check "Taskbar Registry (Taskband)" -Status $(if ($taskbarReg) {"INFO"} else {"WARN"}) `
    -Detail $(if ($taskbarReg) { "Taskband key present" } else { "Taskband key not found (may be normal on first login)" })

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 8 — TPM / SECURITY
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[8/10] TPM + Security State..." -ForegroundColor Yellow
$rows_security = @()

# TPM
$tpm = Get-CimInstance -Namespace "Root\CIMv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction SilentlyContinue
if ($tpm) {
    $tpmStatus = if ($tpm.IsEnabled_InitialValue -and $tpm.IsActivated_InitialValue) { "PASS" } else { "WARN" }
    $rows_security += New-CheckRow -Check "TPM State" -Status $tpmStatus `
        -Detail "Enabled: $($tpm.IsEnabled_InitialValue)  |  Activated: $($tpm.IsActivated_InitialValue)  |  Version: $($tpm.SpecVersion)"
} else {
    $rows_security += New-CheckRow -Check "TPM State" -Status "WARN" `
        -Detail "TPM/vTPM not detected or WMI query failed" `
        -Recommendation "Horizon full clone pools can provision vTPM per-clone at creation. Template itself need not have vTPM."
}

# Secure Boot
$secureBoot = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue
$rows_security += New-CheckRow -Check "Secure Boot" `
    -Status $(if ($secureBoot) { "PASS" } else { "INFO" }) `
    -Detail $(if ($secureBoot -eq $true) { "Enabled" } elseif ($secureBoot -eq $false) { "Disabled" } else { "Not supported / BIOS mode" })

# Defender / AV
$av = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
$avStatus = if ($av) { "INFO" } else { "WARN" }
$rows_security += New-CheckRow -Check "Antivirus Product" -Status $avStatus `
    -Detail $(if ($av) { ($av | ForEach-Object { $_.displayName }) -join ", " } else { "No AV detected via SecurityCenter2" })

# Windows Defender Service
$wdSvc = Test-ServiceStatus -ServiceName "WinDefend"
$rows_security += New-CheckRow -Check "Windows Defender Service" -Status $wdSvc.Status -Detail $wdSvc.Detail

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 9 — NETWORK
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[9/10] Network..." -ForegroundColor Yellow
$rows_network = @()

# IP / NIC
$nics = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" }
foreach ($nic in $nics) {
    $rows_network += New-CheckRow -Check "NIC: $($nic.InterfaceAlias)" -Status "INFO" `
        -Detail "$($nic.IPAddress) / $($nic.PrefixLength)"
}

# DNS
$dns = Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" -and $_.ServerAddresses.Count -gt 0 }
foreach ($d in $dns) {
    $rows_network += New-CheckRow -Check "DNS: $($d.InterfaceAlias)" -Status "INFO" `
        -Detail ($d.ServerAddresses -join ", ")
}

# Domain Controller connectivity
if ($cs.PartOfDomain) {
    $dcTest = Test-NetConnection -ComputerName $cs.Domain -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue
    $rows_network += New-CheckRow -Check "Domain Controller LDAP (389)" `
        -Status $(if ($dcTest) { "PASS" } else { "FAIL" }) `
        -Detail "$($cs.Domain):389 — $(if ($dcTest) {'Reachable'} else {'UNREACHABLE'})" `
        -Recommendation $(if (-not $dcTest) { "Check DNS and firewall for DC connectivity." } else { "" })
}

# Horizon Connection Server (optional)
if ($HorizonConnectionServer -and $HorizonConnectionServer -ne "") {
    $horizonConn = Test-NetConnection -ComputerName $HorizonConnectionServer -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue
    $rows_network += New-CheckRow -Check "Horizon Connection Server 443" `
        -Status $(if ($horizonConn) { "PASS" } else { "FAIL" }) `
        -Detail "$HorizonConnectionServer`:443 — $(if ($horizonConn) {'Reachable'} else {'UNREACHABLE'})"
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 10 — SUMMARY SCORE
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "[10/10] Calculating Summary..." -ForegroundColor Yellow

$allRows = @($rows_system + $rows_horizon + $rows_appvol + $rows_nvidia + $rows_appx + $rows_events + $rows_shell + $rows_security + $rows_network)
$totalFail = ([regex]::Matches($allRows -join "", "badge fail")).Count
$totalWarn = ([regex]::Matches($allRows -join "", "badge warn")).Count
$totalPass = ([regex]::Matches($allRows -join "", "badge pass")).Count

$overallStatus = if ($totalFail -gt 0) { "CRITICAL" } elseif ($totalWarn -gt 3) { "DEGRADED" } elseif ($totalWarn -gt 0) { "WARNING" } else { "HEALTHY" }
$overallColor  = switch ($overallStatus) {
    "CRITICAL"  { "#ff3b3b" }
    "DEGRADED"  { "#ff8c00" }
    "WARNING"   { "#f5c518" }
    "HEALTHY"   { "#00e676" }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>VDI Health Report — $env:COMPUTERNAME</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Share+Tech+Mono&family=Exo+2:wght@300;400;600;800&display=swap');

  :root {
    --bg:        #0a0e17;
    --surface:   #111827;
    --surface2:  #1a2233;
    --border:    #1e3a5f;
    --accent:    #00b4ff;
    --accent2:   #00e676;
    --text:      #c9d8e8;
    --text-dim:  #5a7a9a;
    --pass:      #00e676;
    --fail:      #ff3b3b;
    --warn:      #f5c518;
    --info:      #00b4ff;
    --skip:      #5a7a9a;
    --mono:      'Share Tech Mono', monospace;
    --sans:      'Exo 2', sans-serif;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    background: var(--bg);
    color: var(--text);
    font-family: var(--sans);
    font-size: 13px;
    line-height: 1.6;
    padding: 0;
  }

  /* ── HEADER ── */
  .header {
    background: linear-gradient(135deg, #0a1628 0%, #0d1f3c 50%, #091420 100%);
    border-bottom: 1px solid var(--border);
    padding: 32px 40px 24px;
    position: relative;
    overflow: hidden;
  }
  .header::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background: repeating-linear-gradient(
      0deg, transparent, transparent 39px,
      rgba(0,180,255,0.03) 39px, rgba(0,180,255,0.03) 40px
    ),
    repeating-linear-gradient(
      90deg, transparent, transparent 39px,
      rgba(0,180,255,0.03) 39px, rgba(0,180,255,0.03) 40px
    );
    pointer-events: none;
  }
  .header-top { display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 16px; }
  .header h1 {
    font-family: var(--sans);
    font-weight: 800;
    font-size: 22px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--accent);
    text-shadow: 0 0 20px rgba(0,180,255,0.4);
  }
  .header h2 {
    font-family: var(--mono);
    font-size: 13px;
    color: var(--text-dim);
    margin-top: 4px;
    letter-spacing: 0.05em;
  }
  .overall-badge {
    font-family: var(--sans);
    font-weight: 800;
    font-size: 15px;
    letter-spacing: 0.12em;
    padding: 10px 24px;
    border-radius: 4px;
    border: 2px solid;
    text-transform: uppercase;
    color: $overallColor;
    border-color: $overallColor;
    box-shadow: 0 0 20px ${overallColor}40;
    text-shadow: 0 0 10px ${overallColor}80;
    white-space: nowrap;
  }
  .meta-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 12px;
    margin-top: 20px;
  }
  .meta-item {
    background: rgba(0,180,255,0.05);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 8px 12px;
  }
  .meta-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.1em; color: var(--text-dim); }
  .meta-value { font-family: var(--mono); font-size: 12px; color: var(--accent); margin-top: 2px; }

  /* ── SCORE BAR ── */
  .score-bar {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 16px 40px;
    display: flex;
    align-items: center;
    gap: 32px;
    flex-wrap: wrap;
  }
  .score-item { display: flex; align-items: center; gap: 8px; }
  .score-num { font-family: var(--mono); font-size: 22px; font-weight: bold; }
  .score-lbl { font-size: 11px; text-transform: uppercase; letter-spacing: 0.08em; color: var(--text-dim); }
  .score-num.pass { color: var(--pass); }
  .score-num.warn { color: var(--warn); }
  .score-num.fail { color: var(--fail); }
  .score-divider { width: 1px; height: 40px; background: var(--border); }

  /* ── CONTENT ── */
  .content { padding: 24px 40px; }

  /* ── SECTION ── */
  .section {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 6px;
    margin-bottom: 20px;
    overflow: hidden;
  }
  .section-header {
    background: linear-gradient(90deg, var(--surface2), var(--surface));
    border-bottom: 1px solid var(--border);
    padding: 12px 16px;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .section-icon { font-size: 16px; }
  .section-title {
    font-weight: 600;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--accent);
  }
  .section-count {
    margin-left: auto;
    font-family: var(--mono);
    font-size: 11px;
    color: var(--text-dim);
  }

  /* ── TABLE ── */
  table { width: 100%; border-collapse: collapse; }
  th {
    background: var(--surface2);
    color: var(--text-dim);
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    padding: 8px 14px;
    text-align: left;
    border-bottom: 1px solid var(--border);
  }
  td {
    padding: 9px 14px;
    border-bottom: 1px solid rgba(30,58,95,0.5);
    vertical-align: top;
  }
  tr:last-child td { border-bottom: none; }
  .check-name { font-weight: 600; font-size: 12px; color: var(--text); white-space: nowrap; width: 260px; }
  td:nth-child(3) { font-family: var(--mono); font-size: 11px; color: var(--text-dim); word-break: break-word; }
  .rec {
    margin-top: 5px;
    font-family: var(--sans);
    font-size: 11px;
    color: var(--warn);
    opacity: 0.85;
  }
  tr.row-fail   { background: rgba(255,59,59,0.04); }
  tr.row-warn   { background: rgba(245,197,24,0.04); }
  tr.row-pass   { background: rgba(0,230,118,0.02); }
  tr:hover      { background: rgba(0,180,255,0.05) !important; }

  /* ── BADGES ── */
  .badge {
    display: inline-block;
    font-family: var(--mono);
    font-size: 10px;
    font-weight: bold;
    padding: 3px 8px;
    border-radius: 3px;
    letter-spacing: 0.06em;
    white-space: nowrap;
  }
  .badge.pass { background: rgba(0,230,118,0.12); color: var(--pass); border: 1px solid rgba(0,230,118,0.3); }
  .badge.fail { background: rgba(255,59,59,0.12); color: var(--fail); border: 1px solid rgba(255,59,59,0.3); }
  .badge.warn { background: rgba(245,197,24,0.12); color: var(--warn); border: 1px solid rgba(245,197,24,0.3); }
  .badge.info { background: rgba(0,180,255,0.12); color: var(--info); border: 1px solid rgba(0,180,255,0.3); }
  .badge.skip { background: rgba(90,122,154,0.12); color: var(--skip); border: 1px solid rgba(90,122,154,0.3); }

  /* ── FOOTER ── */
  .footer {
    border-top: 1px solid var(--border);
    padding: 16px 40px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-family: var(--mono);
    font-size: 11px;
    color: var(--text-dim);
    background: var(--surface);
  }
  .footer-brand { color: var(--accent); opacity: 0.6; }

  @media print {
    body { background: #fff; color: #000; }
    .header { background: #f0f4f8 !important; }
    .section, table, td, th { border-color: #ccc !important; }
  }
</style>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div>
      <div style="font-family:'Share Tech Mono',monospace;font-size:11px;color:#5a7a9a;margin-bottom:4px;letter-spacing:0.15em;">
        OMNISSA HORIZON 2512 // NOVA POOL // FULL CLONE
      </div>
      <h1>⬡ VDI Health Report</h1>
      <h2>$env:COMPUTERNAME  //  Windows 11 24H2  //  NVIDIA vGPU 582.x</h2>
    </div>
    <div class="overall-badge">$overallStatus</div>
  </div>
  <div class="meta-grid">
    <div class="meta-item"><div class="meta-label">Report Generated</div><div class="meta-value">$RunTime</div></div>
    <div class="meta-item"><div class="meta-label">Computer</div><div class="meta-value">$env:COMPUTERNAME</div></div>
    <div class="meta-item"><div class="meta-label">Run As</div><div class="meta-value">$env:USERNAME</div></div>
    <div class="meta-item"><div class="meta-label">App Vol Manager</div><div class="meta-value">$AppVolManagerFQDN</div></div>
    <div class="meta-item"><div class="meta-label">Expected GPU Driver</div><div class="meta-value">$ExpectedDriverVersion</div></div>
  </div>
</div>

<div class="score-bar">
  <div class="score-item">
    <div class="score-num pass">$totalPass</div>
    <div class="score-lbl">Passed</div>
  </div>
  <div class="score-divider"></div>
  <div class="score-item">
    <div class="score-num warn">$totalWarn</div>
    <div class="score-lbl">Warnings</div>
  </div>
  <div class="score-divider"></div>
  <div class="score-item">
    <div class="score-num fail">$totalFail</div>
    <div class="score-lbl">Failed</div>
  </div>
  <div class="score-divider"></div>
  <div style="font-family:'Share Tech Mono',monospace;font-size:11px;color:#5a7a9a;">
    Overall: <span style="color:$overallColor;font-weight:bold;">$overallStatus</span>
  </div>
</div>

<div class="content">

  <div class="section">
    <div class="section-header">
      <span class="section-icon">🖥</span>
      <span class="section-title">System Baseline</span>
      <span class="section-count">OS · Domain · Activation · Disk · Sysprep</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_system -join "`n")
    </table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">🔷</span>
      <span class="section-title">VMware Tools + Horizon Agent</span>
      <span class="section-count">Services · Version · Guest Customization</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_horizon -join "`n")
    </table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">📦</span>
      <span class="section-title">App Volumes Agent</span>
      <span class="section-count">svservice · Version · Manager Connectivity · Log Path</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_appvol -join "`n")
    </table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">🎮</span>
      <span class="section-title">NVIDIA vGPU</span>
      <span class="section-count">Driver Version · GPU Profile · Control Panel Client · ECC</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_nvidia -join "`n")
    </table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">📱</span>
      <span class="section-title">AppX Provisioned State</span>
      <span class="section-count">Sysprep Blockers · Per-User Drift · Windows Store</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_appx -join "`n")
    </table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">📋</span>
      <span class="section-title">DCOM + Event Log Audit</span>
      <span class="section-count">10016 · WscDataProtection · Application + System Errors</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_events -join "`n")
    </table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">🪟</span>
      <span class="section-title">Shell + Profile Integrity</span>
      <span class="section-count">Shell Folders · OneDrive · Explorer</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_shell -join "`n")
    </table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">🔒</span>
      <span class="section-title">TPM + Security</span>
      <span class="section-count">vTPM · Secure Boot · Defender · AV</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_security -join "`n")
    </table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">🌐</span>
      <span class="section-title">Network</span>
      <span class="section-count">NIC · DNS · Domain Controller · App Vol Manager</span>
    </div>
    <table>
      <tr><th>Check</th><th>Status</th><th>Detail</th></tr>
      $($rows_network -join "`n")
    </table>
  </div>

</div><!-- /content -->

<div class="footer">
  <div>Report: $ReportFile</div>
  <div class="footer-brand">VDI Health Report v1.0 // Nova Pool // Omnissa Horizon 2512</div>
</div>

</body>
</html>
"@

# ─────────────────────────────────────────────────────────────────────────────
# WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────
$html | Out-File -FilePath $ReportFile -Encoding UTF8 -Force

Write-Host "`n[VDI Health Check] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "[VDI Health Check] COMPLETE" -ForegroundColor Green
Write-Host "[VDI Health Check] Status  : $overallStatus" -ForegroundColor $(if ($overallStatus -eq "HEALTHY") {"Green"} elseif ($overallStatus -eq "WARNING") {"Yellow"} else {"Red"})
Write-Host "[VDI Health Check] PASS    : $totalPass" -ForegroundColor Green
Write-Host "[VDI Health Check] WARN    : $totalWarn" -ForegroundColor Yellow
Write-Host "[VDI Health Check] FAIL    : $totalFail" -ForegroundColor Red
Write-Host "[VDI Health Check] Report  : $ReportFile" -ForegroundColor Cyan
Write-Host "[VDI Health Check] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor Cyan

# Auto-open in default browser
Start-Process $ReportFile
