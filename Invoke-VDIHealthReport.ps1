#Requires -RunAsAdministrator
<#
.SYNOPSIS
    VDI Golden Image Health Check Report — Generic / Interactive

.DESCRIPTION
    Generates a timestamped HTML health report for any Horizon VDI environment.
    Prompts the user for all environment-specific values at runtime.
   

    Error handling: every section is wrapped in try/catch.
    Errors are logged, reported in the HTML, and execution always continues.

    Checks covered:
      - System baseline (OS, domain, activation, disk, sysprep)
      - AppX provisioned vs per-user drift + sysprep blockers
      - DCOM 10016 audit + VBS/Credential Guard (Event 124)
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
      - Script execution errors (self-audit)

.PARAMETER Silent
    Skip interactive prompts. Connectivity checks requiring FQDNs are skipped.

.PARAMETER OutputPath
    Override output folder. Prompted if not provided.

.PARAMETER StopOnSectionFail
    If set, aborts the entire script on the first section-level exception.
    Default behavior: log error and continue all remaining sections.

.EXAMPLE
    .\Invoke-VDIHealthReport.ps1
    .\Invoke-VDIHealthReport.ps1 -Silent
    .\Invoke-VDIHealthReport.ps1 -OutputPath "D:\Reports"
    .\Invoke-VDIHealthReport.ps1 -Silent -StopOnSectionFail

.NOTES
    Version : 1.2
    Generic  : No environment-specific values hardcoded
    Changes  : Full try/catch error handling; VBS Event 124 check added
#>

[CmdletBinding()]
param(
    [switch]$Silent,
    [string]$OutputPath = "",
    [switch]$StopOnSectionFail
)

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL ERROR HANDLING CONFIG
# ErrorActionPreference = Stop so try/catch catches everything.
# Each section wraps risky calls — failures are logged and execution continues.
# ─────────────────────────────────────────────────────────────────────────────
$ErrorActionPreference = "Stop"
$ScriptErrors          = [System.Collections.Generic.List[psobject]]::new()
$SectionTimings        = [System.Collections.Generic.List[psobject]]::new()

function Register-ScriptError {
    param(
        [string]$Section,
        [string]$Check,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $msg = if ($ErrorRecord) { $ErrorRecord.Exception.Message } else { "Unknown error" }
    $ScriptErrors.Add([pscustomobject]@{
        Section   = $Section
        Check     = $Check
        Message   = $msg
        Time      = (Get-Date -Format "HH:mm:ss")
    })
    Write-Host "    [!] ERROR in $Section — $Check`: $msg" -ForegroundColor Red
}

function Invoke-SafeCheck {
    <#
    .SYNOPSIS
        Wraps a scriptblock in try/catch.
        On error: logs to $ScriptErrors, returns a WARN row, continues.
    #>
    param(
        [string]$Section,
        [string]$Check,
        [scriptblock]$ScriptBlock
    )
    try {
        return & $ScriptBlock
    } catch {
        Register-ScriptError -Section $Section -Check $Check -ErrorRecord $_
        return New-CheckRow $Check "WARN" "Check encountered an error — see Script Errors section" `
            "Error: $($_.Exception.Message)"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║          VDI GOLDEN IMAGE HEALTH CHECK REPORT           ║" -ForegroundColor Cyan
    Write-Host "  ║          Omnissa Horizon  |  Generic Edition  v1.2      ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Prompt-Input {
    param(
        [string]$Prompt,
        [string]$Default = "",
        [switch]$Optional
    )
    $label = if ($Default)   { "$Prompt [default: $Default]" } `
             elseif ($Optional) { "$Prompt [press Enter to skip]" } `
             else             { $Prompt }
    Write-Host "  ► " -NoNewline -ForegroundColor Cyan
    Write-Host $label -NoNewline -ForegroundColor White
    Write-Host ": " -NoNewline
    $ErrorActionPreference = "SilentlyContinue"
    $val = Read-Host
    $ErrorActionPreference = "Stop"
    if ([string]::IsNullOrWhiteSpace($val)) { return $Default }
    return $val.Trim()
}

function Write-SectionHeader { param([string]$Text); Write-Host ""; Write-Host "  ─── $Text" -ForegroundColor Yellow }
function Write-Progress-Step { param([string]$N,[string]$Text); Write-Host "  [$N] $Text..." -ForegroundColor Cyan }

# ─────────────────────────────────────────────────────────────────────────────
# INTERACTIVE SETUP
# ─────────────────────────────────────────────────────────────────────────────
Write-Banner

try {
    if (-not $Silent) {
        Write-Host "  Fields with defaults: press Enter to accept." -ForegroundColor Gray
        Write-Host "  Optional fields: press Enter to skip." -ForegroundColor Gray

        Write-SectionHeader "ENVIRONMENT IDENTITY"
        $EnvName           = Prompt-Input "Environment / Pool Name (e.g. Production, Dev, QA)" -Default "VDI"
        $PoolType          = Prompt-Input "Pool Type (FullClone / InstantClone / LinkedClone)"  -Default "FullClone"

        Write-SectionHeader "OUTPUT"
        if ([string]::IsNullOrWhiteSpace($OutputPath)) {
            $OutputPath    = Prompt-Input "Report output folder" -Default "C:\Temp\VDIHealthReports"
        }

        Write-SectionHeader "NVIDIA vGPU"
        $ExpectedDriverVer = Prompt-Input "Expected vGPU guest driver version" -Default "582.16"

        Write-SectionHeader "APP VOLUMES MANAGER"
        $AppVolManagerFQDN = Prompt-Input "App Volumes Manager FQDN or IP" -Optional
        $AppVolManagerPort = if ($AppVolManagerFQDN) { Prompt-Input "App Volumes Manager port" -Default "443" } else { "443" }

        Write-SectionHeader "HORIZON CONNECTION SERVER (optional)"
        $HorizonCS         = Prompt-Input "Horizon Connection Server FQDN or IP" -Optional

        Write-SectionHeader "ADDITIONAL CONNECTIVITY CHECKS (optional)"
        $ExtraHost1        = Prompt-Input "Additional host to test" -Optional
        $ExtraHost2        = Prompt-Input "Additional host #2"      -Optional

        Write-Host ""
        Write-Host "  Configuration confirmed. Starting health checks..." -ForegroundColor Green
        Write-Host ""
    } else {
        $EnvName="VDI"; $PoolType="FullClone"; $ExpectedDriverVer="582.16"
        $AppVolManagerFQDN=""; $AppVolManagerPort="443"; $HorizonCS=""
        $ExtraHost1=""; $ExtraHost2=""
        if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = "C:\Temp\VDIHealthReports" }
        Write-Host "  [Silent mode] Running with auto-detected values..." -ForegroundColor Yellow
        Write-Host ""
    }
} catch {
    Write-Host "  [FATAL] Setup failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT SETUP
# ─────────────────────────────────────────────────────────────────────────────
try {
    if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
} catch {
    Write-Host "  [WARN] Could not create output dir '$OutputPath', falling back to C:\Temp" -ForegroundColor Yellow
    $OutputPath = "C:\Temp"
    if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
}

$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = Join-Path $OutputPath "VDIHealth_$($env:COMPUTERNAME)_$Timestamp.html"
$RunTime    = Get-Date -Format "dddd, MMMM dd yyyy  HH:mm:ss"
$RunStart   = Get-Date

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
        "ERR"  { return '<span class="badge fail">ERR</span>'  }
        default{ return '<span class="badge info">UNKN</span>' }
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
    try {
        $svc = Get-Service -Name $ServiceName -ErrorAction Stop
        if ($svc.Status -eq $ExpectedState) {
            return @{ Status="PASS"; Detail="$ServiceName — $($svc.Status)" }
        }
        return @{ Status="FAIL"; Detail="$ServiceName — Expected: $ExpectedState | Actual: $($svc.Status)" }
    } catch {
        return @{ Status="FAIL"; Detail="Service not found or inaccessible: $ServiceName" }
    }
}

function Test-PortConnection {
    param([string]$HostName, [int]$Port, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($HostName)) {
        return @{ Status="SKIP"; Detail="$Label — no host provided"; Recommendation="" }
    }
    try {
        $result = Test-NetConnection -ComputerName $HostName -Port $Port -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction Stop
        return @{
            Status         = if ($result) {"PASS"} else {"FAIL"}
            Detail         = "$Label — ${HostName}:${Port} — $(if ($result) {'Reachable'} else {'UNREACHABLE'})"
            Recommendation = if (-not $result) {"Check DNS resolution and firewall rules for ${HostName}:${Port}"} else {""}
        }
    } catch {
        return @{ Status="WARN"; Detail="$Label — ${HostName}:${Port} — Test failed: $($_.Exception.Message)"; Recommendation="Verify host is reachable and DNS resolves correctly." }
    }
}

function Start-SectionTimer { param([string]$Name); return @{ Name=$Name; Start=(Get-Date) } }
function Stop-SectionTimer  {
    param([hashtable]$Timer)
    $elapsed = ((Get-Date) - $Timer.Start).TotalSeconds
    $SectionTimings.Add([pscustomobject]@{ Section=$Timer.Name; Seconds=[math]::Round($elapsed,2) })
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — SYSTEM BASELINE
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "1/10" "System Baseline"
$rows_system = @()
$t = Start-SectionTimer "System Baseline"
try {
    $os        = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $cs        = Get-CimInstance Win32_ComputerSystem  -ErrorAction Stop
    $buildNum  = [int]$os.BuildNumber
    $uptime    = (Get-Date) - $os.LastBootUpTime
    $ramGB     = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)

    $rows_system += New-CheckRow "OS Version"   "INFO" "$($os.Caption) — Build $buildNum | $($os.OSArchitecture)"
    $rows_system += New-CheckRow "Hostname"     "INFO" $env:COMPUTERNAME
    $rows_system += New-CheckRow "Environment"  "INFO" "Name: $EnvName | Pool Type: $PoolType"

    $bStatus = if ($buildNum -ge 26100) {"PASS"} else {"WARN"}
    $rows_system += New-CheckRow "Windows 11 24H2 Build" $bStatus "Build $buildNum  (24H2 minimum = 26100)" `
        $(if ($bStatus -eq "WARN") {"Build below 24H2 baseline. Verify ISO source."} else {""})

    $dStatus = if ($cs.PartOfDomain) {"PASS"} else {"WARN"}
    $rows_system += New-CheckRow "Domain Join" $dStatus `
        $(if ($cs.PartOfDomain) {"Joined: $($cs.Domain)"} else {"NOT domain joined"}) `
        $(if (-not $cs.PartOfDomain) {"Verify Guest Customization completed."} else {""})

    try {
        $lic = Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%' AND LicenseStatus=1" -ErrorAction Stop | Select-Object -First 1
        $aStatus = if ($lic) {"PASS"} else {"WARN"}
        $rows_system += New-CheckRow "Windows Activation" $aStatus `
            $(if ($lic) {"Activated — $($lic.Name)"} else {"Not activated or unable to verify"}) `
            $(if (-not $lic) {"Check KMS connectivity or MAK activation."} else {""})
    } catch {
        $rows_system += New-CheckRow "Windows Activation" "WARN" "Query failed: $($_.Exception.Message)"
        Register-ScriptError "System Baseline" "Windows Activation" $_
    }

    $rows_system += New-CheckRow "Last Boot / Uptime" "INFO" `
        "Booted: $($os.LastBootUpTime.ToString('MM/dd/yyyy HH:mm:ss'))  |  Uptime: $([math]::Floor($uptime.TotalHours))h $($uptime.Minutes)m"

    $ramStatus = if ($ramGB -ge 4) {"PASS"} else {"WARN"}
    $rows_system += New-CheckRow "Physical RAM" $ramStatus "$ramGB GB" `
        $(if ($ramGB -lt 4) {"Minimum 4GB recommended for VDI desktop."} else {""})

    try {
        $disk      = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
        $diskFree  = [math]::Round($disk.FreeSpace / 1GB, 1)
        $diskTotal = [math]::Round($disk.Size / 1GB, 1)
        $diskPct   = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 0)
        $diskStatus= if ($diskPct -ge 20) {"PASS"} elseif ($diskPct -ge 10) {"WARN"} else {"FAIL"}
        $rows_system += New-CheckRow "Disk C: Free Space" $diskStatus "$diskFree GB free of $diskTotal GB ($diskPct% free)" `
            $(if ($diskStatus -eq "FAIL") {"Critical — under 10% free. Image may be oversized."} `
              elseif ($diskStatus -eq "WARN") {"Consider cleanup or expanding template disk."} else {""})
    } catch {
        $rows_system += New-CheckRow "Disk C: Free Space" "WARN" "Query failed: $($_.Exception.Message)"
        Register-ScriptError "System Baseline" "Disk C" $_
    }

    try {
        $setupKey  = Get-ItemProperty "HKLM:\SYSTEM\Setup" -ErrorAction Stop
        $oobe      = $setupKey.OOBEInProgress
        $setupType = $setupKey.SetupType
        $spStatus  = if ($oobe -eq 0 -and $setupType -eq 0) {"PASS"} else {"WARN"}
        $rows_system += New-CheckRow "Sysprep / OOBE State" $spStatus `
            "OOBEInProgress=$oobe | SetupType=$setupType  (both must be 0 on deployed clone)" `
            $(if ($spStatus -eq "WARN") {"VM may still be in generalize/OOBE state. Verify Guest Customization."} else {""})
    } catch {
        $rows_system += New-CheckRow "Sysprep / OOBE State" "WARN" "Registry query failed: $($_.Exception.Message)"
        Register-ScriptError "System Baseline" "Sysprep State" $_
    }

} catch {
    $rows_system += New-CheckRow "System Baseline" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "System Baseline" "Section Init" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — VMWARE TOOLS + HORIZON AGENT
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "2/10" "VMware Tools + Horizon Agent"
$rows_horizon = @()
$t = Start-SectionTimer "Horizon Agent"
try {
    $toolsResult = Test-ServiceStatus -ServiceName "VMTools"
    $rows_horizon += New-CheckRow "VMware Tools Service" $toolsResult.Status $toolsResult.Detail `
        $(if ($toolsResult.Status -eq "FAIL") {"VMware Tools required for Guest Customization and Horizon."} else {""})

    try {
        $toolsReg = Get-ItemProperty "HKLM:\SOFTWARE\VMware, Inc.\VMware Tools" -ErrorAction Stop
        $rows_horizon += New-CheckRow "VMware Tools Version" "INFO" $toolsReg.Version
    } catch {
        $rows_horizon += New-CheckRow "VMware Tools Version" "WARN" "Registry key not found — Tools may not be installed."
    }

    foreach ($svc in @(
        @{ Name="vmware-viewagent";  Label="Horizon View Agent";         Required=$true  },
        @{ Name="wsnm";              Label="Horizon WSNM (Blast/PCoIP)"; Required=$true  },
        @{ Name="CSVD";              Label="Horizon CSVD (USB/Scanner)"; Required=$false },
        @{ Name="vmware-view-usbd"; Label="Horizon USB Arbitrator";     Required=$false }
    )) {
        try {
            $r = Test-ServiceStatus -ServiceName $svc.Name
            $finalStatus = if ($r.Status -eq "FAIL" -and -not $svc.Required) {"WARN"} else {$r.Status}
            $rows_horizon += New-CheckRow $svc.Label $finalStatus $r.Detail
        } catch {
            $rows_horizon += New-CheckRow $svc.Label "WARN" "Service check failed: $($_.Exception.Message)"
            Register-ScriptError "Horizon Agent" $svc.Label $_
        }
    }

    try {
        $hReg = Get-ItemProperty "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\Agent" -ErrorAction Stop
        $rows_horizon += New-CheckRow "Horizon Agent Version" "INFO" $hReg.ProductVersion
    } catch {
        $rows_horizon += New-CheckRow "Horizon Agent Version" "WARN" "Registry key not found. Verify Horizon Agent is installed."
    }

    try {
        $gcLog = "C:\Windows\Temp\vmware-imc\guestcust.log"
        if (Test-Path $gcLog -ErrorAction Stop) {
            $gcTail   = (Get-Content $gcLog -Tail 5 -ErrorAction Stop) -join " | "
            $gcStatus = if ($gcTail -match "success|Customization finished") {"PASS"} `
                        elseif ($gcTail -match "error|fail") {"FAIL"} else {"INFO"}
            $rows_horizon += New-CheckRow "Guest Customization Log (last 5 lines)" $gcStatus $gcTail `
                $(if ($gcStatus -eq "FAIL") {"Review: C:\Windows\Temp\vmware-imc\guestcust.log"} else {""})
        } else {
            $rows_horizon += New-CheckRow "Guest Customization Log" "WARN" "Not found: $gcLog" `
                "Log absent — customization may not have run, or VM was not deployed via Horizon pool."
        }
    } catch {
        $rows_horizon += New-CheckRow "Guest Customization Log" "WARN" "Read failed: $($_.Exception.Message)"
        Register-ScriptError "Horizon Agent" "GuestCust Log" $_
    }

} catch {
    $rows_horizon += New-CheckRow "Horizon Agent Section" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "Horizon Agent" "Section" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — APP VOLUMES AGENT
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "3/10" "App Volumes Agent"
$rows_appvol = @()
$t = Start-SectionTimer "App Volumes"
try {
    $svResult = Test-ServiceStatus -ServiceName "svservice"
    $rows_appvol += New-CheckRow "App Volumes Agent (svservice)" $svResult.Status $svResult.Detail `
        $(if ($svResult.Status -eq "FAIL") {"svservice not running. Reinstall App Volumes Agent."} else {""})

    try {
        $avReg = Get-ItemProperty "HKLM:\SOFTWARE\CloudVolumes\Agent" -ErrorAction Stop
        $rows_appvol += New-CheckRow "App Volumes Agent Version" "INFO" $avReg.Version
    } catch {
        try {
            $avReg = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\CloudVolumes\Agent" -ErrorAction Stop
            $rows_appvol += New-CheckRow "App Volumes Agent Version" "INFO" "$($avReg.Version) (WOW64)"
        } catch {
            $rows_appvol += New-CheckRow "App Volumes Agent Version" "WARN" "Registry key not found — agent may not be installed."
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($AppVolManagerFQDN)) {
        try {
            $mgrResult = Test-PortConnection -HostName $AppVolManagerFQDN -Port ([int]$AppVolManagerPort) -Label "App Volumes Manager"
            $rows_appvol += New-CheckRow "App Volumes Manager Connectivity" $mgrResult.Status $mgrResult.Detail $mgrResult.Recommendation
        } catch {
            $rows_appvol += New-CheckRow "App Volumes Manager Connectivity" "WARN" "Connectivity test failed: $($_.Exception.Message)"
            Register-ScriptError "App Volumes" "Manager Connectivity" $_
        }
    } else {
        $rows_appvol += New-CheckRow "App Volumes Manager Connectivity" "SKIP" "No manager host provided."
    }

    foreach ($p in @("C:\Program Files (x86)\CloudVolumes\Agent\log", "C:\ProgramData\CloudVolumes\Logs")) {
        try {
            $pStatus = if (Test-Path $p -ErrorAction Stop) {"PASS"} else {"WARN"}
            $rows_appvol += New-CheckRow "App Volumes Log Path" $pStatus $p `
                $(if ($pStatus -eq "WARN") {"Log directory missing — agent may not be installed correctly."} else {""})
        } catch {
            $rows_appvol += New-CheckRow "App Volumes Log Path" "WARN" "Path check failed: $p"
            Register-ScriptError "App Volumes" "Log Path: $p" $_
        }
    }

} catch {
    $rows_appvol += New-CheckRow "App Volumes Section" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "App Volumes" "Section" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — NVIDIA vGPU
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "4/10" "NVIDIA vGPU"
$rows_nvidia = @()
$t = Start-SectionTimer "NVIDIA vGPU"
try {
    $nvidiaSmi = "C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"

    if (Test-Path $nvidiaSmi) {
        try {
            $driverVer    = (& $nvidiaSmi --query-gpu=driver_version --format=csv,noheader 2>$null).Trim()
            $driverStatus = if ($driverVer -eq $ExpectedDriverVer) {"PASS"} elseif ($driverVer) {"WARN"} else {"FAIL"}
            $rows_nvidia += New-CheckRow "NVIDIA Driver Version" $driverStatus `
                "Installed: $driverVer  |  Expected: $ExpectedDriverVer" `
                $(if ($driverStatus -eq "WARN") {"Driver version mismatch. Source correct driver from NVIDIA licensing portal."} `
                  elseif ($driverStatus -eq "FAIL") {"Driver not detected. Install vGPU guest driver from NVIDIA licensing portal."} else {""})
        } catch {
            $rows_nvidia += New-CheckRow "NVIDIA Driver Version" "WARN" "nvidia-smi query failed: $($_.Exception.Message)"
            Register-ScriptError "NVIDIA vGPU" "Driver Version Query" $_
        }

        try {
            $gpuName = (& $nvidiaSmi --query-gpu=name --format=csv,noheader 2>$null).Trim()
            $rows_nvidia += New-CheckRow "GPU Name / Profile" "INFO" $gpuName
        } catch {
            $rows_nvidia += New-CheckRow "GPU Name / Profile" "WARN" "Query failed."
        }

        try {
            $gpuState = (& $nvidiaSmi --query-gpu=pstate,temperature.gpu,utilization.gpu --format=csv,noheader 2>$null).Trim()
            $rows_nvidia += New-CheckRow "GPU P-State / Temp / Utilization" "INFO" $gpuState
        } catch {
            $rows_nvidia += New-CheckRow "GPU P-State / Temp / Utilization" "WARN" "Query failed."
        }

        try {
            $gpuErrors = (& $nvidiaSmi --query-gpu=ecc.errors.corrected.volatile.total,ecc.errors.uncorrected.volatile.total --format=csv,noheader 2>$null).Trim()
            if ($gpuErrors -and $gpuErrors -notmatch "N/A") {
                $errVals   = $gpuErrors -split ","
                $eccStatus = if ([int]$errVals[1].Trim() -gt 0) {"WARN"} else {"PASS"}
                $rows_nvidia += New-CheckRow "GPU ECC Errors (Uncorrected)" $eccStatus `
                    "Corrected: $($errVals[0].Trim())  |  Uncorrected: $($errVals[1].Trim())"
            }
        } catch {
            $rows_nvidia += New-CheckRow "GPU ECC Errors" "WARN" "ECC query failed (may not be supported on this vGPU profile)."
        }
    } else {
        $rows_nvidia += New-CheckRow "nvidia-smi.exe" "FAIL" "Not found at: $nvidiaSmi" `
            "Install vGPU guest driver from NVIDIA licensing portal (nvid.nvidia.com). Do NOT use consumer/GeForce driver."
    }

    $nvCpPath = "C:\Program Files\NVIDIA Corporation\Control Panel Client\nvcplui.exe"
    $cpStatus = if (Test-Path $nvCpPath) {"PASS"} else {"FAIL"}
    $rows_nvidia += New-CheckRow "NVIDIA Control Panel Client" $cpStatus `
        $(if ($cpStatus -eq "PASS") {$nvCpPath} else {"NOT FOUND: $nvCpPath"}) `
        $(if ($cpStatus -eq "FAIL") {"Control Panel Client missing. Source full vGPU guest driver package from NVIDIA licensing portal — consumer driver packages exclude this component."} else {""})

} catch {
    $rows_nvidia += New-CheckRow "NVIDIA vGPU Section" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "NVIDIA vGPU" "Section" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — APPX PROVISIONED STATE
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "5/10" "AppX Provisioned State"
$rows_appx = @()
$t = Start-SectionTimer "AppX State"
try {
    $provisionedPkgs  = Get-AppxProvisionedPackage -Online -ErrorAction Stop
    $provisionedNames = $provisionedPkgs | Select-Object -ExpandProperty PackageName
    $userPkgs         = Get-AppxPackage -AllUsers -ErrorAction Stop

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

    try {
        $store       = Get-AppxPackage -AllUsers -Name "Microsoft.WindowsStore" -ErrorAction Stop
        $storeStatus = if ($store -and $store.Status -eq "Ok") {"PASS"} elseif ($store) {"WARN"} else {"FAIL"}
        $rows_appx += New-CheckRow "Windows Store" $storeStatus `
            $(if ($store) {"Status: $($store.Status)  |  Version: $($store.Version)"} else {"Not installed"}) `
            $(if ($storeStatus -ne "PASS") {"Re-register Store AppX or re-provision via DISM /Add-Capability."} else {""})
    } catch {
        $rows_appx += New-CheckRow "Windows Store" "WARN" "Query failed: $($_.Exception.Message)"
        Register-ScriptError "AppX" "Windows Store" $_
    }

} catch {
    $rows_appx += New-CheckRow "AppX Section" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "AppX" "Section" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 — DCOM + EVENT LOG + VBS
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "6/10" "DCOM + Event Log + VBS Audit"
$rows_events = @()
$t = Start-SectionTimer "DCOM + Events"
try {
    # DCOM 10016
    try {
        $dcom10016  = Get-WinEvent -LogName System -FilterHashtable @{Id=10016; StartTime=(Get-Date).AddHours(-24)} -ErrorAction Stop
        $dcomCount  = $dcom10016.Count
    } catch [System.Exception] {
        if ($_.Exception.Message -match "No events were found") { $dcomCount = 0; $dcom10016 = $null }
        else { throw }
    }
    $dcomStatus = if ($dcomCount -eq 0) {"PASS"} elseif ($dcomCount -lt 10) {"WARN"} else {"FAIL"}
    $rows_events += New-CheckRow "DCOM 10016 Errors (Last 24h)" $dcomStatus "$dcomCount occurrences" `
        $(if ($dcomStatus -ne "PASS") {"Fix DCOM permissions via dcomcnfg, or disable wscsvc if Security Center not required."} else {""})

    if ($dcom10016) {
        $wscDcom = $dcom10016 | Where-Object {$_.Message -like "*WscDataProtection*"} | Select-Object -First 1
        if ($wscDcom) {
            $rows_events += New-CheckRow "DCOM 10016 — WscDataProtection" "WARN" `
                "WscDataProtection LocalLaunch permission denied to SYSTEM" `
                "Grant NT AUTHORITY\SYSTEM Local Launch + Local Activation via dcomcnfg, OR disable wscsvc if not needed."
        }
    }

    # VBS / Credential Guard — Event 124 Kernel-Boot
    try {
        $vbs124 = Get-WinEvent -LogName System -FilterHashtable @{Id=124; StartTime=(Get-Date).AddDays(-1)} -ErrorAction Stop |
            Where-Object {$_.ProviderName -like "*Kernel-Boot*"} | Select-Object -First 3
        $vbs124Count = if ($vbs124) {$vbs124.Count} else {0}
    } catch {
        $vbs124Count = 0; $vbs124 = $null
    }
    $vbsStatus = if ($vbs124Count -eq 0) {"PASS"} else {"WARN"}
    $rows_events += New-CheckRow "VBS / Credential Guard (Event 124)" $vbsStatus `
        $(if ($vbs124Count -gt 0) {"$vbs124Count occurrence(s) — 'VBS enablement policy check at phase 0 failed: not supported'"} `
          else {"No Event 124 detected — VBS not conflicting"}) `
        $(if ($vbs124Count -gt 0) {"Disable VBS via GPO or registry: HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard — EnableVirtualizationBasedSecurity = 0"} else {""})

    # Application errors since boot
    $bootTime = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).LastBootUpTime
    try {
        $appErrors = Get-WinEvent -LogName Application -FilterHashtable @{Level=2; StartTime=$bootTime} -ErrorAction Stop | Select-Object -First 15
        $appErrCnt = $appErrors.Count
    } catch {
        $appErrCnt = 0; $appErrors = $null
    }
    $aeStatus = if ($appErrCnt -eq 0) {"PASS"} elseif ($appErrCnt -lt 5) {"WARN"} else {"FAIL"}
    $rows_events += New-CheckRow "Application Errors (Since Boot)" $aeStatus `
        "$appErrCnt errors since boot ($($bootTime.ToString('MM/dd HH:mm')))"

    if ($appErrors) {
        $errList = ($appErrors | Select-Object -First 5 | ForEach-Object {
            "$($_.TimeCreated.ToString('HH:mm:ss')) | $($_.ProviderName) | $($_.Message.Substring(0,[Math]::Min(120,$_.Message.Length)))"
        }) -join "<br>"
        $rows_events += New-CheckRow "Top Application Errors" "INFO" $errList
    }

    # System errors since boot
    try {
        $sysErrors = Get-WinEvent -LogName System -FilterHashtable @{Level=2; StartTime=$bootTime} -ErrorAction Stop | Select-Object -First 10
        $sysErrCnt = $sysErrors.Count
    } catch {
        $sysErrCnt = 0
    }
    $seStatus = if ($sysErrCnt -eq 0) {"PASS"} elseif ($sysErrCnt -lt 5) {"WARN"} else {"FAIL"}
    $rows_events += New-CheckRow "System Errors (Since Boot)" $seStatus "$sysErrCnt errors since boot"

} catch {
    $rows_events += New-CheckRow "DCOM / Event Log Section" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "DCOM + Events" "Section" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 — SHELL + PROFILE INTEGRITY
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "7/10" "Shell + Profile Integrity"
$rows_shell = @()
$t = Start-SectionTimer "Shell Integrity"
try {
    try {
        $shellFolders = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -ErrorAction Stop
        $rows_shell += New-CheckRow "HKCU Shell Folders Key" "PASS" "Present"
    } catch {
        $rows_shell += New-CheckRow "HKCU Shell Folders Key" "FAIL" "MISSING — may cause Explorer errors at login" `
            "Profile corruption likely. Rebuild golden image profile."
    }

    try {
        $odRun = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -ErrorAction Stop).OneDrive
        $rows_shell += New-CheckRow "OneDrive Autorun (HKLM Run)" "INFO" `
            $(if ($odRun) {"Present: $odRun"} else {"Not in HKLM Run (may be in HKCU — normal)"})
    } catch {
        $rows_shell += New-CheckRow "OneDrive Autorun" "WARN" "Registry query failed."
        Register-ScriptError "Shell" "OneDrive Autorun" $_
    }

    try {
        $explorerProc = Get-Process -Name "explorer" -ErrorAction Stop
        $rows_shell += New-CheckRow "Explorer.exe Process" "PASS" "Running (PID: $($explorerProc.Id -join ', '))"
    } catch {
        $rows_shell += New-CheckRow "Explorer.exe Process" "FAIL" "NOT running or query failed" `
            "Shell crash or process not started. Check Event Log for Explorer crash entries."
    }

} catch {
    $rows_shell += New-CheckRow "Shell Section" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "Shell" "Section" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 8 — TPM + SECURITY
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "8/10" "TPM + Security"
$rows_security = @()
$t = Start-SectionTimer "Security"
try {
    try {
        $tpm = Get-CimInstance -Namespace "Root\CIMv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction Stop
        if ($tpm) {
            $tpmStatus = if ($tpm.IsEnabled_InitialValue -and $tpm.IsActivated_InitialValue) {"PASS"} else {"WARN"}
            $rows_security += New-CheckRow "TPM / vTPM State" $tpmStatus `
                "Enabled: $($tpm.IsEnabled_InitialValue)  |  Activated: $($tpm.IsActivated_InitialValue)  |  Spec: $($tpm.SpecVersion)"
        } else {
            $rows_security += New-CheckRow "TPM / vTPM State" "WARN" "Not detected" `
                "Horizon full clone pools provision vTPM per-clone at creation — template does not require vTPM."
        }
    } catch {
        $rows_security += New-CheckRow "TPM / vTPM State" "WARN" "WMI query failed: $($_.Exception.Message)" `
            "Horizon full clone pools provision vTPM per-clone at creation."
    }

    try {
        $secureBoot = Confirm-SecureBootUEFI -ErrorAction Stop
        $rows_security += New-CheckRow "Secure Boot" `
            $(if ($secureBoot) {"PASS"} else {"INFO"}) `
            $(if ($secureBoot -eq $true) {"Enabled"} elseif ($secureBoot -eq $false) {"Disabled"} else {"Not supported / BIOS mode"})
    } catch {
        $rows_security += New-CheckRow "Secure Boot" "INFO" "Query failed — may be BIOS/legacy boot mode."
    }

    try {
        $av = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction Stop
        $rows_security += New-CheckRow "Antivirus Product" $(if ($av) {"INFO"} else {"WARN"}) `
            $(if ($av) {($av | ForEach-Object {$_.displayName}) -join ", "} else {"No AV detected via SecurityCenter2"})
    } catch {
        $rows_security += New-CheckRow "Antivirus Product" "WARN" "SecurityCenter2 query failed."
        Register-ScriptError "Security" "AV Query" $_
    }

    $wdResult = Test-ServiceStatus -ServiceName "WinDefend"
    $rows_security += New-CheckRow "Windows Defender Service" $wdResult.Status $wdResult.Detail

    $wscResult = Test-ServiceStatus -ServiceName "wscsvc"
    $rows_security += New-CheckRow "Windows Security Center (wscsvc)" $wscResult.Status $wscResult.Detail `
        $(if ($wscResult.Status -eq "PASS") {"wscsvc generates DCOM 10016 errors — disable only if confirmed safe in this environment."} else {""})

    # VBS registry state check
    try {
        $dvgReg = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard" -ErrorAction Stop
        $vbsEnabled = $dvgReg.EnableVirtualizationBasedSecurity
        $vbsRegStatus = if ($vbsEnabled -eq 0) {"PASS"} else {"WARN"}
        $rows_security += New-CheckRow "VBS Registry (DeviceGuard)" $vbsRegStatus `
            "EnableVirtualizationBasedSecurity = $vbsEnabled  (0 = disabled)" `
            $(if ($vbsEnabled -ne 0) {"Disable VBS: Set EnableVirtualizationBasedSecurity = 0 in HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard to prevent Event 124 Kernel-Boot errors."} else {""})
    } catch {
        $rows_security += New-CheckRow "VBS Registry (DeviceGuard)" "INFO" "DeviceGuard key not present — VBS likely not configured."
    }

} catch {
    $rows_security += New-CheckRow "Security Section" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "Security" "Section" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 9 — NETWORK
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "9/10" "Network"
$rows_network = @()
$t = Start-SectionTimer "Network"
try {
    try {
        $nics = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop | Where-Object {$_.InterfaceAlias -notlike "*Loopback*"}
        foreach ($nic in $nics) {
            $rows_network += New-CheckRow "NIC: $($nic.InterfaceAlias)" "INFO" "$($nic.IPAddress) / $($nic.PrefixLength)"
        }
    } catch {
        $rows_network += New-CheckRow "NIC Information" "WARN" "Query failed: $($_.Exception.Message)"
        Register-ScriptError "Network" "NIC Query" $_
    }

    try {
        $dns = Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction Stop |
            Where-Object {$_.InterfaceAlias -notlike "*Loopback*" -and $_.ServerAddresses.Count -gt 0}
        foreach ($d in $dns) {
            $rows_network += New-CheckRow "DNS: $($d.InterfaceAlias)" "INFO" ($d.ServerAddresses -join ", ")
        }
    } catch {
        $rows_network += New-CheckRow "DNS Configuration" "WARN" "Query failed: $($_.Exception.Message)"
        Register-ScriptError "Network" "DNS Query" $_
    }

    if ($cs -and $cs.PartOfDomain) {
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
            try {
                $r = Test-NetConnection -ComputerName $eh -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction Stop
                $rows_network += New-CheckRow "Additional Host: $eh" $(if ($r) {"PASS"} else {"FAIL"}) `
                    $(if ($r) {"Reachable"} else {"UNREACHABLE"})
            } catch {
                $rows_network += New-CheckRow "Additional Host: $eh" "WARN" "Test failed: $($_.Exception.Message)"
                Register-ScriptError "Network" "Host: $eh" $_
            }
        }
    }

} catch {
    $rows_network += New-CheckRow "Network Section" "FAIL" "Section failed: $($_.Exception.Message)"
    Register-ScriptError "Network" "Section" $_
    if ($StopOnSectionFail) { throw }
}
Stop-SectionTimer $t

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 10 — SCRIPT EXECUTION ERRORS
# ─────────────────────────────────────────────────────────────────────────────
Write-Progress-Step "10/10" "Building Report"
$rows_scripterrors = @()
if ($ScriptErrors.Count -gt 0) {
    foreach ($e in $ScriptErrors) {
        $rows_scripterrors += New-CheckRow "$($e.Section) — $($e.Check)" "WARN" `
            "[$($e.Time)] $($e.Message)"
    }
} else {
    $rows_scripterrors += New-CheckRow "Script Execution" "PASS" "No errors encountered during script execution."
}

# ─────────────────────────────────────────────────────────────────────────────
# SCORING
# ─────────────────────────────────────────────────────────────────────────────
$allRows       = @($rows_system+$rows_horizon+$rows_appvol+$rows_nvidia+$rows_appx+$rows_events+$rows_shell+$rows_security+$rows_network)
$totalFail     = ([regex]::Matches($allRows -join "", "badge fail")).Count
$totalWarn     = ([regex]::Matches($allRows -join "", "badge warn")).Count
$totalPass     = ([regex]::Matches($allRows -join "", "badge pass")).Count
$overallStatus = if ($totalFail -gt 0) {"CRITICAL"} elseif ($totalWarn -gt 3) {"DEGRADED"} elseif ($totalWarn -gt 0) {"WARNING"} else {"HEALTHY"}
$overallColor  = switch ($overallStatus) {
    "CRITICAL" {"#ff3b3b"} "DEGRADED" {"#ff8c00"} "WARNING" {"#f5c518"} default {"#00e676"}
}
$runDuration   = [math]::Round(((Get-Date) - $RunStart).TotalSeconds, 1)
$timingRows    = ($SectionTimings | ForEach-Object { "<tr><td class='check-name'>$($_.Section)</td><td></td><td style='font-family:monospace;font-size:11px;color:#5a7a9a;'>$($_.Seconds)s</td></tr>" }) -join ""
$errorCount    = $ScriptErrors.Count

# ─────────────────────────────────────────────────────────────────────────────
# HTML OUTPUT
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
tr.row-fail{background:rgba(255,59,59,.06)}tr.row-warn{background:rgba(245,197,24,.04)}tr.row-pass{background:rgba(0,230,118,.02)}
tr:hover{background:rgba(0,180,255,.05)!important}
.badge{display:inline-block;font-family:var(--mono);font-size:10px;font-weight:bold;padding:3px 9px;border-radius:3px;letter-spacing:.08em;white-space:nowrap}
.badge.pass{background:rgba(0,230,118,.12);color:var(--pass);border:1px solid rgba(0,230,118,.3)}
.badge.fail{background:rgba(255,59,59,.12);color:var(--fail);border:1px solid rgba(255,59,59,.3)}
.badge.warn{background:rgba(245,197,24,.12);color:var(--warn);border:1px solid rgba(245,197,24,.3)}
.badge.info{background:rgba(0,180,255,.12);color:var(--info);border:1px solid rgba(0,180,255,.3)}
.badge.skip{background:rgba(90,122,154,.12);color:var(--skip);border:1px solid rgba(90,122,154,.3)}
.error-banner{background:rgba(255,59,59,.08);border:1px solid rgba(255,59,59,.25);border-radius:6px;padding:10px 16px;margin-bottom:20px;font-family:var(--mono);font-size:11px;color:var(--fail)}
.footer{border-top:1px solid var(--border);padding:16px 40px;display:flex;justify-content:space-between;align-items:center;font-family:var(--mono);font-size:11px;color:var(--text-dim);background:var(--surface)}
@media print{body{background:#fff;color:#000}.section,td,th{border-color:#ccc!important}}
</style>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div>
      <div style="font-family:monospace;font-size:11px;color:#5a7a9a;margin-bottom:4px;letter-spacing:.15em;">
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
    <div class="meta-item"><div class="meta-label">Run Duration</div><div class="meta-value">${runDuration}s</div></div>
    <div class="meta-item"><div class="meta-label">Script Errors</div><div class="meta-value" style="color:$(if($errorCount -gt 0){'#f5c518'}else{'#00e676'})">$errorCount</div></div>
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
  <div class="score-divider"></div>
  <div style="font-family:monospace;font-size:11px;color:$(if($errorCount -gt 0){'#f5c518'}else{'#5a7a9a'})">Script Errors: <span style="font-weight:bold">$errorCount</span></div>
</div>

<div class="content">

$(if ($errorCount -gt 0) {
"<div class='error-banner'>⚠ $errorCount script error(s) occurred during execution. Some checks may be incomplete. See Script Execution Errors section below.</div>"
})

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
    <div class="section-header"><span class="section-icon">📋</span><span class="section-title">DCOM + Event Log + VBS</span><span class="section-count">10016 · WscDataProtection · Event 124 VBS · App + System Errors</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_events -join "")</table>
  </div>

  <div class="section">
    <div class="section-header"><span class="section-icon">🪟</span><span class="section-title">Shell + Profile Integrity</span><span class="section-count">Shell Folders · OneDrive · Explorer</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_shell -join "")</table>
  </div>

  <div class="section">
    <div class="section-header"><span class="section-icon">🔒</span><span class="section-title">TPM + Security</span><span class="section-count">vTPM · Secure Boot · Defender · AV · wscsvc · VBS Registry</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_security -join "")</table>
  </div>

  <div class="section">
    <div class="section-header"><span class="section-icon">🌐</span><span class="section-title">Network</span><span class="section-count">NIC · DNS · Domain Controller · Horizon CS · App Vol Manager</span></div>
    <table><tr><th>Check</th><th>Status</th><th>Detail</th></tr>$($rows_network -join "")</table>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="section-icon">$(if ($errorCount -gt 0) {'⚠'} else {'✔'})</span>
      <span class="section-title">Script Execution Errors</span>
      <span class="section-count">Self-audit — errors logged but execution continued</span>
    </div>
    <table><tr><th>Section — Check</th><th>Status</th><th>Error Message</th></tr>$($rows_scripterrors -join "")</table>
  </div>

  <div class="section">
    <div class="section-header"><span class="section-icon">⏱</span><span class="section-title">Section Timings</span><span class="section-count">Total: ${runDuration}s</span></div>
    <table><tr><th>Section</th><th></th><th>Duration</th></tr>$timingRows</table>
  </div>

</div>

<div class="footer">
  <div>$ReportFile</div>
  <div>VDI Health Report v1.2 — Generic Edition | Errors: $errorCount | Duration: ${runDuration}s</div>
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
Write-Host "  ║  COMPLETE — $overallStatus" -ForegroundColor $statusColor
Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor DarkGray
Write-Host "  ║  PASS          : $totalPass" -ForegroundColor Green
Write-Host "  ║  WARN          : $totalWarn" -ForegroundColor Yellow
Write-Host "  ║  FAIL          : $totalFail" -ForegroundColor Red
Write-Host "  ║  SCRIPT ERRORS : $errorCount" -ForegroundColor $(if ($errorCount -gt 0) {"Yellow"} else {"Green"})
Write-Host "  ║  DURATION      : ${runDuration}s" -ForegroundColor Gray
Write-Host "  ╠══════════════════════════════════════════════════════════╣" -ForegroundColor DarkGray
Write-Host "  ║  $ReportFile" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor DarkGray
Write-Host ""

if ($errorCount -gt 0) {
    Write-Host "  Script errors encountered:" -ForegroundColor Yellow
    foreach ($e in $ScriptErrors) {
        Write-Host "    [$($e.Time)] $($e.Section) — $($e.Check): $($e.Message)" -ForegroundColor Yellow
    }
    Write-Host ""
}

Start-Process $ReportFile
