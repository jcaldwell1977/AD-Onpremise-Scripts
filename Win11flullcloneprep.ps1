<#
    Horizon Full Clone Template Prep — Windows 11
    Run as Administrator on the template VM
    Take a snapshot BEFORE running this
#>

$ErrorActionPreference = 'Continue'
$log = "C:\Windows\Temp\TemplatePrep_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $log

Write-Host "=== Horizon Full Clone Template Prep ===" -ForegroundColor Cyan
Write-Host "Log: $log`n"

# ---------------------------------------------------------------
# 1. Remove problem AppX packages (sysprep killers on Win11)
# ---------------------------------------------------------------
Write-Host "[1/8] Removing AppX packages known to break sysprep..." -ForegroundColor Yellow

$pkgs = @(
    'Microsoft.Wallet','Microsoft.Teams','MicrosoftTeams','Clipchamp.Clipchamp',
    'Microsoft.BingNews','Microsoft.BingWeather','Microsoft.GamingApp',
    'Microsoft.XboxGamingOverlay','Microsoft.XboxIdentityProvider',
    'Microsoft.XboxSpeechToTextOverlay','Microsoft.Xbox.TCUI',
    'Microsoft.MicrosoftOfficeHub','Microsoft.YourPhone',
    'Microsoft.WindowsCommunicationsApps','Microsoft.MixedReality.Portal',
    'Microsoft.GetHelp','Microsoft.Getstarted','Microsoft.People',
    'Microsoft.WindowsFeedbackHub','Microsoft.Microsoft3DViewer',
    'Microsoft.MicrosoftSolitaireCollection','Microsoft.SkypeApp',
    'Microsoft.ZuneMusic','Microsoft.ZuneVideo','Microsoft.WindowsMaps',
    'Microsoft.MicrosoftStickyNotes','Microsoft.PowerAutomateDesktop',
    'MicrosoftCorporationII.QuickAssist','Microsoft.OutlookForWindows'
)

foreach ($p in $pkgs) {
    Get-AppxPackage -AllUsers $p -ErrorAction SilentlyContinue |
        Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    Get-AppxProvisionedPackage -Online |
        Where-Object DisplayName -eq $p |
        Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    Write-Host "  Removed: $p"
}

# ---------------------------------------------------------------
# 2. Disable Store auto-updates (prevents AppX from re-breaking)
# ---------------------------------------------------------------
Write-Host "`n[2/8] Disabling Store auto-updates..." -ForegroundColor Yellow

$storePath = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore'
if (-not (Test-Path $storePath)) { New-Item -Path $storePath -Force | Out-Null }
Set-ItemProperty -Path $storePath -Name 'AutoDownload' -Value 2 -Type DWord
Set-ItemProperty -Path $storePath -Name 'DisableOSUpgrade' -Value 1 -Type DWord

# Disable Content Delivery Manager auto-installs (fresh user accounts won't get random apps)
$cdmPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent'
if (-not (Test-Path $cdmPath)) { New-Item -Path $cdmPath -Force | Out-Null }
Set-ItemProperty -Path $cdmPath -Name 'DisableWindowsConsumerFeatures' -Value 1 -Type DWord

# ---------------------------------------------------------------
# 3. Disable Windows Update for sysprep + reset state
# ---------------------------------------------------------------
Write-Host "`n[3/8] Stopping Windows Update services..." -ForegroundColor Yellow

Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
Set-Service -Name wuauserv -StartupType Disabled

# Clear pending update state that can block sysprep
Remove-Item 'C:\Windows\SoftwareDistribution\Download\*' -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------
# 4. Disk cleanup
# ---------------------------------------------------------------
Write-Host "`n[4/8] Running disk cleanup..." -ForegroundColor Yellow

# Clear temp
Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item 'C:\Windows\Temp\*' -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item 'C:\Windows\Prefetch\*' -Recurse -Force -ErrorAction SilentlyContinue

# Clear event logs
Get-WinEvent -ListLog * -ErrorAction SilentlyContinue |
    Where-Object { $_.RecordCount -gt 0 } |
    ForEach-Object { wevtutil cl $_.LogName 2>$null }

# Component cleanup (irreversible — locks in current state, shrinks WinSxS)
Dism /Online /Cleanup-Image /StartComponentCleanup /ResetBase

# ---------------------------------------------------------------
# 5. VDI optimizations
# ---------------------------------------------------------------
Write-Host "`n[5/8] Applying VDI optimizations..." -ForegroundColor Yellow

# Disable hibernation
powercfg -h off

# High performance power plan
powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

# Disable System Restore
Disable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue

# Disable scheduled defrag (SSD/VDI doesn't want it)
Disable-ScheduledTask -TaskName "\Microsoft\Windows\Defrag\ScheduledDefrag" -ErrorAction SilentlyContinue

# Disable Superfetch/SysMain (bad for VDI)
Stop-Service -Name SysMain -Force -ErrorAction SilentlyContinue
Set-Service -Name SysMain -StartupType Disabled

# Disable Windows Search indexing (optional — depends on user needs)
# Stop-Service -Name WSearch -Force -ErrorAction SilentlyContinue
# Set-Service -Name WSearch -StartupType Disabled

# Disable telemetry
$telPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'
if (-not (Test-Path $telPath)) { New-Item -Path $telPath -Force | Out-Null }
Set-ItemProperty -Path $telPath -Name 'AllowTelemetry' -Value 0 -Type DWord

# ---------------------------------------------------------------
# 6. Disable problematic scheduled tasks
# ---------------------------------------------------------------
Write-Host "`n[6/8] Disabling problematic scheduled tasks..." -ForegroundColor Yellow

$tasksToDisable = @(
    '\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser',
    '\Microsoft\Windows\Application Experience\ProgramDataUpdater',
    '\Microsoft\Windows\Autochk\Proxy',
    '\Microsoft\Windows\Customer Experience Improvement Program\Consolidator',
    '\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip',
    '\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector',
    '\Microsoft\Windows\Maps\MapsUpdateTask',
    '\Microsoft\Windows\Feedback\Siuf\DmClient',
    '\Microsoft\Windows\Windows Error Reporting\QueueReporting'
)
foreach ($t in $tasksToDisable) {
    Disable-ScheduledTask -TaskPath (Split-Path $t) -TaskName (Split-Path $t -Leaf) -ErrorAction SilentlyContinue | Out-Null
}

# ---------------------------------------------------------------
# 7. Network/DNS reset for clean sysprep
# ---------------------------------------------------------------
Write-Host "`n[7/8] Clearing network state..." -ForegroundColor Yellow

ipconfig /flushdns | Out-Null
# Don't release IP yet — you'll lose RDP. Sysprep handles this.

# ---------------------------------------------------------------
# 8. Pre-sysprep verification
# ---------------------------------------------------------------
Write-Host "`n[8/8] Pre-sysprep checks..." -ForegroundColor Yellow

# Check for any remaining per-user AppX packages that aren't provisioned
$orphans = Get-AppxPackage -AllUsers |
    Where-Object { $_.NonRemovable -eq $false -and $_.PackageUserInformation.InstallState -contains 'Installed' } |
    Where-Object {
        $name = $_.Name
        -not (Get-AppxProvisionedPackage -Online | Where-Object DisplayName -eq $name)
    }

if ($orphans) {
    Write-Host "`nWARNING: Found per-user AppX packages NOT provisioned for all users:" -ForegroundColor Red
    $orphans | Select-Object Name, PackageFullName | Format-Table -AutoSize
    Write-Host "These WILL cause sysprep to fail. Review before proceeding." -ForegroundColor Red
} else {
    Write-Host "  No orphan AppX packages detected." -ForegroundColor Green
}

# Check Horizon Agent
$agent = Get-ItemProperty 'HKLM:\SOFTWARE\Omnissa, Inc.\Omnissa Horizon Agent\Installer' -ErrorAction SilentlyContinue
if (-not $agent) {
    $agent = Get-ItemProperty 'HKLM:\SOFTWARE\VMware, Inc.\VMware VDM\Agent' -ErrorAction SilentlyContinue
}
if ($agent) {
    Write-Host "  Horizon Agent detected." -ForegroundColor Green
} else {
    Write-Host "  WARNING: Horizon Agent NOT detected. Install before sysprep." -ForegroundColor Red
}

Write-Host "`n=== Prep complete ===" -ForegroundColor Cyan
Write-Host "Next steps:"
Write-Host "  1. Review any warnings above"
Write-Host "  2. Shut down cleanly"
Write-Host "  3. Take snapshot 'pre-sysprep'"
Write-Host "  4. Power on, run: C:\Windows\System32\Sysprep\sysprep.exe /generalize /shutdown /oobe"
Write-Host "  5. After VM powers off, take final snapshot for the pool`n"

Stop-Transcript
