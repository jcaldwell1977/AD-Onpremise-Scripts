# =============================================================================
# VEEAM BACKUP & REPLICATION v13 – FULL AS-BUILT REPORT GENERATOR
# Covers EVERY configuration in detail – Infrastructure, Jobs, Global Settings
# Style: Professional As-Built (Isuzu-standard layout)
# =============================================================================

param(
    [string]$OutputPath = "C:\AsBuiltReports",
    [string]$ReportName = "Veeam-v13-AsBuilt-Report",
    [switch]$IncludeDiagrams,
    [switch]$HealthCheck
)

# Ensure PowerShell 7 & Veeam module
if ($PSVersionTable.PSVersion.Major -lt 7) { throw "Run in PowerShell 7 (pwsh.exe) for Veeam v13" }
Import-Module Veeam.Backup.PowerShell -ErrorAction SilentlyContinue

$ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$Html = @"
<html><head><title>$ReportName</title>
<style>
    body {font-family:Arial; margin:40px; background:#f8f8f8;}
    h1,h2,h3 {color:#003087;}
    table {width:100%; border-collapse:collapse; margin:15px 0;}
    th,td {border:1px solid #ccc; padding:8px; text-align:left;}
    th {background:#003087; color:white;}
    .section {margin:30px 0; padding:20px; background:white; border-radius:8px; box-shadow:0 2px 10px rgba(0,0,0,0.1);}
</style></head><body>
<h1 style="text-align:center;">VEEAM BACKUP & REPLICATION v13 – AS-BUILT REPORT</h1>
<h2 style="text-align:center;">$ReportName<br>Generated: $ReportDate</h2>
<hr>

<div class='section'>
<h2>1. Backup Server & Global Settings</h2>
"@
$Html += (Get-VBRServer | Where-Object {$_.Type -eq "Local"} | Select-Object Name, Version, DNSName, IsConnected | ConvertTo-Html -Fragment)

# Global Settings (all available)
$GlobalSettings = @{
    "Security & Encryption" = Get-VBRSecurityOptions
    "Notifications"         = Get-VBRNotificationOptions
    "Email"                 = Get-VBREmailOptions
    "SNMP"                  = Get-VBRSNMPService
    "History"               = Get-VBRJobHistoryOptions
}
$Html += "<h3>Global Options Summary</h3><table><tr><th>Category</th><th>Key Settings</th></tr>"
foreach ($cat in $GlobalSettings.Keys) {
    $Html += "<tr><td>$cat</td><td>$(($GlobalSettings[$cat] | Out-String).Trim())</td></tr>"
}
$Html += "</table></div>"

# 2. Licenses
$Html += "<div class='section'><h2>2. Licenses</h2>"
$Html += (Get-VBRLicense | Select-Object Edition, Status, ExpirationDate, UsedLicenses, TotalLicenses | ConvertTo-Html -Fragment)
$Html += "</div>"

# 3. Backup Infrastructure
$Html += "<div class='section'><h2>3. Backup Infrastructure</h2>"

# Proxies
$Html += "<h3>Backup Proxies</h3>" + (Get-VBRViProxy + Get-VBRHvProxy | Select-Object Name, Type, HostName, MaxTasks, Status | ConvertTo-Html -Fragment)

# Repositories
$Html += "<h3>Backup Repositories</h3>" + (Get-VBRBackupRepository | Select-Object Name, Type, Path, CapacityGB, FreeSpaceGB, Status | ConvertTo-Html -Fragment)

# Scale-Out Backup Repositories
$Html += "<h3>Scale-Out Backup Repositories (SOBR)</h3>" + (Get-VBRSOBR | Select-Object Name, PerformanceTier, CapacityTier, ArchiveTier, Status | ConvertTo-Html -Fragment)

$Html += "</div>"

# 4. All Jobs – FULL DETAIL
$Html += "<div class='section'><h2>4. Backup & Replication Jobs (Full Detail)</h2>"
$AllJobs = Get-VBRJob
foreach ($job in $AllJobs) {
    $options = $job.GetOptions()
    $Html += "<h3>Job: $($job.Name) [Type: $($job.JobType)]</h3>"
    $Html += "<table><tr><th>Property</th><th>Value</th></tr>"
    $Html += "<tr><td>Name</td><td>$($job.Name)</td></tr>"
    $Html += "<tr><td>Description</td><td>$($job.Description)</td></tr>"
    $Html += "<tr><td>Schedule</td><td>$($job.ScheduleOptions)</td></tr>"
    $Html += "<tr><td>Source</td><td>$($job.Source)</td></tr>"
    $Html += "<tr><td>Target</td><td>$($job.Target)</td></tr>"
    $Html += "<tr><td>Retention</td><td>$($options.RetentionPolicy)</td></tr>"
    $Html += "<tr><td>Advanced</td><td>$($options | Out-String)</td></tr>"
    $Html += "</table>"
}
$Html += "</div>"

# 5. Cloud Connect, Tape, SureBackup, Inventory (full coverage)
$Html += "<div class='section'><h2>5. Cloud Connect Infrastructure</h2>" + (Get-VBRCloudTenant | Select-Object Name, Status, Resources | ConvertTo-Html -Fragment) + "</div>"
$Html += "<div class='section'><h2>6. Tape Infrastructure</h2>" + (Get-VBRTapeLibrary + Get-VBRTapeMediaPool | ConvertTo-Html -Fragment) + "</div>"
$Html += "<div class='section'><h2>7. SureBackup & Virtual Labs</h2>" + (Get-VBRSureBackupJob | ConvertTo-Html -Fragment) + "</div>"
$Html += "<div class='section'><h2>8. Inventory (VI, File Shares, Physical)</h2>" + (Get-VBRViServer + Get-VBRFileShare + Get-VBRPhysicalInfrastructure | ConvertTo-Html -Fragment) + "</div>"

# Footer
$Html += "<hr><p style='text-align:center;'>Veeam v13 As-Built Report – Complete Configuration Export<br>End of Document</p></body></html>"

# Save report
$FilePath = "$OutputPath\$ReportName-$(Get-Date -Format 'yyyyMMdd-HHmm').html"
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
$Html | Out-File -FilePath $FilePath -Encoding UTF8
Write-Host "✅ As-Built Report generated: $FilePath" -ForegroundColor Green
Write-Host "Open in browser or import to Word for PDF." -ForegroundColor Cyan
