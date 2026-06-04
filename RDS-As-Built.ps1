<#
.SYNOPSIS
    Generates a detailed As-Built documentation for a Remote Desktop Services (RDS) environment.

.OUTPUT
    Creates a nicely formatted HTML report + CSV exports in the specified folder.

.EXAMPLE
    .\Get-RDSAsBuilt.ps1 -OutputPath "C:\RDS_AsBuilt"
#>

param (
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "C:\RDS_AsBuilt"
)

# Create output folder
$ReportDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
$ReportFolder = Join-Path $OutputPath "RDS_AsBuilt_$ReportDate"
New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null

$HtmlPath = Join-Path $ReportFolder "RDS_AsBuilt_Report.html"
$LogPath   = Join-Path $ReportFolder "RDS_AsBuilt_Log.txt"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Tee-Object -FilePath $LogPath -Append
}

Write-Log "Starting RDS As-Built documentation..."

# Import RDS Module
Import-Module RemoteDesktop -ErrorAction SilentlyContinue
if (-not (Get-Module -Name RemoteDesktop)) {
    Write-Log "ERROR: RemoteDesktop module not found. Please run this script on an RDS Connection Broker server."
    exit 1
}

# HTML Header
$Html = @"
<!DOCTYPE html>
<html>
<head>
    <title>RDS Environment As-Built - $env:COMPUTERNAME</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #003366; }
        table { border-collapse: collapse; width: 100%; margin: 15px 0; }
        th, td { border: 1px solid #999; padding: 8px; text-align: left; }
        th { background-color: #003366; color: white; }
        .section { margin-top: 30px; }
    </style>
</head>
<body>
    <h1>Remote Desktop Services - As-Built Documentation</h1>
    <p><strong>Generated:</strong> $(Get-Date)</p>
    <p><strong>Server:</strong> $env:COMPUTERNAME</p>
"@

# 1. RDS Deployment Overview
Write-Log "Gathering RDS Deployment Information..."
$Deployment = Get-RDDeployment -ErrorAction SilentlyContinue

$Html += @"
    <div class="section">
        <h2>1. RDS Deployment Overview</h2>
        <table>
            <tr><th>Property</th><th>Value</th></tr>
"@
if ($Deployment) {
    $Deployment.PSObject.Properties | ForEach-Object {
        $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
    }
} else {
    $Html += "<tr><td colspan='2'>No deployment information found.</td></tr>"
}
$Html += "</table></div>"

# 2. Connection Brokers
Write-Log "Gathering Connection Brokers..."
$CBs = Get-RDConnectionBrokerHighAvailability -ErrorAction SilentlyContinue
$Html += @"
    <div class="section">
        <h2>2. Connection Broker Servers</h2>
        <table>
            <tr><th>Server</th><th>Role</th><th>Status</th></tr>
"@
foreach ($cb in (Get-RDServer | Where-Object {$_.Roles -like "*Connection Broker*"})) {
    $Html += "<tr><td>$($cb.Server)</td><td>Connection Broker</td><td>Active</td></tr>"
}
$Html += "</table></div>"

# 3. Session Hosts
Write-Log "Gathering Session Hosts..."
$SessionHosts = Get-RDSessionHost
$Html += @"
    <div class="section">
        <h2>3. Session Host Servers</h2>
        <table>
            <tr><th>Server</th><th>Status</th><th>Number of Sessions</th><th>Drain Mode</th></tr>
"@
foreach ($sh in $SessionHosts) {
    $sessions = (Get-RDUserSession -CollectionName "*" -ConnectionBroker $sh.SessionHost | Measure-Object).Count
    $Html += "<tr><td>$($sh.SessionHost)</td><td>$($sh.Status)</td><td>$sessions</td><td>$($sh.DrainMode)</td></tr>"
}
$Html += "</table></div>"

# 4. Collections
Write-Log "Gathering Session Collections..."
$Collections = Get-RDSessionCollection

$Html += @"
    <div class="section">
        <h2>4. RDS Collections</h2>
"@
foreach ($col in $Collections) {
    $Html += "<h3>Collection: $($col.CollectionName)</h3>"
    $Html += "<table><tr><th>Property</th><th>Value</th></tr>"
    $col.PSObject.Properties | ForEach-Object {
        $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
    }
    $Html += "</table>"

    # Users in collection
    $Users = Get-RDUserSession -CollectionName $col.CollectionName -ErrorAction SilentlyContinue
    if ($Users) {
        $Html += "<h4>Active Sessions</h4><table><tr><th>User</th><th>Server</th><th>Session State</th><th>Client IP</th></tr>"
        foreach ($u in $Users) {
            $Html += "<tr><td>$($u.UserName)</td><td>$($u.HostServer)</td><td>$($u.SessionState)</td><td>$($u.ClientIPAddress)</td></tr>"
        }
        $Html += "</table>"
    }
}
$Html += "</div>"

# 5. RD Gateway
Write-Log "Gathering RD Gateway configuration..."
$Gateway = Get-RDGatewayServer | Select-Object * -ErrorAction SilentlyContinue

$Html += @"
    <div class="section">
        <h2>5. RD Gateway Servers</h2>
        <table>
            <tr><th>Server</th><th>Status</th><th>Logon Method</th></tr>
"@
if ($Gateway) {
    foreach ($gw in $Gateway) {
        $Html += "<tr><td>$($gw.Server)</td><td>$($gw.Status)</td><td>$($gw.LogonMethod)</td></tr>"
    }
} else {
    $Html += "<tr><td colspan='3'>No RD Gateway configured or accessible.</td></tr>"
}
$Html += "</table></div>"

# 6. Licensing
Write-Log "Gathering Licensing Information..."
$Licensing = Get-RDLicenseConfiguration -ErrorAction SilentlyContinue

$Html += @"
    <div class="section">
        <h2>6. RDS Licensing</h2>
        <table>
            <tr><th>Property</th><th>Value</th></tr>
"@
if ($Licensing) {
    $Licensing.PSObject.Properties | ForEach-Object {
        $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
    }
}
$Html += "</table></div>"

# 7. Certificates
Write-Log "Gathering Certificates..."
$Html += @"
    <div class="section">
        <h2>7. RDS Certificates</h2>
        <table>
            <tr><th>Role</th><th>Thumbprint</th><th>Subject</th><th>Expiration</th></tr>
"@
Get-RDCertificate | ForEach-Object {
    $Html += "<tr><td>$($_.Role)</td><td>$($_.Thumbprint)</td><td>$($_.Subject)</td><td>$($_.ExpiresOn)</td></tr>"
}
$Html += "</table></div>"

# Close HTML
$Html += "</body></html>"

# Save HTML Report
$Html | Out-File -FilePath $HtmlPath -Encoding UTF8

# Export CSVs
Get-RDSessionHost | Export-Csv "$ReportFolder\SessionHosts.csv" -NoTypeInformation
Get-RDSessionCollection | Export-Csv "$ReportFolder\Collections.csv" -NoTypeInformation
Get-RDUserSession -ErrorAction SilentlyContinue | Export-Csv "$ReportFolder\CurrentSessions.csv" -NoTypeInformation

Write-Log "RDS As-Built documentation completed successfully!"
Write-Log "Report saved to: $HtmlPath"

# Open the report
Invoke-Item $HtmlPath

Write-Host "`nRDS As-Built report has been generated and opened." -ForegroundColor Green
