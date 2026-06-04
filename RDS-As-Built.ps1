<#
.SYNOPSIS
     Detailed RDS As-Built Documentation including Collections, Assignments, and FSLogix
#>

param (
    [string]$OutputPath = "C:\RDS_AsBuilt"
)

$ReportDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
$ReportFolder = Join-Path $OutputPath "RDS_AsBuilt_Detailed_$ReportDate"
New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null

$HtmlPath = Join-Path $ReportFolder "RDS_AsBuilt_Full_Report.html"
$LogPath  = Join-Path $ReportFolder "RDS_AsBuilt_Log.txt"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Tee-Object -FilePath $LogPath -Append
}

Write-Log "Starting Extremely Detailed RDS As-Built..."

Import-Module RemoteDesktop -ErrorAction SilentlyContinue
if (-not (Get-Module RemoteDesktop)) {
    Write-Log "ERROR: RemoteDesktop module not available. Run on Connection Broker."
    exit 1
}

# ========================== HTML Start ==========================
$Html = @"
<!DOCTYPE html>
<html>
<head>
    <title>RDS Environment - Detailed As-Built</title>
    <style>
        body { font-family: Segoe UI, Arial; margin: 30px; background: #f9f9f9; }
        h1, h2, h3 { color: #003366; }
        table { border-collapse: collapse; width: 100%; margin: 15px 0; }
        th, td { border: 1px solid #555; padding: 8px; }
        th { background: #003366; color: white; }
        .section { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 25px; }
    </style>
</head>
<body>
    <h1>Remote Desktop Services - Detailed As-Built Documentation</h1>
    <p><strong>Generated:</strong> $(Get-Date) | <strong>Server:</strong> $env:COMPUTERNAME</p>
"@

# 1. Deployment Overview
$Deployment = Get-RDDeployment
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
}
$Html += "</table></div>"

# 2. RDS Servers & Roles
$Html += @"
<div class="section">
    <h2>2. RDS Servers and Roles</h2>
    <table>
        <tr><th>Server</th><th>Roles</th></tr>
"@
Get-RDServer | ForEach-Object {
    $Html += "<tr><td>$($_.Server)</td><td>$($_.Roles -join ', ')</td></tr>"
}
$Html += "</table></div>"

# 3. Connection Brokers
$Html += @"
<div class="section">
    <h2>3. Connection Broker High Availability</h2>
"@
Get-RDConnectionBrokerHighAvailability | ForEach-Object {
    $_.PSObject.Properties | ForEach-Object {
        $Html += "<p><strong>$($_.Name):</strong> $($_.Value)</p>"
    }
}
$Html += "</div>"

# 4. Session Collections - Extremely Detailed
Write-Log "Gathering detailed Collection information..."
$Collections = Get-RDSessionCollection

$Html += @"
<div class="section">
    <h2>4. RDS Session Collections</h2>
"@

foreach ($col in $Collections) {
    $ColName = $col.CollectionName
    $Config = Get-RDSessionCollectionConfiguration -CollectionName $ColName
    $UserGroups = Get-RDSessionCollectionConfiguration -CollectionName $ColName -UserGroup | Select-Object -ExpandProperty UserGroup

    $Html += "<h3>Collection: $ColName</h3>"
    $Html += "<table><tr><th>Property</th><th>Value</th></tr>"

    $Config.PSObject.Properties | Where-Object {$_.Name -notlike "*UserGroup*"} | ForEach-Object {
        $value = if ($_.Value -is [array]) { $_.Value -join "; " } else { $_.Value }
        $Html += "<tr><td>$($_.Name)</td><td>$value</td></tr>"
    }
    $Html += "</table>"

    # User/Group Assignments
    $Html += "<h4>User Group Assignments</h4><table><tr><th>Assigned Groups</th></tr>"
    if ($UserGroups) {
        $UserGroups | ForEach-Object { $Html += "<tr><td>$_</td></tr>" }
    } else {
        $Html += "<tr><td>No groups assigned</td></tr>"
    }
    $Html += "</table>"

    # Published RemoteApps
    $Apps = Get-RDRemoteApp -CollectionName $ColName -ErrorAction SilentlyContinue
    if ($Apps) {
        $Html += "<h4>Published RemoteApps</h4><table><tr><th>Display Name</th><th>Alias</th><th>Path</th><th>Show in Web</th></tr>"
        $Apps | ForEach-Object {
            $Html += "<tr><td>$($_.DisplayName)</td><td>$($_.Alias)</td><td>$($_.FilePath)</td><td>$($_.ShowInWebAccess)</td></tr>"
        }
        $Html += "</table>"
    }
}

$Html += "</div>"

# 5. FSLogix Detection (Very Detailed)
Write-Log "Scanning for FSLogix configuration..."
$SessionHosts = (Get-RDSessionHost).SessionHost

$Html += @"
<div class="section">
    <h2>5. FSLogix Configuration</h2>
"@

foreach ($sh in $SessionHosts) {
    $Html += "<h3>Session Host: $sh</h3>"
    
    # Profiles
    $ProfilesReg = Invoke-Command -ComputerName $sh -ScriptBlock {
        Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -ErrorAction SilentlyContinue
    } -ErrorAction SilentlyContinue

    if ($ProfilesReg) {
        $Html += "<h4>FSLogix Profiles</h4><table><tr><th>Setting</th><th>Value</th></tr>"
        $ProfilesReg.PSObject.Properties | Where-Object {$_.Name -notlike "PS*"} | ForEach-Object {
            $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
        }
        $Html += "</table>"
    } else {
        $Html += "<p><strong>No FSLogix Profiles configuration found.</strong></p>"
    }

    # Office Containers
    $OfficeReg = Invoke-Command -ComputerName $sh -ScriptBlock {
        Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\OfficeContainers" -ErrorAction SilentlyContinue
    } -ErrorAction SilentlyContinue

    if ($OfficeReg) {
        $Html += "<h4>FSLogix Office Containers</h4><table><tr><th>Setting</th><th>Value</th></tr>"
        $OfficeReg.PSObject.Properties | Where-Object {$_.Name -notlike "PS*"} | ForEach-Object {
            $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
        }
        $Html += "</table>"
    }
}

$Html += "</div>"

# Finalize HTML
$Html += "</body></html>"

# Save Report
$Html | Out-File -FilePath $HtmlPath -Encoding UTF8

# Export CSVs for raw data
Get-RDSessionCollection | Export-Csv "$ReportFolder\Collections.csv" -NoTypeInformation
Get-RDSessionCollectionConfiguration | Export-Csv "$ReportFolder\Collection_Configs.csv" -NoTypeInformation
Get-RDRemoteApp | Export-Csv "$ReportFolder\RemoteApps.csv" -NoTypeInformation

Write-Log "Detailed RDS As-Built completed!"
Write-Host "Full detailed report generated at: $HtmlPath" -ForegroundColor Green

Invoke-Item $HtmlPath
