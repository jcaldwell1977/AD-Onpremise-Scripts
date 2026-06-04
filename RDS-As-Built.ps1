<#
.SYNOPSIS
    Extremely Detailed RDS As-Built with FSLogix Profile Disks + Full RD Gateway Config
#>

param (
    [string]$OutputPath = "C:\RDS_AsBuilt"
)

$ReportDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
$ReportFolder = Join-Path $OutputPath "RDS_AsBuilt_UltraDetailed_$ReportDate"
New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null

$HtmlPath = Join-Path $ReportFolder "RDS_AsBuilt_Full_Report.html"
$LogPath  = Join-Path $ReportFolder "RDS_AsBuilt_Log.txt"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Tee-Object -FilePath $LogPath -Append
}

Write-Log "Starting Ultra-Detailed RDS As-Built..."

Import-Module RemoteDesktop -ErrorAction SilentlyContinue
if (-not (Get-Module RemoteDesktop)) {
    Write-Log "ERROR: RemoteDesktop module not available."
    exit 1
}

# ========================== HTML Start ==========================
$Html = @"
<!DOCTYPE html>
<html>
<head>
    <title>RDS Ultra-Detailed As-Built</title>
    <style>
        body { font-family: Segoe UI, Arial; margin: 30px; background: #f9f9f9; }
        h1, h2, h3 { color: #003366; }
        table { border-collapse: collapse; width: 100%; margin: 15px 0; }
        th, td { border: 1px solid #555; padding: 9px; text-align: left; }
        th { background: #003366; color: white; }
        .section { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin-bottom: 30px; }
    </style>
</head>
<body>
    <h1>Remote Desktop Services - Ultra Detailed As-Built</h1>
    <p><strong>Generated:</strong> $(Get-Date) | <strong>Server:</strong> $env:COMPUTERNAME</p>
"@

# 1. Deployment Overview
$Deployment = Get-RDDeployment
$Html += @"
<div class="section">
    <h2>1. RDS Deployment Overview</h2>
    <table><tr><th>Property</th><th>Value</th></tr>
"@
$Deployment.PSObject.Properties | ForEach-Object {
    $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
}
$Html += "</table></div>"

# 2. RDS Servers & Roles
$Html += @"
<div class="section">
    <h2>2. RDS Servers and Roles</h2>
    <table><tr><th>Server</th><th>Roles</th></tr>
"@
Get-RDServer | ForEach-Object {
    $Html += "<tr><td>$($_.Server)</td><td>$($_.Roles -join ', ')</td></tr>"
}
$Html += "</table></div>"

# 3. RD Gateway - Detailed Configuration
Write-Log "Gathering detailed RD Gateway configuration..."
$Html += @"
<div class="section">
    <h2>3. RD Gateway Configuration</h2>
"@

$Gateways = Get-RDGatewayServer
if ($Gateways) {
    foreach ($gw in $Gateways) {
        $Html += "<h3>Gateway Server: $($gw.Server)</h3>"
        
        # Basic Gateway Info
        $Html += "<table><tr><th>Property</th><th>Value</th></tr>"
        $gw.PSObject.Properties | ForEach-Object {
            $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
        }
        $Html += "</table>"

        # Advanced Gateway Settings
        try {
            $GwConfig = Get-RDGatewayConfiguration -ErrorAction Stop
            $Html += "<h4>Gateway Advanced Settings</h4><table><tr><th>Property</th><th>Value</th></tr>"
            $GwConfig.PSObject.Properties | ForEach-Object {
                $value = if ($null -eq $_.Value) { "N/A" } else { $_.Value }
                $Html += "<tr><td>$($_.Name)</td><td>$value</td></tr>"
            }
            $Html += "</table>"
        } catch {}
    }
} else {
    $Html += "<p>No RD Gateway servers configured.</p>"
}
$Html += "</div>"

# 4. Session Collections (Detailed)
$Collections = Get-RDSessionCollection
$Html += @"
<div class="section">
    <h2>4. RDS Session Collections & Assignments</h2>
"@

foreach ($col in $Collections) {
    $ColName = $col.CollectionName
    $Config = Get-RDSessionCollectionConfiguration -CollectionName $ColName
    $UserGroups = (Get-RDSessionCollectionConfiguration -CollectionName $ColName -UserGroup).UserGroup

    $Html += "<h3>Collection: $ColName</h3>"
    $Html += "<table><tr><th>Property</th><th>Value</th></tr>"
    $Config.PSObject.Properties | Where-Object {$_.Name -notlike "*UserGroup*"} | ForEach-Object {
        $val = if ($_.Value -is [array]) { $_.Value -join "; " } else { $_.Value }
        $Html += "<tr><td>$($_.Name)</td><td>$val</td></tr>"
    }
    $Html += "</table>"

    # User/Group Assignments
    $Html += "<h4>Assigned User Groups</h4><table><tr><th>Group</th></tr>"
    if ($UserGroups) {
        $UserGroups | ForEach-Object { $Html += "<tr><td>$_</td></tr>" }
    } else {
        $Html += "<tr><td>None</td></tr>"
    }
    $Html += "</table>"

    # Published RemoteApps
    $Apps = Get-RDRemoteApp -CollectionName $ColName -ErrorAction SilentlyContinue
    if ($Apps.Count -gt 0) {
        $Html += "<h4>Published RemoteApps</h4><table><tr><th>Display Name</th><th>Alias</th><th>Path</th></tr>"
        $Apps | ForEach-Object {
            $Html += "<tr><td>$($_.DisplayName)</td><td>$($_.Alias)</td><td>$($_.FilePath)</td></tr>"
        }
        $Html += "</table>"
    }
}
$Html += "</div>"

# 5. FSLogix - Extremely Detailed (Profiles + Office + Disk Info)
Write-Log "Scanning FSLogix configuration across Session Hosts..."
$SessionHosts = (Get-RDSessionHost).SessionHost

$Html += @"
<div class="section">
    <h2>5. FSLogix Profile Disk Configuration</h2>
"@

foreach ($sh in $SessionHosts) {
    $Html += "<h3>Session Host: $sh</h3>"

    # FSLogix Profiles
    $Profiles = Invoke-Command -ComputerName $sh -ScriptBlock {
        Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -ErrorAction SilentlyContinue
    } -ErrorAction SilentlyContinue

    if ($Profiles) {
        $Html += "<h4>FSLogix Profiles Settings</h4><table><tr><th>Setting</th><th>Value</th></tr>"
        $Profiles.PSObject.Properties | Where-Object {$_.Name -notlike "PS*"} | ForEach-Object {
            $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
        }
        $Html += "</table>"

        # Profile Disk Details
        if ($Profiles.VHDLocations) {
            $Html += "<h4>Profile Disk Locations</h4><p>$($Profiles.VHDLocations)</p>"
        }
        if ($Profiles.VHDXSizeBytes) {
            $sizeGB = [math]::Round($Profiles.VHDXSizeBytes / 1GB, 2)
            $Html += "<p><strong>Max Profile Disk Size:</strong> $sizeGB GB</p>"
        }
    } else {
        $Html += "<p>No FSLogix Profiles configuration found on this host.</p>"
    }

    # FSLogix Office Container
    $Office = Invoke-Command -ComputerName $sh -ScriptBlock {
        Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\OfficeContainers" -ErrorAction SilentlyContinue
    } -ErrorAction SilentlyContinue

    if ($Office) {
        $Html += "<h4>FSLogix Office Container Settings</h4><table><tr><th>Setting</th><th>Value</th></tr>"
        $Office.PSObject.Properties | Where-Object {$_.Name -notlike "PS*"} | ForEach-Object {
            $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
        }
        $Html += "</table>"
    }
}

$Html += "</div>"

# Close HTML
$Html += "</body></html>"

# Save Report
$Html | Out-File -FilePath $HtmlPath -Encoding UTF8

# Export useful CSVs
Get-RDSessionCollectionConfiguration | Export-Csv "$ReportFolder\Collection_Configs.csv" -NoTypeInformation
Get-RDRemoteApp | Export-Csv "$ReportFolder\RemoteApps.csv" -NoTypeInformation
Get-RDGatewayServer | Export-Csv "$ReportFolder\RDGateway_Servers.csv" -NoTypeInformation

Write-Log "Ultra-detailed RDS As-Built completed successfully!"
Write-Host "`nDetailed RDS As-Built report generated and opened." -ForegroundColor Green
Invoke-Item $HtmlPath
