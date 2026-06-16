<#
.SYNOPSIS
   Detailed Professional RDS As-Built Documentation (HTML)
   Includes: Deployment (via servers/roles), RD Gateway, RD Web Access, Collections, FSLogix
#>

param (
    [string]$OutputPath = "C:\RDS_AsBuilt"
)

$ReportDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
$ReportFolder = Join-Path $OutputPath "RDS_AsBuilt_Professional_$ReportDate"
New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null
$HtmlPath = Join-Path $ReportFolder "RDS_AsBuilt_Professional_Report.html"
$LogPath = Join-Path $ReportFolder "RDS_AsBuilt_Log.txt"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Tee-Object -FilePath $LogPath -Append
}

Write-Log "Starting Professional RDS As-Built HTML Report..."

Import-Module RemoteDesktop -ErrorAction SilentlyContinue
if (-not (Get-Module RemoteDesktop)) {
    Write-Host "ERROR: Run this script on a Connection Broker server (RemoteDesktop module required)." -ForegroundColor Red
    exit 1
}

# ========================== PROFESSIONAL HTML HEADER ==========================
$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>RDS Environment - Professional As-Built</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;500;600&display=swap');
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 40px;
            background: #f8f9fa;
            color: #333;
            line-height: 1.6;
        }
        .container {
            max-width: 1300px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 4px 25px rgba(0,0,0,0.12);
        }
        h1 {
            color: #003366;
            text-align: center;
            border-bottom: 4px solid #003366;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        h2 {
            color: #003366;
            border-left: 6px solid #003366;
            padding-left: 18px;
            background: #f0f4f8;
            padding-top: 12px;
            padding-bottom: 12px;
            margin-top: 45px;
        }
        h3 {
            color: #0055aa;
            margin-top: 35px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 18px 0;
        }
        th, td {
            padding: 12px 15px;
            border: 1px solid #ccc;
            text-align: left;
        }
        th {
            background-color: #003366;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .section {
            margin-bottom: 55px;
        }
        .header-info {
            text-align: center;
            margin-bottom: 40px;
            color: #555;
            font-size: 1.1em;
        }
        .footer {
            text-align: center;
            margin-top: 80px;
            color: #777;
            font-size: 0.95em;
        }
        .warning { color: #d9534f; font-style: italic; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Remote Desktop Services (RDS)<br>As-Built Documentation</h1>
        <div class="header-info">
            <p><strong>Generated:</strong> $(Get-Date -Format "MMMM dd, yyyy HH:mm") &nbsp;&nbsp;|&nbsp;&nbsp;
               <strong>Server:</strong> $env:COMPUTERNAME</p>
        </div>
"@
Write-Log "HTML header generated."

# 1. Deployment Overview (using Get-RDServer)
$Html += @"
        <div class="section">
            <h2>1. Deployment Overview (RDS Servers & Roles)</h2>
            <table>
                <tr><th>Server Name</th><th>Roles</th></tr>
"@
try {
    $Servers = Get-RDServer -ErrorAction Stop
    foreach ($srv in $Servers) {
        $roles = if ($srv.Roles) { $srv.Roles -join ', ' } else { "None" }
        $Html += "<tr><td>$($srv.Server)</td><td>$roles</td></tr>"
    }
} catch {
    $Html += "<tr><td colspan='2' class='warning'>Error retrieving servers: $($_.Exception.Message)</td></tr>"
}
$Html += "</table></div>"

# 2. RD Gateway
$Html += @"
        <div class="section">
            <h2>2. RD Gateway Configuration</h2>
"@
try {
    $GatewayConfig = Get-RDDeploymentGatewayConfiguration -ErrorAction Stop
    if ($GatewayConfig) {
        $Html += "<h3>Gateway Settings</h3>"
        $Html += "<table><tr><th>Property</th><th>Value</th></tr>"
        $GatewayConfig.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" } | ForEach-Object {
            $val = if ($_.Value -is [array]) { $_.Value -join "; " } else { $_.Value }
            $Html += "<tr><td>$($_.Name)</td><td>$val</td></tr>"
        }
        $Html += "</table>"
    } else {
        $Html += "<p>No RD Gateway configured or accessible.</p>"
    }
} catch {
    $Html += "<p class='warning'>Error retrieving Gateway config: $($_.Exception.Message)</p>"
}
$Html += "</div>"

# 3. RD Web Access
$Html += @"
        <div class="section">
            <h2>3. Remote Desktop Web Access</h2>
"@
$WebServers = $Servers | Where-Object { $_.Roles -like "*RD-Web-Access*" }
if ($WebServers) {
    foreach ($wa in $WebServers) {
        $Html += "<h3>Web Access Server: $($wa.Server)</h3>"
        $Html += "<table><tr><th>Property</th><th>Value</th></tr>"
        $wa.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" } | ForEach-Object {
            $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
        }
        $Html += "</table>"
    }
} else {
    $Html += "<p>No RD Web Access role found.</p>"
}
$Html += "</div>"

# 4. Session Collections
$Html += @"
        <div class="section">
            <h2>4. RDS Session Collections & Assignments</h2>
"@
try {
    $Collections = Get-RDSessionCollection -ErrorAction Stop
    foreach ($col in $Collections) {
        $ColName = $col.CollectionName
        $Config = Get-RDSessionCollectionConfiguration -CollectionName $ColName -ErrorAction SilentlyContinue
        $Groups = (Get-RDSessionCollectionConfiguration -CollectionName $ColName -UserGroup -ErrorAction SilentlyContinue).UserGroup

        $Html += "<h3>Collection: $ColName</h3>"
        if ($Config) {
            $Html += "<table><tr><th>Property</th><th>Value</th></tr>"
            $Config.PSObject.Properties | Where-Object { $_.Name -notlike "*UserGroup*" -and $_.Name -notlike "PS*" } | ForEach-Object {
                $val = if ($_.Value -is [array]) { $_.Value -join "; " } else { $_.Value }
                $Html += "<tr><td>$($_.Name)</td><td>$val</td></tr>"
            }
            $Html += "</table>"
        }

        $Html += "<h4>Assigned User Groups</h4><table><tr><th>Group Name</th></tr>"
        if ($Groups) {
            $Groups | ForEach-Object { $Html += "<tr><td>$_</td></tr>" }
        } else {
            $Html += "<tr><td>None</td></tr>"
        }
        $Html += "</table>"
    }
} catch {
    $Html += "<p class='warning'>Error retrieving collections: $($_.Exception.Message)</p>"
}
$Html += "</div>"

# 5. FSLogix Profile Configuration
$Html += @"
        <div class="section">
            <h2>5. FSLogix Profile Disk Configuration</h2>
"@
$SessionHosts = (Get-RDSessionHost -ErrorAction SilentlyContinue).SessionHost
if (-not $SessionHosts) { $SessionHosts = @($env:COMPUTERNAME) }  # fallback

foreach ($sh in $SessionHosts) {
    $Html += "<h3>Session Host: $sh</h3>"
    try {
        $Profiles = Invoke-Command -ComputerName $sh -ScriptBlock {
            Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles" -ErrorAction SilentlyContinue
        } -ErrorAction Stop
        if ($Profiles) {
            $Html += "<table><tr><th>Setting</th><th>Value</th></tr>"
            $Profiles.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" } | ForEach-Object {
                $Html += "<tr><td>$($_.Name)</td><td>$($_.Value)</td></tr>"
            }
            $Html += "</table>"
        } else {
            $Html += "<p>No FSLogix Profiles configuration found on this host.</p>"
        }
    } catch {
        $Html += "<p class='warning'>Could not query FSLogix on $sh: $($_.Exception.Message)</p>"
    }
}
$Html += "</div>"

# Footer
$Html += @"
        <div class="footer">
            <p><strong>Professional RDS As-Built Documentation</strong> — Confidential</p>
            <p>Generated by PowerShell Script • $(Get-Date -Format "MMMM dd, yyyy")</p>
        </div>
    </div>
</body>
</html>
"@

# Save the Report
$Html | Out-File -FilePath $HtmlPath -Encoding UTF8

# Export CSVs
Get-RDServer | Export-Csv "$ReportFolder\RDS_Servers.csv" -NoTypeInformation
Get-RDSessionCollectionConfiguration | Export-Csv "$ReportFolder\Collection_Configs.csv" -NoTypeInformation -ErrorAction SilentlyContinue

Write-Log "Professional HTML As-Built report generated successfully!"
Write-Host "`n✅ Professional RDS As-Built HTML report has been generated!" -ForegroundColor Green
Write-Host "📁 Location: $HtmlPath" -ForegroundColor Cyan
Invoke-Item $HtmlPath
