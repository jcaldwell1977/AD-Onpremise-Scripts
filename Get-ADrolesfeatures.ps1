#Requires -RunAsAdministrator
# AD Consolidation - Server Role/Feature Inventory Script
# Scans domain-joined Windows servers for key roles and features relevant to domain migration

Import-Module ActiveDirectory -ErrorAction SilentlyContinue
Import-Module ServerManager -ErrorAction SilentlyContinue

# Output file (timestamped in C:\Temp)
$OutputCSV = "C:\Temp\AD-Server-Role-Inventory-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"

# Get all enabled Windows servers from the current domain
Write-Host "Querying AD for enabled Windows servers..." -ForegroundColor Cyan
$Servers = Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' `
           -Properties Name, DNSHostName, OperatingSystem, OperatingSystemVersion, IPv4Address, LastLogonDate `
           -ErrorAction Stop | Sort-Object Name

if ($Servers.Count -eq 0) {
    Write-Warning "No servers found in AD. Check filter or permissions."
    exit
}

Write-Host "Found $($Servers.Count) servers. Starting role scan..." -ForegroundColor Green

$Results = @()

foreach ($Server in $Servers) {
    $ComputerName = $Server.Name
    Write-Host "Scanning $ComputerName ..." -ForegroundColor Yellow -NoNewline

    $Result = [PSCustomObject]@{
        ServerName             = $ComputerName
        DNSHostName            = $Server.DNSHostName
        IPAddress              = $Server.IPv4Address
        OS                     = $Server.OperatingSystem
        OSVersion              = $Server.OperatingSystemVersion
        LastLogon              = $Server.LastLogonDate
        IsOnline               = $false
        DHCPInstalled          = "N/A"
        DNSInstalled           = "N/A"
        FileServices           = "N/A"
        PrintServices          = "N/A"
        IIS_WebServer          = "N/A"
        AD_CertificateServices = "N/A"   # CA
        SQL_Detected           = "N/A"
        NPS_Role               = "N/A"   # Network Policy Server (RADIUS)
        RDS_Role               = "N/A"   # Remote Desktop Services
        FailoverClustering     = "N/A"
        DFS_Namespace          = "N/A"
        DFS_Replication        = "N/A"
        WSUS                   = "N/A"   # Windows Server Update Services
        WDS                    = "N/A"   # Windows Deployment Services
        Notes_Errors           = ""
    }

    # Test connectivity
    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        $Result.IsOnline = $true

        try {
            # Get installed features/roles
            $Features = Get-WindowsFeature -ComputerName $ComputerName -ErrorAction Stop

            # DHCP Server
            $Result.DHCPInstalled = if ($Features | Where-Object { $_.Name -eq "DHCP" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # DNS Server
            $Result.DNSInstalled = if ($Features | Where-Object { $_.Name -eq "DNS" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # File and Storage Services (broad)
            $Result.FileServices = if ($Features | Where-Object { $_.Name -like "FS*" -and $_.InstallState -eq "Installed" }) { "Installed (FS*)" } else { "Not Installed" }

            # Print and Document Services
            $Result.PrintServices = if ($Features | Where-Object { $_.Name -eq "Print-Services" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # Web Server (IIS)
            $Result.IIS_WebServer = if ($Features | Where-Object { $_.Name -eq "Web-Server" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # Active Directory Certificate Services (CA)
            $Result.AD_CertificateServices = if ($Features | Where-Object { $_.Name -eq "AD-Certificate" -and $_.InstallState -eq "Installed" }) { "Installed (CA)" } else { "Not Installed" }

            # Network Policy Server (RADIUS for WiFi/VPN)
            $Result.NPS_Role = if ($Features | Where-Object { $_.Name -like "*NPAS*" -or $_.Name -eq "NPAS-Policy-Server" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # Remote Desktop Services (broad check)
            $Result.RDS_Role = if ($Features | Where-Object { $_.Name -like "RDS*" -and $_.InstallState -eq "Installed" }) { "Installed (RDS*)" } else { "Not Installed" }

            # Failover Clustering
            $Result.FailoverClustering = if ($Features | Where-Object { $_.Name -eq "Failover-Clustering" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # DFS Namespace
            $Result.DFS_Namespace = if ($Features | Where-Object { $_.Name -eq "FS-DFS-Namespace" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # DFS Replication
            $Result.DFS_Replication = if ($Features | Where-Object { $_.Name -eq "FS-DFS-Replication" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # WSUS (Windows Server Update Services)
            $Result.WSUS = if ($Features | Where-Object { $_.Name -eq "UpdateServices" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # WDS (Windows Deployment Services)
            $Result.WDS = if ($Features | Where-Object { $_.Name -eq "WDS" -and $_.InstallState -eq "Installed" }) { "Installed" } else { "Not Installed" }

            # SQL Server detection (feature + service check)
            $SQLFeature = $Features | Where-Object { $_.Name -like "*SQL*" -and $_.InstallState -eq "Installed" }
            $SQLService = Get-Service -ComputerName $ComputerName -Name "MSSQL*" -ErrorAction SilentlyContinue
            if ($SQLFeature -or $SQLService) {
                $Result.SQL_Detected = "Detected (Feature: $($SQLFeature.Count > 0); Service: $($SQLService.Count > 0))"
            } else {
                $Result.SQL_Detected = "Not Detected"
            }
        }
        catch {
            $Result.Notes_Errors = "Error: $($_.Exception.Message)"
        }
    } else {
        $Result.Notes_Errors = "Offline / Unreachable"
    }

    $Results += $Result
    Write-Host " Done" -ForegroundColor Green
}

# Export to CSV
$Results | Export-Csv -Path $OutputCSV -NoTypeInformation -Encoding UTF8

Write-Host "`nScan complete!" -ForegroundColor Cyan
Write-Host "Results saved to: $OutputCSV" -ForegroundColor Green
Write-Host "Total servers scanned: $($Servers.Count)"
Write-Host "Online servers: $($Results.Where({$_.IsOnline}).Count)"
