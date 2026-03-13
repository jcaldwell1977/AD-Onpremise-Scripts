#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory, ServerManager, DhcpServer
<#
    Check Domain Controllers for:
    - Certificate Authority (AD CS)
    - Network Policy Server (NPS)
    - DHCP Server role
    - Print Server role
    - Non-standard file shares (excluding SYSVOL, NETLOGON, and typical admin/hidden shares)
    
    Also checks SYSVOL replication method: FRS vs DFSR
    Now includes:
    - List all domain controllers with OS version
    - Prompt user to scan domain for IIS and SQL Server instances
    - Generate report if user chooses to scan
#>

Write-Host "Checking Domain Controllers for sensitive roles, file shares, SYSVOL replication, and DHCP servers..." -ForegroundColor Cyan
Write-Host "---------------------------------------------------------------`n"

# Get all domain controllers with OS version
Write-Host "Retrieving all domain controllers and their OS versions..." -ForegroundColor Cyan
$DCs = Get-ADDomainController -Filter * -Properties OperatingSystem | Sort-Object Name
if (-not $DCs) {
    Write-Warning "No domain controllers found."
    exit
}

Write-Host "Found $($DCs.Count) Domain Controller(s):" -ForegroundColor Yellow
$DCs | Select-Object Name, IPv4Address, OperatingSystem, Site | Format-Table -AutoSize

Write-Host "`nScanning each DC for roles and shares..." -ForegroundColor Yellow
$results = @()

foreach ($dc in $DCs) {
    $computerName = $dc.Name
    Write-Host "Scanning $computerName ..." -NoNewline -ForegroundColor Gray

    try {
        # Get installed roles/features remotely
        $features = Get-WindowsFeature -ComputerName $computerName -ErrorAction Stop |
            Where-Object { $_.Installed -eq $true } |
            Select-Object -ExpandProperty Name

        $hasCA = $features -contains "AD-Certificate"
        $hasNPS = $features -contains "NPAS" -or $features -contains "NPAS-Policy-Server"
        $hasDHCP = $features -contains "DHCP"
        $hasPrint = $features -contains "Print-Server"

        # Get all file shares (excluding default SYSVOL, NETLOGON, and typical admin/hidden shares)
        $shares = Get-SmbShare -CimSession $computerName -ErrorAction Stop |
            Where-Object {
                $_.Name -notin @("SYSVOL", "NETLOGON", "ADMIN$", "C$", "IPC$", "print$") -and
                $_.Special -eq $false -and
                $_.Path -notlike "*\Windows\*" 
            } |
            Select-Object -Property Name, Path, Description -First 10

        $shareCount = $shares.Count
        $shareList = if ($shareCount -gt 0) { $shares.Name -join ", " } else { "None found" }

        $status = [PSCustomObject]@{
            DCName    = $computerName
            HasCA     = $hasCA
            HasNPS    = $hasNPS
            HasDHCP   = $hasDHCP
            HasPrint  = $hasPrint
            AnyRisky  = $hasCA -or $hasNPS -or $hasDHCP -or $hasPrint
            RolesFound = ($features | Where-Object { $_ -in @("AD-Certificate","NPAS","NPAS-Policy-Server","DHCP","Print-Server") }) -join ", "
            UserShares = "$shareCount share(s): $shareList"
        }

        if ($status.AnyRisky -or $shareCount -gt 0) {
            Write-Host " FOUND" -ForegroundColor Red
        } else {
            Write-Host " clean" -ForegroundColor Green
        }

        $results += $status
    }
    catch {
        Write-Host " ERROR" -ForegroundColor Red
        Write-Warning "Could not query $computerName : $($_.Exception.Message)"
        $results += [PSCustomObject]@{
            DCName     = $computerName
            HasCA      = "Error"
            HasNPS     = "Error"
            HasDHCP    = "Error"
            HasPrint   = "Error"
            AnyRisky   = "Error"
            RolesFound = $_.Exception.Message
            UserShares = "Error querying shares"
        }
    }
}

Write-Host "`nSummary - Domain Controllers with sensitive roles or user file shares:" -ForegroundColor Cyan
$riskyDCs = $results | Where-Object { $_.AnyRisky -eq $true -or $_.AnyRisky -eq "Error" -or $_.UserShares -notlike "*None found*" }
if ($riskyDCs) {
    $riskyDCs | Format-Table DCName, HasCA, HasNPS, HasDHCP, HasPrint, UserShares -AutoSize
} else {
    Write-Host "No domain controllers have Certificate Authority, NPS, DHCP, Print Server roles, or non-standard user file shares." -ForegroundColor Green
}

# ------------------------------
# NEW: Check Registered DHCP Servers and Ping Test
# ------------------------------
Write-Host "`nChecking for registered DHCP servers in the domain..." -ForegroundColor Cyan

try {
    $dhcpServers = Get-DhcpServerInDC -ErrorAction Stop | Select-Object -ExpandProperty DnsName -Unique
    if ($dhcpServers) {
        Write-Host "Found $($dhcpServers.Count) registered DHCP server(s):" -ForegroundColor Yellow
        $dhcpServers | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }

        Write-Host "`nTesting ping response for each DHCP server..." -ForegroundColor Yellow
        $pingResults = @()
        foreach ($dhcp in $dhcpServers) {
            Write-Host "Pinging $dhcp ..." -NoNewline -ForegroundColor Gray
            $pingResult = Test-Connection -ComputerName $dhcp -Count 2 -Quiet -ErrorAction SilentlyContinue
            if ($pingResult) {
                Write-Host " RESPONDS" -ForegroundColor Green
                $pingResults += [PSCustomObject]@{ Server = $dhcp; Responds = "Yes" }
            } else {
                Write-Host " NO RESPONSE" -ForegroundColor Red
                $pingResults += [PSCustomObject]@{ Server = $dhcp; Responds = "No" }
            }
        }
    } else {
        Write-Host "No registered DHCP servers found in the domain." -ForegroundColor Green
    }
}
catch {
    Write-Warning "Failed to retrieve registered DHCP servers: $($_.Exception.Message)"
    Write-Host "  (This may occur if no DHCP servers are registered or permissions issue)" -ForegroundColor Yellow
}

# ------------------------------
# NEW: Prompt to Scan for IIS and SQL Servers
# ------------------------------
Write-Host "`nWould you like to scan the domain for IIS and SQL Server instances?" -ForegroundColor Cyan
$scanChoice = Read-Host "Enter Y for yes, N for no (default N)"

if ($scanChoice -eq "Y" -or $scanChoice -eq "y") {
    Write-Host "Scanning domain for Windows Server computers with IIS and/or SQL Server..." -ForegroundColor Yellow

    # Get all Windows Server computers (exclude DCs for speed, or include if you want)
    $servers = Get-ADComputer -Filter 'OperatingSystem -like "*Server*"' -Properties OperatingSystem, DNSHostName |
               Where-Object { $_.OperatingSystem -notlike "*Domain Controller*" } |
               Sort-Object Name

    if (-not $servers) {
        Write-Host "No Windows Server computers (excluding DCs) found in the domain." -ForegroundColor Yellow
    } else {
        Write-Host "Found $($servers.Count) Windows Server computer(s) to scan..." -ForegroundColor Yellow

        $iisServers = @()
        $sqlServers = @()

        foreach ($server in $servers) {
            $name = $server.DNSHostName
            Write-Host "Scanning $name ..." -NoNewline -ForegroundColor Gray

            try {
                $session = New-CimSession -ComputerName $name -ErrorAction Stop

                # Check for IIS (Web-Server role)
                $iis = Get-WindowsFeature -Name Web-Server -CimSession $session -ErrorAction Stop
                if ($iis.Installed) {
                    $iisServers += $name
                    Write-Host " IIS found" -ForegroundColor Red
                } else {
                    Write-Host " no IIS" -ForegroundColor Gray
                }

                # Check for SQL Server (common services)
                $sqlServices = Get-Service -Name "MSSQL*" -CimSession $session -ErrorAction Stop
                if ($sqlServices) {
                    $sqlServers += $name
                    Write-Host " SQL found" -ForegroundColor Red
                } else {
                    Write-Host " no SQL" -ForegroundColor Gray
                }

                Remove-CimSession $session
            }
            catch {
                Write-Host " ERROR" -ForegroundColor Red
                Write-Warning "Failed to scan $name : $($_.Exception.Message)"
            }
        }

        # Generate IIS & SQL Report
        Write-Host "`nIIS & SQL Server Scan Report:" -ForegroundColor Cyan

        if ($iisServers) {
            Write-Host "Servers with IIS installed:" -ForegroundColor Red
            $iisServers | ForEach-Object { Write-Host "  - $_" }
        } else {
            Write-Host "No servers with IIS found." -ForegroundColor Green
        }

        if ($sqlServers) {
            Write-Host "`nServers with SQL Server detected:" -ForegroundColor Red
            $sqlServers | ForEach-Object { Write-Host "  - $_" }
        } else {
            Write-Host "No servers with SQL Server found." -ForegroundColor Green
        }

        # Optional: Save report to file
        $reportPath = "C:\Temp\IIS_SQL_Report_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"
        "IIS Servers:`n$($iisServers -join "`n")`n`nSQL Servers:`n$($sqlServers -join "`n")" | Out-File $reportPath -Encoding UTF8
        Write-Host "Report saved to: $reportPath" -ForegroundColor Green
    }
} else {
    Write-Host "IIS & SQL scan skipped by user." -ForegroundColor Yellow
}

# ------------------------------
# Check SYSVOL replication method (FRS vs DFSR) + extra validation
# ------------------------------
Write-Host "`nChecking SYSVOL replication method (FRS vs DFSR)..." -ForegroundColor Cyan
try {
    $dfsrmig = Get-DfsrMigrationState -ErrorAction Stop
    if ($dfsrmig.MigrationState -eq "Eliminated") {
        Write-Host "SYSVOL is using " -NoNewline
        Write-Host "DFSR" -ForegroundColor Green -NoNewline
        Write-Host " (migration completed)"

        # Check FRS service is disabled
        $frsService = Get-Service -Name NTFRS -ErrorAction SilentlyContinue
        if ($frsService -and $frsService.Status -ne "Stopped") {
            Write-Host "WARNING: FRS service is not disabled (Status: $($frsService.Status)) - DFSR migration may not be fully complete" -ForegroundColor Yellow
        } else {
            Write-Host "FRS service is disabled (as expected for DFSR)" -ForegroundColor Green
        }

        # Check DFSR service is running
        $dfsrService = Get-Service -Name DFSR -ErrorAction SilentlyContinue
        if ($dfsrService -and $dfsrService.Status -eq "Running") {
            Write-Host "DFSR service is running" -ForegroundColor Green
        } else {
            Write-Host "WARNING: DFSR service is not running (Status: $($dfsrService.Status))" -ForegroundColor Yellow
        }

        # Check SYSVOL and NETLOGON shares exist
        $sysvolShare = Get-SmbShare -Name SYSVOL -ErrorAction SilentlyContinue
        $netlogonShare = Get-SmbShare -Name NETLOGON -ErrorAction SilentlyContinue
        if ($sysvolShare -and $netlogonShare) {
            Write-Host "SYSVOL and NETLOGON shares are present" -ForegroundColor Green
        } else {
            Write-Host "WARNING: SYSVOL or NETLOGON share is missing" -ForegroundColor Red
        }
    }
    elseif ($dfsrmig.MigrationState -match "Prepared|Redirected|Started") {
        Write-Host "SYSVOL migration to DFSR is " -NoNewline
        Write-Host "in progress" -ForegroundColor Yellow -NoNewline
        Write-Host " (state: $($dfsrmig.MigrationState))"
    }
    else {
        Write-Host "SYSVOL appears to be using " -NoNewline
        Write-Host "FRS" -ForegroundColor Red -NoNewline
        Write-Host " (classic File Replication Service)"
    }
}
catch {
    Write-Host "SYSVOL replication method: " -NoNewline
    Write-Host "FRS (likely)" -ForegroundColor Red
    Write-Host " (Get-DfsrMigrationState failed – domain probably not migrated to DFSR)"
    Write-Host " Error: $($_.Exception.Message)" -ForegroundColor DarkGray
}

# ------------------------------
# NEW: Run dcdiag Report (basic health check)
# ------------------------------
Write-Host "`nRunning dcdiag for domain health report..." -ForegroundColor Cyan
$dcdiagOutput = dcdiag /q /test:sysvol 2>&1  # Quiet mode, focus on SYSVOL test
if ($dcdiagOutput -match "passed test" -or $dcdiagOutput -eq $null) {
    Write-Host "dcdiag SYSVOL test: Passed (no errors reported)" -ForegroundColor Green
} else {
    Write-Host "dcdiag SYSVOL test: " -NoNewline
    Write-Host "Failed / Warnings found" -ForegroundColor Red
    Write-Host "Summary of dcdiag issues:" -ForegroundColor Yellow
    $dcdiagOutput | ForEach-Object { Write-Host $_ -ForegroundColor Red }
}

Write-Host "`nScan complete." -ForegroundColor Cyan
