#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory, ServerManager

<#
    Check Domain Controllers for:
    - Certificate Authority (AD CS)
    - Network Policy Server (NPS)
    - DHCP Server role
    - Print Server role
    - Non-standard file shares (excluding SYSVOL, NETLOGON, and typical admin/hidden shares)
    
    Also checks SYSVOL replication method: FRS vs DFSR
#>

Write-Host "Checking Domain Controllers for sensitive roles, file shares, and SYSVOL replication..." -ForegroundColor Cyan
Write-Host "---------------------------------------------------------------`n"

# Get all domain controllers
$DCs = Get-ADDomainController -Filter * | Sort-Object Name

if (-not $DCs) {
    Write-Warning "No domain controllers found."
    exit
}

Write-Host "Found $($DCs.Count) Domain Controller(s):"
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

        $hasCA     = $features -contains "AD-Certificate"
        $hasNPS    = $features -contains "NPAS" -or $features -contains "NPAS-Policy-Server"
        $hasDHCP   = $features -contains "DHCP"
        $hasPrint  = $features -contains "Print-Server"   # Print Server role

        # Get all file shares (excluding default SYSVOL, NETLOGON, and typical admin/hidden shares)
        $shares = Get-SmbShare -CimSession $computerName -ErrorAction Stop |
            Where-Object {
                $_.Name -notin @("SYSVOL", "NETLOGON", "ADMIN$", "C$", "IPC$", "print$") -and
                $_.Special -eq $false -and
                $_.Path -notlike "*\Windows\*"   # exclude hidden/admin shares
            } |
            Select-Object -Property Name, Path, Description -First 10

        $shareCount = $shares.Count
        $shareList  = if ($shareCount -gt 0) { $shares.Name -join ", " } else { "None found" }

        $status = [PSCustomObject]@{
            DCName     = $computerName
            HasCA      = $hasCA
            HasNPS     = $hasNPS
            HasDHCP    = $hasDHCP
            HasPrint   = $hasPrint
            AnyRisky   = $hasCA -or $hasNPS -or $hasDHCP -or $hasPrint
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

# Check SYSVOL replication method (FRS vs DFSR)
Write-Host "`nChecking SYSVOL replication method (FRS vs DFSR)..." -ForegroundColor Cyan

try {
    $dfsrmig = Get-DfsrMigrationState -ErrorAction Stop
    if ($dfsrmig.MigrationState -eq "Eliminated") {
        Write-Host "SYSVOL is using " -NoNewline
        Write-Host "DFSR" -ForegroundColor Green -NoNewline
        Write-Host " (migration completed)"
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
    Write-Host "  (Get-DfsrMigrationState failed – domain probably not migrated to DFSR)"
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor DarkGray
}

Write-Host "`nScan complete." -ForegroundColor Cyan
