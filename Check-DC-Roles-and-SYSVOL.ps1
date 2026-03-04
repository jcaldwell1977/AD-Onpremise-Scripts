#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory, ServerManager, GroupPolicy

<#
    Check Domain Controllers for:
    - Certificate Authority (AD CS)
    - Network Policy Server (NPS)
    - DHCP Server role

    Also checks SYSVOL replication method: FRS vs DFSR
#>

Write-Host "Checking Domain Controllers for sensitive roles and SYSVOL replication method..." -ForegroundColor Cyan
Write-Host "---------------------------------------------------------------`n"

# Get all domain controllers
$DCs = Get-ADDomainController -Filter * | Sort-Object Name

if (-not $DCs) {
    Write-Warning "No domain controllers found."
    exit
}

Write-Host "Found $($DCs.Count) Domain Controller(s):"
$DCs | Select-Object Name, IPv4Address, OperatingSystem, Site | Format-Table -AutoSize

Write-Host "`nChecking installed roles on each DC..." -ForegroundColor Yellow

$results = @()

foreach ($dc in $DCs) {
    $computerName = $dc.Name

    Write-Host "Scanning $computerName ..." -NoNewline -ForegroundColor Gray

    try {
        # Use ServerManager module to get roles/features remotely
        $roles = Get-WindowsFeature -ComputerName $computerName -ErrorAction Stop |
            Where-Object { $_.Installed -eq $true } |
            Select-Object -ExpandProperty Name

        $hasCA   = $roles -contains "AD-Certificate"
        $hasNPS  = $roles -contains "NPAS" -or $roles -contains "NPAS-Policy-Server"
        $hasDHCP = $roles -contains "DHCP"

        $status = [PSCustomObject]@{
            DCName     = $computerName
            HasCA      = $hasCA
            HasNPS     = $hasNPS
            HasDHCP    = $hasDHCP
            AnyRisky   = $hasCA -or $hasNPS -or $hasDHCP
            RolesFound = ($roles | Where-Object { $_ -in @("AD-Certificate","NPAS","NPAS-Policy-Server","DHCP") }) -join ", "
        }

        if ($status.AnyRisky) {
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
            AnyRisky   = "Error"
            RolesFound = $_.Exception.Message
        }
    }
}

Write-Host "`nSummary - Domain Controllers with sensitive roles:" -ForegroundColor Cyan
$riskyDCs = $results | Where-Object { $_.AnyRisky -eq $true -or $_.AnyRisky -eq "Error" }

if ($riskyDCs) {
    $riskyDCs | Format-Table DCName, HasCA, HasNPS, HasDHCP, RolesFound -AutoSize
} else {
    Write-Host "No domain controllers have Certificate Authority, NPS, or DHCP roles installed." -ForegroundColor Green
}

# ────────────────────────────────────────────────────────────────
# Check SYSVOL replication method (FRS vs DFSR)
# ────────────────────────────────────────────────────────────────

Write-Host "`nChecking SYSVOL replication method (FRS vs DFSR)..." -ForegroundColor Cyan

try {
    $dfsrmigState = Get-DfsrMigrationState -ErrorAction Stop

    if ($dfsrmigState.MigrationState -eq "Eliminated") {
        Write-Host "SYSVOL is using " -NoNewline
        Write-Host "DFSR" -ForegroundColor Green -NoNewline
        Write-Host " (migration completed)" -ForegroundColor Gray
    }
    elseif ($dfsrmigState.MigrationState -match "Prepared|Redirected|Started") {
        Write-Host "SYSVOL migration to DFSR is " -NoNewline
        Write-Host "in progress" -ForegroundColor Yellow -NoNewline
        Write-Host " (state: $($dfsrmigState.MigrationState))" -ForegroundColor Gray
    }
    else {
        # If Get-DfsrMigrationState fails or returns nothing → likely still FRS
        $frs = Get-ADDomain | Select-Object -ExpandProperty FSMORoleOwner
        $sysvolPath = "\\$($env:USERDNSDOMAIN)\sysvol"
        if (Test-Path $sysvolPath -ErrorAction SilentlyContinue) {
            Write-Host "SYSVOL appears to be using " -NoNewline
            Write-Host "FRS" -ForegroundColor Red -NoNewline
            Write-Host " (classic File Replication Service)" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "SYSVOL replication method: " -NoNewline
    Write-Host "FRS (likely)" -ForegroundColor Red
    Write-Host "  (Get-DfsrMigrationState failed – domain probably not migrated to DFSR)"
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor DarkGray
}

Write-Host "`nDone." -ForegroundColor Cyan
