#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory, ServerManager

<#
    Check Domain Controllers for:
    - Certificate Authority (AD CS)
    - Network Policy Server (NPS)
    - DHCP Server role
    - Print Server role (NEW)
    - Non-standard file shares (excluding SYSVOL and NETLOGON) (NEW)

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
        $hasPrint  = $features -contains "Print-Server"   # <-- NEW: Print Server role

        # Get all file shares (excluding default SYSVOL and NETLOGON)
        $shares = Get-SmbShare -CimSession $computerName -ErrorAction Stop |
            Where-Object {
                $_.Name -notin @("SYSVOL", "NETLOGON", "ADMIN$", "C$", "IPC$", "print$") -and
                $_.Special -eq $false -and
                $_.Path -notlike "*\Windows\*"   # exclude hidden/admin shares
            } |
            Select-Object -Property Name, Path, Description -First 10  
