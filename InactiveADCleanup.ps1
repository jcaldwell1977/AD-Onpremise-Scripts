#Requires -RunAsAdministrator
# AD Consolidation - Stale Object Cleanup Script (Users + Computers)
# Supports 30/60/90/120/150/180 days + exclusions by OU/Group/User

param(
    [ValidateSet(30,60,90,120,150,180)]
    [int]$Days = 30,

    [string[]]$ExcludeOU = @(),          # e.g. "OU=ServiceAccounts,DC=corp,DC=cfsbrands,DC=com"
    [string[]]$ExcludeGroup = @(),       # Group names (samAccountName)
    [string[]]$ExcludeUser = @(),        # Specific samAccountNames

    [switch]$Disable,                    # WARNING: Only use after reviewing report!
    [switch]$ComputersOnly,
    [switch]$UsersOnly
)

Import-Module ActiveDirectory -ErrorAction SilentlyContinue

$OutputCSV = "C:\Temp\Stale-Objects-Report-Days$Days-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$CutoffDate = (Get-Date).AddDays(-$Days)

Write-Host "=== Stale Object Scan ($Days days inactive) ===" -ForegroundColor Cyan
Write-Host "Cutoff date: $CutoffDate" -ForegroundColor Yellow

# Build filter
$Filter = "LastLogonTimestamp -lt $($CutoffDate.ToFileTime()) -and Enabled -eq 'True'"

# Get objects
$Objects = @()

if (-not $ComputersOnly) {
    Write-Host "Scanning Users..." -ForegroundColor Gray
    $Objects += Get-ADUser -Filter $Filter -Properties LastLogonTimestamp, DistinguishedName, MemberOf, Description -ErrorAction SilentlyContinue
}

if (-not $UsersOnly) {
    Write-Host "Scanning Computers..." -ForegroundColor Gray
    $Objects += Get-ADComputer -Filter $Filter -Properties LastLogonTimestamp, DistinguishedName, Description -ErrorAction SilentlyContinue
}

Write-Host "Found $($Objects.Count) stale objects before exclusions." -ForegroundColor Green

# Apply exclusions
$FinalList = @()
$total = $Objects.Count
$current = 0

foreach ($obj in $Objects) {
    $current++
    $percent = [math]::Round(($current / $total) * 100)
    Write-Progress -Activity "Applying Exclusions" -Status "Checking $($obj.SamAccountName)" -PercentComplete $percent

    $DN = $obj.DistinguishedName
    $Sam = $obj.SamAccountName

    # OU exclusion
    if ($ExcludeOU | Where-Object { $DN -like "*$_*" }) { continue }

    # Group exclusion
    if ($obj.MemberOf) {
        $groups = $obj.MemberOf | ForEach-Object { (Get-ADGroup $_).SamAccountName }
        if ($groups | Where-Object { $ExcludeGroup -contains $_ }) { continue }
    }

    # Specific user exclusion
    if ($ExcludeUser -contains $Sam) { continue }

    $FinalList += [PSCustomObject]@{
        Type                = if ($obj.ObjectClass -eq "user") { "User" } else { "Computer" }
        SamAccountName      = $Sam
        Name                = $obj.Name
        OU                  = ($DN -split ",OU=" | Select-Object -Skip 1) -join ", "
        LastLogonDate       = if ($obj.LastLogonTimestamp) { [DateTime]::FromFileTime($obj.LastLogonTimestamp) } else { "Never" }
        DaysInactive        = $Days
        Description         = $obj.Description
        DistinguishedName   = $DN
    }
}

# Clear progress
Write-Progress -Activity "Applying Exclusions" -Completed

# Export report
$FinalList | Export-Csv -Path $OutputCSV -NoTypeInformation -Encoding UTF8

Write-Host "`n=== Scan Complete ===" -ForegroundColor Cyan
Write-Host "Total stale objects after exclusions: $($FinalList.Count)" -ForegroundColor Green
Write-Host "Report saved: $OutputCSV" -ForegroundColor Yellow

# Optional disable
if ($Disable) {
    Write-Host "DISABLING objects (as requested)..." -ForegroundColor Red
    foreach ($item in $FinalList) {
        if ($item.Type -eq "User") {
            Disable-ADAccount -Identity $item.SamAccountName
            Write-Host "Disabled User: $($item.SamAccountName)" -ForegroundColor Red
        } else {
            Disable-ADAccount -Identity $item.SamAccountName
            Write-Host "Disabled Computer: $($item.SamAccountName)" -ForegroundColor Red
        }
    }
    Write-Host "All objects disabled!" -ForegroundColor Red
} else {
    Write-Host "Run again with -Disable to actually disable these objects (after review!)." -ForegroundColor Yellow
}
