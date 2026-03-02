# InactiveADCleanup.ps1
# PowerShell script to identify, report, disable, and move inactive user accounts in Active Directory.
# Features:
# - Interactive prompts for timeframe (30/60/90 days) and action (report or disable/move).
# - Searches for inactive users based on LastLogonTimestamp.
# - Generates CSV report with details like Name, SamAccountName, LastLogonDate, DistinguishedName.
# - Creates timeframe-specific OU (e.g., "30DayDisabled") under domain root if it doesn't exist.
# - Disables accounts and moves them to the corresponding OU.
# - Additional features: 
#   - Dry-run mode (WhatIf) for testing without changes.
#   - Logging to a timestamped file.
#   - Confirmation prompt before disabling/moving.
#   - Option to exclude specific OUs or accounts (configurable in variables).
#   - Error handling and verbose output.
# - Requires: ActiveDirectory module (install via Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 on Windows 10+).
# - Run as admin with domain admin privileges.

# Import required module
Import-Module ActiveDirectory -ErrorAction Stop

# Configurable variables
$logPath = "$PSScriptRoot\InactiveADCleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$reportPath = "$PSScriptRoot\InactiveUsersReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$excludeOUs = @("CN=System,DC=yourdomain,DC=com")  # Add DN patterns to exclude (e.g., built-in OUs)
$excludeAccounts = @("Administrator", "Guest")  # SamAccountNames to exclude
$domainDN = (Get-ADDomain).DistinguishedName  # Domain root DN
$ouPrefix = "DayDisabled"  # e.g., 30DayDisabled

# Function to log messages
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File -FilePath $logPath -Append
    Write-Verbose $Message
}

# Start logging
Write-Log "Script started by $($env:USERNAME)."

# Prompt for timeframe
Write-Host "Select inactivity timeframe:"
Write-Host "1: 30 days"
Write-Host "2: 60 days"
Write-Host "3: 90 days"
$choice = Read-Host "Enter choice (1/2/3)"
switch ($choice) {
    "1" { $days = 30 }
    "2" { $days = 60 }
    "3" { $days = 90 }
    default { Write-Log "Invalid choice. Exiting." -Level "ERROR"; exit }
}
$inactiveDate = (Get-Date).AddDays(-$days)
$ouName = "$($days)$ouPrefix"
$ouPath = "OU=$ouName,$domainDN"
Write-Log "Selected timeframe: $days days. Inactive since: $inactiveDate. Target OU: $ouPath."

# Prompt for action
Write-Host "Select action:"
Write-Host "1: Generate report only"
Write-Host "2: Disable and move accounts (with confirmation)"
$actionChoice = Read-Host "Enter choice (1/2)"
$doAction = ($actionChoice -eq "2")

# Prompt for dry-run
$whatIf = $false
$dryRunChoice = Read-Host "Enable dry-run mode (WhatIf)? (Y/N) [Default: N]"
if ($dryRunChoice -eq "Y") { $whatIf = $true }
Write-Log "Dry-run mode: $whatIf."

# Search for inactive users
Write-Log "Searching for inactive users..."
$inactiveUsers = Search-ADAccount -AccountInactive -TimeSpan (New-TimeSpan -Days $days) -UsersOnly -SearchBase $domainDN |
    Where-Object {
        $_.SamAccountName -notin $excludeAccounts -and
        $excludeOUs -notcontains $_.DistinguishedName.Split(',', 2)[1]  # Basic OU exclusion
    } |
    Get-ADUser -Properties LastLogonDate, DistinguishedName |
    Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, Enabled

Write-Log "Found $($inactiveUsers.Count) inactive users."

if ($inactiveUsers.Count -eq 0) {
    Write-Host "No inactive users found. Exiting."
    Write-Log "No inactive users found. Script ended."
    exit
}

# Generate report regardless of action
Write-Log "Generating report: $reportPath"
$inactiveUsers | Export-Csv -Path $reportPath -NoTypeInformation
Write-Host "Report saved to: $reportPath"

if (-not $doAction) {
    Write-Log "Report-only mode. Script ended."
    exit
}

# Confirm action
$confirm = Read-Host "Proceed with disabling and moving $($inactiveUsers.Count) accounts? (Y/N)"
if ($confirm -ne "Y") {
    Write-Log "Action cancelled by user. Script ended."
    exit
}

# Create OU if it doesn't exist
try {
    if (-not (Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -SearchBase $domainDN)) {
        New-ADOrganizationalUnit -Name $ouName -Path $domainDN -ProtectedFromAccidentalDeletion $true -WhatIf:$whatIf
        Write-Log "Created OU: $ouPath"
    } else {
        Write-Log "OU already exists: $ouPath"
    }
} catch {
    Write-Log "Error creating OU: $_" -Level "ERROR"
    exit
}

# Disable and move accounts
foreach ($user in $inactiveUsers) {
    try {
        if ($user.Enabled) {
            Disable-ADAccount -Identity $user.DistinguishedName -WhatIf:$whatIf -Confirm:$false
            Write-Log "Disabled: $($user.SamAccountName)"
        }
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $ouPath -WhatIf:$whatIf -Confirm:$false
        Write-Log "Moved: $($user.SamAccountName) to $ouPath"
    } catch {
        Write-Log "Error processing $($user.SamAccountName): $_" -Level "ERROR"
    }
}

Write-Log "Script completed successfully."
Write-Host "Actions completed. Check log: $logPath"
