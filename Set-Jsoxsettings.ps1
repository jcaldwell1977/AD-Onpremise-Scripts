#Requires -RunAsAdministrator
# -----------------------------------------
# Enterprise Windows Hardening Script – Detailed HTML Report
# SOX-focused, rollback-enabled, error-resilient
# -----------------------------------------

# === Paths & Variables ===
$reportFolder    = "C:\CIS_Hardening_SOX\Reports"
$reportFile      = Join-Path $reportFolder "Hardening_Report_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').html"
$logPath         = "C:\CIS_Hardening_SOX\hardening_transcript.txt"
$rollbackFolder  = "C:\CIS_Hardening_SOX\Backup"
$rollbackScript  = Join-Path $rollbackFolder "Rollback.ps1"

$rollbackContent = @()
$results         = @()

# Create folders
foreach ($folder in @($reportFolder, $rollbackFolder)) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

# Start logging
try {
    Start-Transcript -Path $logPath -Append -Force
    Write-Host "Hardening script started at $(Get-Date)"
}
catch {
    Write-Warning "Failed to start transcript: $($_.Exception.Message)"
}

# ------------------------------
# Helper: Add result to report
# ------------------------------
function Add-Result {
    param(
        [string]$Section,
        [string]$Status,          # Success, Warning, Error, Skipped, Info
        [string]$Action,
        [string]$Before = "N/A",
        [string]$After  = "N/A",
        [string]$Result = "",
        [string]$Details = ""
    )
    $results += [PSCustomObject]@{
        Section = $Section
        Status  = $Status
        Action  = $Action
        Before  = $Before
        After   = $After
        Result  = $Result
        Details = $Details -replace "`n","<br>"
    }
}

# ------------------------------
# 1. Domain Join Detection (SOX: Access Control)
# ------------------------------
$domainJoined = $false
try {
    $domainInfo = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "Domain" -ErrorAction Stop
    $domainJoined = ($domainInfo.Domain -and $domainInfo.Domain -ne "")
    Add-Result -Section "System Information" -Status "Success" -Action "Domain join check" -After "Domain joined: $domainJoined"
} catch {
    Add-Result -Section "System Information" -Status "Warning" -Action "Domain join check" -Details $_.Exception.Message
    $domainJoined = $false
}
Write-Host "Domain joined? $domainJoined"

# ------------------------------
# 2. Backup Audit Policies (SOX: Auditing)
# ------------------------------
$auditPolBackup = Join-Path $rollbackFolder "AuditPol_Backup.txt"

if (Get-Command AuditPol -ErrorAction SilentlyContinue) {
    if (Test-Path $auditPolBackup) { Remove-Item $auditPolBackup -Force -ErrorAction SilentlyContinue }
    $backupSuccess = $false
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            AuditPol /backup /file:"$auditPolBackup"
            if (Test-Path $auditPolBackup) {
                Add-Result -Section "Auditing" -Status "Success" -Action "Backup AuditPol settings" -After "Backup created: $auditPolBackup"
                $rollbackContent += "AuditPol /restore /file:`"$auditPolBackup`""
                $backupSuccess = $true
                break
            }
        } catch {
            Write-Warning "AuditPol backup attempt $attempt failed: $($_.Exception.Message)"
            Start-Sleep -Seconds 2
        }
    }
    if (-not $backupSuccess) {
        Add-Result -Section "Auditing" -Status "Error" -Action "Backup AuditPol settings" -Details "Backup failed after retries"
    }
} else {
    Add-Result -Section "Auditing" -Status "Skipped" -Action "Backup AuditPol settings" -Details "AuditPol command not available"
}

# ------------------------------
# 3. SOXAdmin Account (SOX: Access Control)
# ------------------------------
$soxUser = "SOXAdmin"
$password = ConvertTo-SecureString "P@ssw0rd123!" -AsPlainText -Force  # ← CHANGE IN PRODUCTION!
try {
    $userExists = Get-LocalUser -Name $soxUser -ErrorAction SilentlyContinue
    if (-not $userExists) {
        New-LocalUser -Name $soxUser -Password $password -FullName "SOX Compliance Admin" -ErrorAction Stop
        Add-LocalGroupMember -Group "Administrators" -Member $soxUser -ErrorAction Stop
        $rollbackContent += "Remove-LocalUser -Name '$soxUser' -ErrorAction SilentlyContinue"
        Add-Result -Section "Access Control" -Status "Success" -Action "Create SOXAdmin account" -After "Created and added to Administrators"
    } else {
        Add-Result -Section "Access Control" -Status "Info" -Action "Create SOXAdmin account" -Details "Account already exists"
    }
} catch {
    Add-Result -Section "Access Control" -Status "Error" -Action "Create SOXAdmin account" -Details $_.Exception.Message
}

# ------------------------------
# 4. Disable Built-in Accounts (SOX: Least Privilege)
# ------------------------------
$accountsToDisable = @("Guest", "DefaultAccount", "WDAGUtilityAccount")
foreach ($acct in $accountsToDisable) {
    try {
        $user = Get-LocalUser -Name $acct -ErrorAction SilentlyContinue
        if ($user) {
            if ($user.Enabled) {
                Disable-LocalUser -Name $acct -ErrorAction Stop
                $rollbackContent += "Enable-LocalUser -Name '$acct' -ErrorAction SilentlyContinue"
                Add-Result -Section "Access Control" -Status "Success" -Action "Disable $acct" -Before "Enabled" -After "Disabled"
            } else {
                Add-Result -Section "Access Control" -Status "Info" -Action "Disable $acct" -Details "Already disabled"
            }
        } else {
            Add-Result -Section "Access Control" -Status "Skipped" -Action "Disable $acct" -Details "Account not found"
        }
    } catch {
        Add-Result -Section "Access Control" -Status "Warning" -Action "Disable $acct" -Details $_.Exception.Message
    }
}

# ------------------------------
# 5. Audit Policies (SOX: Comprehensive Auditing)
# ------------------------------
$auditSettings = @(
    @{Sub="Logon"; Success="enable"; Failure="enable"},
    @{Sub="Logoff"; Success="enable"; Failure="enable"},
    @{Sub="Account Lockout"; Success="enable"; Failure="enable"},
    @{Sub="User Account Management"; Success="enable"; Failure="enable"},
    @{Sub="Security Group Management"; Success="enable"; Failure="enable"},
    @{Sub="Computer Account Management"; Success="enable"; Failure="enable"},
    @{Sub="Other Account Management Events"; Success="enable"; Failure="enable"},
    @{Sub="Policy Change"; Success="enable"; Failure="enable"},
    @{Sub="Audit Policy Change"; Success="enable"; Failure="enable"},
    @{Sub="Authentication Policy Change"; Success="enable"; Failure="enable"},
    @{Sub="Authorization Policy Change"; Success="enable"; Failure="enable"},
    @{Sub="MPSSVC Rule-Level Policy Change"; Success="enable"; Failure="enable"},
    @{Sub="Filtering Platform Policy Change"; Success="enable"; Failure="enable"},
    @{Sub="Object Access"; Success="enable"; Failure="enable"}
)

foreach ($item in $auditSettings) {
    try {
        AuditPol /set /subcategory:"$($item.Sub)" /success:$($item.Success) /failure:$($item.Failure)
        Add-Result -Section "Auditing" -Status "Success" -Action "Set $($item.Sub) audit" -After "Success + Failure enabled"
    } catch {
        Add-Result -Section "Auditing" -Status "Warning" -Action "Set $($item.Sub) audit" -Details $_.Exception.Message
    }
}

# ------------------------------
# 6. Event Log Sizes (SOX: Retain logs for compliance)
# ------------------------------
$logs = @{
    "Security"    = 4194304   # 4 MB
    "Application" = 4194304   # 4 MB
    "System"      = 4194304   # 4 MB
    "Windows PowerShell" = 4194304
}

foreach ($log in $logs.Keys) {
    try {
        wevtutil sl $log /ms:$($logs[$log])
        Add-Result -Section "Logging" -Status "Success" -Action "Set max size for ${log}" -After "$($logs[$log]) KB"
    } catch {
        Add-Result -Section "Logging" -Status "Warning" -Action "Set max size for ${log}" -Details $_.Exception.Message
    }
}

# ------------------------------
# 7. PowerShell Logging (SOX: Script Monitoring)
# ------------------------------
$psScriptBlockPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging"
$psTransPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\Transcription"

foreach ($path in @($psScriptBlockPath, $psTransPath)) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}

try {
    $beforeScript = (Get-ItemProperty -Path $psScriptBlockPath -Name "EnableScriptBlockLogging" -ErrorAction SilentlyContinue).EnableScriptBlockLogging
    $beforeTrans  = (Get-ItemProperty -Path $psTransPath -Name "EnableTranscripting" -ErrorAction SilentlyContinue).EnableTranscripting

    Set-ItemProperty -Path $psScriptBlockPath -Name "EnableScriptBlockLogging" -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $psTransPath -Name "EnableTranscripting" -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $psTransPath -Name "EnableInvocationHeader" -Value 1 -Type DWord -Force

    Add-Result -Section "Script Monitoring" -Status "Success" -Action "Enable PowerShell logging" `
               -Before "ScriptBlock: $beforeScript, Transcription: $beforeTrans" `
               -After "ScriptBlock: 1, Transcription: 1, Invocation Header: 1"
} catch {
    Add-Result -Section "Script Monitoring" -Status "Error" -Action "Enable PowerShell logging" -Details $_.Exception.Message
}

# ------------------------------
# 8. Disable SMBv1 & Insecure Protocols
# ------------------------------
try {
    $beforeSMB1 = (Get-SmbServerConfiguration).EnableSMB1Protocol
    Set-SmbServerConfiguration -EnableSMB1Protocol $false -Force -ErrorAction Stop
    Add-Result -Section "Protocols" -Status "Success" -Action "Disable SMBv1" -Before $beforeSMB1 -After $false
} catch {
    Add-Result -Section "Protocols" -Status "Warning" -Action "Disable SMBv1" -Details $_.Exception.Message
}

try {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "LmCompatibilityLevel" -Value 5 -Type DWord -Force
    Add-Result -Section "Protocols" -Status "Success" -Action "Set LmCompatibilityLevel to 5 (NTLMv2 only)" -After "5"
} catch {
    Add-Result -Section "Protocols" -Status "Warning" -Action "Set LmCompatibilityLevel" -Details $_.Exception.Message
}

# ------------------------------
# 9. Credential Guard & LSA Protection
# ------------------------------
$deviceGuardPath = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard"
if (-not (Test-Path $deviceGuardPath)) { New-Item -Path $deviceGuardPath -Force | Out-Null }

try {
    $beforeVBS = (Get-ItemProperty -Path $deviceGuardPath -Name "EnableVirtualizationBasedSecurity" -ErrorAction SilentlyContinue).EnableVirtualizationBasedSecurity
    $beforePSF = (Get-ItemProperty -Path $deviceGuardPath -Name "RequirePlatformSecurityFeatures" -ErrorAction SilentlyContinue).RequirePlatformSecurityFeatures

    Set-ItemProperty -Path $deviceGuardPath -Name "EnableVirtualizationBasedSecurity" -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $deviceGuardPath -Name "RequirePlatformSecurityFeatures" -Value 1 -Type DWord -Force
    Add-Result -Section "Credential Protection" -Status "Success" -Action "Enable Credential Guard" -Before "VBS: $beforeVBS, PSF: $beforePSF" -After "VBS: 1, PSF: 1 (reboot required)"
} catch {
    Add-Result -Section "Credential Protection" -Status "Error" -Action "Enable Credential Guard" -Details $_.Exception.Message
}

try {
    $beforeLSA = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "RunAsPPL" -ErrorAction SilentlyContinue).RunAsPPL
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "RunAsPPL" -Value 1 -Type DWord -Force
    Add-Result -Section "Process Protection" -Status "Success" -Action "Enable LSA Protection" -Before $beforeLSA -After "1"
} catch {
    Add-Result -Section "Process Protection" -Status "Error" -Action "Enable LSA Protection" -Details $_.Exception.Message
}

# ------------------------------
# 10. Exploit Protection & ASR Rules
# ------------------------------
try {
    $asrRuleId = "D4F940AB-401B-4EFC-AADC-AD5F3C50688A"  # Block Office child processes
    Add-MpPreference -AttackSurfaceReductionRules_Ids $asrRuleId -AttackSurfaceReductionRules_Actions Enabled -ErrorAction Stop
    Add-Result -Section "Vulnerability Mitigation" -Status "Success" -Action "Enable ASR rule (Block Office child processes)" -Result "Applied"
} catch {
    Add-Result -Section "Vulnerability Mitigation" -Status "Warning" -Action "Enable ASR rule" -Details $_.Exception.Message
}

# ------------------------------
# 11. Disable Unnecessary Services
# ------------------------------
$servicesToDisable = @("Fax", "XblGameSave", "XblAuthManager", "TabletInputService", "WMPNetworkSvc", "WSearch", "WbioSrvc", "WerSvc", "WlanSvc", "WpcMonSvc", "WpnService")

foreach ($svc in $servicesToDisable) {
    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
        try {
            Stop-Service -Name $svc -Force -ErrorAction Stop
            Set-Service -Name $svc -StartupType Disabled -ErrorAction Stop
            $rollbackContent += "Set-Service -Name '$svc' -StartupType Manual -ErrorAction SilentlyContinue; Start-Service -Name '$svc' -ErrorAction SilentlyContinue"
            Add-Result -Section "Services" -Status "Success" -Action "Disable service $svc" -Result "Disabled"
        } catch {
            Add-Result -Section "Services" -Status "Warning" -Action "Disable service $svc" -Details $_.Exception.Message
        }
    } else {
        Add-Result -Section "Services" -Status "Skipped" -Action "Disable service $svc" -Details "Service not found"
    }
}

# ------------------------------
# 12. BitLocker (Data Encryption)
# ------------------------------
try {
    $tpm = Get-Tpm -ErrorAction Stop
    if ($tpm.TpmReady) {
        Enable-BitLocker -MountPoint "C:" -TpmProtector -ErrorAction Stop
        Add-Result -Section "Encryption" -Status "Success" -Action "Enable BitLocker" -Result "Enabled with TPM"
    } else {
        Add-Result -Section "Encryption" -Status "Skipped" -Action "Enable BitLocker" -Details "TPM not ready"
    }
} catch {
    Add-Result -Section "Encryption" -Status "Warning" -Action "Enable BitLocker" -Details $_.Exception.Message
}

# ------------------------------
# 13. Legal Notice (Access Notification)
# ------------------------------
try {
    $policyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    Set-ItemProperty -Path $policyPath -Name "legalnoticecaption" -Value "SOX Compliance Notice" -Force
    Set-ItemProperty -Path $policyPath -Name "legalnoticetext" -Value "This system is monitored and logged for financial compliance. Unauthorized access is prohibited." -Force
    Add-Result -Section "Access Notification" -Status "Success" -Action "Set legal notice" -Result "Applied"
} catch {
    Add-Result -Section "Access Notification" -Status "Error" -Action "Set legal notice" -Details $_.Exception.Message
}

# ------------------------------
# 14. UAC & Credential Guard Enhancements
# ------------------------------
try {
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value 1 -Type DWord -Force
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value 2 -Type DWord -Force  # Prompt for consent on secure desktop
    Add-Result -Section "Security" -Status "Success" -Action "Enable UAC" -Result "Enabled with secure desktop prompt"
} catch {
    Add-Result -Section "Security" -Status "Warning" -Action "Enable UAC" -Details $_.Exception.Message
}

# ------------------------------
# Generate Rollback Script
# ------------------------------
try {
    $rollbackContent | Out-File -FilePath $rollbackScript -Encoding UTF8 -Force
    Add-Result -Section "Rollback" -Status "Success" -Action "Create rollback script" -Details $rollbackScript
} catch {
    Add-Result -Section "Rollback" -Status "Error" -Action "Create rollback script" -Details $_.Exception.Message
}

# ------------------------------
# Generate Detailed HTML Report
# ------------------------------
$htmlHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SOX Compliance Hardening Report</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; margin: 40px; background: #f8f9fa; color: #2d3436; line-height: 1.6; }
        h1 { color: #2c3e50; border-bottom: 3px solid #0984e3; padding-bottom: 12px; }
        h2 { color: #0984e3; margin-top: 2.5em; }
        table { width: 100%; border-collapse: collapse; margin: 25px 0; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
        th, td { padding: 14px 18px; text-align: left; border-bottom: 1px solid #dfe6e9; }
        th { background: #0984e3; color: white; font-weight: 600; }
        tr:nth-child(even) { background: #f1f5f9; }
        .success { color: #27ae60; font-weight: bold; }
        .warning { color: #e67e22; font-weight: bold; }
        .error { color: #e74c3c; font-weight: bold; }
        .skipped { color: #7f8c8d; font-style: italic; }
        .info { color: #3498db; }
        .footer { margin-top: 50px; text-align: center; color: #636e72; font-size: 0.95em; }
        a { color: #0984e3; text-decoration: none; }
        a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <h1>SOX Compliance Hardening Report</h1>
    <p><strong>Hostname:</strong> $env:COMPUTERNAME</p>
    <p><strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <p><strong>Domain Joined:</strong> $domainJoined</p>
    <p><strong>Windows Version:</strong> $([System.Environment]::OSVersion.VersionString)</p>
    <hr>
    <h2>Hardening Results Summary</h2>
    <table>
        <thead>
            <tr>
                <th>Section</th>
                <th>Status</th>
                <th>Action / Setting</th>
                <th>Before</th>
                <th>After / Result</th>
                <th>Details / Error</th>
            </tr>
        </thead>
        <tbody>
"@

$htmlBody = ""
foreach ($r in $results) {
    $statusClass = switch ($r.Status) {
        "Success" { "success" }
        "Warning" { "warning" }
        "Error"   { "error" }
        "Skipped" { "skipped" }
        "Info"    { "info" }
        default   { "" }
    }
    $htmlBody += @"
            <tr class="$statusClass">
                <td>$($r.Section)</td>
                <td>$($r.Status)</td>
                <td>$($r.Action)</td>
                <td>$($r.Before)</td>
                <td>$($r.After)</td>
                <td>$($r.Details)</td>
            </tr>
"@
}

$htmlFooter = @"
        </tbody>
    </table>
    <div class="footer">
        <p>Generated by SOX Hardening Script | Full log: <a href="file:///$logPath">$logPath</a></p>
        <p>Rollback script: <a href="file:///$rollbackScript">$rollbackScript</a></p>
        <p>Review and test changes before production use.</p>
    </div>
</body>
</html>
"@

# Save HTML report
$htmlReport = $htmlHeader + $htmlBody + $htmlFooter
$htmlReport | Out-File -FilePath $reportFile -Encoding UTF8 -Force
Add-Result -Section "Reporting" -Status "Success" -Action "Generate HTML report" -Details $reportFile

# ------------------------------
# Final Cleanup & Output
# ------------------------------
try {
    Stop-Transcript -ErrorAction SilentlyContinue
} catch {
    Write-Warning "Failed to stop transcript: $($_.Exception.Message)"
}

Write-Host "`nHardening script completed at $(Get-Date)"
Write-Host "Detailed HTML Report saved to: $reportFile"
Write-Host "Open in any browser for formatted summary."
Write-Host "Full log: $logPath"
Write-Host "Rollback script: $rollbackScript"
