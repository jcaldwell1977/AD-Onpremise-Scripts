# =============================================================================
# VEEAM BACKUP & REPLICATION v13 - FULL AS-BUILT REPORT GENERATOR
# Produced by The Redesign Group - Technology & Cybersecurity Consulting
# https://redesign-group.com | Data Protection Practice
# Version 3.1 | Includes Backup Copy Job Audit + IP Resolution
# =============================================================================
# Usage:
#   pwsh -File Veeam-v13-AsBuilt-Redesign.ps1
#   pwsh -File Veeam-v13-AsBuilt-Redesign.ps1 -CustomerName "Isuzu" -HealthCheck
# =============================================================================

param(
    [string]$OutputPath   = "C:\AsBuiltReports",
    [string]$ReportName   = "Veeam-v13-AsBuilt-Report",
    [string]$CustomerName = "Customer",
    [string]$PreparedBy   = "The Redesign Group",
    [switch]$IncludeDiagrams,
    [switch]$HealthCheck
)

if ($PSVersionTable.PSVersion.Major -lt 7) { throw "Run in PowerShell 7 (pwsh.exe) for Veeam v13" }
Import-Module Veeam.Backup.PowerShell -ErrorAction SilentlyContinue

$ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm"

# Resolve a hostname to its IPv4 address
function Resolve-HostIP {
    param([string]$Hostname)
    if ([string]::IsNullOrWhiteSpace($Hostname)) { return 'N/A' }
    try {
        $r = [System.Net.Dns]::GetHostAddresses($Hostname) |
             Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
             Select-Object -First 1
        if ($r) { return $r.IPAddressToString } else { return 'Unresolvable' }
    } catch { return 'Unresolvable' }
}

# =============================================================================
# CSS - Redesign Group brand palette
# Dark background, teal/green accent, Syne + DM Mono typography
# =============================================================================
$CSS = @"
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&family=DM+Sans:wght@300;400;500;600&display=swap');
:root {
    --rg-black:#0a0a0a; --rg-dark:#111111; --rg-dark2:#181818; --rg-dark3:#222222;
    --rg-border:#2a2a2a; --rg-borderl:#333333;
    --rg-teal:#00c4a0; --rg-green:#4ade80; --rg-accent:#00c4a0;
    --rg-glow:rgba(0,196,160,0.14);
    --rg-text:#f0f0f0; --rg-text2:#a0a0a0; --rg-text3:#555555;
    --success:#4ade80; --warn:#fbbf24; --fail:#f87171; --info:#60a5fa; --veeam:#00b4d8;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'DM Sans',Arial,sans-serif;background:var(--rg-black);color:var(--rg-text);font-size:13px;line-height:1.65;-webkit-font-smoothing:antialiased;}
.accent-bar{height:3px;background:linear-gradient(90deg,#00c4a0 0%,#4ade80 100%);}
/* Header */
.rg-header{background:var(--rg-dark);border-bottom:1px solid var(--rg-border);padding:44px 60px 36px;position:relative;overflow:hidden;}
.rg-header::before{content:'';position:absolute;inset:0;background:radial-gradient(ellipse 55% 80% at 85% 50%,rgba(0,196,160,0.07) 0%,transparent 70%),radial-gradient(ellipse 35% 60% at 10% 30%,rgba(74,222,128,0.04) 0%,transparent 60%);pointer-events:none;}
.rg-header::after{content:'';position:absolute;inset:0;background-image:radial-gradient(circle,rgba(255,255,255,0.035) 1px,transparent 1px);background-size:24px 24px;pointer-events:none;}
.header-inner{position:relative;z-index:1;display:flex;justify-content:space-between;align-items:flex-start;gap:32px;flex-wrap:wrap;}
.header-left{display:flex;flex-direction:column;gap:18px;}
.rg-wordmark{display:flex;flex-direction:column;gap:2px;}
.rg-wordmark .mark{font-family:'Syne',Arial,sans-serif;font-weight:800;font-size:13px;letter-spacing:3px;text-transform:uppercase;color:var(--rg-accent);}
.rg-wordmark .sub{font-family:'DM Mono',monospace;font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--rg-text2);}
.chip-row{display:flex;gap:8px;flex-wrap:wrap;}
.chip{font-family:'DM Mono',monospace;font-size:9px;letter-spacing:1.5px;text-transform:uppercase;padding:3px 10px;border-radius:2px;border:1px solid;}
.chip-v{color:var(--veeam);border-color:rgba(0,180,216,0.4);background:rgba(0,180,216,0.06);}
.chip-rg{color:var(--rg-accent);border-color:rgba(0,196,160,0.4);background:rgba(0,196,160,0.06);}
.chip-ab{color:var(--rg-text2);border-color:var(--rg-borderl);background:transparent;}
.report-title{font-family:'Syne',Arial,sans-serif;font-weight:700;font-size:26px;letter-spacing:-0.3px;color:var(--rg-text);line-height:1.15;}
.header-meta{font-family:'DM Mono',monospace;font-size:11px;color:var(--rg-text2);display:flex;flex-wrap:wrap;gap:6px 20px;}
.header-meta span::before{content:'';display:inline-block;width:4px;height:4px;border-radius:50%;background:var(--rg-accent);margin-right:6px;vertical-align:middle;}
.header-right{display:flex;flex-direction:column;gap:6px;align-items:flex-end;text-align:right;}
.cust-label{font-family:'DM Mono',monospace;font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--rg-text3);}
.cust-name{font-family:'Syne',Arial,sans-serif;font-size:20px;font-weight:700;color:var(--rg-text);}
.prep-by{font-family:'DM Mono',monospace;font-size:10px;color:var(--rg-text2);}
/* Layout */
.layout{display:grid;grid-template-columns:220px 1fr;min-height:calc(100vh - 200px);}
/* TOC */
.toc{background:var(--rg-dark);border-right:1px solid var(--rg-border);padding:24px 0;position:sticky;top:0;height:100vh;overflow-y:auto;}
.toc-title{padding:0 18px 14px;border-bottom:1px solid var(--rg-border);margin-bottom:10px;font-family:'DM Mono',monospace;font-size:9px;letter-spacing:2.5px;text-transform:uppercase;color:var(--rg-text3);}
.toc a{display:flex;align-items:center;gap:10px;color:var(--rg-text2);text-decoration:none;padding:7px 18px;font-size:12px;border-left:2px solid transparent;transition:all 0.15s;}
.toc a:hover{color:var(--rg-text);background:rgba(255,255,255,0.03);border-left-color:var(--rg-borderl);}
.toc a.active{color:var(--rg-accent);background:var(--rg-glow);border-left-color:var(--rg-accent);}
.toc a .tn{font-family:'DM Mono',monospace;font-size:9px;color:var(--rg-text3);width:18px;flex-shrink:0;}
.toc a.active .tn{color:var(--rg-accent);}
/* Main */
.main-content{padding:36px 52px;}
/* Section */
.section{margin-bottom:40px;scroll-margin-top:20px;background:var(--rg-dark2);border:1px solid var(--rg-border);border-radius:6px;overflow:hidden;}
.section-head{display:flex;align-items:center;gap:12px;padding:14px 20px;background:var(--rg-dark3);border-bottom:1px solid var(--rg-border);}
.sec-num{font-family:'DM Mono',monospace;font-size:9px;color:var(--rg-accent);background:var(--rg-glow);border:1px solid rgba(0,196,160,0.25);width:26px;height:26px;border-radius:4px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.section-head h2{font-family:'Syne',Arial,sans-serif;font-size:16px;font-weight:700;color:var(--rg-text);}
.section-body{padding:20px;}
.subsection{margin-bottom:22px;}
.subsection h3{font-family:'DM Sans',Arial,sans-serif;font-size:10px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;color:var(--rg-text2);margin-bottom:10px;display:flex;align-items:center;gap:10px;}
.subsection h3::after{content:'';flex:1;height:1px;background:var(--rg-border);}
/* Tables */
.tw{overflow-x:auto;border-radius:4px;border:1px solid var(--rg-border);}
table{width:100%;border-collapse:collapse;font-size:12px;font-family:'DM Mono',monospace;background:var(--rg-dark2);}
th{background:var(--rg-dark3);color:var(--rg-text3);font-size:9px;letter-spacing:1.5px;text-transform:uppercase;padding:9px 12px;text-align:left;border-bottom:1px solid var(--rg-border);white-space:nowrap;}
td{padding:8px 12px;border-bottom:1px solid var(--rg-border);color:var(--rg-text);vertical-align:top;word-break:break-word;max-width:360px;}
tr:last-child td{border-bottom:none;}
tr:hover td{background:rgba(255,255,255,0.015);}
td.ip{color:var(--rg-accent);font-weight:500;}
/* Summary cards */
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(138px,1fr));gap:12px;margin-bottom:28px;}
.card{background:var(--rg-dark2);border:1px solid var(--rg-border);border-radius:6px;padding:18px 14px 14px;position:relative;overflow:hidden;}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,#00c4a0,#4ade80);opacity:0.7;}
.card .val{font-family:'Syne',Arial,sans-serif;font-size:32px;font-weight:800;color:var(--rg-text);line-height:1;}
.card .lbl{font-family:'DM Mono',monospace;font-size:9px;letter-spacing:1.5px;text-transform:uppercase;color:var(--rg-text3);margin-top:7px;}
/* Badges */
.badge{display:inline-block;font-family:'DM Mono',monospace;font-size:9px;padding:2px 8px;border-radius:2px;letter-spacing:0.5px;text-transform:uppercase;border:1px solid;}
.b-ok{background:rgba(74,222,128,0.08);color:#4ade80;border-color:rgba(74,222,128,0.3);}
.b-warn{background:rgba(251,191,36,0.08);color:#fbbf24;border-color:rgba(251,191,36,0.3);}
.b-fail{background:rgba(248,113,113,0.08);color:#f87171;border-color:rgba(248,113,113,0.3);}
.b-neu{background:rgba(96,96,96,0.12);color:#888;border-color:#333;}
/* Health banner */
.hbanner{background:rgba(248,113,113,0.06);border:1px solid rgba(248,113,113,0.25);border-left:3px solid #f87171;border-radius:4px;padding:16px 20px;margin-bottom:20px;}
.hbanner h3{color:#f87171;font-family:'DM Mono',monospace;font-size:10px;letter-spacing:2px;text-transform:uppercase;margin-bottom:10px;}
.hbanner ul{padding-left:18px;}
.hbanner li{color:#f87171;font-size:12px;margin-bottom:4px;font-family:'DM Mono',monospace;}
/* Copy job audit info banner */
.abanner{background:rgba(0,196,160,0.06);border:1px solid rgba(0,196,160,0.2);border-left:3px solid var(--rg-accent);border-radius:4px;padding:12px 16px;margin-bottom:14px;font-family:'DM Mono',monospace;font-size:11px;color:var(--rg-text2);}
.abanner strong{color:var(--rg-accent);}
/* Job cards */
.job-card{background:var(--rg-dark2);border:1px solid var(--rg-border);border-radius:5px;margin-bottom:10px;overflow:hidden;}
.job-card-head{display:flex;align-items:center;justify-content:space-between;background:var(--rg-dark3);padding:11px 16px;cursor:pointer;user-select:none;gap:12px;}
.job-card-head:hover{background:rgba(255,255,255,0.025);}
.job-card-title{font-family:'DM Sans',Arial,sans-serif;font-weight:600;font-size:13px;color:var(--rg-text);}
.job-card-body{padding:16px;display:none;}
.job-card.open .job-card-body{display:block;}
.jmeta{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:12px;font-family:'DM Mono',monospace;font-size:11px;color:var(--rg-text2);}
.jmeta span::before{content:'';display:inline-block;width:3px;height:3px;border-radius:50%;background:var(--rg-accent);margin-right:5px;vertical-align:middle;}
.ilabel{font-family:'DM Mono',monospace;font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--rg-text3);margin:14px 0 8px;}
/* Audit trail table */
.audit-wrap{border:1px solid var(--rg-border);border-left:3px solid var(--rg-accent);border-radius:4px;overflow:hidden;}
/* Misc */
.empty{color:var(--rg-text3);font-style:italic;font-size:12px;padding:8px 0;font-family:'DM Mono',monospace;}
.err{color:#f87171;font-size:12px;font-family:'DM Mono',monospace;padding:6px 0;}
pre{font-family:'DM Mono',monospace;font-size:11px;color:var(--rg-text2);white-space:pre-wrap;}
/* Footer */
.rg-footer{background:var(--rg-dark);border-top:1px solid var(--rg-border);padding:18px 52px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;font-family:'DM Mono',monospace;font-size:10px;color:var(--rg-text3);}
.rg-footer .fb strong{color:var(--rg-accent);font-size:11px;display:block;margin-bottom:2px;}
.rg-footer .fc{text-align:center;align-self:center;}
.rg-footer .fr{text-align:right;align-self:center;}
/* Print */
@media print{
    body{background:#fff;color:#000;}
    .toc{display:none;}
    .layout{display:block;}
    .main-content{padding:0;}
    .job-card-body{display:block !important;}
    th{background:#ddd;color:#000;}
    td{color:#111;}
    .section{border:1px solid #ccc;background:#fff;}
    .section-head{background:#eee;}
}
"@

# =============================================================================
# JavaScript - collapsible cards, sticky TOC, table wrapping, IP highlighting
# =============================================================================
$JS = @"
document.querySelectorAll('.job-card-head').forEach(h => {
    h.addEventListener('click', () => h.parentElement.classList.toggle('open'));
});
const secs = document.querySelectorAll('.section');
const links = document.querySelectorAll('.toc a');
const obs = new IntersectionObserver(entries => {
    entries.forEach(e => {
        if (e.isIntersecting) {
            links.forEach(l => l.classList.remove('active'));
            const l = document.querySelector('.toc a[href="#' + e.target.id + '"]');
            if (l) l.classList.add('active');
        }
    });
}, { rootMargin: '-20% 0px -70% 0px' });
secs.forEach(s => obs.observe(s));
document.querySelectorAll('table').forEach(t => {
    if (!t.parentElement.classList.contains('tw') && !t.parentElement.classList.contains('audit-wrap')) {
        const w = document.createElement('div');
        w.className = 'tw';
        t.parentNode.insertBefore(w, t);
        w.appendChild(t);
    }
});
document.querySelectorAll('td').forEach(td => {
    if (/^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(td.textContent.trim())) {
        td.classList.add('ip');
    }
});
"@

# =============================================================================
# DATA COLLECTION
# =============================================================================

Write-Host "Collecting Veeam configuration data..." -ForegroundColor Cyan

$BkpServer   = $null; try { $BkpServer   = Get-VBRServer | Where-Object { $_.Type -eq 'Local' } | Select-Object -First 1 } catch { }
$License     = $null; try { $License     = Get-VBRInstalledLicense } catch { }
$AllJobs     = @();   try { $AllJobs     = @(Get-VBRJob) } catch { }
$BackupJobs  = $AllJobs | Where-Object { $_.JobType -eq 'Backup' }
$CopyJobs    = $AllJobs | Where-Object { $_.JobType -eq 'BackupSync' }
$ReplJobs    = $AllJobs | Where-Object { $_.JobType -in @('Replica','SimpleTransactionLog') }
$NasJobs     = @(); try { $NasJobs     = @(Get-VBRNASBackupJob) }         catch { }
$TapeJobs    = @(); try { $TapeJobs    = @(Get-VBRTapeJob) }              catch { }
$ViProxies   = @(); try { $ViProxies   = @(Get-VBRViProxy) }             catch { }
$HvProxies   = @(); try { $HvProxies   = @(Get-VBRHvProxy) }             catch { }
$NasProxies  = @(); try { $NasProxies  = @(Get-VBRNASProxyServer) }      catch { }
$Repos       = @(); try { $Repos       = @(Get-VBRBackupRepository) }    catch { }
$SOBRs       = @(); try { $SOBRs       = @(Get-VBRBackupRepository -ScaleOut) } catch { }
$ExtRepos    = @(); try { $ExtRepos    = @(Get-VBRExternalRepository) }  catch { }
$CDPPolicies = @(); try { $CDPPolicies = @(Get-VBRCDPPolicy) }           catch { }
$SBJobs      = @(); try { $SBJobs      = @(Get-VBRSureBackupJob) }       catch { }
$VLabs       = @(); try { $VLabs       = @(Get-VBRVirtualLab) }          catch { }
$AppGroups   = @(); try { $AppGroups   = @(Get-VBRApplicationGroup) }    catch { }
$TapeLibs    = @(); try { $TapeLibs    = @(Get-VBRTapeLibrary) }         catch { }
$TapePools   = @(); try { $TapePools   = @(Get-VBRTapeMediaPool) }       catch { }
$CloudTenants= @(); try { $CloudTenants= @(Get-VBRCloudTenant) }         catch { }
$ViServers   = @(); try { $ViServers   = @(Get-VBRServer | Where-Object { $_.Type -in @('VC','ESXi') }) } catch { }
$HvServers   = @(); try { $HvServers   = @(Get-VBRServer | Where-Object { $_.Type -eq 'HvServer' }) }    catch { }
$FileShares  = @(); try { $FileShares  = @(Get-VBRNASFileShare) }        catch { }
$Credentials = @(); try { $Credentials = @(Get-VBRCredentials | Select-Object Name, Description) } catch { }
$NotifOpts   = $null; try { $NotifOpts = Get-VBRNotificationOptions }    catch { }
$EmailOpts   = $null; try { $EmailOpts = Get-VBREmailOptions }           catch { }

# Build backup copy job audit records
$CopyJobAudit = $CopyJobs | ForEach-Object {
    $j = $_
    $last = $null
    try { $last = Get-VBRJobSession -Job $j -Last 1 -ErrorAction SilentlyContinue } catch { }
    $srcRepo = try { ($j.GetSourceRepository()).Name } catch { 'N/A' }
    $tgtRepo = try { $j.GetTargetRepository().Name   } catch { 'N/A' }
    [PSCustomObject]@{
        Name          = $j.Name
        Enabled       = $j.IsScheduleEnabled
        'Source Repo' = $srcRepo
        'Target Repo' = $tgtRepo
        'Last Result' = if ($last) { $last.Result } else { 'No Sessions' }
        'Last Run'    = if ($last) { $last.EndTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
        'Xfer GB'     = if ($last) { [math]::Round($last.Progress.TransferedSize / 1GB, 2) } else { 0 }
        Description   = $j.Description
    }
}

# Health check warnings
$Warnings = [System.Collections.Generic.List[string]]::new()
if ($HealthCheck) {
    foreach ($j in $AllJobs) {
        if ($j.IsScheduleEnabled) {
            try {
                $s = Get-VBRJobSession -Job $j -Last 1 -ErrorAction SilentlyContinue
                if ($s -and $s.Result -eq 'Failed') {
                    $Warnings.Add("Job '$($j.Name)' last run FAILED ($($s.EndTime.ToString('yyyy-MM-dd HH:mm'))).")
                }
            } catch { }
        }
    }
    foreach ($r in $Repos) {
        if ($r.TotalSpace -gt 0) {
            $pct = [math]::Round(($r.FreeSpace / $r.TotalSpace) * 100, 1)
            if ($pct -lt 10)  { $Warnings.Add("Repo '$($r.Name)' critically low: $pct% free.") }
            elseif ($pct -lt 20) { $Warnings.Add("Repo '$($r.Name)' below 20% free ($pct%).") }
        }
    }
    foreach ($ca in $CopyJobAudit) {
        if ($ca.'Last Result' -eq 'Failed') { $Warnings.Add("Backup Copy Job '$($ca.Name)' last run FAILED.") }
    }
    if ($License -and $License.ExpirationDate -lt (Get-Date).AddDays(30)) {
        $Warnings.Add("License expires soon: $($License.ExpirationDate.ToString('yyyy-MM-dd'))")
    }
}

$totalJobs  = $AllJobs.Count + $NasJobs.Count + $TapeJobs.Count
$enabledJobs= ($AllJobs | Where-Object { $_.IsScheduleEnabled }).Count
$totalRepos = $Repos.Count + $SOBRs.Count
$totalProx  = $ViProxies.Count + $HvProxies.Count + $NasProxies.Count

Write-Host "Data collected. Building HTML report..." -ForegroundColor Cyan

# =============================================================================
# HTML GENERATION
# =============================================================================

$sb = [System.Text.StringBuilder]::new()

[void]$sb.AppendLine("<!DOCTYPE html><html lang='en'><head>")
[void]$sb.AppendLine("<meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1.0'>")
[void]$sb.AppendLine("<title>$ReportName - $CustomerName</title>")
[void]$sb.AppendLine("<style>$CSS</style></head><body>")

# Accent bar + Header
[void]$sb.AppendLine("<div class='accent-bar'></div>")
[void]$sb.AppendLine("<div class='rg-header'><div class='header-inner'>")
[void]$sb.AppendLine("  <div class='header-left'>")
[void]$sb.AppendLine("    <div class='rg-wordmark'><span class='mark'>The Redesign Group</span><span class='sub'>Technology &amp; Cybersecurity Consulting</span></div>")
[void]$sb.AppendLine("    <div class='chip-row'><span class='chip chip-v'>Veeam B&amp;R v13</span><span class='chip chip-rg'>Data Protection</span><span class='chip chip-ab'>As-Built Report</span></div>")
[void]$sb.AppendLine("    <div class='report-title'>$ReportName</div>")
[void]$sb.AppendLine("    <div class='header-meta'><span>Generated: $ReportDate</span><span>Server: $($BkpServer.Name ?? 'Unknown')</span></div>")
[void]$sb.AppendLine("  </div>")
[void]$sb.AppendLine("  <div class='header-right'>")
[void]$sb.AppendLine("    <span class='cust-label'>Prepared for</span>")
[void]$sb.AppendLine("    <span class='cust-name'>$CustomerName</span>")
[void]$sb.AppendLine("    <span class='prep-by'>Prepared by: $PreparedBy</span>")
[void]$sb.AppendLine("  </div>")
[void]$sb.AppendLine("</div></div>")

# Layout + TOC
[void]$sb.AppendLine("<div class='layout'>")
[void]$sb.AppendLine("<nav class='toc'>")
[void]$sb.AppendLine("  <div class='toc-title'>Contents</div>")
[void]$sb.AppendLine("  <a href='#sec-dash'><span class='tn'>*</span>Dashboard</a>")
[void]$sb.AppendLine("  <a href='#sec-server'><span class='tn'>01</span>Backup Server</a>")
[void]$sb.AppendLine("  <a href='#sec-license'><span class='tn'>02</span>License</a>")
[void]$sb.AppendLine("  <a href='#sec-infra'><span class='tn'>03</span>Infrastructure</a>")
[void]$sb.AppendLine("  <a href='#sec-repos'><span class='tn'>04</span>Repositories</a>")
[void]$sb.AppendLine("  <a href='#sec-jobs'><span class='tn'>05</span>Backup Jobs</a>")
[void]$sb.AppendLine("  <a href='#sec-copy'><span class='tn'>06</span>Backup Copy Jobs</a>")
[void]$sb.AppendLine("  <a href='#sec-nas'><span class='tn'>07</span>NAS Backup</a>")
[void]$sb.AppendLine("  <a href='#sec-repl'><span class='tn'>08</span>Replication &amp; CDP</a>")
[void]$sb.AppendLine("  <a href='#sec-sb'><span class='tn'>09</span>SureBackup</a>")
[void]$sb.AppendLine("  <a href='#sec-tape'><span class='tn'>10</span>Tape</a>")
[void]$sb.AppendLine("  <a href='#sec-cloud'><span class='tn'>11</span>Cloud Connect</a>")
[void]$sb.AppendLine("  <a href='#sec-inv'><span class='tn'>12</span>Inventory</a>")
[void]$sb.AppendLine("  <a href='#sec-global'><span class='tn'>13</span>Global Settings</a>")
[void]$sb.AppendLine("</nav>")
[void]$sb.AppendLine("<div class='main-content'>")

# Health warnings
if ($HealthCheck -and $Warnings.Count -gt 0) {
    [void]$sb.AppendLine("<div class='hbanner'><h3>Health Check Warnings</h3><ul>")
    foreach ($w in $Warnings) { [void]$sb.AppendLine("<li>$w</li>") }
    [void]$sb.AppendLine("</ul></div>")
}

# ---- Dashboard ----
[void]$sb.AppendLine("<div class='section' id='sec-dash'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>*</span><h2>Summary Dashboard</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'><div class='cards'>")
@(
    @{V=$totalJobs;                               L="Total Jobs"},
    @{V=$enabledJobs;                             L="Scheduled"},
    @{V=$BackupJobs.Count;                        L="Backup Jobs"},
    @{V=$CopyJobs.Count;                          L="Copy Jobs"},
    @{V=$totalRepos;                              L="Repositories"},
    @{V=$totalProx;                               L="Proxies"},
    @{V=($ViServers.Count + $HvServers.Count);    L="Hypervisors"},
    @{V=$TapeLibs.Count;                          L="Tape Libraries"}
) | ForEach-Object {
    [void]$sb.AppendLine("<div class='card'><div class='val'>$($_.V)</div><div class='lbl'>$($_.L)</div></div>")
}
[void]$sb.AppendLine("  </div></div></div>")

# ---- Section 01: Backup Server & Global Settings ----
[void]$sb.AppendLine("<div class='section' id='sec-server'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>01</span><h2>Backup Server &amp; Global Settings</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")

$srvTable = Get-VBRServer | Where-Object { $_.Type -eq 'Local' } | ForEach-Object {
    [PSCustomObject]@{
        Name        = $_.Name
        'IP Address'= Resolve-HostIP $_.Name
        Version     = $_.Version
        DNS         = $_.DNSName
        Connected   = $_.IsConnected
    }
}
[void]$sb.AppendLine(($srvTable | ConvertTo-Html -Fragment))

[void]$sb.AppendLine("<div class='subsection' style='margin-top:20px'><h3>Global Options</h3>")
[void]$sb.AppendLine("<table><tr><th>Category</th><th>Key Settings</th></tr>")
$gSecOpts  = $null; try { $gSecOpts  = Get-VBRSecurityOptions }    catch { }
$gNotif    = $null; try { $gNotif    = Get-VBRNotificationOptions } catch { }
$gEmail    = $null; try { $gEmail    = Get-VBREmailOptions }        catch { }
$gSNMP     = $null; try { $gSNMP     = Get-VBRSNMPService }         catch { }
$gHistory  = $null; try { $gHistory  = Get-VBRJobHistoryOptions }   catch { }
$globalCats = [ordered]@{
    "Security and Encryption" = $gSecOpts
    "Notifications"           = $gNotif
    "Email"                   = $gEmail
    "SNMP"                    = $gSNMP
    "Job History"             = $gHistory
}
foreach ($cat in $globalCats.Keys) {
    $val = if ($globalCats[$cat]) { ($globalCats[$cat] | Out-String).Trim() } else { 'N/A' }
    [void]$sb.AppendLine("<tr><td>$cat</td><td><pre>$val</pre></td></tr>")
}
[void]$sb.AppendLine("</table></div>")
[void]$sb.AppendLine("  </div></div>")

# ---- Section 02: License ----
[void]$sb.AppendLine("<div class='section' id='sec-license'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>02</span><h2>License</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")
if ($License) {
    $licData = [PSCustomObject]@{
        Edition           = $License.Edition
        Type              = $License.LicenseType
        Status            = $License.Status
        'Expiration Date' = $License.ExpirationDate.ToString('yyyy-MM-dd')
        'Support Expiry'  = $License.SupportExpirationDate.ToString('yyyy-MM-dd')
        'Used / Total'    = "$($License.UsedLicensesNumber) / $($License.TotalLicensesNumber)"
        'Used Sockets'    = $License.UsedSocketsNumber
    }
    [void]$sb.AppendLine(($licData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='err'>Could not retrieve license information.</p>") }
[void]$sb.AppendLine("  </div></div>")

# ---- Section 03: Infrastructure ----
[void]$sb.AppendLine("<div class='section' id='sec-infra'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>03</span><h2>Backup Infrastructure</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")

[void]$sb.AppendLine("<div class='subsection'><h3>VMware Backup Proxies</h3>")
if ($ViProxies.Count -gt 0) {
    $vpData = $ViProxies | ForEach-Object {
        [PSCustomObject]@{
            Name            = $_.Name
            Host            = $_.Host.Name
            'IP Address'    = Resolve-HostIP $_.Host.Name
            'Transport Mode'= $_.Options.TransportMode
            'Max Tasks'     = $_.MaxTasksCount
            Status          = if ($_.IsDisabled) { 'Disabled' } else { 'Enabled' }
        }
    }
    [void]$sb.AppendLine(($vpData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No VMware proxies configured.</p>") }
[void]$sb.AppendLine("</div>")

[void]$sb.AppendLine("<div class='subsection'><h3>Hyper-V Off-Host Proxies</h3>")
if ($HvProxies.Count -gt 0) {
    $hvpData = $HvProxies | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Host        = $_.Host.Name
            'IP Address'= Resolve-HostIP $_.Host.Name
            'Max Tasks' = $_.MaxTasksCount
        }
    }
    [void]$sb.AppendLine(($hvpData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No Hyper-V proxies configured.</p>") }
[void]$sb.AppendLine("</div>")

[void]$sb.AppendLine("<div class='subsection'><h3>NAS Backup Proxies</h3>")
if ($NasProxies.Count -gt 0) {
    $npData = $NasProxies | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Server      = $_.Server.Name
            'IP Address'= Resolve-HostIP $_.Server.Name
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($npData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No NAS proxies configured.</p>") }
[void]$sb.AppendLine("</div>")

[void]$sb.AppendLine("  </div></div>")

# ---- Section 04: Repositories ----
[void]$sb.AppendLine("<div class='section' id='sec-repos'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>04</span><h2>Repositories &amp; Storage</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")

[void]$sb.AppendLine("<div class='subsection'><h3>Backup Repositories</h3>")
if ($Repos.Count -gt 0) {
    $repoData = $Repos | ForEach-Object {
        $freeGB  = [math]::Round($_.FreeSpace  / 1GB, 1)
        $totalGB = [math]::Round($_.TotalSpace / 1GB, 1)
        $pct     = if ($totalGB -gt 0) { [math]::Round(($freeGB / $totalGB) * 100, 1) } else { 'N/A' }
        [PSCustomObject]@{
            Name         = $_.Name
            Type         = $_.Type
            Host         = $_.Host.Name
            'IP Address' = Resolve-HostIP $_.Host.Name
            Path         = $_.Path
            'Total GB'   = $totalGB
            'Free GB'    = $freeGB
            'Free %'     = $pct
            'Per-VM'     = $_.UsePerVMBackupFiles
            Immutability = $_.ImmutabilityEnabled
        }
    }
    [void]$sb.AppendLine(($repoData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No repositories found.</p>") }
[void]$sb.AppendLine("</div>")

[void]$sb.AppendLine("<div class='subsection'><h3>Scale-Out Backup Repositories (SOBR)</h3>")
if ($SOBRs.Count -gt 0) {
    $sobrData = $SOBRs | ForEach-Object {
        [PSCustomObject]@{
            Name            = $_.Name
            Policy          = $_.PolicyType
            'Active Extents'= ($_.Extents | Where-Object { $_.IsActive } | Measure-Object).Count
            'Capacity Tier' = $_.CapacityTier.Enabled
            'Archive Tier'  = $_.ArchiveTier.Enabled
        }
    }
    [void]$sb.AppendLine(($sobrData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No SOBRs configured.</p>") }
[void]$sb.AppendLine("</div>")

[void]$sb.AppendLine("<div class='subsection'><h3>External / Object Storage</h3>")
if ($ExtRepos.Count -gt 0) {
    [void]$sb.AppendLine(($ExtRepos | Select-Object Name, Type, Description | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No external repositories configured.</p>") }
[void]$sb.AppendLine("</div>")

[void]$sb.AppendLine("  </div></div>")

# ---- Section 05: Backup Jobs ----
[void]$sb.AppendLine("<div class='section' id='sec-jobs'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>05</span><h2>Backup Jobs</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")

if ($BackupJobs.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Overview</h3>")
    $jOverview = $BackupJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            Repository  = try { $_.GetTargetRepository().Name } catch { 'N/A' }
            Objects     = try { ($_.GetObjectsInJob() | Measure-Object).Count } catch { 0 }
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($jOverview | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")

    [void]$sb.AppendLine("<div class='subsection'><h3>Detailed Configuration - click to expand</h3>")
    foreach ($job in $BackupJobs) {
        $opts  = $null; try { $opts  = $job.GetOptions()         } catch { }
        $sched = $null; try { $sched = $job.ScheduleOptions      } catch { }
        $objs  = @();   try { $objs  = @($job.GetObjectsInJob()) } catch { }
        $badge = if ($job.IsScheduleEnabled) { "<span class='badge b-ok'>Enabled</span>" } else { "<span class='badge b-neu'>Disabled</span>" }

        [void]$sb.AppendLine("<div class='job-card'>")
        [void]$sb.AppendLine("  <div class='job-card-head'><span class='job-card-title'>$($job.Name)</span>$badge</div>")
        [void]$sb.AppendLine("  <div class='job-card-body'>")
        [void]$sb.AppendLine("    <div class='jmeta'><span>Type: $($job.JobType)</span><span>Repo: $(try{$job.GetTargetRepository().Name}catch{'N/A'})</span><span>Objects: $($objs.Count)</span></div>")

        if ($objs.Count -gt 0) {
            [void]$sb.AppendLine("<div class='ilabel'>Source Objects</div>")
            $od = $objs | ForEach-Object { [PSCustomObject]@{ Name=$_.Name; Type=$_.Type; Location=$_.Location } }
            [void]$sb.AppendLine(($od | ConvertTo-Html -Fragment))
        }
        if ($opts) {
            [void]$sb.AppendLine("<div class='ilabel'>Retention and Storage</div>")
            $rd = [PSCustomObject]@{
                'Retention Type' = $opts.RetentionPolicy.Type
                'Restore Points' = $opts.RetentionPolicy.Quantity
                'GFS Enabled'   = $opts.RetentionPolicy.IsGFSEnabled
                Dedup            = $opts.JobOptions.EnableDeduplication
                Compression      = $opts.JobOptions.CompressionType
                Encryption       = $opts.JobOptions.EncryptionEnabled
            }
            [void]$sb.AppendLine(($rd | ConvertTo-Html -Fragment))
        }
        if ($sched) {
            [void]$sb.AppendLine("<div class='ilabel'>Schedule</div>")
            $sd = [PSCustomObject]@{
                Enabled       = $sched.Enabled
                Type          = $sched.Type
                'Starts At'   = $sched.StartDateTime
                'Retry Count' = $sched.RetryCount
                'Retry Wait'  = "$($sched.RetryTimeout) min"
            }
            [void]$sb.AppendLine(($sd | ConvertTo-Html -Fragment))
        }
        [void]$sb.AppendLine("  </div></div>")
    }
    [void]$sb.AppendLine("</div>")
} else { [void]$sb.AppendLine("<p class='empty'>No backup jobs found.</p>") }

[void]$sb.AppendLine("  </div></div>")

# ---- Section 06: Backup Copy Jobs - FULL AUDIT ----
[void]$sb.AppendLine("<div class='section' id='sec-copy'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>06</span><h2>Backup Copy Jobs - Audit</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")

if ($CopyJobs.Count -gt 0) {

    # Alert for any failed copy jobs
    $failedCopy = $CopyJobAudit | Where-Object { $_.'Last Result' -eq 'Failed' }
    if ($failedCopy.Count -gt 0) {
        [void]$sb.AppendLine("<div class='hbanner'><h3>Failed Backup Copy Jobs</h3><ul>")
        foreach ($fc in $failedCopy) {
            [void]$sb.AppendLine("<li>$($fc.Name) - Last ran: $($fc.'Last Run')</li>")
        }
        [void]$sb.AppendLine("</ul></div>")
    }

    [void]$sb.AppendLine("<div class='subsection'><h3>Copy Job Summary</h3>")
    [void]$sb.AppendLine("<div class='abanner'><strong>Audit scope:</strong> Source repo, target repo, last session result, GB transferred, and 5-session history per job.</div>")
    [void]$sb.AppendLine(($CopyJobAudit | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")

    [void]$sb.AppendLine("<div class='subsection'><h3>Detailed Configuration and Session Audit - click to expand</h3>")
    foreach ($job in $CopyJobs) {
        $opts     = $null; try { $opts     = $job.GetOptions()    } catch { }
        $sched    = $null; try { $sched    = $job.ScheduleOptions } catch { }
        $sessions = @();   try { $sessions = @(Get-VBRJobSession -Job $job -Last 5 -ErrorAction SilentlyContinue) } catch { }

        $badge   = if ($job.IsScheduleEnabled) { "<span class='badge b-ok'>Enabled</span>" } else { "<span class='badge b-neu'>Disabled</span>" }
        $srcRepo = try { ($job.GetSourceRepository()).Name } catch { 'N/A' }
        $tgtRepo = try { $job.GetTargetRepository().Name   } catch { 'N/A' }

        [void]$sb.AppendLine("<div class='job-card'>")
        [void]$sb.AppendLine("  <div class='job-card-head'><span class='job-card-title'>$($job.Name)</span>$badge</div>")
        [void]$sb.AppendLine("  <div class='job-card-body'>")
        [void]$sb.AppendLine("    <div class='jmeta'><span>Type: Backup Copy</span><span>Source: $srcRepo</span><span>Target: $tgtRepo</span></div>")

        if ($opts) {
            [void]$sb.AppendLine("<div class='ilabel'>Retention and Storage</div>")
            $rd = [PSCustomObject]@{
                'Restore Points' = $opts.RetentionPolicy.Quantity
                'GFS Enabled'    = $opts.RetentionPolicy.IsGFSEnabled
                'Weekly GFS'     = $opts.RetentionPolicy.WeeklyFullSchedule.RepeatCount
                'Monthly GFS'    = $opts.RetentionPolicy.MonthlyFullSchedule.RepeatCount
                'Yearly GFS'     = $opts.RetentionPolicy.YearlyFullSchedule.RepeatCount
                Encryption       = $opts.JobOptions.EncryptionEnabled
            }
            [void]$sb.AppendLine(($rd | ConvertTo-Html -Fragment))
        }

        if ($sched) {
            [void]$sb.AppendLine("<div class='ilabel'>Schedule</div>")
            $sd = [PSCustomObject]@{
                Enabled       = $sched.Enabled
                Type          = $sched.Type
                'Retry Count' = $sched.RetryCount
                'Retry Wait'  = "$($sched.RetryTimeout) min"
            }
            [void]$sb.AppendLine(($sd | ConvertTo-Html -Fragment))
        }

        [void]$sb.AppendLine("<div class='ilabel'>Last 5 Sessions - Audit Trail</div>")
        if ($sessions.Count -gt 0) {
            [void]$sb.AppendLine("<div class='audit-wrap'>")
            $sessData = $sessions | ForEach-Object {
                [PSCustomObject]@{
                    Start         = $_.CreationTime.ToString('yyyy-MM-dd HH:mm')
                    End           = $_.EndTime.ToString('yyyy-MM-dd HH:mm')
                    Result        = $_.Result
                    'Duration min'= [math]::Round(($_.EndTime - $_.CreationTime).TotalMinutes, 1)
                    'Xfer GB'     = [math]::Round($_.Progress.TransferedSize / 1GB, 3)
                    State         = $_.State
                }
            }
            [void]$sb.AppendLine(($sessData | ConvertTo-Html -Fragment))
            [void]$sb.AppendLine("</div>")
        } else {
            [void]$sb.AppendLine("<p class='empty'>No session history found.</p>")
        }

        [void]$sb.AppendLine("  </div></div>")
    }
    [void]$sb.AppendLine("</div>")
} else {
    [void]$sb.AppendLine("<p class='empty'>No backup copy jobs configured.</p>")
}

[void]$sb.AppendLine("  </div></div>")

# ---- Section 07: NAS Backup ----
[void]$sb.AppendLine("<div class='section' id='sec-nas'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>07</span><h2>NAS Backup</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")
if ($NasJobs.Count -gt 0) {
    $njData = $NasJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            Repository  = $_.BackupRepository.Name
            'Copy Repo' = $_.CopyBackupRepository.Name
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($njData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No NAS backup jobs configured.</p>") }
if ($FileShares.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection' style='margin-top:16px'><h3>File Shares</h3>")
    $fsData = $FileShares | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; Path=$_.Path; Type=$_.ShareType; 'Cache Repo'=$_.CacheRepository.Name }
    }
    [void]$sb.AppendLine(($fsData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
[void]$sb.AppendLine("  </div></div>")

# ---- Section 08: Replication & CDP ----
[void]$sb.AppendLine("<div class='section' id='sec-repl'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>08</span><h2>Replication &amp; CDP</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")
if ($ReplJobs.Count -gt 0) {
    $replData = $ReplJobs | ForEach-Object {
        [PSCustomObject]@{
            Name            = $_.Name
            Enabled         = $_.IsScheduleEnabled
            Target          = $_.Target
            'Restore Points'= try { $_.GetOptions().RetentionPolicy.Quantity } catch { 'N/A' }
            Description     = $_.Description
        }
    }
    [void]$sb.AppendLine(($replData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No replication jobs configured.</p>") }
if ($CDPPolicies.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection' style='margin-top:16px'><h3>CDP Policies</h3>")
    [void]$sb.AppendLine(($CDPPolicies | Select-Object Name, IsEnabled, Description | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
[void]$sb.AppendLine("  </div></div>")

# ---- Section 09: SureBackup ----
[void]$sb.AppendLine("<div class='section' id='sec-sb'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>09</span><h2>SureBackup &amp; Recovery Verification</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")
if ($VLabs.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Virtual Labs</h3>")
    $vlData = $VLabs | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; Host=$_.Host.Name; Status=$_.Status; Proxy=$_.ProxyAppliance.Name }
    }
    [void]$sb.AppendLine(($vlData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
if ($AppGroups.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Application Groups</h3>")
    $agData = $AppGroups | ForEach-Object {
        $vms = try { @($_.GetApplications()) } catch { @() }
        [PSCustomObject]@{ Name=$_.Name; VMs=$vms.Count; Description=$_.Description }
    }
    [void]$sb.AppendLine(($agData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
if ($SBJobs.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>SureBackup Jobs</h3>")
    $sbData = $SBJobs | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; Enabled=$_.IsScheduleEnabled; 'Virtual Lab'=$_.VirtualLab.Name; Description=$_.Description }
    }
    [void]$sb.AppendLine(($sbData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
} else { [void]$sb.AppendLine("<p class='empty'>No SureBackup jobs configured.</p>") }
[void]$sb.AppendLine("  </div></div>")

# ---- Section 10: Tape ----
[void]$sb.AppendLine("<div class='section' id='sec-tape'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>10</span><h2>Tape Infrastructure</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")
if ($TapeLibs.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Tape Libraries</h3>")
    $tlData = $TapeLibs | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; State=$_.State; Model=$_.Model; Drives=($_.Drives | Measure-Object).Count; Slots=$_.TotalSlots }
    }
    [void]$sb.AppendLine(($tlData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
if ($TapePools.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Tape Media Pools</h3>")
    $tpData = $TapePools | ForEach-Object {
        [PSCustomObject]@{
            Name         = $_.Name
            Type         = $_.Type
            'Media Count'= try { ($_.GetTapeMedias() | Measure-Object).Count } catch { 'N/A' }
            Retention    = $_.RetentionPolicy
        }
    }
    [void]$sb.AppendLine(($tpData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
if ($TapeJobs.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Tape Jobs</h3>")
    $tjData = $TapeJobs | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; Type=$_.Type; Enabled=$_.Enabled; 'Media Pool'=$_.MediaPool.Name; Description=$_.Description }
    }
    [void]$sb.AppendLine(($tjData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
} else { [void]$sb.AppendLine("<p class='empty'>No tape infrastructure configured.</p>") }
[void]$sb.AppendLine("  </div></div>")

# ---- Section 11: Cloud Connect ----
[void]$sb.AppendLine("<div class='section' id='sec-cloud'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>11</span><h2>Cloud Connect</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")
if ($CloudTenants.Count -gt 0) {
    $ctData = $CloudTenants | ForEach-Object {
        [PSCustomObject]@{
            Name           = $_.Name
            Enabled        = $_.Enabled
            'Lease Expiry' = $_.LeaseExpirationDate.ToString('yyyy-MM-dd')
            Description    = $_.Description
        }
    }
    [void]$sb.AppendLine(($ctData | ConvertTo-Html -Fragment))
} else { [void]$sb.AppendLine("<p class='empty'>No cloud tenants configured.</p>") }
[void]$sb.AppendLine("  </div></div>")

# ---- Section 12: Inventory ----
[void]$sb.AppendLine("<div class='section' id='sec-inv'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>12</span><h2>Inventory</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")
if ($ViServers.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>VMware vSphere Servers</h3>")
    $viSrvData = $ViServers | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; 'IP Address'=Resolve-HostIP $_.Name; Type=$_.Type; DNS=$_.DNSName; Connected=$_.IsConnected }
    }
    [void]$sb.AppendLine(($viSrvData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
if ($HvServers.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Hyper-V Servers</h3>")
    $hvSrvData = $HvServers | ForEach-Object {
        [PSCustomObject]@{ Name=$_.Name; 'IP Address'=Resolve-HostIP $_.Name; DNS=$_.DNSName; Connected=$_.IsConnected }
    }
    [void]$sb.AppendLine(($hvSrvData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
if ($Credentials.Count -gt 0) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Stored Credentials - Names Only</h3>")
    [void]$sb.AppendLine(($Credentials | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
[void]$sb.AppendLine("  </div></div>")

# ---- Section 13: Global Settings ----
[void]$sb.AppendLine("<div class='section' id='sec-global'>")
[void]$sb.AppendLine("  <div class='section-head'><span class='sec-num'>13</span><h2>Global Settings</h2></div>")
[void]$sb.AppendLine("  <div class='section-body'>")
if ($NotifOpts) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Notification Options</h3>")
    $nData = [PSCustomObject]@{
        'Send on Success' = $NotifOpts.SendSuccessEmail
        'Send on Warning' = $NotifOpts.SendWarningEmail
        'Send on Failure' = $NotifOpts.SendFailureEmail
        'Notify on Retry' = $NotifOpts.SendNotificationOnLastRetryFailure
    }
    [void]$sb.AppendLine(($nData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
if ($EmailOpts) {
    [void]$sb.AppendLine("<div class='subsection'><h3>Email Settings</h3>")
    $eData = [PSCustomObject]@{
        'SMTP Server' = $EmailOpts.SMTPServer
        'SMTP Port'   = $EmailOpts.SMTPPort
        From          = $EmailOpts.From
        To            = $EmailOpts.To
        SSL           = $EmailOpts.EnableSSL
        'Use Auth'    = $EmailOpts.UseAuthentication
    }
    [void]$sb.AppendLine(($eData | ConvertTo-Html -Fragment))
    [void]$sb.AppendLine("</div>")
}
[void]$sb.AppendLine("  </div></div>")

# ---- Close layout + Footer ----
[void]$sb.AppendLine("</div></div>")
[void]$sb.AppendLine("<div class='accent-bar'></div>")
[void]$sb.AppendLine("<div class='rg-footer'>")
[void]$sb.AppendLine("  <div class='fb'><strong>The Redesign Group</strong><span>Technology &amp; Cybersecurity Consulting</span><span>redesign-group.com</span></div>")
[void]$sb.AppendLine("  <div class='fc'>Veeam B&amp;R v13 As-Built Report v3.1 | $ReportDate</div>")
[void]$sb.AppendLine("  <div class='fr'>$CustomerName<br>Server: $($BkpServer.Name ?? 'Unknown')</div>")
[void]$sb.AppendLine("</div>")
[void]$sb.AppendLine("<script>$JS</script>")
[void]$sb.AppendLine("</body></html>")

# =============================================================================
# SAVE
# =============================================================================

$FilePath = "$OutputPath\$ReportName-$CustomerName-$(Get-Date -Format 'yyyyMMdd-HHmm').html"
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
$sb.ToString() | Out-File -FilePath $FilePath -Encoding UTF8

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  Report saved: $FilePath" -ForegroundColor Green
Write-Host "  Customer: $CustomerName  |  Prepared by: $PreparedBy" -ForegroundColor Cyan
if ($HealthCheck -and $Warnings.Count -gt 0) {
    Write-Host "  $($Warnings.Count) health warning(s) found - review the report." -ForegroundColor Yellow
}
Write-Host "  Open in any browser. Ctrl+P to export as PDF." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
