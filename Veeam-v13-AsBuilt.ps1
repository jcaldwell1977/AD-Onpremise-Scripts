# =============================================================================
# VEEAM BACKUP & REPLICATION v13 – AS-BUILT REPORT GENERATOR
# Version 2.0 | Production-Grade | Full Coverage
#
# Usage:
#   pwsh -File Veeam-v13-AsBuilt.ps1
#   pwsh -File Veeam-v13-AsBuilt.ps1 -OutputPath "D:\Reports" -HealthCheck
#
# Requirements:
#   - PowerShell 7.2+
#   - Veeam Backup & Replication v13 Console installed on the machine running this script
#   - Veeam.Backup.PowerShell module (auto-loaded by Veeam installer)
#   - Run as a user with Veeam Backup Administrator role
# =============================================================================

#Requires -Version 7.2

param(
    [string]$OutputPath   = "C:\AsBuiltReports",
    [string]$ReportTitle  = "Veeam Backup & Replication v13 – As-Built Report",
    [switch]$HealthCheck,           # Emit warnings for degraded state (low disk, failed jobs, etc.)
    [switch]$SkipJobDetails         # Omit per-job deep-dive (speeds up large environments)
)

# ─── 0. ENVIRONMENT GUARD ────────────────────────────────────────────────────

Write-Host "⏳  Loading Veeam PowerShell module..." -ForegroundColor Cyan
try {
    Import-Module Veeam.Backup.PowerShell -ErrorAction Stop
} catch {
    Write-Error "Cannot load Veeam.Backup.PowerShell module. Ensure the Veeam B&R Console is installed on this machine.`n$_"
    exit 1
}

$VeeamModule = Get-Module Veeam.Backup.PowerShell
Write-Host "✅  Module loaded: $($VeeamModule.Version)" -ForegroundColor Green

# ─── 1. HELPERS ──────────────────────────────────────────────────────────────

$ReportDate    = Get-Date -Format "dddd, dd MMMM yyyy HH:mm"
$ReportDateISO = Get-Date -Format "yyyyMMdd-HHmm"
$Warnings      = [System.Collections.Generic.List[string]]::new()

function ConvertTo-HtmlTable {
    <#
    .SYNOPSIS  Converts an array of objects or a single object into an HTML <table> fragment.
               Handles empty collections gracefully.
    #>
    param(
        [Parameter(ValueFromPipeline)]$InputObject,
        [string[]]$Properties,
        [string]$EmptyMessage = "No data found."
    )
    begin   { $rows = [System.Collections.Generic.List[object]]::new() }
    process { if ($null -ne $InputObject) { foreach ($item in $InputObject) { $rows.Add($item) } } }
    end {
        if ($rows.Count -eq 0) { return "<p class='empty'>$EmptyMessage</p>" }
        $sel = if ($Properties) { $rows | Select-Object $Properties } else { $rows }
        # ConvertTo-Html -Fragment produces a clean table; strip surrounding XML declaration
        ($sel | ConvertTo-Html -Fragment) -replace '<!DOCTYPE[^>]*>', '' -replace '<html>|</html>|<body>|</body>|<head>|</head>', ''
    }
}

function Get-StatusBadge {
    param([string]$Status)
    $map = @{
        'Success'  = 'badge-success'
        'Warning'  = 'badge-warn'
        'Failed'   = 'badge-fail'
        'Running'  = 'badge-info'
        'None'     = 'badge-neutral'
        'Disabled' = 'badge-neutral'
        'Valid'    = 'badge-success'
        'Invalid'  = 'badge-fail'
        'Expired'  = 'badge-fail'
    }
    $cls = $map[$Status]
    if (-not $cls) { $cls = 'badge-neutral' }
    return "<span class='badge $cls'>$Status</span>"
}

function Add-HealthWarning {
    param([string]$Message)
    $script:Warnings.Add($Message)
}

# ─── 2. HTML SKELETON ────────────────────────────────────────────────────────

$CSS = @'
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

:root {
    --bg:         #0d1117;
    --surface:    #161b22;
    --surface2:   #21262d;
    --border:     #30363d;
    --accent:     #00b4d8;
    --accent2:    #f7c948;
    --text:       #e6edf3;
    --text-muted: #8b949e;
    --success:    #3fb950;
    --warn:       #f7c948;
    --fail:       #f85149;
    --info:       #58a6ff;
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
    font-family: 'IBM Plex Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    font-size: 13px;
    line-height: 1.6;
}

/* ── Header ── */
.report-header {
    background: linear-gradient(135deg, #0d1117 0%, #0a2a3d 60%, #003d5b 100%);
    border-bottom: 2px solid var(--accent);
    padding: 48px 60px 36px;
    position: relative;
    overflow: hidden;
}
.report-header::before {
    content: '';
    position: absolute;
    inset: 0;
    background: repeating-linear-gradient(
        90deg,
        transparent,
        transparent 59px,
        rgba(0,180,216,0.04) 60px
    );
}
.report-header .logo-row {
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 24px;
}
.veeam-badge {
    background: var(--accent);
    color: #000;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    font-size: 11px;
    letter-spacing: 2px;
    padding: 4px 10px;
    border-radius: 2px;
}
.report-header h1 {
    font-size: 26px;
    font-weight: 700;
    letter-spacing: -0.5px;
    color: var(--text);
}
.report-header .meta {
    color: var(--text-muted);
    font-size: 12px;
    margin-top: 6px;
    font-family: 'IBM Plex Mono', monospace;
}

/* ── Navigation TOC ── */
.toc {
    background: var(--surface);
    border-right: 1px solid var(--border);
    padding: 24px;
}
.toc h3 {
    font-size: 10px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 12px;
}
.toc a {
    display: block;
    color: var(--text-muted);
    text-decoration: none;
    padding: 4px 8px;
    border-radius: 4px;
    font-size: 12px;
    transition: all 0.15s;
}
.toc a:hover, .toc a.active { color: var(--accent); background: rgba(0,180,216,0.08); }

/* ── Layout ── */
.layout {
    display: grid;
    grid-template-columns: 220px 1fr;
    min-height: calc(100vh - 200px);
}
.main-content { padding: 32px 48px; }

/* ── Sections ── */
.section {
    margin-bottom: 40px;
    scroll-margin-top: 20px;
}
.section-header {
    display: flex;
    align-items: center;
    gap: 12px;
    border-bottom: 1px solid var(--border);
    padding-bottom: 10px;
    margin-bottom: 20px;
}
.section-num {
    background: var(--accent);
    color: #000;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    font-size: 11px;
    width: 24px;
    height: 24px;
    border-radius: 4px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}
.section h2 {
    font-size: 16px;
    font-weight: 600;
    color: var(--text);
}
.subsection { margin: 20px 0; }
.subsection h3 {
    font-size: 13px;
    font-weight: 600;
    color: var(--accent);
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    gap: 8px;
}
.subsection h3::before {
    content: '';
    display: inline-block;
    width: 3px;
    height: 14px;
    background: var(--accent);
    border-radius: 2px;
}

/* ── Tables ── */
table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
    font-family: 'IBM Plex Mono', monospace;
    background: var(--surface);
    border-radius: 6px;
    overflow: hidden;
    border: 1px solid var(--border);
}
th {
    background: var(--surface2);
    color: var(--text-muted);
    font-size: 10px;
    letter-spacing: 1px;
    text-transform: uppercase;
    padding: 10px 12px;
    text-align: left;
    border-bottom: 1px solid var(--border);
}
td {
    padding: 9px 12px;
    border-bottom: 1px solid var(--border);
    vertical-align: top;
    color: var(--text);
    word-break: break-word;
    max-width: 400px;
}
tr:last-child td { border-bottom: none; }
tr:hover td { background: rgba(255,255,255,0.02); }

/* ── Badges ── */
.badge {
    display: inline-block;
    font-size: 10px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 2px;
    letter-spacing: 0.5px;
}
.badge-success { background: rgba(63,185,80,0.15); color: var(--success); border: 1px solid rgba(63,185,80,0.3); }
.badge-warn    { background: rgba(247,201,72,0.15); color: var(--warn);    border: 1px solid rgba(247,201,72,0.3); }
.badge-fail    { background: rgba(248,81,73,0.15);  color: var(--fail);    border: 1px solid rgba(248,81,73,0.3); }
.badge-info    { background: rgba(88,166,255,0.15); color: var(--info);    border: 1px solid rgba(88,166,255,0.3); }
.badge-neutral { background: rgba(139,148,158,0.15); color: var(--text-muted); border: 1px solid rgba(139,148,158,0.3); }

/* ── Summary Cards ── */
.summary-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
    gap: 12px;
    margin-bottom: 32px;
}
.summary-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 16px;
    text-align: center;
}
.summary-card .val {
    font-size: 28px;
    font-weight: 700;
    font-family: 'IBM Plex Mono', monospace;
    color: var(--accent);
    line-height: 1;
}
.summary-card .lbl {
    font-size: 10px;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-top: 6px;
}

/* ── Health Warning Banner ── */
.health-banner {
    background: rgba(248,81,73,0.08);
    border: 1px solid rgba(248,81,73,0.3);
    border-radius: 6px;
    padding: 16px 20px;
    margin-bottom: 28px;
}
.health-banner h3 {
    color: var(--fail);
    font-size: 12px;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 10px;
}
.health-banner ul { padding-left: 18px; }
.health-banner li { color: var(--fail); font-size: 12px; margin-bottom: 4px; }

/* ── Job Card ── */
.job-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 6px;
    margin-bottom: 16px;
    overflow: hidden;
}
.job-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: var(--surface2);
    padding: 12px 16px;
    cursor: pointer;
    user-select: none;
}
.job-card-header:hover { background: rgba(255,255,255,0.04); }
.job-card-title {
    font-weight: 600;
    font-size: 13px;
    font-family: 'IBM Plex Sans', sans-serif;
}
.job-card-body { padding: 16px; display: none; }
.job-card.open .job-card-body { display: block; }
.job-card-meta {
    display: flex;
    gap: 16px;
    flex-wrap: wrap;
    margin-bottom: 12px;
    font-size: 11px;
    color: var(--text-muted);
    font-family: 'IBM Plex Mono', monospace;
}

/* ── Misc ── */
.empty  { color: var(--text-muted); font-style: italic; font-size: 12px; padding: 8px 0; }
.error  { color: var(--fail); font-size: 12px; font-family: 'IBM Plex Mono', monospace; }
.note   { color: var(--text-muted); font-size: 11px; margin-top: 6px; }
hr      { border: none; border-top: 1px solid var(--border); margin: 32px 0; }
pre     { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--text-muted); white-space: pre-wrap; }

/* ── Footer ── */
.report-footer {
    background: var(--surface);
    border-top: 1px solid var(--border);
    padding: 20px 60px;
    font-size: 11px;
    color: var(--text-muted);
    font-family: 'IBM Plex Mono', monospace;
    display: flex;
    justify-content: space-between;
}

/* ── Print ── */
@media print {
    body { background: #fff; color: #000; }
    .toc, .job-card-header { display: none; }
    .layout { display: block; }
    .main-content { padding: 0; }
    .job-card-body { display: block !important; }
    th { background: #e0e0e0; color: #000; }
    .badge { border: 1px solid #999; background: #eee; color: #000; }
    .report-header { background: #003d5b; }
    .section { page-break-inside: avoid; }
    table { border: 1px solid #ccc; }
}
'@

$JS = @'
document.querySelectorAll('.job-card-header').forEach(h => {
    h.addEventListener('click', () => h.parentElement.classList.toggle('open'));
});

// Sticky TOC highlight
const sections = document.querySelectorAll('.section');
const tocLinks = document.querySelectorAll('.toc a');
const observer = new IntersectionObserver(entries => {
    entries.forEach(e => {
        if (e.isIntersecting) {
            tocLinks.forEach(l => l.classList.remove('active'));
            const link = document.querySelector('.toc a[href="#' + e.target.id + '"]');
            if (link) link.classList.add('active');
        }
    });
}, { threshold: 0.2 });
sections.forEach(s => observer.observe(s));
'@

# ─── 3. DATA COLLECTION ──────────────────────────────────────────────────────

Write-Host "⏳  Collecting Veeam configuration data..." -ForegroundColor Cyan

# ── 3.1 Backup Server ──
$BkpServer = $null
try   { $BkpServer = Get-VBRServer | Where-Object { $_.IsLocal -or $_.Type -eq 'Local' } | Select-Object -First 1 }
catch { Write-Warning "Could not retrieve local server info: $_" }

# ── 3.2 License ──
$License = $null
try {
    $License = Get-VBRInstalledLicense
    if ($HealthCheck) {
        if ($License.Status -eq 'Expired')          { Add-HealthWarning "License is EXPIRED." }
        elseif ($License.ExpirationDate -lt (Get-Date).AddDays(30)) {
            Add-HealthWarning "License expires within 30 days: $($License.ExpirationDate.ToString('yyyy-MM-dd'))"
        }
    }
} catch { Write-Warning "License query failed: $_" }

# ── 3.3 Proxies ──
$ViProxies  = @(); try { $ViProxies  = @(Get-VBRViProxy)  } catch { Write-Warning "VI Proxies: $_" }
$HvProxies  = @(); try { $HvProxies  = @(Get-VBRHvProxy)  } catch { Write-Warning "HV Proxies: $_" }
$NasProxies = @(); try { $NasProxies = @(Get-VBRNASProxyServer) } catch { Write-Warning "NAS Proxies: $_" }

# ── 3.4 Repositories ──
$Repos = @()
try {
    $Repos = @(Get-VBRBackupRepository)
    if ($HealthCheck) {
        foreach ($r in $Repos) {
            if ($r.TotalSpace -gt 0) {
                $pctFree = [math]::Round(($r.FreeSpace / $r.TotalSpace) * 100, 1)
                if ($pctFree -lt 10) { Add-HealthWarning "Repository '$($r.Name)' is critically low on space ($pctFree% free)." }
                elseif ($pctFree -lt 20) { Add-HealthWarning "Repository '$($r.Name)' is below 20% free space ($pctFree%)." }
            }
        }
    }
} catch { Write-Warning "Repositories: $_" }

# ── 3.5 SOBR ──
$SOBRs = @(); try { $SOBRs = @(Get-VBRBackupRepository -ScaleOut) } catch { Write-Warning "SOBR: $_" }

# ── 3.6 External Repositories (Cloud & Immutable Object Storage) ──
$ExtRepos = @(); try { $ExtRepos = @(Get-VBRExternalRepository) } catch { Write-Warning "External Repos: $_" }

# ── 3.7 Jobs ──
$AllJobs = @(); try { $AllJobs = @(Get-VBRJob) } catch { Write-Warning "Jobs: $_" }

$FailedJobs = @()
if ($HealthCheck) {
    foreach ($j in $AllJobs) {
        if ($j.IsScheduleEnabled) {
            try {
                $lastSession = Get-VBRJobSession -Job $j -Last 1 -ErrorAction SilentlyContinue
                if ($lastSession -and $lastSession.Result -eq 'Failed') {
                    Add-HealthWarning "Job '$($j.Name)' last run FAILED ($(($lastSession.EndTime).ToString('yyyy-MM-dd HH:mm')))."
                    $FailedJobs += $j.Name
                }
            } catch { }
        }
    }
}

# ── 3.8 NAS Backup Jobs ──
$NasJobs = @(); try { $NasJobs = @(Get-VBRNASBackupJob) } catch { Write-Warning "NAS Backup Jobs: $_" }

# ── 3.9 Replication Jobs (included in Get-VBRJob but typed separately for clarity) ──
$ReplJobs = $AllJobs | Where-Object { $_.JobType -in @('Replica','SimpleTransactionLog') }

# ── 3.10 CDP Policies ──
$CDPPolicies = @(); try { $CDPPolicies = @(Get-VBRCDPPolicy) } catch { Write-Warning "CDP: $_" }

# ── 3.11 Backup Copy Jobs ──
$CopyJobs = $AllJobs | Where-Object { $_.JobType -eq 'BackupSync' }

# ── 3.12 SureBackup ──
$SBJobs = @(); try { $SBJobs = @(Get-VBRSureBackupJob) } catch { Write-Warning "SureBackup: $_" }
$VLabs  = @(); try { $VLabs  = @(Get-VBRVirtualLab)    } catch { Write-Warning "Virtual Labs: $_" }
$AppGroups = @(); try { $AppGroups = @(Get-VBRApplicationGroup) } catch { Write-Warning "App Groups: $_" }

# ── 3.13 Tape ──
$TapeLibraries  = @(); try { $TapeLibraries  = @(Get-VBRTapeLibrary)   } catch { Write-Warning "Tape Libraries: $_" }
$TapeMediaPools = @(); try { $TapeMediaPools = @(Get-VBRTapeMediaPool)  } catch { Write-Warning "Tape Media Pools: $_" }
$TapeJobs       = @(); try { $TapeJobs       = @(Get-VBRTapeJob)        } catch { Write-Warning "Tape Jobs: $_" }

# ── 3.14 Cloud Connect ──
$CloudTenants   = @(); try { $CloudTenants   = @(Get-VBRCloudTenant)     } catch { Write-Warning "Cloud Tenants: $_" }
$CloudHardware  = @(); try { $CloudHardware  = @(Get-VBRCloudHardwarePlan)} catch { Write-Warning "Cloud HW Plans: $_" }

# ── 3.15 vSphere / Hyper-V Infrastructure ──
$ViServers = @(); try { $ViServers = @(Get-VBRServer | Where-Object { $_.Type -in @('VC','ESXi') }) } catch { Write-Warning "VI Servers: $_" }
$HvServers = @(); try { $HvServers = @(Get-VBRServer | Where-Object { $_.Type -eq 'HvServer' })    } catch { Write-Warning "HV Servers: $_" }

# ── 3.16 NAS / File Shares ──
$FileShares = @(); try { $FileShares = @(Get-VBRNASFileShare) } catch { Write-Warning "File Shares: $_" }

# ── 3.17 Credentials ──
$Credentials = @(); try { $Credentials = @(Get-VBRCredentials | Select-Object Name, Description, ChangePasswordTo) } catch { Write-Warning "Credentials: $_" }

# ── 3.18 Global Notification Settings ──
$NotifOpts = $null;  try { $NotifOpts = Get-VBRNotificationOptions }  catch { }
$EmailOpts = $null;  try { $EmailOpts = Get-VBREmailOptions }          catch { }

Write-Host "✅  Data collection complete." -ForegroundColor Green

# ─── 4. SUMMARY CARD VALUES ──────────────────────────────────────────────────

$totalJobs      = $AllJobs.Count + $NasJobs.Count + $TapeJobs.Count
$enabledJobs    = ($AllJobs | Where-Object { $_.IsScheduleEnabled }).Count
$totalRepos     = $Repos.Count + $SOBRs.Count
$totalProxies   = $ViProxies.Count + $HvProxies.Count + $NasProxies.Count
$totalTape      = $TapeLibraries.Count

# ─── 5. HTML GENERATION ──────────────────────────────────────────────────────

Write-Host "⏳  Building HTML report..." -ForegroundColor Cyan

$sb = [System.Text.StringBuilder]::new()

# HTML Head
[void]$sb.AppendLine(@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$ReportTitle</title>
<style>$CSS</style>
</head>
<body>
"@)

# ── Header ──
[void]$sb.AppendLine(@"
<div class="report-header">
  <div class="logo-row">
    <span class="veeam-badge">VEEAM</span>
    <span style="color:var(--text-muted);font-size:12px;font-family:'IBM Plex Mono',monospace;">BACKUP &amp; REPLICATION v13</span>
  </div>
  <h1>$ReportTitle</h1>
  <div class="meta">Generated: $ReportDate &nbsp;|&nbsp; Server: $($BkpServer.Name ?? 'Unknown') &nbsp;|&nbsp; Module: $($VeeamModule.Version)</div>
</div>
"@)

# ── TOC + Layout ──
[void]$sb.AppendLine('<div class="layout">')
[void]$sb.AppendLine(@'
<nav class="toc">
  <h3>Contents</h3>
  <a href="#sec-summary">Summary Dashboard</a>
  <a href="#sec-server">1. Backup Server</a>
  <a href="#sec-license">2. License</a>
  <a href="#sec-infra">3. Infrastructure</a>
  <a href="#sec-repos">4. Repositories</a>
  <a href="#sec-jobs">5. Backup Jobs</a>
  <a href="#sec-nas">6. NAS Backup</a>
  <a href="#sec-replication">7. Replication &amp; CDP</a>
  <a href="#sec-surebackup">8. SureBackup</a>
  <a href="#sec-tape">9. Tape</a>
  <a href="#sec-cloud">10. Cloud Connect</a>
  <a href="#sec-inventory">11. Inventory</a>
  <a href="#sec-global">12. Global Settings</a>
</nav>
'@)

[void]$sb.AppendLine('<div class="main-content">')

# ── Health Warnings ──
if ($HealthCheck -and $Warnings.Count -gt 0) {
    [void]$sb.AppendLine('<div class="health-banner"><h3>⚠ Health Check Warnings</h3><ul>')
    foreach ($w in $Warnings) { [void]$sb.AppendLine("<li>$w</li>") }
    [void]$sb.AppendLine('</ul></div>')
}

# ── Summary Dashboard ──
[void]$sb.AppendLine('<div class="section" id="sec-summary"><div class="section-header"><span class="section-num">◈</span><h2>Summary Dashboard</h2></div>')
[void]$sb.AppendLine('<div class="summary-grid">')

$cards = @(
    @{ Val = $totalJobs;     Lbl = "Total Jobs" },
    @{ Val = $enabledJobs;   Lbl = "Scheduled Jobs" },
    @{ Val = $totalRepos;    Lbl = "Repositories" },
    @{ Val = $totalProxies;  Lbl = "Proxies" },
    @{ Val = $ViServers.Count + $HvServers.Count; Lbl = "Hypervisor Hosts" },
    @{ Val = $FileShares.Count; Lbl = "File Shares" },
    @{ Val = $TapeLibraries.Count; Lbl = "Tape Libraries" },
    @{ Val = $CloudTenants.Count;  Lbl = "Cloud Tenants" }
)
foreach ($c in $cards) {
    [void]$sb.AppendLine("<div class='summary-card'><div class='val'>$($c.Val)</div><div class='lbl'>$($c.Lbl)</div></div>")
}
[void]$sb.AppendLine('</div></div>')

# ── Section 1: Backup Server ──
[void]$sb.AppendLine('<div class="section" id="sec-server"><div class="section-header"><span class="section-num">1</span><h2>Backup Server</h2></div>')
if ($BkpServer) {
    $serverData = [PSCustomObject]@{
        Name        = $BkpServer.Name
        DNS         = $BkpServer.DNSName
        Type        = $BkpServer.Type
        Description = $BkpServer.Description
        IsConnected = $BkpServer.IsConnected
    }
    [void]$sb.AppendLine(($serverData | ConvertTo-HtmlTable))
} else {
    [void]$sb.AppendLine("<p class='error'>Could not retrieve local server information.</p>")
}
[void]$sb.AppendLine('</div>')

# ── Section 2: License ──
[void]$sb.AppendLine('<div class="section" id="sec-license"><div class="section-header"><span class="section-num">2</span><h2>License</h2></div>')
if ($License) {
    $licData = [PSCustomObject]@{
        Edition       = $License.Edition
        Type          = $License.LicenseType
        Status        = $License.Status
        ExpirationDate = $License.ExpirationDate?.ToString('yyyy-MM-dd')
        'Support Expiry' = $License.SupportExpirationDate?.ToString('yyyy-MM-dd')
        'Licensed VMs/Agents' = "$($License.UsedLicensesNumber) / $($License.TotalLicensesNumber)"
        'Licensed Sockets' = $License.UsedSocketsNumber
    }
    [void]$sb.AppendLine(($licData | ConvertTo-HtmlTable))
    if ($HealthCheck -and $License.Status -ne 'Valid') {
        [void]$sb.AppendLine("<p class='error'>⚠ License status is $($License.Status)</p>")
    }
} else {
    [void]$sb.AppendLine("<p class='error'>Could not retrieve license information.</p>")
}
[void]$sb.AppendLine('</div>')

# ── Section 3: Infrastructure (Proxies) ──
[void]$sb.AppendLine('<div class="section" id="sec-infra"><div class="section-header"><span class="section-num">3</span><h2>Backup Infrastructure</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>VMware Backup Proxies</h3>')
if ($ViProxies.Count -gt 0) {
    $viData = $ViProxies | ForEach-Object {
        [PSCustomObject]@{
            Name          = $_.Name
            Host          = $_.Host.Name
            'Transport Mode' = $_.Options.TransportMode
            'Max Tasks'   = $_.MaxTasksCount
            Status        = $_.IsDisabled ? 'Disabled' : 'Enabled'
        }
    }
    [void]$sb.AppendLine(($viData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No VMware proxies configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Hyper-V Backup Proxies (Off-Host Processing Servers)</h3>')
if ($HvProxies.Count -gt 0) {
    $hvData = $HvProxies | ForEach-Object {
        [PSCustomObject]@{
            Name       = $_.Name
            Host       = $_.Host.Name
            'Max Tasks' = $_.MaxTasksCount
        }
    }
    [void]$sb.AppendLine(($hvData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No Hyper-V off-host proxies configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>NAS Backup Proxies</h3>')
if ($NasProxies.Count -gt 0) {
    [void]$sb.AppendLine(($NasProxies | Select-Object Name, Description, @{N='Server';E={$_.Server.Name}} | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No NAS proxies configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 4: Repositories ──
[void]$sb.AppendLine('<div class="section" id="sec-repos"><div class="section-header"><span class="section-num">4</span><h2>Repositories &amp; Storage</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Backup Repositories</h3>')
if ($Repos.Count -gt 0) {
    $repoData = $Repos | ForEach-Object {
        $freeGB   = [math]::Round($_.FreeSpace  / 1GB, 1)
        $totalGB  = [math]::Round($_.TotalSpace / 1GB, 1)
        $pctFree  = if ($totalGB -gt 0) { [math]::Round(($freeGB / $totalGB) * 100, 1) } else { 'N/A' }
        [PSCustomObject]@{
            Name        = $_.Name
            Type        = $_.Type
            Path        = $_.Path
            Host        = $_.Host.Name
            'Total (GB)' = $totalGB
            'Free (GB)'  = $freeGB
            'Free %'     = $pctFree
            'Per-VM Backup' = $_.UsePerVMBackupFiles
            Immutability  = $_.ImmutabilityEnabled
        }
    }
    [void]$sb.AppendLine(($repoData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No backup repositories found.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Scale-Out Backup Repositories (SOBR)</h3>')
if ($SOBRs.Count -gt 0) {
    $sobrData = $SOBRs | ForEach-Object {
        [PSCustomObject]@{
            Name                  = $_.Name
            'Policy'              = $_.PolicyType
            'Performance Extents' = ($_.Extents | Where-Object { $_.IsActive } | Measure-Object).Count
            'Capacity Tier'       = $_.CapacityTier.Enabled
            'Archive Tier'        = $_.ArchiveTier.Enabled
        }
    }
    [void]$sb.AppendLine(($sobrData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No Scale-Out Repositories configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>External / Object Storage Repositories</h3>')
if ($ExtRepos.Count -gt 0) {
    [void]$sb.AppendLine(($ExtRepos | Select-Object Name, Type, Description | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No external repositories configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 5: Backup Jobs ──
[void]$sb.AppendLine('<div class="section" id="sec-jobs"><div class="section-header"><span class="section-num">5</span><h2>Backup Jobs</h2></div>')

$BackupJobs = $AllJobs | Where-Object { $_.JobType -eq 'Backup' }

if ($BackupJobs.Count -gt 0) {
    [void]$sb.AppendLine('<div class="subsection"><h3>Job Overview</h3>')
    $jobOverview = $BackupJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            Repository  = $_.GetTargetRepository()?.Name
            Objects     = ($_.GetObjectsInJob() | Measure-Object).Count
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($jobOverview | ConvertTo-HtmlTable))
    [void]$sb.AppendLine('</div>')

    if (-not $SkipJobDetails) {
        [void]$sb.AppendLine('<div class="subsection"><h3>Detailed Job Configuration (click to expand)</h3>')
        foreach ($job in $BackupJobs) {
            $opts = $null; try { $opts = $job.GetOptions() } catch { }
            $sched = $null; try { $sched = $job.ScheduleOptions } catch { }
            $objs  = @(); try { $objs = @($job.GetObjectsInJob()) } catch { }

            $statusBadge = if ($job.IsScheduleEnabled) { Get-StatusBadge 'Success' } else { Get-StatusBadge 'Disabled' }

            [void]$sb.AppendLine("<div class='job-card'>")
            [void]$sb.AppendLine("<div class='job-card-header'><span class='job-card-title'>$([System.Web.HttpUtility]::HtmlEncode($job.Name))</span>$statusBadge</div>")
            [void]$sb.AppendLine("<div class='job-card-body'>")
            [void]$sb.AppendLine("<div class='job-card-meta'><span>Type: $($job.JobType)</span><span>Repo: $($job.GetTargetRepository()?.Name)</span><span>Objects: $($objs.Count)</span></div>")

            # Source Objects
            if ($objs.Count -gt 0) {
                [void]$sb.AppendLine("<p style='margin-bottom:8px;font-weight:600;font-size:11px;color:var(--text-muted);'>SOURCE OBJECTS</p>")
                $objData = $objs | ForEach-Object {
                    [PSCustomObject]@{ Name = $_.Name; Type = $_.Type; Location = $_.Location }
                }
                [void]$sb.AppendLine(($objData | ConvertTo-HtmlTable -EmptyMessage "No source objects."))
            }

            # Retention
            if ($opts) {
                [void]$sb.AppendLine("<p style='margin:12px 0 8px;font-weight:600;font-size:11px;color:var(--text-muted);'>RETENTION &amp; STORAGE</p>")
                $retData = [PSCustomObject]@{
                    'Retention Type'    = $opts.RetentionPolicy.Type
                    'Retention Count'   = $opts.RetentionPolicy.Quantity
                    'GFS Weekly'        = $opts.RetentionPolicy.IsGFSEnabled
                    'GFS Weekly Cnt'    = $opts.RetentionPolicy.WeeklyFullSchedule?.RepeatCount
                    'GFS Monthly Cnt'   = $opts.RetentionPolicy.MonthlyFullSchedule?.RepeatCount
                    'GFS Yearly Cnt'    = $opts.RetentionPolicy.YearlyFullSchedule?.RepeatCount
                    'Dedup'             = $opts.JobOptions.EnableDeduplication
                    'Compression'       = $opts.JobOptions.CompressionType
                    'StorageOpt'        = $opts.JobOptions.BackupStorageOptions?.StorageOptimization
                    'Encryption'        = $opts.JobOptions.EncryptionEnabled
                }
                [void]$sb.AppendLine(($retData | ConvertTo-HtmlTable))

                # Guest Processing
                [void]$sb.AppendLine("<p style='margin:12px 0 8px;font-weight:600;font-size:11px;color:var(--text-muted);'>GUEST PROCESSING</p>")
                $guestData = [PSCustomObject]@{
                    'App-Aware'         = $opts.JobOptions.GenerationPolicy?.IsAppAwareEnabled
                    'Guest Interaction'  = $opts.JobOptions.GenerationPolicy?.IsGuestInteractionEnabled
                    'SQL Logs'          = $opts.JobOptions.GenerationPolicy?.SqlBackupMode
                    'Oracle Logs'       = $opts.JobOptions.GenerationPolicy?.OracleBackupMode
                    'Guest OS Creds'    = $opts.JobOptions.GenerationPolicy?.UseGuestCredentials
                    'Index Files'       = $opts.JobOptions.GenerationPolicy?.FileSystemIndexingScope
                }
                [void]$sb.AppendLine(($guestData | ConvertTo-HtmlTable))
            }

            # Schedule
            if ($sched) {
                [void]$sb.AppendLine("<p style='margin:12px 0 8px;font-weight:600;font-size:11px;color:var(--text-muted);'>SCHEDULE</p>")
                $schedData = [PSCustomObject]@{
                    'Enabled'       = $sched.Enabled
                    'Run At'        = $sched.StartDateTime
                    'Daily Kind'    = $sched.Type
                    'Retry Count'   = $sched.RetryCount
                    'Retry Wait'    = "$($sched.RetryTimeout) min"
                    'Backup Window' = $sched.BackupWindowEnabled
                }
                [void]$sb.AppendLine(($schedData | ConvertTo-HtmlTable))
            }

            [void]$sb.AppendLine("</div></div>")
        }
        [void]$sb.AppendLine('</div>')
    }
} else {
    [void]$sb.AppendLine("<p class='empty'>No backup jobs found.</p>")
}

# Backup Copy Jobs
[void]$sb.AppendLine('<div class="subsection"><h3>Backup Copy Jobs</h3>')
if ($CopyJobs.Count -gt 0) {
    $copyData = $CopyJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            Source      = $_.Source
            Repository  = $_.GetTargetRepository()?.Name
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($copyData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No backup copy jobs configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 6: NAS Backup ──
[void]$sb.AppendLine('<div class="section" id="sec-nas"><div class="section-header"><span class="section-num">6</span><h2>NAS Backup</h2></div>')
[void]$sb.AppendLine('<div class="subsection"><h3>NAS Backup Jobs</h3>')
if ($NasJobs.Count -gt 0) {
    $nasJobData = $NasJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            Repository  = $_.BackupRepository?.Name
            CopyRepo    = $_.CopyBackupRepository?.Name
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($nasJobData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No NAS backup jobs configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>File Shares</h3>')
if ($FileShares.Count -gt 0) {
    $shareData = $FileShares | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Path        = $_.Path
            Type        = $_.ShareType
            CacheRepo   = $_.CacheRepository?.Name
        }
    }
    [void]$sb.AppendLine(($shareData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No file shares configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 7: Replication & CDP ──
[void]$sb.AppendLine('<div class="section" id="sec-replication"><div class="section-header"><span class="section-num">7</span><h2>Replication &amp; CDP</h2></div>')
[void]$sb.AppendLine('<div class="subsection"><h3>Replication Jobs</h3>')
if ($ReplJobs.Count -gt 0) {
    $replData = $ReplJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            Target      = $_.Target
            'Restore Points' = $_.GetOptions()?.RetentionPolicy?.Quantity
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($replData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No replication jobs configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>CDP (Continuous Data Protection) Policies</h3>')
if ($CDPPolicies.Count -gt 0) {
    $cdpData = $CDPPolicies | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsEnabled
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($cdpData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No CDP policies configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 8: SureBackup ──
[void]$sb.AppendLine('<div class="section" id="sec-surebackup"><div class="section-header"><span class="section-num">8</span><h2>SureBackup &amp; Recovery Verification</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Virtual Labs</h3>')
if ($VLabs.Count -gt 0) {
    $vlabData = $VLabs | ForEach-Object {
        [PSCustomObject]@{
            Name   = $_.Name
            Host   = $_.Host.Name
            Status = $_.Status
            Proxy  = $_.ProxyAppliance?.Name
        }
    }
    [void]$sb.AppendLine(($vlabData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No virtual labs configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Application Groups</h3>')
if ($AppGroups.Count -gt 0) {
    $agData = $AppGroups | ForEach-Object {
        $vms = try { @($_.GetApplications()) } catch { @() }
        [PSCustomObject]@{
            Name        = $_.Name
            'VMs'       = $vms.Count
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($agData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No application groups configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>SureBackup Jobs</h3>')
if ($SBJobs.Count -gt 0) {
    $sbData = $SBJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            'Virtual Lab' = $_.VirtualLab?.Name
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($sbData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No SureBackup jobs configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 9: Tape ──
[void]$sb.AppendLine('<div class="section" id="sec-tape"><div class="section-header"><span class="section-num">9</span><h2>Tape Infrastructure</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Tape Libraries</h3>')
if ($TapeLibraries.Count -gt 0) {
    $tapeLibData = $TapeLibraries | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            State       = $_.State
            Model       = $_.Model
            Drives      = ($_.Drives | Measure-Object).Count
            Slots       = $_.TotalSlots
        }
    }
    [void]$sb.AppendLine(($tapeLibData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No tape libraries configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Tape Media Pools</h3>')
if ($TapeMediaPools.Count -gt 0) {
    $tapePoolData = $TapeMediaPools | ForEach-Object {
        [PSCustomObject]@{
            Name           = $_.Name
            Type           = $_.Type
            'Media Count'  = ($_.GetTapeMedias() | Measure-Object).Count
            'Retention'    = $_.RetentionPolicy
        }
    }
    [void]$sb.AppendLine(($tapePoolData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No tape media pools configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Tape Jobs</h3>')
if ($TapeJobs.Count -gt 0) {
    $tapeJobData = $TapeJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Type        = $_.Type
            Enabled     = $_.Enabled
            'Media Pool' = $_.MediaPool?.Name
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($tapeJobData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No tape jobs configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 10: Cloud Connect ──
[void]$sb.AppendLine('<div class="section" id="sec-cloud"><div class="section-header"><span class="section-num">10</span><h2>Cloud Connect</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Cloud Tenants</h3>')
if ($CloudTenants.Count -gt 0) {
    $tenantData = $CloudTenants | ForEach-Object {
        [PSCustomObject]@{
            Name              = $_.Name
            Enabled           = $_.Enabled
            'Lease Expiration' = $_.LeaseExpirationDate?.ToString('yyyy-MM-dd')
            Description       = $_.Description
        }
    }
    [void]$sb.AppendLine(($tenantData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No cloud tenants configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Cloud Hardware Plans</h3>')
if ($CloudHardware.Count -gt 0) {
    [void]$sb.AppendLine(($CloudHardware | Select-Object Name, 'CPU', 'Memory', 'Storage' | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No hardware plans configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 11: Inventory ──
[void]$sb.AppendLine('<div class="section" id="sec-inventory"><div class="section-header"><span class="section-num">11</span><h2>Inventory</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>VMware vSphere Servers</h3>')
if ($ViServers.Count -gt 0) {
    $viSrvData = $ViServers | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Type        = $_.Type
            DNS         = $_.DNSName
            IsConnected = $_.IsConnected
        }
    }
    [void]$sb.AppendLine(($viSrvData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No vSphere servers added.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Hyper-V Servers</h3>')
if ($HvServers.Count -gt 0) {
    $hvSrvData = $HvServers | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            DNS         = $_.DNSName
            IsConnected = $_.IsConnected
        }
    }
    [void]$sb.AppendLine(($hvSrvData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No Hyper-V servers added.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Credentials (Names Only – Passwords Never Exported)</h3>')
if ($Credentials.Count -gt 0) {
    [void]$sb.AppendLine(($Credentials | ConvertTo-HtmlTable -EmptyMessage "No credentials found."))
} else { [void]$sb.AppendLine("<p class='empty'>No credentials stored.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 12: Global Settings ──
[void]$sb.AppendLine('<div class="section" id="sec-global"><div class="section-header"><span class="section-num">12</span><h2>Global Settings</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Notification Options</h3>')
if ($NotifOpts) {
    $notifData = [PSCustomObject]@{
        'Send Success'    = $NotifOpts.SendSuccessEmail
        'Send Warning'    = $NotifOpts.SendWarningEmail
        'Send Failure'    = $NotifOpts.SendFailureEmail
        'Notify Waiting'  = $NotifOpts.SendNotificationOnLastRetryFailure
    }
    [void]$sb.AppendLine(($notifData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>Could not retrieve notification settings.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Email Settings</h3>')
if ($EmailOpts) {
    $emailData = [PSCustomObject]@{
        'SMTP Server'   = $EmailOpts.SMTPServer
        'SMTP Port'     = $EmailOpts.SMTPPort
        'From'          = $EmailOpts.From
        'To'            = $EmailOpts.To
        'SSL'           = $EmailOpts.EnableSSL
        'Use Auth'      = $EmailOpts.UseAuthentication
    }
    [void]$sb.AppendLine(($emailData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>Email notifications not configured or cmdlet unavailable.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Footer ──
[void]$sb.AppendLine(@"
</div></div>
<div class="report-footer">
  <span>Veeam B&amp;R v13 As-Built Report v2.0</span>
  <span>$ReportDate</span>
  <span>$($BkpServer.Name ?? 'Unknown Server')</span>
</div>
<script>$JS</script>
</body></html>
"@)

# ─── 6. WRITE OUTPUT ─────────────────────────────────────────────────────────

New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
$FilePath = Join-Path $OutputPath "Veeam-v13-AsBuilt-$ReportDateISO.html"
$sb.ToString() | Out-File -FilePath $FilePath -Encoding UTF8 -Force

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  ✅  Report saved to: $FilePath" -ForegroundColor Green
if ($HealthCheck -and $Warnings.Count -gt 0) {
    Write-Host "  ⚠   $($Warnings.Count) health warning(s) found – review the report." -ForegroundColor Yellow
}
Write-Host "  📄  Open in any browser. Use Ctrl+P to print / export PDF." -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
