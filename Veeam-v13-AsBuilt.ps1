# =============================================================================
# VEEAM BACKUP & REPLICATION v13 - AS-BUILT REPORT GENERATOR
# Version 3.0 | Redesign Group Branded | Full Coverage
#
# Produced by The Redesign Group - Global Technology & Cybersecurity Consulting
# https://redesign-group.com | Data Protection Practice
#
# Usage:
#   pwsh -File Veeam-v13-AsBuilt-Redesign.ps1
#   pwsh -File Veeam-v13-AsBuilt-Redesign.ps1 -OutputPath "D:\Reports" -HealthCheck
#   pwsh -File Veeam-v13-AsBuilt-Redesign.ps1 -CustomerName "Acme Corp" -HealthCheck
#
# Requirements:
#   - PowerShell 7.2+
#   - Veeam Backup & Replication v13 Console installed on this machine
#   - Veeam.Backup.PowerShell module (auto-loaded by Veeam installer)
#   - Run as a user with Veeam Backup Administrator role
# =============================================================================

#Requires -Version 7.2

param(
    [string]$OutputPath   = "C:\AsBuiltReports",
    [string]$ReportTitle  = "Veeam Backup & Replication v13 - As-Built Report",
    [string]$CustomerName = "Customer",          # Client/customer name shown in header
    [string]$PreparedBy   = "The Redesign Group",
    [switch]$HealthCheck,
    [switch]$SkipJobDetails
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
    $cls = $map[$Status]; if (-not $cls) { $cls = 'badge-neutral' }
    return "<span class='badge $cls'>$Status</span>"
}

function Add-HealthWarning { param([string]$Message); $script:Warnings.Add($Message) }

# Resolve IP address for a given hostname (returns IP string or 'Unresolvable')
function Resolve-HostIP {
    param([string]$Hostname)
    if ([string]::IsNullOrWhiteSpace($Hostname)) { return 'N/A' }
    try {
        $result = [System.Net.Dns]::GetHostAddresses($Hostname) |
                  Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                  Select-Object -First 1
        if ($result) { return $result.IPAddressToString } else { return 'Unresolvable' }
    } catch { return 'Unresolvable' }
}

# ─── 2. CSS / JS ─────────────────────────────────────────────────────────────

$CSS = @'
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&family=DM+Sans:wght@300;400;500;600&display=swap');

/* ── Redesign Group Brand Tokens ── */
:root {
    /* Core palette - derived from redesign-group.com */
    --rg-black:       #0a0a0a;
    --rg-dark:        #111111;
    --rg-dark-2:      #181818;
    --rg-dark-3:      #222222;
    --rg-border:      #2a2a2a;
    --rg-border-lt:   #333333;

    /* Redesign accent - their gradient goes teal->green */
    --rg-teal:        #00c4a0;
    --rg-green:       #4ade80;
    --rg-accent:      #00c4a0;
    --rg-accent-glow: rgba(0,196,160,0.18);

    /* Text */
    --rg-text:        #f0f0f0;
    --rg-text-2:      #a0a0a0;
    --rg-text-3:      #606060;

    /* Semantic */
    --success:   #4ade80;
    --warn:      #fbbf24;
    --fail:      #f87171;
    --info:      #60a5fa;

    /* Veeam reference colour */
    --veeam:     #00b4d8;
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
    font-family: 'DM Sans', sans-serif;
    background: var(--rg-black);
    color: var(--rg-text);
    font-size: 13px;
    line-height: 1.65;
    -webkit-font-smoothing: antialiased;
}

/* ────────────────────────────────────────────────
   HEADER
──────────────────────────────────────────────── */
.report-header {
    background: var(--rg-dark);
    border-bottom: 1px solid var(--rg-border);
    padding: 0;
    position: relative;
    overflow: hidden;
}

/* Subtle animated gradient mesh - nod to RG hero */
.report-header::before {
    content: '';
    position: absolute;
    inset: 0;
    background:
        radial-gradient(ellipse 60% 80% at 80% 50%, rgba(0,196,160,0.07) 0%, transparent 70%),
        radial-gradient(ellipse 40% 60% at 10% 30%, rgba(74,222,128,0.04) 0%, transparent 60%);
    pointer-events: none;
}

/* Dot-grid texture */
.report-header::after {
    content: '';
    position: absolute;
    inset: 0;
    background-image: radial-gradient(circle, rgba(255,255,255,0.04) 1px, transparent 1px);
    background-size: 24px 24px;
    pointer-events: none;
}

.header-inner {
    position: relative;
    z-index: 1;
    padding: 48px 60px 40px;
    display: grid;
    grid-template-columns: 1fr auto;
    align-items: start;
    gap: 32px;
}

/* Redesign Group wordmark (SVG inline text) */
.rg-logo {
    display: flex;
    flex-direction: column;
    gap: 2px;
}
.rg-logo-mark {
    font-family: 'Syne', sans-serif;
    font-weight: 800;
    font-size: 13px;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: var(--rg-accent);
}
.rg-logo-sub {
    font-family: 'DM Sans', sans-serif;
    font-weight: 300;
    font-size: 10px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--rg-text-2);
}

.header-left { display: flex; flex-direction: column; gap: 20px; }

.header-chips {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
}
.chip {
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    font-weight: 500;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 4px 10px;
    border-radius: 2px;
    border: 1px solid;
}
.chip-veeam  { color: var(--veeam);   border-color: rgba(0,180,216,0.4);  background: rgba(0,180,216,0.06); }
.chip-rg     { color: var(--rg-accent); border-color: rgba(0,196,160,0.4); background: rgba(0,196,160,0.06); }
.chip-asbuilt{ color: var(--rg-text-2); border-color: var(--rg-border-lt); background: transparent; }

.report-header h1 {
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    font-size: 30px;
    letter-spacing: -0.5px;
    color: var(--rg-text);
    line-height: 1.15;
}
.report-header h1 em {
    font-style: normal;
    color: var(--rg-accent);
}

.header-meta {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    color: var(--rg-text-2);
    display: flex;
    flex-wrap: wrap;
    gap: 8px 20px;
}
.header-meta span { display: flex; align-items: center; gap: 5px; }
.header-meta span::before {
    content: '';
    display: inline-block;
    width: 4px;
    height: 4px;
    border-radius: 50%;
    background: var(--rg-accent);
}

/* Header right column: customer block */
.header-right {
    text-align: right;
    display: flex;
    flex-direction: column;
    gap: 8px;
    align-items: flex-end;
}
.header-right .customer-label {
    font-size: 9px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--rg-text-3);
    font-family: 'DM Mono', monospace;
}
.header-right .customer-name {
    font-family: 'Syne', sans-serif;
    font-size: 18px;
    font-weight: 700;
    color: var(--rg-text);
}
.header-right .prepared-by {
    font-size: 10px;
    color: var(--rg-text-2);
    font-family: 'DM Mono', monospace;
}

/* Accent bar at top */
.header-accent-bar {
    height: 3px;
    background: linear-gradient(90deg, var(--rg-teal) 0%, var(--rg-green) 100%);
}

/* ────────────────────────────────────────────────
   LAYOUT
──────────────────────────────────────────────── */
.layout {
    display: grid;
    grid-template-columns: 230px 1fr;
    min-height: calc(100vh - 220px);
}

/* ── Sidebar TOC ── */
.toc {
    background: var(--rg-dark);
    border-right: 1px solid var(--rg-border);
    padding: 28px 0;
    position: sticky;
    top: 0;
    height: 100vh;
    overflow-y: auto;
}
.toc-header {
    padding: 0 20px 16px;
    border-bottom: 1px solid var(--rg-border);
    margin-bottom: 12px;
}
.toc-header span {
    font-size: 9px;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    color: var(--rg-text-3);
    font-family: 'DM Mono', monospace;
}
.toc a {
    display: flex;
    align-items: center;
    gap: 10px;
    color: var(--rg-text-2);
    text-decoration: none;
    padding: 7px 20px;
    font-size: 12px;
    font-family: 'DM Sans', sans-serif;
    font-weight: 400;
    transition: all 0.15s;
    border-left: 2px solid transparent;
}
.toc a:hover  { color: var(--rg-text); background: rgba(255,255,255,0.03); border-left-color: var(--rg-border-lt); }
.toc a.active { color: var(--rg-accent); background: var(--rg-accent-glow); border-left-color: var(--rg-accent); }
.toc a .toc-num {
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    color: var(--rg-text-3);
    width: 16px;
    flex-shrink: 0;
}
.toc a.active .toc-num { color: var(--rg-accent); }

/* ── Main content ── */
.main-content { padding: 40px 52px; max-width: 1200px; }

/* ────────────────────────────────────────────────
   SECTIONS
──────────────────────────────────────────────── */
.section {
    margin-bottom: 52px;
    scroll-margin-top: 24px;
}
.section-header {
    display: flex;
    align-items: center;
    gap: 14px;
    padding-bottom: 14px;
    margin-bottom: 24px;
    border-bottom: 1px solid var(--rg-border);
}
.section-num {
    font-family: 'DM Mono', monospace;
    font-size: 9px;
    font-weight: 500;
    color: var(--rg-accent);
    background: var(--rg-accent-glow);
    border: 1px solid rgba(0,196,160,0.25);
    width: 28px;
    height: 28px;
    border-radius: 4px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    letter-spacing: 0;
}
.section h2 {
    font-family: 'Syne', sans-serif;
    font-size: 17px;
    font-weight: 700;
    color: var(--rg-text);
    letter-spacing: -0.2px;
}

.subsection { margin: 24px 0; }
.subsection h3 {
    font-family: 'DM Sans', sans-serif;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: var(--rg-text-2);
    margin-bottom: 12px;
    display: flex;
    align-items: center;
    gap: 10px;
}
.subsection h3::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--rg-border);
}

/* ────────────────────────────────────────────────
   TABLES
──────────────────────────────────────────────── */
.table-wrap { overflow-x: auto; border-radius: 6px; border: 1px solid var(--rg-border); }

table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
    font-family: 'DM Mono', monospace;
    background: var(--rg-dark-2);
}
th {
    background: var(--rg-dark-3);
    color: var(--rg-text-3);
    font-size: 9px;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 10px 14px;
    text-align: left;
    border-bottom: 1px solid var(--rg-border);
    white-space: nowrap;
}
td {
    padding: 9px 14px;
    border-bottom: 1px solid var(--rg-border);
    vertical-align: top;
    color: var(--rg-text);
    word-break: break-word;
    max-width: 380px;
    font-family: 'DM Mono', monospace;
    font-size: 12px;
}
tr:last-child td { border-bottom: none; }
tr:hover td { background: rgba(255,255,255,0.015); }

/* IP address cells */
td.ip { color: var(--rg-accent); font-weight: 500; }

/* ────────────────────────────────────────────────
   BADGES
──────────────────────────────────────────────── */
.badge {
    display: inline-block;
    font-size: 9px;
    font-family: 'DM Mono', monospace;
    font-weight: 500;
    padding: 2px 8px;
    border-radius: 2px;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    border: 1px solid;
}
.badge-success { background: rgba(74,222,128,0.08);  color: var(--success); border-color: rgba(74,222,128,0.25); }
.badge-warn    { background: rgba(251,191,36,0.08);  color: var(--warn);    border-color: rgba(251,191,36,0.25); }
.badge-fail    { background: rgba(248,113,113,0.08); color: var(--fail);    border-color: rgba(248,113,113,0.25); }
.badge-info    { background: rgba(96,165,250,0.08);  color: var(--info);    border-color: rgba(96,165,250,0.25); }
.badge-neutral { background: rgba(96,96,96,0.12);    color: var(--rg-text-2); border-color: var(--rg-border-lt); }

/* ────────────────────────────────────────────────
   SUMMARY CARDS
──────────────────────────────────────────────── */
.summary-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(148px, 1fr));
    gap: 14px;
    margin-bottom: 40px;
}
.summary-card {
    background: var(--rg-dark-2);
    border: 1px solid var(--rg-border);
    border-radius: 6px;
    padding: 20px 16px 16px;
    position: relative;
    overflow: hidden;
    transition: border-color 0.2s;
}
.summary-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, var(--rg-teal), var(--rg-green));
    opacity: 0.6;
}
.summary-card:hover { border-color: rgba(0,196,160,0.3); }
.summary-card .val {
    font-family: 'Syne', sans-serif;
    font-size: 34px;
    font-weight: 800;
    color: var(--rg-text);
    line-height: 1;
}
.summary-card .lbl {
    font-size: 9px;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: var(--rg-text-3);
    margin-top: 8px;
    font-family: 'DM Mono', monospace;
}

/* ────────────────────────────────────────────────
   HEALTH BANNER
──────────────────────────────────────────────── */
.health-banner {
    background: rgba(248,113,113,0.06);
    border: 1px solid rgba(248,113,113,0.25);
    border-left: 3px solid var(--fail);
    border-radius: 4px;
    padding: 18px 22px;
    margin-bottom: 32px;
}
.health-banner h3 {
    color: var(--fail);
    font-size: 10px;
    letter-spacing: 2px;
    text-transform: uppercase;
    font-family: 'DM Mono', monospace;
    margin-bottom: 12px;
}
.health-banner ul { padding-left: 18px; }
.health-banner li { color: var(--fail); font-size: 12px; margin-bottom: 5px; font-family: 'DM Mono', monospace; }

/* ────────────────────────────────────────────────
   JOB CARDS
──────────────────────────────────────────────── */
.job-card {
    background: var(--rg-dark-2);
    border: 1px solid var(--rg-border);
    border-radius: 6px;
    margin-bottom: 12px;
    overflow: hidden;
    transition: border-color 0.15s;
}
.job-card:hover { border-color: var(--rg-border-lt); }
.job-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: var(--rg-dark-3);
    padding: 13px 18px;
    cursor: pointer;
    user-select: none;
    gap: 12px;
}
.job-card-header:hover { background: rgba(255,255,255,0.025); }
.job-card-title {
    font-family: 'DM Sans', sans-serif;
    font-weight: 600;
    font-size: 13px;
    color: var(--rg-text);
}
.job-card-body { padding: 18px; display: none; }
.job-card.open .job-card-body { display: block; }
.job-card-meta {
    display: flex;
    gap: 18px;
    flex-wrap: wrap;
    margin-bottom: 14px;
    font-size: 11px;
    color: var(--rg-text-2);
    font-family: 'DM Mono', monospace;
}
.job-card-meta span { display: flex; align-items: center; gap: 4px; }
.job-card-meta span::before {
    content: '';
    width: 3px; height: 3px;
    border-radius: 50%;
    background: var(--rg-accent);
    flex-shrink: 0;
}

.inner-label {
    margin: 14px 0 8px;
    font-size: 9px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--rg-text-3);
    font-family: 'DM Mono', monospace;
    font-weight: 500;
}

/* ────────────────────────────────────────────────
   AUDIT TABLE (backup copy audit)
──────────────────────────────────────────────── */
.audit-table-wrap {
    background: var(--rg-dark);
    border: 1px solid var(--rg-border);
    border-left: 3px solid var(--rg-accent);
    border-radius: 4px;
    overflow: hidden;
    margin-bottom: 16px;
}
.audit-table-wrap table {
    background: transparent;
}

/* ────────────────────────────────────────────────
   MISC
──────────────────────────────────────────────── */
.empty { color: var(--rg-text-3); font-style: italic; font-size: 12px; padding: 10px 0; font-family: 'DM Mono', monospace; }
.error { color: var(--fail); font-size: 12px; font-family: 'DM Mono', monospace; padding: 8px 0; }
.note  { color: var(--rg-text-3); font-size: 11px; margin-top: 6px; font-family: 'DM Mono', monospace; }
hr     { border: none; border-top: 1px solid var(--rg-border); margin: 36px 0; }
pre    { font-family: 'DM Mono', monospace; font-size: 11px; color: var(--rg-text-2); white-space: pre-wrap; }

/* Section separator pill */
.section-pill {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: var(--rg-accent-glow);
    border: 1px solid rgba(0,196,160,0.2);
    border-radius: 20px;
    padding: 3px 10px 3px 6px;
    font-size: 10px;
    color: var(--rg-accent);
    font-family: 'DM Mono', monospace;
    margin-bottom: 18px;
}
.section-pill::before {
    content: '';
    width: 6px; height: 6px;
    border-radius: 50%;
    background: var(--rg-accent);
}

/* ────────────────────────────────────────────────
   FOOTER
──────────────────────────────────────────────── */
.report-footer {
    background: var(--rg-dark);
    border-top: 1px solid var(--rg-border);
    padding: 20px 52px;
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 12px;
    font-family: 'DM Mono', monospace;
    font-size: 10px;
    color: var(--rg-text-3);
}
.report-footer .footer-brand {
    display: flex;
    flex-direction: column;
    gap: 2px;
}
.report-footer .footer-brand strong {
    color: var(--rg-accent);
    font-size: 11px;
}
.report-footer .footer-center { text-align: center; align-self: center; }
.report-footer .footer-right  { text-align: right; align-self: center; }
.footer-accent-bar {
    height: 2px;
    background: linear-gradient(90deg, var(--rg-teal) 0%, var(--rg-green) 100%);
    opacity: 0.5;
}

/* ────────────────────────────────────────────────
   PRINT
──────────────────────────────────────────────── */
@media print {
    body { background: #fff; color: #000; }
    .toc { display: none; }
    .layout { display: block; }
    .main-content { padding: 0; max-width: none; }
    .job-card-body { display: block !important; }
    th { background: #e8e8e8; color: #333; }
    td { color: #111; }
    .badge { border: 1px solid #999; background: #eee; color: #000; }
    .report-header { background: #111; }
    .section { page-break-inside: avoid; }
    table { border: 1px solid #ccc; }
    .summary-card { border: 1px solid #ccc; background: #f9f9f9; }
    .summary-card .val { color: #000; }
}
'@

$JS = @'
// Job card toggle
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
            const link = document.querySelector(`.toc a[href="#${e.target.id}"]`);
            if (link) link.classList.add('active');
        }
    });
}, { rootMargin: '-20% 0px -70% 0px' });
sections.forEach(s => observer.observe(s));

// Wrap all tables that aren't already wrapped
document.querySelectorAll('table').forEach(t => {
    if (!t.parentElement.classList.contains('table-wrap') &&
        !t.parentElement.classList.contains('audit-table-wrap')) {
        const wrap = document.createElement('div');
        wrap.className = 'table-wrap';
        t.parentNode.insertBefore(wrap, t);
        wrap.appendChild(t);
    }
});

// Colour IP address cells
document.querySelectorAll('td').forEach(td => {
    if (/^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(td.textContent.trim())) {
        td.classList.add('ip');
    }
});
'@

# ─── 3. DATA COLLECTION ──────────────────────────────────────────────────────

Write-Host "⏳  Collecting Veeam configuration data..." -ForegroundColor Cyan

# ── 3.1 Backup Server ──
$BkpServer = $null
try { $BkpServer = Get-VBRServer | Where-Object { $_.IsLocal -or $_.Type -eq 'Local' } | Select-Object -First 1 }
catch { Write-Warning "Could not retrieve local server info: $_" }

# ── 3.2 License ──
$License = $null
try {
    $License = Get-VBRInstalledLicense
    if ($HealthCheck) {
        if ($License.Status -eq 'Expired') { Add-HealthWarning "License is EXPIRED." }
        elseif ($License.ExpirationDate -lt (Get-Date).AddDays(30)) {
            Add-HealthWarning "License expires within 30 days: $($License.ExpirationDate.ToString('yyyy-MM-dd'))"
        }
    }
} catch { Write-Warning "License query failed: $_" }

# ── 3.3 Proxies (with IP resolution) ──
$ViProxies  = @(); try { $ViProxies  = @(Get-VBRViProxy)  } catch { Write-Warning "VI Proxies: $_" }
$HvProxies  = @(); try { $HvProxies  = @(Get-VBRHvProxy)  } catch { Write-Warning "HV Proxies: $_" }
$NasProxies = @(); try { $NasProxies = @(Get-VBRNASProxyServer) } catch { Write-Warning "NAS Proxies: $_" }

# Resolve proxy IPs at collection time
$ViProxyData = $ViProxies | ForEach-Object {
    $hostName = $_.Host.Name
    [PSCustomObject]@{
        Name             = $_.Name
        Host             = $hostName
        'IP Address'     = Resolve-HostIP $hostName
        'Transport Mode' = $_.Options.TransportMode
        'Max Tasks'      = $_.MaxTasksCount
        Status           = if ($_.IsDisabled) { 'Disabled' } else { 'Enabled' }
    }
}
$HvProxyData = $HvProxies | ForEach-Object {
    $hostName = $_.Host.Name
    [PSCustomObject]@{
        Name         = $_.Name
        Host         = $hostName
        'IP Address' = Resolve-HostIP $hostName
        'Max Tasks'  = $_.MaxTasksCount
    }
}
$NasProxyData = $NasProxies | ForEach-Object {
    $srvName = $_.Server.Name
    [PSCustomObject]@{
        Name         = $_.Name
        Server       = $srvName
        'IP Address' = Resolve-HostIP $srvName
        Description  = $_.Description
    }
}

# ── 3.4 Repositories (with IP resolution) ──
$Repos = @()
try {
    $Repos = @(Get-VBRBackupRepository)
    if ($HealthCheck) {
        foreach ($r in $Repos) {
            if ($r.TotalSpace -gt 0) {
                $pctFree = [math]::Round(($r.FreeSpace / $r.TotalSpace) * 100, 1)
                if ($pctFree -lt 10)  { Add-HealthWarning "Repository '$($r.Name)' critically low on space ($pctFree% free)." }
                elseif ($pctFree -lt 20) { Add-HealthWarning "Repository '$($r.Name)' below 20% free space ($pctFree%)." }
            }
        }
    }
} catch { Write-Warning "Repositories: $_" }

$RepoData = $Repos | ForEach-Object {
    $freeGB  = [math]::Round($_.FreeSpace  / 1GB, 1)
    $totalGB = [math]::Round($_.TotalSpace / 1GB, 1)
    $pctFree = if ($totalGB -gt 0) { [math]::Round(($freeGB / $totalGB) * 100, 1) } else { 'N/A' }
    $hostName = $_.Host.Name
    [PSCustomObject]@{
        Name           = $_.Name
        Type           = $_.Type
        Host           = $hostName
        'IP Address'   = Resolve-HostIP $hostName
        Path           = $_.Path
        'Total (GB)'   = $totalGB
        'Free (GB)'    = $freeGB
        'Free %'       = $pctFree
        'Per-VM'       = $_.UsePerVMBackupFiles
        Immutability   = $_.ImmutabilityEnabled
    }
}

# ── 3.5 SOBR ──
$SOBRs = @(); try { $SOBRs = @(Get-VBRBackupRepository -ScaleOut) } catch { Write-Warning "SOBR: $_" }

# ── 3.6 External Repos ──
$ExtRepos = @(); try { $ExtRepos = @(Get-VBRExternalRepository) } catch { Write-Warning "External Repos: $_" }

# ── 3.7 All Jobs ──
$AllJobs = @(); try { $AllJobs = @(Get-VBRJob) } catch { Write-Warning "Jobs: $_" }

if ($HealthCheck) {
    foreach ($j in $AllJobs) {
        if ($j.IsScheduleEnabled) {
            try {
                $lastSession = Get-VBRJobSession -Job $j -Last 1 -ErrorAction SilentlyContinue
                if ($lastSession -and $lastSession.Result -eq 'Failed') {
                    Add-HealthWarning "Job '$($j.Name)' last run FAILED ($(($lastSession.EndTime).ToString('yyyy-MM-dd HH:mm')))."
                }
            } catch { }
        }
    }
}

# ── 3.8 Job type splits ──
$BackupJobs = $AllJobs | Where-Object { $_.JobType -eq 'Backup' }
$CopyJobs   = $AllJobs | Where-Object { $_.JobType -eq 'BackupSync' }
$ReplJobs   = $AllJobs | Where-Object { $_.JobType -in @('Replica','SimpleTransactionLog') }

# ── 3.9 NAS Backup Jobs ──
$NasJobs = @(); try { $NasJobs = @(Get-VBRNASBackupJob) } catch { Write-Warning "NAS Backup Jobs: $_" }

# ── 3.10 CDP Policies ──
$CDPPolicies = @(); try { $CDPPolicies = @(Get-VBRCDPPolicy) } catch { Write-Warning "CDP: $_" }

# ── 3.11 SureBackup ──
$SBJobs    = @(); try { $SBJobs    = @(Get-VBRSureBackupJob)     } catch { Write-Warning "SureBackup: $_" }
$VLabs     = @(); try { $VLabs     = @(Get-VBRVirtualLab)        } catch { Write-Warning "Virtual Labs: $_" }
$AppGroups = @(); try { $AppGroups = @(Get-VBRApplicationGroup)  } catch { Write-Warning "App Groups: $_" }

# ── 3.12 Tape ──
$TapeLibraries  = @(); try { $TapeLibraries  = @(Get-VBRTapeLibrary)   } catch { Write-Warning "Tape Libraries: $_" }
$TapeMediaPools = @(); try { $TapeMediaPools = @(Get-VBRTapeMediaPool)  } catch { Write-Warning "Tape Media Pools: $_" }
$TapeJobs       = @(); try { $TapeJobs       = @(Get-VBRTapeJob)        } catch { Write-Warning "Tape Jobs: $_" }

# ── 3.13 Cloud Connect ──
$CloudTenants  = @(); try { $CloudTenants  = @(Get-VBRCloudTenant)      } catch { Write-Warning "Cloud Tenants: $_" }
$CloudHardware = @(); try { $CloudHardware = @(Get-VBRCloudHardwarePlan)} catch { Write-Warning "Cloud HW Plans: $_" }

# ── 3.14 vSphere / Hyper-V Infrastructure ──
$ViServers = @(); try { $ViServers = @(Get-VBRServer | Where-Object { $_.Type -in @('VC','ESXi') }) } catch { Write-Warning "VI Servers: $_" }
$HvServers = @(); try { $HvServers = @(Get-VBRServer | Where-Object { $_.Type -eq 'HvServer' })    } catch { Write-Warning "HV Servers: $_" }

# ── 3.15 NAS File Shares ──
$FileShares = @(); try { $FileShares = @(Get-VBRNASFileShare) } catch { Write-Warning "File Shares: $_" }

# ── 3.16 Credentials ──
$Credentials = @(); try { $Credentials = @(Get-VBRCredentials | Select-Object Name, Description) } catch { Write-Warning "Credentials: $_" }

# ── 3.17 Notifications ──
$NotifOpts = $null; try { $NotifOpts = Get-VBRNotificationOptions } catch { }
$EmailOpts = $null; try { $EmailOpts = Get-VBREmailOptions         } catch { }

# ── 3.18 Backup Copy Job Session Audit ──
# Pull last session result for each copy job
$CopyJobAudit = $CopyJobs | ForEach-Object {
    $job = $_
    $lastSession = $null
    try { $lastSession = Get-VBRJobSession -Job $job -Last 1 -ErrorAction SilentlyContinue } catch { }
    $srcRepo = try { ($job.GetSourceRepository()).Name } catch { 'N/A' }
    [PSCustomObject]@{
        Name           = $job.Name
        Enabled        = $job.IsScheduleEnabled
        'Source Repo'  = $srcRepo
        'Target Repo'  = try { $job.GetTargetRepository()?.Name } catch { 'N/A' }
        'Last Result'  = if ($lastSession) { $lastSession.Result } else { 'No Sessions' }
        'Last Run'     = if ($lastSession) { $lastSession.EndTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
        'Transfer (GB)'= if ($lastSession) { [math]::Round($lastSession.Progress.TransferedSize / 1GB, 2) } else { 0 }
        Description    = $job.Description
    }
}

Write-Host "✅  Data collection complete." -ForegroundColor Green

# ─── 4. SUMMARY VALUES ───────────────────────────────────────────────────────

$totalJobs    = $AllJobs.Count + $NasJobs.Count + $TapeJobs.Count
$enabledJobs  = ($AllJobs | Where-Object { $_.IsScheduleEnabled }).Count
$totalRepos   = $Repos.Count + $SOBRs.Count
$totalProxies = $ViProxies.Count + $HvProxies.Count + $NasProxies.Count

# ─── 5. HTML GENERATION ──────────────────────────────────────────────────────

Write-Host "⏳  Building HTML report..." -ForegroundColor Cyan
$sb = [System.Text.StringBuilder]::new()

[void]$sb.AppendLine(@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$ReportTitle - $CustomerName</title>
<style>$CSS</style>
</head>
<body>
"@)

# ── Header ──
[void]$sb.AppendLine(@"
<div class="header-accent-bar"></div>
<div class="report-header">
  <div class="header-inner">
    <div class="header-left">
      <div class="rg-logo">
        <span class="rg-logo-mark">The Redesign Group</span>
        <span class="rg-logo-sub">Technology &amp; Cybersecurity Consulting</span>
      </div>
      <div class="header-chips">
        <span class="chip chip-veeam">Veeam B&amp;R v13</span>
        <span class="chip chip-rg">Data Protection</span>
        <span class="chip chip-asbuilt">As-Built Report</span>
      </div>
      <h1>$([System.Web.HttpUtility]::HtmlEncode($ReportTitle))</h1>
      <div class="header-meta">
        <span>Generated: $ReportDate</span>
        <span>Server: $($BkpServer.Name ?? 'Unknown')</span>
        <span>Module: $($VeeamModule.Version)</span>
      </div>
    </div>
    <div class="header-right">
      <span class="customer-label">Prepared for</span>
      <span class="customer-name">$([System.Web.HttpUtility]::HtmlEncode($CustomerName))</span>
      <span class="prepared-by">Prepared by: $([System.Web.HttpUtility]::HtmlEncode($PreparedBy))</span>
    </div>
  </div>
</div>
"@)

# ── Layout + TOC ──
[void]$sb.AppendLine('<div class="layout">')
[void]$sb.AppendLine(@'
<nav class="toc">
  <div class="toc-header"><span>Contents</span></div>
  <a href="#sec-summary"><span class="toc-num">◈</span>Summary Dashboard</a>
  <a href="#sec-server"><span class="toc-num">01</span>Backup Server</a>
  <a href="#sec-license"><span class="toc-num">02</span>License</a>
  <a href="#sec-infra"><span class="toc-num">03</span>Infrastructure</a>
  <a href="#sec-repos"><span class="toc-num">04</span>Repositories</a>
  <a href="#sec-jobs"><span class="toc-num">05</span>Backup Jobs</a>
  <a href="#sec-copyjobs"><span class="toc-num">06</span>Backup Copy Jobs</a>
  <a href="#sec-nas"><span class="toc-num">07</span>NAS Backup</a>
  <a href="#sec-replication"><span class="toc-num">08</span>Replication &amp; CDP</a>
  <a href="#sec-surebackup"><span class="toc-num">09</span>SureBackup</a>
  <a href="#sec-tape"><span class="toc-num">10</span>Tape</a>
  <a href="#sec-cloud"><span class="toc-num">11</span>Cloud Connect</a>
  <a href="#sec-inventory"><span class="toc-num">12</span>Inventory</a>
  <a href="#sec-global"><span class="toc-num">13</span>Global Settings</a>
</nav>
'@)

[void]$sb.AppendLine('<div class="main-content">')

# ── Health Warnings ──
if ($HealthCheck -and $Warnings.Count -gt 0) {
    [void]$sb.AppendLine('<div class="health-banner"><h3>▲ Health Check Warnings</h3><ul>')
    foreach ($w in $Warnings) { [void]$sb.AppendLine("<li>$([System.Web.HttpUtility]::HtmlEncode($w))</li>") }
    [void]$sb.AppendLine('</ul></div>')
}

# ── Summary Dashboard ──
[void]$sb.AppendLine('<div class="section" id="sec-summary"><div class="section-header"><span class="section-num">◈</span><h2>Summary Dashboard</h2></div>')
[void]$sb.AppendLine('<div class="summary-grid">')
$cards = @(
    @{ Val = $totalJobs;                                              Lbl = "Total Jobs" },
    @{ Val = $enabledJobs;                                            Lbl = "Scheduled" },
    @{ Val = $BackupJobs.Count;                                       Lbl = "Backup Jobs" },
    @{ Val = $CopyJobs.Count;                                         Lbl = "Copy Jobs" },
    @{ Val = $totalRepos;                                             Lbl = "Repositories" },
    @{ Val = $totalProxies;                                           Lbl = "Proxies" },
    @{ Val = ($ViServers.Count + $HvServers.Count);                   Lbl = "Hypervisors" },
    @{ Val = $TapeLibraries.Count;                                    Lbl = "Tape Libraries" }
)
foreach ($c in $cards) {
    [void]$sb.AppendLine("<div class='summary-card'><div class='val'>$($c.Val)</div><div class='lbl'>$($c.Lbl)</div></div>")
}
[void]$sb.AppendLine('</div></div>')

# ── Section 1: Backup Server ──
[void]$sb.AppendLine('<div class="section" id="sec-server"><div class="section-header"><span class="section-num">01</span><h2>Backup Server</h2></div>')
if ($BkpServer) {
    $serverData = [PSCustomObject]@{
        Name         = $BkpServer.Name
        'IP Address' = Resolve-HostIP $BkpServer.Name
        DNS          = $BkpServer.DNSName
        Type         = $BkpServer.Type
        Description  = $BkpServer.Description
        'Connected'  = $BkpServer.IsConnected
    }
    [void]$sb.AppendLine(($serverData | ConvertTo-HtmlTable))
} else {
    [void]$sb.AppendLine("<p class='error'>Could not retrieve local server information.</p>")
}
[void]$sb.AppendLine('</div>')

# ── Section 2: License ──
[void]$sb.AppendLine('<div class="section" id="sec-license"><div class="section-header"><span class="section-num">02</span><h2>License</h2></div>')
if ($License) {
    $licData = [PSCustomObject]@{
        Edition                  = $License.Edition
        Type                     = $License.LicenseType
        Status                   = $License.Status
        'Expiration Date'        = $License.ExpirationDate?.ToString('yyyy-MM-dd')
        'Support Expiry'         = $License.SupportExpirationDate?.ToString('yyyy-MM-dd')
        'Used / Total Licenses'  = "$($License.UsedLicensesNumber) / $($License.TotalLicensesNumber)"
        'Used Sockets'           = $License.UsedSocketsNumber
    }
    [void]$sb.AppendLine(($licData | ConvertTo-HtmlTable))
} else {
    [void]$sb.AppendLine("<p class='error'>Could not retrieve license information.</p>")
}
[void]$sb.AppendLine('</div>')

# ── Section 3: Infrastructure (Proxies) ──
[void]$sb.AppendLine('<div class="section" id="sec-infra"><div class="section-header"><span class="section-num">03</span><h2>Backup Infrastructure</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>VMware Backup Proxies</h3>')
if ($ViProxyData.Count -gt 0) {
    [void]$sb.AppendLine(($ViProxyData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No VMware proxies configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Hyper-V Off-Host Processing Servers</h3>')
if ($HvProxyData.Count -gt 0) {
    [void]$sb.AppendLine(($HvProxyData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No Hyper-V off-host proxies configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>NAS Backup Proxies</h3>')
if ($NasProxyData.Count -gt 0) {
    [void]$sb.AppendLine(($NasProxyData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No NAS proxies configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 4: Repositories ──
[void]$sb.AppendLine('<div class="section" id="sec-repos"><div class="section-header"><span class="section-num">04</span><h2>Repositories &amp; Storage</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Backup Repositories</h3>')
if ($RepoData.Count -gt 0) {
    [void]$sb.AppendLine(($RepoData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No backup repositories found.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Scale-Out Backup Repositories (SOBR)</h3>')
if ($SOBRs.Count -gt 0) {
    $sobrData = $SOBRs | ForEach-Object {
        [PSCustomObject]@{
            Name                  = $_.Name
            Policy                = $_.PolicyType
            'Active Extents'      = ($_.Extents | Where-Object { $_.IsActive } | Measure-Object).Count
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
[void]$sb.AppendLine('<div class="section" id="sec-jobs"><div class="section-header"><span class="section-num">05</span><h2>Backup Jobs</h2></div>')

if ($BackupJobs.Count -gt 0) {
    [void]$sb.AppendLine('<div class="subsection"><h3>Job Overview</h3>')
    $jobOverview = $BackupJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            Repository  = try { $_.GetTargetRepository()?.Name } catch { 'N/A' }
            Objects     = try { ($_.GetObjectsInJob() | Measure-Object).Count } catch { 0 }
            Description = $_.Description
        }
    }
    [void]$sb.AppendLine(($jobOverview | ConvertTo-HtmlTable))
    [void]$sb.AppendLine('</div>')

    if (-not $SkipJobDetails) {
        [void]$sb.AppendLine('<div class="subsection"><h3>Detailed Job Configuration</h3>')
        foreach ($job in $BackupJobs) {
            $opts  = $null; try { $opts  = $job.GetOptions()     } catch { }
            $sched = $null; try { $sched = $job.ScheduleOptions  } catch { }
            $objs  = @();   try { $objs  = @($job.GetObjectsInJob()) } catch { }
            $statusBadge = if ($job.IsScheduleEnabled) { Get-StatusBadge 'Success' } else { Get-StatusBadge 'Disabled' }

            [void]$sb.AppendLine("<div class='job-card'>")
            [void]$sb.AppendLine("<div class='job-card-header'><span class='job-card-title'>$([System.Web.HttpUtility]::HtmlEncode($job.Name))</span>$statusBadge</div>")
            [void]$sb.AppendLine("<div class='job-card-body'>")
            [void]$sb.AppendLine("<div class='job-card-meta'><span>Type: $($job.JobType)</span><span>Repo: $(try{$job.GetTargetRepository()?.Name}catch{'N/A'})</span><span>Objects: $($objs.Count)</span></div>")

            if ($objs.Count -gt 0) {
                [void]$sb.AppendLine("<div class='inner-label'>Source Objects</div>")
                $objData = $objs | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Type = $_.Type; Location = $_.Location } }
                [void]$sb.AppendLine(($objData | ConvertTo-HtmlTable -EmptyMessage "No source objects."))
            }

            if ($opts) {
                [void]$sb.AppendLine("<div class='inner-label'>Retention &amp; Storage</div>")
                $retData = [PSCustomObject]@{
                    'Retention Type'  = $opts.RetentionPolicy.Type
                    'Retention Count' = $opts.RetentionPolicy.Quantity
                    'GFS Enabled'     = $opts.RetentionPolicy.IsGFSEnabled
                    'Weekly GFS'      = $opts.RetentionPolicy.WeeklyFullSchedule?.RepeatCount
                    'Monthly GFS'     = $opts.RetentionPolicy.MonthlyFullSchedule?.RepeatCount
                    'Yearly GFS'      = $opts.RetentionPolicy.YearlyFullSchedule?.RepeatCount
                    'Dedup'           = $opts.JobOptions.EnableDeduplication
                    'Compression'     = $opts.JobOptions.CompressionType
                    'Encryption'      = $opts.JobOptions.EncryptionEnabled
                }
                [void]$sb.AppendLine(($retData | ConvertTo-HtmlTable))

                [void]$sb.AppendLine("<div class='inner-label'>Guest Processing</div>")
                $guestData = [PSCustomObject]@{
                    'App-Aware'       = $opts.JobOptions.GenerationPolicy?.IsAppAwareEnabled
                    'SQL Log Mode'    = $opts.JobOptions.GenerationPolicy?.SqlBackupMode
                    'Oracle Log Mode' = $opts.JobOptions.GenerationPolicy?.OracleBackupMode
                    'Index Files'     = $opts.JobOptions.GenerationPolicy?.FileSystemIndexingScope
                }
                [void]$sb.AppendLine(($guestData | ConvertTo-HtmlTable))
            }

            if ($sched) {
                [void]$sb.AppendLine("<div class='inner-label'>Schedule</div>")
                $schedData = [PSCustomObject]@{
                    Enabled        = $sched.Enabled
                    'Runs At'      = $sched.StartDateTime
                    'Type'         = $sched.Type
                    'Retry Count'  = $sched.RetryCount
                    'Retry Wait'   = "$($sched.RetryTimeout) min"
                    'Backup Window'= $sched.BackupWindowEnabled
                }
                [void]$sb.AppendLine(($schedData | ConvertTo-HtmlTable))
            }

            [void]$sb.AppendLine('</div></div>')
        }
        [void]$sb.AppendLine('</div>')
    }
} else {
    [void]$sb.AppendLine("<p class='empty'>No backup jobs found.</p>")
}
[void]$sb.AppendLine('</div>')

# ── Section 6: Backup Copy Jobs (FULL AUDIT) ──
[void]$sb.AppendLine('<div class="section" id="sec-copyjobs"><div class="section-header"><span class="section-num">06</span><h2>Backup Copy Jobs - Audit</h2></div>')
[void]$sb.AppendLine('<p class="section-pill">Full configuration &amp; last-session audit</p>')

if ($CopyJobs.Count -gt 0) {

    # Overview table
    [void]$sb.AppendLine('<div class="subsection"><h3>Copy Job Summary</h3>')
    [void]$sb.AppendLine(($CopyJobAudit | ConvertTo-HtmlTable))
    [void]$sb.AppendLine('</div>')

    # Health check: any failed copy jobs?
    $failedCopy = $CopyJobAudit | Where-Object { $_.'Last Result' -eq 'Failed' }
    if ($failedCopy.Count -gt 0) {
        [void]$sb.AppendLine('<div class="health-banner"><h3>▲ Failed Backup Copy Jobs</h3><ul>')
        foreach ($fc in $failedCopy) {
            [void]$sb.AppendLine("<li>$([System.Web.HttpUtility]::HtmlEncode($fc.Name)) — Last ran: $($fc.'Last Run')</li>")
        }
        [void]$sb.AppendLine('</ul></div>')
    }

    if (-not $SkipJobDetails) {
        [void]$sb.AppendLine('<div class="subsection"><h3>Detailed Copy Job Configuration</h3>')
        foreach ($job in $CopyJobs) {
            $opts  = $null; try { $opts  = $job.GetOptions()       } catch { }
            $sched = $null; try { $sched = $job.ScheduleOptions    } catch { }

            # Last 5 sessions for audit trail
            $sessions = @()
            try { $sessions = @(Get-VBRJobSession -Job $job -Last 5 -ErrorAction SilentlyContinue) } catch { }

            $statusBadge = if ($job.IsScheduleEnabled) { Get-StatusBadge 'Success' } else { Get-StatusBadge 'Disabled' }

            [void]$sb.AppendLine("<div class='job-card'>")
            [void]$sb.AppendLine("<div class='job-card-header'><span class='job-card-title'>$([System.Web.HttpUtility]::HtmlEncode($job.Name))</span>$statusBadge</div>")
            [void]$sb.AppendLine("<div class='job-card-body'>")
            [void]$sb.AppendLine("<div class='job-card-meta'>")
            [void]$sb.AppendLine("<span>Type: BackupCopy</span>")
            [void]$sb.AppendLine("<span>Target: $(try{$job.GetTargetRepository()?.Name}catch{'N/A'})</span>")
            [void]$sb.AppendLine("<span>Enabled: $($job.IsScheduleEnabled)</span>")
            [void]$sb.AppendLine("</div>")

            # Retention
            if ($opts) {
                [void]$sb.AppendLine("<div class='inner-label'>Retention &amp; Storage</div>")
                $retData = [PSCustomObject]@{
                    'Restore Points'  = $opts.RetentionPolicy.Quantity
                    'GFS Enabled'     = $opts.RetentionPolicy.IsGFSEnabled
                    'Weekly GFS'      = $opts.RetentionPolicy.WeeklyFullSchedule?.RepeatCount
                    'Monthly GFS'     = $opts.RetentionPolicy.MonthlyFullSchedule?.RepeatCount
                    'Yearly GFS'      = $opts.RetentionPolicy.YearlyFullSchedule?.RepeatCount
                    Encryption        = $opts.JobOptions.EncryptionEnabled
                }
                [void]$sb.AppendLine(($retData | ConvertTo-HtmlTable))
            }

            # Schedule
            if ($sched) {
                [void]$sb.AppendLine("<div class='inner-label'>Schedule</div>")
                $schedData = [PSCustomObject]@{
                    Enabled        = $sched.Enabled
                    Type           = $sched.Type
                    'Retry Count'  = $sched.RetryCount
                    'Retry Wait'   = "$($sched.RetryTimeout) min"
                    'Backup Window'= $sched.BackupWindowEnabled
                }
                [void]$sb.AppendLine(($schedData | ConvertTo-HtmlTable))
            }

            # Session audit trail
            [void]$sb.AppendLine("<div class='inner-label'>Last 5 Sessions (Audit Trail)</div>")
            if ($sessions.Count -gt 0) {
                [void]$sb.AppendLine("<div class='audit-table-wrap'>")
                $sessionData = $sessions | ForEach-Object {
                    [PSCustomObject]@{
                        'Start Time'     = $_.CreationTime.ToString('yyyy-MM-dd HH:mm')
                        'End Time'       = $_.EndTime.ToString('yyyy-MM-dd HH:mm')
                        Result           = $_.Result
                        'Transferred GB' = [math]::Round($_.Progress.TransferedSize / 1GB, 3)
                        'Duration (min)' = [math]::Round(($_.EndTime - $_.CreationTime).TotalMinutes, 1)
                        State            = $_.State
                    }
                }
                [void]$sb.AppendLine(($sessionData | ConvertTo-HtmlTable))
                [void]$sb.AppendLine("</div>")
            } else {
                [void]$sb.AppendLine("<p class='empty'>No session history found for this job.</p>")
            }

            [void]$sb.AppendLine('</div></div>')
        }
        [void]$sb.AppendLine('</div>')
    }
} else {
    [void]$sb.AppendLine("<p class='empty'>No backup copy jobs configured.</p>")
}
[void]$sb.AppendLine('</div>')

# ── Section 7: NAS Backup ──
[void]$sb.AppendLine('<div class="section" id="sec-nas"><div class="section-header"><span class="section-num">07</span><h2>NAS Backup</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>NAS Backup Jobs</h3>')
if ($NasJobs.Count -gt 0) {
    $nasJobData = $NasJobs | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Enabled     = $_.IsScheduleEnabled
            Repository  = $_.BackupRepository?.Name
            'Copy Repo' = $_.CopyBackupRepository?.Name
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
            Name       = $_.Name
            Path       = $_.Path
            Type       = $_.ShareType
            'Cache Repo'= $_.CacheRepository?.Name
        }
    }
    [void]$sb.AppendLine(($shareData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No file shares configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 8: Replication & CDP ──
[void]$sb.AppendLine('<div class="section" id="sec-replication"><div class="section-header"><span class="section-num">08</span><h2>Replication &amp; CDP</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Replication Jobs</h3>')
if ($ReplJobs.Count -gt 0) {
    $replData = $ReplJobs | ForEach-Object {
        [PSCustomObject]@{
            Name            = $_.Name
            Enabled         = $_.IsScheduleEnabled
            Target          = $_.Target
            'Restore Points'= try { $_.GetOptions()?.RetentionPolicy?.Quantity } catch { 'N/A' }
            Description     = $_.Description
        }
    }
    [void]$sb.AppendLine(($replData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No replication jobs configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>CDP Policies</h3>')
if ($CDPPolicies.Count -gt 0) {
    $cdpData = $CDPPolicies | ForEach-Object {
        [PSCustomObject]@{ Name = $_.Name; Enabled = $_.IsEnabled; Description = $_.Description }
    }
    [void]$sb.AppendLine(($cdpData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No CDP policies configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 9: SureBackup ──
[void]$sb.AppendLine('<div class="section" id="sec-surebackup"><div class="section-header"><span class="section-num">09</span><h2>SureBackup &amp; Recovery Verification</h2></div>')

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
        [PSCustomObject]@{ Name = $_.Name; VMs = $vms.Count; Description = $_.Description }
    }
    [void]$sb.AppendLine(($agData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No application groups configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>SureBackup Jobs</h3>')
if ($SBJobs.Count -gt 0) {
    $sbData = $SBJobs | ForEach-Object {
        [PSCustomObject]@{
            Name          = $_.Name
            Enabled       = $_.IsScheduleEnabled
            'Virtual Lab' = $_.VirtualLab?.Name
            Description   = $_.Description
        }
    }
    [void]$sb.AppendLine(($sbData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No SureBackup jobs configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 10: Tape ──
[void]$sb.AppendLine('<div class="section" id="sec-tape"><div class="section-header"><span class="section-num">10</span><h2>Tape Infrastructure</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Tape Libraries</h3>')
if ($TapeLibraries.Count -gt 0) {
    $tapeLibData = $TapeLibraries | ForEach-Object {
        [PSCustomObject]@{
            Name   = $_.Name
            State  = $_.State
            Model  = $_.Model
            Drives = ($_.Drives | Measure-Object).Count
            Slots  = $_.TotalSlots
        }
    }
    [void]$sb.AppendLine(($tapeLibData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No tape libraries configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Tape Media Pools</h3>')
if ($TapeMediaPools.Count -gt 0) {
    $tapePoolData = $TapeMediaPools | ForEach-Object {
        [PSCustomObject]@{
            Name          = $_.Name
            Type          = $_.Type
            'Media Count' = try { ($_.GetTapeMedias() | Measure-Object).Count } catch { 'N/A' }
            Retention     = $_.RetentionPolicy
        }
    }
    [void]$sb.AppendLine(($tapePoolData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No tape media pools configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Tape Jobs</h3>')
if ($TapeJobs.Count -gt 0) {
    $tapeJobData = $TapeJobs | ForEach-Object {
        [PSCustomObject]@{
            Name         = $_.Name
            Type         = $_.Type
            Enabled      = $_.Enabled
            'Media Pool' = $_.MediaPool?.Name
            Description  = $_.Description
        }
    }
    [void]$sb.AppendLine(($tapeJobData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No tape jobs configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 11: Cloud Connect ──
[void]$sb.AppendLine('<div class="section" id="sec-cloud"><div class="section-header"><span class="section-num">11</span><h2>Cloud Connect</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Cloud Tenants</h3>')
if ($CloudTenants.Count -gt 0) {
    $tenantData = $CloudTenants | ForEach-Object {
        [PSCustomObject]@{
            Name               = $_.Name
            Enabled            = $_.Enabled
            'Lease Expiration' = $_.LeaseExpirationDate?.ToString('yyyy-MM-dd')
            Description        = $_.Description
        }
    }
    [void]$sb.AppendLine(($tenantData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No cloud tenants configured.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Cloud Hardware Plans</h3>')
if ($CloudHardware.Count -gt 0) {
    [void]$sb.AppendLine(($CloudHardware | Select-Object Name, CPU, Memory, Storage | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No hardware plans configured.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 12: Inventory ──
[void]$sb.AppendLine('<div class="section" id="sec-inventory"><div class="section-header"><span class="section-num">12</span><h2>Inventory</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>VMware vSphere Servers</h3>')
if ($ViServers.Count -gt 0) {
    $viSrvData = $ViServers | ForEach-Object {
        [PSCustomObject]@{
            Name         = $_.Name
            'IP Address' = Resolve-HostIP $_.Name
            Type         = $_.Type
            DNS          = $_.DNSName
            Connected    = $_.IsConnected
        }
    }
    [void]$sb.AppendLine(($viSrvData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No vSphere servers added.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Hyper-V Servers</h3>')
if ($HvServers.Count -gt 0) {
    $hvSrvData = $HvServers | ForEach-Object {
        [PSCustomObject]@{
            Name         = $_.Name
            'IP Address' = Resolve-HostIP $_.Name
            DNS          = $_.DNSName
            Connected    = $_.IsConnected
        }
    }
    [void]$sb.AppendLine(($hvSrvData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>No Hyper-V servers added.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Stored Credentials (Names Only - No Passwords)</h3>')
if ($Credentials.Count -gt 0) {
    [void]$sb.AppendLine(($Credentials | ConvertTo-HtmlTable -EmptyMessage "No credentials found."))
} else { [void]$sb.AppendLine("<p class='empty'>No credentials stored.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Section 13: Global Settings ──
[void]$sb.AppendLine('<div class="section" id="sec-global"><div class="section-header"><span class="section-num">13</span><h2>Global Settings</h2></div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Notification Options</h3>')
if ($NotifOpts) {
    $notifData = [PSCustomObject]@{
        'Send on Success' = $NotifOpts.SendSuccessEmail
        'Send on Warning' = $NotifOpts.SendWarningEmail
        'Send on Failure' = $NotifOpts.SendFailureEmail
        'Notify on Retry' = $NotifOpts.SendNotificationOnLastRetryFailure
    }
    [void]$sb.AppendLine(($notifData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>Could not retrieve notification settings.</p>") }
[void]$sb.AppendLine('</div>')

[void]$sb.AppendLine('<div class="subsection"><h3>Email Settings</h3>')
if ($EmailOpts) {
    $emailData = [PSCustomObject]@{
        'SMTP Server'    = $EmailOpts.SMTPServer
        'SMTP Port'      = $EmailOpts.SMTPPort
        From             = $EmailOpts.From
        To               = $EmailOpts.To
        SSL              = $EmailOpts.EnableSSL
        'Use Auth'       = $EmailOpts.UseAuthentication
    }
    [void]$sb.AppendLine(($emailData | ConvertTo-HtmlTable))
} else { [void]$sb.AppendLine("<p class='empty'>Email notifications not configured or cmdlet unavailable.</p>") }
[void]$sb.AppendLine('</div></div>')

# ── Footer ──
[void]$sb.AppendLine(@"
</div></div>
<div class="footer-accent-bar"></div>
<div class="report-footer">
  <div class="footer-brand">
    <strong>The Redesign Group</strong>
    <span>Technology &amp; Cybersecurity Consulting</span>
    <span>redesign-group.com</span>
  </div>
  <div class="footer-center">
    Veeam B&amp;R v13 As-Built Report v3.0 &nbsp;|&nbsp; $ReportDate
  </div>
  <div class="footer-right">
    $([System.Web.HttpUtility]::HtmlEncode($CustomerName))<br>
    Server: $($BkpServer.Name ?? 'Unknown')
  </div>
</div>
<script>$JS</script>
</body></html>
"@)

# ─── 6. WRITE OUTPUT ─────────────────────────────────────────────────────────

New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
$FilePath = Join-Path $OutputPath "Veeam-v13-AsBuilt-$CustomerName-$ReportDateISO.html"
$sb.ToString() | Out-File -FilePath $FilePath -Encoding UTF8 -Force

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  ✅  Report saved: $FilePath" -ForegroundColor Green
Write-Host "  🏢  Branded for: $CustomerName  |  Prepared by: $PreparedBy" -ForegroundColor Cyan
if ($HealthCheck -and $Warnings.Count -gt 0) {
    Write-Host "  ⚠   $($Warnings.Count) health warning(s) - review the report." -ForegroundColor Yellow
}
Write-Host "  📄  Open in any browser.  Ctrl+P -> Save as PDF." -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
