#Requires -Version 5.1
<#
.SYNOPSIS
    VCF 9 As-Built Data Collector & Report Generator
    ReDesign Group  |  redesign-group.com

.DESCRIPTION
    Connects to a live VMware Cloud Foundation 9 environment, collects all
    configuration data via SDDC Manager REST API and PowerCLI, writes a
    site-config JSON file, then invokes generate-vcf9-asbuilt.js to produce
    the branded Word (.docx) report.

    Can also run in MANUAL mode — prompts for every value without connecting
    to any live system. Useful for air-gapped environments or for building
    a config file from scratch.

.PARAMETER SddcManagerFqdn
    FQDN or IP of the SDDC Manager appliance.

.PARAMETER SddcCredential
    PSCredential for SDDC Manager login (prompts if not supplied).

.PARAMETER Theme
    Color theme for the report. "dell" = navy/blue (matches template).
    "rdg" = ReDesign Group navy/teal. Default: "dell".

.PARAMETER OutputDocx
    Full path for the generated .docx file.
    Default: .\VCF9_AsBuilt_<customer>_<date>.docx

.PARAMETER ConfigJsonOnly
    Write the JSON config file and exit — do not invoke Node.js.
    Useful for review before generating the document.

.PARAMETER ManualMode
    Skip all live API queries. Prompt for every value manually.

.PARAMETER SkipTlsCheck
    Disable TLS certificate validation (self-signed certs in lab environments).

.PARAMETER NodePath
    Path to the node.exe binary. Default: "node" (assumes PATH).

.PARAMETER ScriptDir
    Directory containing generate-vcf9-asbuilt.js.
    Default: same directory as this script.

.EXAMPLE
    # Live collection — prompts for credentials
    .\Collect-VCF9AsBuilt.ps1 -SddcManagerFqdn sddc-mgr.corp.local

.EXAMPLE
    # Live collection, RDG theme, custom output
    .\Collect-VCF9AsBuilt.ps1 -SddcManagerFqdn sddc-mgr.corp.local -Theme rdg `
        -OutputDocx "C:\Reports\Acme_VCF9_AsBuilt.docx"

.EXAMPLE
    # Manual mode — no live connection required
    .\Collect-VCF9AsBuilt.ps1 -ManualMode

.EXAMPLE
    # Build JSON only (no Node.js / Word)
    .\Collect-VCF9AsBuilt.ps1 -SddcManagerFqdn sddc-mgr.corp.local -ConfigJsonOnly

.NOTES
    Dependencies:
      - Node.js  (https://nodejs.org)  — for .docx generation
      - generate-vcf9-asbuilt.js       — in same directory as this script
      - VMware.PowerCLI module         — optional (richer ESXi/vCenter data)
        Install: Install-Module VMware.PowerCLI -Scope CurrentUser
#>

[CmdletBinding(DefaultParameterSetName = 'Live')]
param (
    [Parameter(ParameterSetName = 'Live', Mandatory = $true)]
    [string]$SddcManagerFqdn,

    [Parameter(ParameterSetName = 'Live')]
    [System.Management.Automation.PSCredential]$SddcCredential,

    [Parameter(ParameterSetName = 'Manual', Mandatory = $true)]
    [switch]$ManualMode,

    [ValidateSet('dell','rdg')]
    [string]$Theme = 'dell',

    [string]$OutputDocx,

    [switch]$ConfigJsonOnly,

    [switch]$SkipTlsCheck,

    [string]$NodePath = 'node',

    [string]$ScriptDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Resolve paths ─────────────────────────────────────────────────────────────
if (-not $ScriptDir) { $ScriptDir = $PSScriptRoot }
$GeneratorScript = Join-Path $ScriptDir 'generate-vcf9-asbuilt.js'

# ── Console helpers ───────────────────────────────────────────────────────────
function Write-Banner {
    $line = '=' * 64
    Write-Host ''
    Write-Host $line -ForegroundColor Cyan
    Write-Host '  VCF 9 As-Built Collector  |  ReDesign Group' -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ''
}

function Write-Section ([string]$Title) {
    Write-Host ''
    Write-Host ('  ── ' + $Title + ' ' + ('-' * [Math]::Max(0, 54 - $Title.Length))) -ForegroundColor DarkCyan
}

function Write-Ok  ([string]$Msg) { Write-Host "  [OK]  $Msg" -ForegroundColor Green  }
function Write-Inf ([string]$Msg) { Write-Host "  [..] $Msg"  -ForegroundColor Gray   }
function Write-Wrn ([string]$Msg) { Write-Host "  [!!] $Msg"  -ForegroundColor Yellow }
function Write-Err ([string]$Msg) { Write-Host "  [XX] $Msg"  -ForegroundColor Red    }

# ── Prompt helper (with default) ──────────────────────────────────────────────
function Read-Value {
    param(
        [string]$Prompt,
        [string]$Default = '',
        [switch]$Secret
    )
    $hint    = if ($Default) { " [$Default]" } else { '' }
    $display = "  ${Prompt}${hint}: "

    if ($Secret) {
        $ss = Read-Host -Prompt $display -AsSecureString
        if ($ss.Length -eq 0 -and $Default) { return $Default }
        return [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss))
    }

    $ans = Read-Host -Prompt $display
    if ([string]::IsNullOrWhiteSpace($ans)) { return $Default }
    return $ans.Trim()
}

# ── TLS helper ────────────────────────────────────────────────────────────────
function Disable-TlsChecks {
    if ($SkipTlsCheck) {
        Write-Wrn 'TLS certificate validation disabled (--SkipTlsCheck)'
        # .NET callback
        $code = @'
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAll : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert,
        WebRequest req, int problem) { return true; }
}
'@
        Add-Type -TypeDefinition $code -ErrorAction SilentlyContinue
        [Net.ServicePointManager]::CertificatePolicy = New-Object TrustAll
        [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }
}

# ── SDDC Manager REST API helpers ─────────────────────────────────────────────
$script:AuthToken = $null
$script:SddcBase  = $null

function Connect-SddcManager {
    param([string]$Fqdn, [System.Management.Automation.PSCredential]$Cred)

    $script:SddcBase = "https://$Fqdn"
    $body = @{ username = $Cred.UserName
               password = $Cred.GetNetworkCredential().Password } | ConvertTo-Json

    try {
        $resp = Invoke-RestMethod -Uri "$script:SddcBase/v1/tokens" `
                    -Method POST -Body $body `
                    -ContentType 'application/json' -ErrorAction Stop
        $script:AuthToken = $resp.accessToken
        Write-Ok "Authenticated to SDDC Manager: $Fqdn"
    } catch {
        Write-Err "Authentication failed: $_"
        throw
    }
}

function Invoke-SddcApi {
    param([string]$Path, [string]$Method = 'GET')
    $headers = @{ Authorization = "Bearer $script:AuthToken"
                  Accept        = 'application/json' }
    try {
        return Invoke-RestMethod -Uri "$script:SddcBase$Path" `
                   -Method $Method -Headers $headers -ErrorAction Stop
    } catch {
        Write-Wrn "API call failed [$Path]: $($_.Exception.Message)"
        return $null
    }
}

# ── vCenter / NSX REST helpers ────────────────────────────────────────────────
function Invoke-VcApi {
    param([string]$VcFqdn, [string]$Path,
          [System.Management.Automation.PSCredential]$Cred)
    $base    = "https://$VcFqdn"
    $headers = @{}
    # Get session token
    try {
        $session = Invoke-RestMethod -Uri "$base/api/session" -Method POST `
                       -Credential $Cred -Authentication Basic `
                       -ContentType 'application/json' -ErrorAction Stop
        $headers['vmware-api-session-id'] = $session
    } catch {
        Write-Wrn "vCenter API session failed [$VcFqdn]: $($_.Exception.Message)"
        return $null
    }
    try {
        return Invoke-RestMethod -Uri "$base$Path" -Headers $headers -ErrorAction Stop
    } catch {
        Write-Wrn "vCenter API call failed [$Path]: $($_.Exception.Message)"
        return $null
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  LIVE DATA COLLECTION
# ══════════════════════════════════════════════════════════════════════════════
function Get-LiveData {
    param([System.Management.Automation.PSCredential]$Cred)

    $d = @{}   # will be converted to JSON config

    # ── SDDC Manager info ─────────────────────────────────────────────────────
    Write-Section 'SDDC Manager'
    $sddcInfo = Invoke-SddcApi '/v1/sddc-manager'
    $sysInfo  = Invoke-SddcApi '/v1/system/primary-management'

    $d.siteId      = if ($sddcInfo) { $sddcInfo.id } else { 'N/A' }
    $d.vcfVersion  = if ($sddcInfo) { $sddcInfo.version } else { '9.0.0.0' }

    $d.sddc = @{
        hostname    = $SddcManagerFqdn
        ip          = try { ([System.Net.Dns]::GetHostAddresses($SddcManagerFqdn) |
                             Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                             Select-Object -First 1).IPAddressToString } catch { 'N/A' }
        version     = if ($sddcInfo) { "$($sddcInfo.version) (build $($sddcInfo.build))" } else { 'N/A' }
        serviceTag  = if ($sddcInfo) { $sddcInfo.serialNumber } else { 'N/A' }
        domain      = ''
        datacenter  = ''
        cluster     = ''
        ntp         = ''
        dns         = ''
        ssoUser     = 'administrator@vsphere.local'
        backupDest  = 'N/A (see SDDC Manager > Administration > Backup)'
    }

    # ── Domains ───────────────────────────────────────────────────────────────
    Write-Section 'Domains'
    $domains = Invoke-SddcApi '/v1/domains'
    $mgmtDomain = $null
    $wldDomains  = @()
    if ($domains -and $domains.elements) {
        $mgmtDomain = $domains.elements | Where-Object { $_.type -eq 'MANAGEMENT' } | Select-Object -First 1
        $wldDomains  = @($domains.elements | Where-Object { $_.type -eq 'WORKLOAD' })
        if ($mgmtDomain) {
            $d.sddc.domain = $mgmtDomain.name
            Write-Ok "Management domain: $($mgmtDomain.name)"
        }
    }

    # ── Clusters ─────────────────────────────────────────────────────────────
    Write-Section 'Clusters'
    $clusters = Invoke-SddcApi '/v1/clusters'
    $mgmtCluster = $null
    $wldCluster  = $null
    if ($clusters -and $clusters.elements) {
        $mgmtCluster = $clusters.elements | Where-Object { $_.domainId -eq $mgmtDomain.id } | Select-Object -First 1
        if ($wldDomains.Count -gt 0) {
            $wldCluster = $clusters.elements | Where-Object { $_.domainId -eq $wldDomains[0].id } | Select-Object -First 1
        }
        if ($mgmtCluster) {
            $d.sddc.cluster = $mgmtCluster.name
            $d.sddc.datacenter = if ($mgmtCluster.datacenterName) { $mgmtCluster.datacenterName } else { 'DC-Primary' }
            Write-Ok "Management cluster: $($mgmtCluster.name) ($($mgmtCluster.primaryDatastoreType))"
        }
    }

    # ── Hosts ─────────────────────────────────────────────────────────────────
    Write-Section 'ESXi Hosts'
    $hosts     = Invoke-SddcApi '/v1/hosts'
    $mgmtHosts = 0
    $wldHosts  = 0
    if ($hosts -and $hosts.elements) {
        $mgmtHosts = ($hosts.elements | Where-Object { $_.domain.id -eq $mgmtDomain.id }).Count
        if ($wldDomains.Count -gt 0) {
            $wldHosts = ($hosts.elements | Where-Object { $_.domain.id -eq $wldDomains[0].id }).Count
        }
        Write-Ok "Hosts — Management: $mgmtHosts | Workload: $wldHosts"

        # Extract NTP/DNS from first management host
        $firstHost = $hosts.elements | Where-Object { $_.domain.id -eq $mgmtDomain.id } | Select-Object -First 1
        if ($firstHost -and $firstHost.networkDetails) {
            $nd = $firstHost.networkDetails
            if ($nd.dnsServers)  { $d.sddc.dns = ($nd.dnsServers -join ', ') }
            if ($nd.ntpServers)  { $d.sddc.ntp = ($nd.ntpServers -join ', ') }
        }
    }

    # ── vCenters ─────────────────────────────────────────────────────────────
    Write-Section 'vCenter Servers'
    $vcenters = Invoke-SddcApi '/v1/vcenters'
    $vcMgmt   = $null
    $vcWld    = $null
    if ($vcenters -and $vcenters.elements) {
        $vcMgmt = $vcenters.elements | Where-Object { $_.domainId -eq $mgmtDomain.id } | Select-Object -First 1
        if ($wldDomains.Count -gt 0) {
            $vcWld = $vcenters.elements | Where-Object { $_.domainId -eq $wldDomains[0].id } | Select-Object -First 1
        }
        if ($vcMgmt) { Write-Ok "vCenter Mgmt: $($vcMgmt.fqdn) v$($vcMgmt.version)" }
        if ($vcWld)  { Write-Ok "vCenter Wld:  $($vcWld.fqdn) v$($vcWld.version)" }
    }

    $d.vCenter = @{
        management = @{
            hostname = if ($vcMgmt) { $vcMgmt.fqdn    } else { 'vcsa-mgmt-01.domain.local' }
            ip       = if ($vcMgmt) { $vcMgmt.ipAddress} else { 'N/A' }
            version  = if ($vcMgmt) { "$($vcMgmt.version) (build $($vcMgmt.build))" } else { 'N/A' }
            cluster  = if ($mgmtCluster) { $mgmtCluster.name } else { 'mgmt-cluster-01' }
            hosts    = "$mgmtHosts"
            ha       = 'Enabled'
            sso      = 'vsphere.local'
        }
        workload = @{
            hostname = if ($vcWld) { $vcWld.fqdn     } else { 'vcsa-wld-01.domain.local' }
            ip       = if ($vcWld) { $vcWld.ipAddress} else { 'N/A' }
            version  = if ($vcWld) { "$($vcWld.version) (build $($vcWld.build))" } else { 'N/A' }
            cluster  = if ($wldCluster) { $wldCluster.name } else { 'wld-cluster-01' }
            hosts    = "$wldHosts"
        }
    }

    # ── NSX ───────────────────────────────────────────────────────────────────
    Write-Section 'NSX'
    $nsxItems = Invoke-SddcApi '/v1/nsxt-clusters'
    $nsxCluster = if ($nsxItems -and $nsxItems.elements) { $nsxItems.elements | Select-Object -First 1 } else { $null }

    $d.nsx = @{
        vip          = if ($nsxCluster) { $nsxCluster.vipFqdn } else { 'nsx-vip.domain.local' }
        manager1     = if ($nsxCluster -and $nsxCluster.nodes.Count -gt 0) { $nsxCluster.nodes[0].fqdn } else { 'nsx-mgr-01.domain.local' }
        manager2     = if ($nsxCluster -and $nsxCluster.nodes.Count -gt 1) { $nsxCluster.nodes[1].fqdn } else { 'nsx-mgr-02.domain.local' }
        manager3     = if ($nsxCluster -and $nsxCluster.nodes.Count -gt 2) { $nsxCluster.nodes[2].fqdn } else { 'nsx-mgr-03.domain.local' }
        version      = if ($nsxCluster) { $nsxCluster.nsxtVersion } else { 'N/A' }
        edgeCluster  = 'edge-cluster-01'
        edgeNode1    = 'nsx-edge-01.domain.local'
        edgeNode2    = 'nsx-edge-02.domain.local'
        tier0        = 'T0-GW-Primary'
        tier1        = 'T1-GW-Workload'
        overlayTZ    = 'nsx-overlay-transportzone'
        vlanTZ       = 'nsx-vlan-transportzone'
        bgpAS        = '65001'
        bgpPeer      = 'N/A (configure after reviewing T0 BGP peers)'
    }
    if ($nsxCluster) { Write-Ok "NSX cluster: $($nsxCluster.vipFqdn) v$($nsxCluster.nsxtVersion)" }

    # Try to get edge clusters from NSX Manager directly
    if ($nsxCluster) {
        $nsxEdges = Invoke-SddcApi '/v1/nsxt-edge-clusters'
        if ($nsxEdges -and $nsxEdges.elements -and $nsxEdges.elements.Count -gt 0) {
            $ec = $nsxEdges.elements[0]
            $d.nsx.edgeCluster = $ec.name
            if ($ec.edgeNodes -and $ec.edgeNodes.Count -gt 0) {
                $d.nsx.edgeNode1 = if ($ec.edgeNodes[0].fqdn) { $ec.edgeNodes[0].fqdn } else { $d.nsx.edgeNode1 }
            }
            if ($ec.edgeNodes -and $ec.edgeNodes.Count -gt 1) {
                $d.nsx.edgeNode2 = if ($ec.edgeNodes[1].fqdn) { $ec.edgeNodes[1].fqdn } else { $d.nsx.edgeNode2 }
            }
            Write-Ok "Edge cluster: $($d.nsx.edgeCluster)"
        }
    }

    # ── vSAN ──────────────────────────────────────────────────────────────────
    Write-Section 'vSAN'
    $vsanData = Invoke-SddcApi "/v1/clusters/$($mgmtCluster.id)/datastores"
    $vsanDs   = $null
    if ($vsanData -and $vsanData.elements) {
        $vsanDs = $vsanData.elements | Where-Object { $_.type -eq 'VSAN' } | Select-Object -First 1
    }

    $d.vsan = @{
        version        = '8.0 U3'
        policy         = 'RAID-5 (Erasure Coding)'
        dedup          = 'Enabled'
        compression    = 'Enabled'
        encryption     = 'Enabled (vSAN Encryption)'
        capacity       = if ($vsanDs) { "$([Math]::Round($vsanDs.capacityGB / 1024, 1)) TiB usable" } else { 'N/A' }
        faultDomains   = '3'
        stretchCluster = 'Disabled'
        fileServices   = 'Disabled'
    }
    Write-Ok "vSAN capacity: $($d.vsan.capacity)"

    # ── Networking ───────────────────────────────────────────────────────────
    Write-Section 'Networking'
    $networks = Invoke-SddcApi '/v1/network-pools'
    $netPool  = if ($networks -and $networks.elements) { $networks.elements | Select-Object -First 1 } else { $null }

    $mgmtNet   = $null; $vmotNet   = $null; $vsanNet   = $null
    $overlayNet = $null; $uplinkNet = $null
    if ($netPool -and $netPool.networks) {
        foreach ($n in $netPool.networks) {
            switch ($n.type) {
                'MANAGEMENT' { $mgmtNet    = $n }
                'VMOTION'    { $vmotNet    = $n }
                'VSAN'       { $vsanNet    = $n }
                'OVERLAY'    { $overlayNet = $n }
                'UPLINK01'   { $uplinkNet  = $n }
            }
        }
    }

    function Net-Str ($n, $label) {
        if (-not $n) { return "VLAN ??? — ???.???.???.0/??"}
        return "VLAN $($n.vlanId) — $($n.subnet)"
    }

    # Derive DNS from sddc if already set
    $dnsServers = if ($d.sddc.dns) { $d.sddc.dns -split ',\s*' } else { @('192.168.10.1','192.168.10.2') }
    $ntpServers = if ($d.sddc.ntp) { $d.sddc.ntp -split ',\s*' } else { @('ntp1.domain.local') }

    $d.networking = @{
        dns1         = if ($dnsServers.Count -gt 0) { $dnsServers[0] } else { '192.168.10.1' }
        dns2         = if ($dnsServers.Count -gt 1) { $dnsServers[1] } else { '192.168.10.2' }
        dns3         = $null
        dnsDomain    = ($SddcManagerFqdn -split '\.', 2)[-1]
        ntp1         = if ($ntpServers.Count -gt 0) { $ntpServers[0] } else { 'ntp1.domain.local' }
        ntp2         = if ($ntpServers.Count -gt 1) { $ntpServers[1] } else { $null }
        timezone     = 'US/Eastern'
        mgmtVlan     = Net-Str $mgmtNet    'VLAN 10'
        mgmtGw       = if ($mgmtNet)    { $mgmtNet.gateway    } else { '192.168.10.1' }
        vmotionVlan  = Net-Str $vmotNet   'VLAN 20'
        vmotionGw    = if ($vmotNet)    { $vmotNet.gateway    } else { '192.168.20.1' }
        vsanVlan     = Net-Str $vsanNet   'VLAN 30'
        vsanGw       = if ($vsanNet)    { $vsanNet.gateway    } else { '192.168.30.1' }
        overlayVlan  = Net-Str $overlayNet 'VLAN 50'
        overlayGw    = if ($overlayNet) { $overlayNet.gateway } else { '192.168.50.1' }
        uplinkVlan   = Net-Str $uplinkNet  'VLAN 40'
        uplinkGw     = if ($uplinkNet)  { $uplinkNet.gateway  } else { '192.168.40.1' }
        workloadVlan = 'VLAN 100-199 (workload segments)'
        defaultGw    = if ($mgmtNet)    { $mgmtNet.gateway    } else { '192.168.10.1' }
        mtu          = '9000 (Jumbo Frames)'
        torSwitch    = 'N/A (verify with network team)'
        dvSwitch     = 'vds-mgmt-01, vds-edge-01'
        uplinks      = '2 x 25GbE per host (active-active LACP)'
        lacp         = 'LACP (active-active, hash: src-dst IP+port)'
    }
    Write-Ok "Network pool: $($netPool.name ?? 'retrieved')"

    # ── Aria Suite ────────────────────────────────────────────────────────────
    Write-Section 'Aria Suite'
    $ariaItems = Invoke-SddcApi '/v1/vrealize'
    $lcm = $null; $li = $null; $ops = $null; $auto = $null

    if ($ariaItems -and $ariaItems.elements) {
        foreach ($a in $ariaItems.elements) {
            switch ($a.type) {
                'VRLI'   { $li   = $a }
                'VROPS'  { $ops  = $a }
                'VRA'    { $auto = $a }
                'VRSLCM' { $lcm  = $a }
            }
        }
    }
    # Also try v1/aria-suite
    if (-not $lcm) {
        $ariaV2 = Invoke-SddcApi '/v1/aria-suite'
        if ($ariaV2 -and $ariaV2.ariaLCM) { $lcm = $ariaV2.ariaLCM }
    }

    $d.aria = @{
        lcm = @{
            hostname = if ($lcm) { $lcm.fqdn      } else { 'aria-lcm-01.domain.local'  }
            ip       = if ($lcm) { $lcm.ipAddress  } else { 'N/A' }
            version  = if ($lcm) { $lcm.version    } else { '8.16' }
        }
        logInsight = @{
            hostname  = if ($li)  { $li.loadBalancerFqdn ?? $li.fqdn } else { 'aria-li-01.domain.local'  }
            ip        = if ($li)  { $li.ipAddress } else { 'N/A' }
            version   = if ($li)  { $li.version   } else { '8.16' }
            cluster   = '3-node cluster'
            retention = '30 days'
        }
        operations = @{
            hostname = if ($ops)  { $ops.loadBalancerFqdn ?? $ops.fqdn } else { 'aria-ops-01.domain.local' }
            ip       = if ($ops)  { $ops.ipAddress } else { 'N/A' }
            version  = if ($ops)  { $ops.version   } else { '8.16' }
            adapters = 'vCenter, NSX, vSAN'
        }
        automation = @{
            hostname = if ($auto) { $auto.loadBalancerFqdn ?? $auto.fqdn } else { 'aria-auto-01.domain.local' }
            ip       = if ($auto) { $auto.ipAddress } else { 'N/A' }
            version  = if ($auto) { $auto.version   } else { '8.16' }
        }
    }
    if ($lcm)  { Write-Ok "Aria LCM:         $($d.aria.lcm.hostname) v$($d.aria.lcm.version)" }
    if ($li)   { Write-Ok "Aria Log Insight: $($d.aria.logInsight.hostname) v$($d.aria.logInsight.version)" }
    if ($ops)  { Write-Ok "Aria Operations:  $($d.aria.operations.hostname) v$($d.aria.operations.version)" }
    if ($auto) { Write-Ok "Aria Automation:  $($d.aria.automation.hostname) v$($d.aria.automation.version)" }

    # ── Licensing ─────────────────────────────────────────────────────────────
    Write-Section 'Licensing'
    $licItems = Invoke-SddcApi '/v1/licenses'
    $licSummary = 'Included in VCF 9 Universal'
    $coreCount  = 'N/A (see SDDC Manager > Administration > Licensing)'
    if ($licItems -and $licItems.elements) {
        $vcfLic = $licItems.elements | Where-Object { $_.productType -eq 'VCF' } | Select-Object -First 1
        if ($vcfLic) {
            $coreCount = "$($vcfLic.quantity) licensed cores"
            Write-Ok "VCF License: $($vcfLic.licenseKey.Substring(0,8))... ($coreCount)"
        }
    }

    $d.licensing = @{
        vcf     = 'VMware Cloud Foundation 9 — Universal License'
        vCenter = $licSummary
        nsx     = $licSummary
        vsan    = $licSummary
        aria    = 'Included in VCF 9 Universal (Aria Suite Enterprise)'
        cores   = $coreCount
        schema  = 'ELMS'
    }

    # ── Security (static defaults — prompt user to verify) ───────────────────
    Write-Section 'Security'
    $certs = Invoke-SddcApi '/v1/certificates?issuedTo=SDDC_MANAGER'
    $certFrom  = 'N/A'
    $certUntil = 'N/A'
    if ($certs -and $certs.elements -and $certs.elements.Count -gt 0) {
        $c = $certs.elements[0]
        $certFrom  = $c.notBefore
        $certUntil = $c.notAfter
        Write-Ok "Certificate valid: $certFrom -> $certUntil"
    }

    $d.security = @{
        tls          = 'TLS 1.2 / 1.3 enforced'
        ca           = 'Internal CA — PKI integrated'
        certFrom     = $certFrom
        certUntil    = $certUntil
        mfa          = 'SAML 2.0 via Workspace ONE Access'
        passwords    = '90-day rotation, 20-char minimum'
        lockdown     = 'Normal Lockdown Mode on all ESXi hosts'
        syslog       = 'Aria Log Insight (centralized)'
        compliance   = 'VMware Security Hardening Guide v9.0'
        backupEncr   = 'Enabled (AES-256)'
        backupEncStrength = 'Medium (AES-256)'
        ssoDomain    = 'vsphere.local'
        idp          = 'Workspace ONE Access (SAML 2.0)'
        adRealm      = 'Configured'
        cifsAuth     = 'Active Directory'
        idfw         = 'Enabled (AD groups)'
        ssh          = 'Disabled on ESXi (Lockdown Mode)'
    }

    # ── SMTP ──────────────────────────────────────────────────────────────────
    Write-Section 'SMTP / Notification'
    $smtpInfo = Invoke-SddcApi '/v1/system/notifications/settings'
    $d.smtp = @{
        server      = if ($smtpInfo) { $smtpInfo.smtpServer   } else { 'mail.domain.local' }
        adminEmail  = if ($smtpInfo) { $smtpInfo.senderEmail  } else { 'vcf-admin@domain.local' }
        autosupport = 'autosupport@domain.local'
        alertEmail  = 'vcf-alerts@domain.local'
        sendTime    = '06:00 daily'
    }
    if ($smtpInfo) { Write-Ok "SMTP: $($d.smtp.server)" }

    # ── Backup ────────────────────────────────────────────────────────────────
    Write-Section 'Backup'
    $backupCfg = Invoke-SddcApi '/v1/system/backup-configuration'
    if ($backupCfg -and $backupCfg.backupLocation) {
        $d.sddc.backupDest = "scp://$($backupCfg.backupLocation.server)$($backupCfg.backupLocation.directory)"
        Write-Ok "Backup dest: $($d.sddc.backupDest)"
    }

    return $d
}

# ══════════════════════════════════════════════════════════════════════════════
#  MANUAL DATA ENTRY
# ══════════════════════════════════════════════════════════════════════════════
function Get-ManualData {

    Write-Host '  Enter values for each field. Press Enter to accept [defaults].' -ForegroundColor Gray
    Write-Host ''

    $d = @{}

    # ── General ───────────────────────────────────────────────────────────────
    Write-Section 'General'
    $d.customer        = Read-Value 'Customer name'                     'Customer Name'
    $d.customerAddress = Read-Value 'Customer address (one line)'       '123 Main St, City, ST 00000, US'
    $d.reviewedBy      = Read-Value 'Reviewed by'                       '—'
    $d.version         = Read-Value 'Document version'                  '1.0'
    $d.vcfVersion      = Read-Value 'VCF version'                       '9.0.0.0'
    $d.siteId          = Read-Value 'Site ID / Service Tag'             'SITE-001'
    $d.theme           = $Theme

    # ── SDDC Manager ─────────────────────────────────────────────────────────
    Write-Section 'SDDC Manager'
    $d.sddc = @{
        hostname   = Read-Value 'SDDC Manager FQDN'          'sddc-mgr.domain.local'
        ip         = Read-Value 'SDDC Manager IP'            '192.168.10.10'
        version    = Read-Value 'SDDC Manager version/build' '9.0.0.0 (build 12345678)'
        serviceTag = Read-Value 'Service tag / serial'       'XXXXXXX'
        domain     = Read-Value 'Management domain name'     'mgmt.domain.local'
        datacenter = Read-Value 'Datacenter name'            'DC-Primary'
        cluster    = Read-Value 'Management cluster name'    'mgmt-cluster-01'
        ntp        = Read-Value 'NTP servers (comma-sep)'    'ntp1.domain.local, ntp2.domain.local'
        dns        = Read-Value 'DNS servers (comma-sep)'    '192.168.10.1, 192.168.10.2'
        ssoUser    = Read-Value 'SSO admin account'          'administrator@vsphere.local'
        backupDest = Read-Value 'Backup destination'         'scp://backup.domain.local/vcf/sddc'
    }

    # ── vCenter ───────────────────────────────────────────────────────────────
    Write-Section 'vCenter — Management Domain'
    $vcMgmtHostname = Read-Value 'vCenter (Mgmt) FQDN'          'vcsa-mgmt-01.domain.local'
    $vcMgmtIp       = Read-Value 'vCenter (Mgmt) IP'            '192.168.10.11'
    $vcMgmtVer      = Read-Value 'vCenter (Mgmt) version/build' '8.0 U3 (build 23456789)'
    $vcMgmtCluster  = Read-Value 'Management cluster name'      'mgmt-cluster-01'
    $vcMgmtHosts    = Read-Value 'Management domain host count' '4'

    Write-Section 'vCenter — Workload Domain'
    $vcWldHostname  = Read-Value 'vCenter (Workload) FQDN'      'vcsa-wld-01.domain.local'
    $vcWldIp        = Read-Value 'vCenter (Workload) IP'        '192.168.10.12'
    $vcWldVer       = Read-Value 'vCenter (Workload) version'   '8.0 U3 (build 23456789)'
    $vcWldCluster   = Read-Value 'Workload cluster name'        'wld-cluster-01'
    $vcWldHosts     = Read-Value 'Workload domain host count'   '8'

    $d.vCenter = @{
        management = @{ hostname=$vcMgmtHostname; ip=$vcMgmtIp; version=$vcMgmtVer
                        cluster=$vcMgmtCluster; hosts=$vcMgmtHosts; ha='Enabled'; sso='vsphere.local' }
        workload   = @{ hostname=$vcWldHostname;  ip=$vcWldIp;  version=$vcWldVer
                        cluster=$vcWldCluster;  hosts=$vcWldHosts }
    }

    # ── NSX ───────────────────────────────────────────────────────────────────
    Write-Section 'NSX Manager'
    $d.nsx = @{
        vip         = Read-Value 'NSX Manager VIP IP'        '192.168.10.20'
        manager1    = Read-Value 'NSX Manager Node 1 FQDN'   'nsx-mgr-01.domain.local'
        manager2    = Read-Value 'NSX Manager Node 2 FQDN'   'nsx-mgr-02.domain.local'
        manager3    = Read-Value 'NSX Manager Node 3 FQDN'   'nsx-mgr-03.domain.local'
        version     = Read-Value 'NSX version/build'         '4.2.0.0 (build 34567890)'
        edgeCluster = Read-Value 'Edge cluster name'         'edge-cluster-01'
        edgeNode1   = Read-Value 'Edge Node 1 FQDN'          'nsx-edge-01.domain.local'
        edgeNode2   = Read-Value 'Edge Node 2 FQDN'          'nsx-edge-02.domain.local'
        tier0       = Read-Value 'Tier-0 Gateway name'       'T0-GW-Primary'
        tier1       = Read-Value 'Tier-1 Gateway name'       'T1-GW-Workload'
        overlayTZ   = Read-Value 'Overlay Transport Zone'    'nsx-overlay-transportzone'
        vlanTZ      = Read-Value 'VLAN Transport Zone'       'nsx-vlan-transportzone'
        bgpAS       = Read-Value 'BGP AS (local)'            '65001'
        bgpPeer     = Read-Value 'BGP upstream peer'         '192.168.10.1 (AS 65000)'
    }

    # ── vSAN ──────────────────────────────────────────────────────────────────
    Write-Section 'vSAN'
    $d.vsan = @{
        version        = Read-Value 'vSAN version'                 '8.0 U3'
        policy         = Read-Value 'Default storage policy'       'RAID-5 (Erasure Coding)'
        dedup          = Read-Value 'Deduplication'                'Enabled'
        compression    = Read-Value 'Compression'                  'Enabled'
        encryption     = Read-Value 'Encryption'                   'Enabled (vSAN Encryption)'
        capacity       = Read-Value 'Total capacity (raw/usable)'  '153.6 TB raw / 76.8 TB usable'
        faultDomains   = Read-Value 'Fault domains'                '3 (Rack-A, Rack-B, Rack-C)'
        stretchCluster = Read-Value 'Stretch cluster'              'Disabled'
        fileServices   = Read-Value 'File services'                'Disabled'
    }

    # ── Networking ────────────────────────────────────────────────────────────
    Write-Section 'Networking — DNS / NTP'
    $dns1   = Read-Value 'Primary DNS IP'     '192.168.10.1'
    $dns2   = Read-Value 'Secondary DNS IP'   '192.168.10.2'
    $dnsDom = Read-Value 'DNS domain'         'domain.local'
    $ntp1   = Read-Value 'Primary NTP server' 'ntp1.domain.local'
    $ntp2   = Read-Value 'Secondary NTP'      'ntp2.domain.local'

    Write-Section 'Networking — VLANs & Subnets'
    $mgmtVlan  = Read-Value 'Management VLAN/subnet'    'VLAN 10 — 192.168.10.0/24'
    $mgmtGw    = Read-Value 'Management gateway'        '192.168.10.1'
    $vmotVlan  = Read-Value 'vMotion VLAN/subnet'       'VLAN 20 — 192.168.20.0/24'
    $vmotGw    = Read-Value 'vMotion gateway'           '192.168.20.1'
    $vsanVlan  = Read-Value 'vSAN VLAN/subnet'          'VLAN 30 — 192.168.30.0/24'
    $vsanGw    = Read-Value 'vSAN gateway'              '192.168.30.1'
    $ovVlan    = Read-Value 'Overlay TEP VLAN/subnet'   'VLAN 50 — 192.168.50.0/24'
    $ovGw      = Read-Value 'Overlay gateway'           '192.168.50.1'
    $uplVlan   = Read-Value 'Uplink VLAN/subnet'        'VLAN 40 — 192.168.40.0/24'
    $uplGw     = Read-Value 'Uplink gateway'            '192.168.40.1'
    $wldVlan   = Read-Value 'Workload VLAN range'       'VLAN 100-199 (workload segments)'
    $tor       = Read-Value 'Top-of-Rack switch model'  'Cisco Nexus 93180YC-FX'
    $dvs       = Read-Value 'Distributed switches'      'vds-mgmt-01 (v8.0.3), vds-edge-01 (v8.0.3)'
    $uplinks   = Read-Value 'Uplinks per host'          '2 x 25GbE per host (active-active LACP)'

    $d.networking = @{
        dns1='192.168.10.1'; dns2=$dns2; dnsDomain=$dnsDom
        ntp1=$ntp1; ntp2=$ntp2; timezone='US/Eastern'
        mgmtVlan=$mgmtVlan; mgmtGw=$mgmtGw
        vmotionVlan=$vmotVlan; vmotionGw=$vmotGw
        vsanVlan=$vsanVlan; vsanGw=$vsanGw
        overlayVlan=$ovVlan; overlayGw=$ovGw
        uplinkVlan=$uplVlan; uplinkGw=$uplGw
        workloadVlan=$wldVlan; defaultGw=$mgmtGw
        mtu='9000 (Jumbo Frames)'; torSwitch=$tor
        dvSwitch=$dvs; uplinks=$uplinks
        lacp='LACP (active-active, hash: src-dst IP+port)'
    }
    $d.networking.dns1 = $dns1   # fix strict mode shadowing above

    # ── Aria Suite ────────────────────────────────────────────────────────────
    Write-Section 'Aria Suite'
    $d.aria = @{
        lcm        = @{ hostname=(Read-Value 'Aria LCM FQDN'           'aria-lcm-01.domain.local')
                        ip      =(Read-Value 'Aria LCM IP'              '192.168.10.30'); version='8.16' }
        logInsight = @{ hostname=(Read-Value 'Aria Log Insight FQDN'    'aria-li-01.domain.local')
                        ip      =(Read-Value 'Aria Log Insight IP'       '192.168.10.31')
                        version='8.16'; cluster='3-node cluster'; retention='30 days' }
        operations = @{ hostname=(Read-Value 'Aria Operations FQDN'     'aria-ops-01.domain.local')
                        ip      =(Read-Value 'Aria Operations IP'        '192.168.10.32')
                        version='8.16'; adapters='vCenter, NSX, vSAN' }
        automation = @{ hostname=(Read-Value 'Aria Automation FQDN'     'aria-auto-01.domain.local')
                        ip      =(Read-Value 'Aria Automation IP'        '192.168.10.33'); version='8.16' }
    }

    # ── SMTP ──────────────────────────────────────────────────────────────────
    Write-Section 'SMTP / Notifications'
    $d.smtp = @{
        server      = Read-Value 'SMTP server'          'mail.domain.local'
        adminEmail  = Read-Value 'Admin email'          'vcf-admin@domain.local'
        autosupport = Read-Value 'Autosupport email'    'autosupport@domain.local'
        alertEmail  = Read-Value 'Alert email'          'vcf-alerts@domain.local'
        sendTime    = '06:00 daily'
    }

    # ── Security ──────────────────────────────────────────────────────────────
    Write-Section 'Security'
    $d.security = @{
        tls='TLS 1.2 / 1.3 enforced'; ca='Internal CA — PKI integrated'
        certFrom='—'; certUntil='—'
        mfa          = Read-Value 'MFA method'           'SAML 2.0 via Workspace ONE Access'
        passwords    = Read-Value 'Password policy'      '90-day rotation, 20-char minimum'
        lockdown     = Read-Value 'ESXi lockdown mode'   'Normal Lockdown Mode on all ESXi hosts'
        syslog       = 'Aria Log Insight (centralized)'
        compliance   = 'VMware Security Hardening Guide v9.0'
        backupEncr   = 'Enabled (AES-256)'
        backupEncStrength = 'Medium (AES-256)'
        ssoDomain    = 'vsphere.local'
        idp          = 'Workspace ONE Access (SAML 2.0)'
        adRealm      = 'Configured'; cifsAuth='Active Directory'
        idfw         = 'Enabled (AD groups)'; ssh='Disabled on ESXi (Lockdown Mode)'
    }

    # ── Licensing ─────────────────────────────────────────────────────────────
    Write-Section 'Licensing'
    $d.licensing = @{
        vcf     = 'VMware Cloud Foundation 9 — Universal License'
        vCenter = 'Included in VCF 9 Universal'
        nsx     = 'Included in VCF 9 Universal'
        vsan    = 'Included in VCF 9 Universal'
        aria    = 'Included in VCF 9 Universal (Aria Suite Enterprise)'
        cores   = Read-Value 'Licensed core count' '384 licensed vCPU cores'
        schema  = 'ELMS'
    }

    return $d
}

# ══════════════════════════════════════════════════════════════════════════════
#  POST-COLLECTION: allow user to fill any gaps / confirm key values
# ══════════════════════════════════════════════════════════════════════════════
function Confirm-Data {
    param([hashtable]$d)

    Write-Section 'Review & Confirm'
    Write-Host '  Confirm or correct key values (Enter = keep existing).' -ForegroundColor Gray

    $d.customer        = Read-Value 'Customer name'        ($d.customer        ?? 'Customer Name')
    $d.customerAddress = Read-Value 'Customer address'     ($d.customerAddress ?? '123 Main St, City, ST 00000, US')
    $d.vcfVersion      = Read-Value 'VCF version'          ($d.vcfVersion      ?? '9.0.0.0')
    $d.siteId          = Read-Value 'Site ID'              ($d.siteId          ?? 'SITE-001')

    # NSX fields often need manual input even in live mode
    $d.nsx.bgpAS   = Read-Value 'NSX BGP AS (local)'      ($d.nsx.bgpAS       ?? '65001')
    $d.nsx.bgpPeer = Read-Value 'NSX BGP upstream peer'   ($d.nsx.bgpPeer     ?? '192.168.10.1 (AS 65000)')
    $d.nsx.tier0   = Read-Value 'Tier-0 Gateway name'     ($d.nsx.tier0       ?? 'T0-GW-Primary')
    $d.nsx.tier1   = Read-Value 'Tier-1 Gateway name'     ($d.nsx.tier1       ?? 'T1-GW-Workload')

    # SMTP
    $d.smtp.server      = Read-Value 'SMTP server'         ($d.smtp.server      ?? 'mail.domain.local')
    $d.smtp.adminEmail  = Read-Value 'Admin email'         ($d.smtp.adminEmail  ?? 'vcf-admin@domain.local')
    $d.smtp.alertEmail  = Read-Value 'Alert email'         ($d.smtp.alertEmail  ?? 'vcf-alerts@domain.local')

    # Security
    $d.security.ca = Read-Value 'Certificate Authority'   ($d.security.ca ?? 'Internal CA — PKI integrated')

    return $d
}

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

Write-Banner
Disable-TlsChecks

# ── Gather data ───────────────────────────────────────────────────────────────
if ($ManualMode) {
    Write-Host '  Mode: MANUAL — no live connection' -ForegroundColor Yellow
    $data = Get-ManualData
} else {
    Write-Host "  Mode: LIVE — connecting to $SddcManagerFqdn" -ForegroundColor Green

    # Get credentials
    if (-not $SddcCredential) {
        Write-Host ''
        $SddcCredential = Get-Credential -Message "SDDC Manager credentials for $SddcManagerFqdn"
    }

    Connect-SddcManager -Fqdn $SddcManagerFqdn -Cred $SddcCredential
    $data = Get-LiveData -Cred $SddcCredential

    # Confirm/fill gaps after live collection
    Write-Host ''
    $proceed = Read-Host '  Review and fill in any missing values? [Y/n]'
    if ($proceed -ne 'n' -and $proceed -ne 'N') {
        $data = Confirm-Data -d $data
    }
}

# ── Set theme ─────────────────────────────────────────────────────────────────
$data.theme = $Theme

# ── Determine output paths ────────────────────────────────────────────────────
$dateStamp   = Get-Date -Format 'yyyy-MM-dd'
$safeName    = ($data.customer -replace '[^\w]', '_')
$jsonOut     = Join-Path $ScriptDir "vcf9-config_${safeName}_${dateStamp}.json"
if (-not $OutputDocx) {
    $OutputDocx = Join-Path $ScriptDir "VCF9_AsBuilt_${safeName}_${dateStamp}.docx"
}

# ── Write JSON config ─────────────────────────────────────────────────────────
Write-Section 'Writing Config JSON'
$data | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonOut -Encoding UTF8
Write-Ok "Config JSON: $jsonOut"

if ($ConfigJsonOnly) {
    Write-Host ''
    Write-Ok 'ConfigJsonOnly flag set — skipping document generation.'
    Write-Host "  Review $jsonOut, then run:" -ForegroundColor Gray
    Write-Host "    node `"$GeneratorScript`" --config `"$jsonOut`" --output `"$OutputDocx`"" -ForegroundColor White
    exit 0
}

# ── Invoke Node.js generator ──────────────────────────────────────────────────
Write-Section 'Generating Word Document'

if (-not (Test-Path $GeneratorScript)) {
    Write-Err "Generator script not found: $GeneratorScript"
    Write-Host "  Download or copy generate-vcf9-asbuilt.js to: $ScriptDir" -ForegroundColor Yellow
    exit 1
}

# Check Node is available
try {
    $nodeVer = & $NodePath --version 2>&1
    Write-Ok "Node.js: $nodeVer"
} catch {
    Write-Err "Node.js not found at '$NodePath'. Install from https://nodejs.org or set -NodePath."
    exit 1
}

# Check docx npm package
$nodeModCheck = & $NodePath -e "require('docx'); process.exit(0)" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Wrn "npm package 'docx' not found. Installing globally..."
    & npm install -g docx
    if ($LASTEXITCODE -ne 0) {
        Write-Err "Failed to install 'docx' npm package. Run: npm install -g docx"
        exit 1
    }
}

$nodeArgs = @(
    "`"$GeneratorScript`"",
    '--config', "`"$jsonOut`"",
    '--output', "`"$OutputDocx`"",
    '--theme',  $Theme,
    '--silent'
)

Write-Inf "Running: node $($nodeArgs -join ' ')"
Write-Host ''

$proc = Start-Process -FilePath $NodePath `
    -ArgumentList $nodeArgs `
    -NoNewWindow -Wait -PassThru

if ($proc.ExitCode -eq 0) {
    Write-Host ''
    Write-Ok "Report generated successfully!"
    Write-Host ''
    Write-Host "  Document : $OutputDocx"   -ForegroundColor White
    Write-Host "  Config   : $jsonOut"      -ForegroundColor Gray
    Write-Host ''

    # Offer to open the document
    if ($IsWindows -or $env:OS -like '*Windows*') {
        $open = Read-Host '  Open document now? [Y/n]'
        if ($open -ne 'n' -and $open -ne 'N') {
            Start-Process $OutputDocx
        }
    }
} else {
    Write-Err "Node.js generator exited with code $($proc.ExitCode)"
    Write-Host "  Check output above for error details." -ForegroundColor Yellow
    exit $proc.ExitCode
}
