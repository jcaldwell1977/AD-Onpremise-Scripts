#!/usr/bin/env node
/**
 * ╔══════════════════════════════════════════════════════════════╗
 *  VCF 9 As-Built Configuration Report Generator
 *  Branded for  ReDesign Group  |  redesign-group.com
 * ╚══════════════════════════════════════════════════════════════╝
 *
 *  Usage:
 *    node generate-vcf9-asbuilt.js
 *    node generate-vcf9-asbuilt.js --config ./site.json
 *    node generate-vcf9-asbuilt.js --output ./reports/vcf9.docx
 *    node generate-vcf9-asbuilt.js --theme rdg       (ReDesign Group navy/teal)
 *    node generate-vcf9-asbuilt.js --theme dell      (Dell/template navy/blue — default)
 *    node generate-vcf9-asbuilt.js --silent           (no prompts, all defaults)
 *
 *  Color themes (selectable via --theme flag or config key "theme"):
 *    "dell"   Navy 00447C / Blue 16609E / Sky 007DB8  (matches Isuzu/Dell template)
 *    "rdg"    Navy 1A3A5C / Teal 0078D4 / Steel 2E75B6 (ReDesign Group brand)
 *
 *  Config JSON accepts (all optional — prompts fire for anything missing):
 *    theme, customer, customerAddress, preparedBy, reviewedBy,
 *    version, docDate, vcfVersion, siteId,
 *    sddc{}, vCenter{management{},workload{}}, nsx{}, vsan{},
 *    networking{}, aria{lcm{},logInsight{},operations{},automation{}},
 *    licensing{}, security{}, smtp{}, contacts[], issues[], datastores[]
 */

"use strict";

const fs       = require("fs");
const path     = require("path");
const readline = require("readline");

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumberElement, PageBreak,
  LevelFormat, TableOfContents, TabStopType, TabStopPosition,
} = require("docx");

// ── CLI ───────────────────────────────────────────────────────────────────────
const args     = process.argv.slice(2);
const getArg   = (f) => { const i = args.indexOf(f); return i > -1 ? args[i + 1] : null; };
const CFG_PATH = getArg("--config") || null;
const OUT_PATH = getArg("--output") || path.join(__dirname, "VCF9_AsBuilt_ReDesignGroup.docx");
const SILENT   = args.includes("--silent");
const CLI_THEME = getArg("--theme") || null;

// ── Load config ───────────────────────────────────────────────────────────────
let FC = {};
if (CFG_PATH) {
  try   { FC = JSON.parse(fs.readFileSync(CFG_PATH, "utf8")); console.log(`✔  Config: ${CFG_PATH}`); }
  catch (e) { console.error(`✖  Config error: ${e.message}`); process.exit(1); }
}

// ── Color Themes ─────────────────────────────────────────────────────────────
const THEMES = {
  // Matches the Isuzu/Dell template exactly
  dell: {
    navy:        "00447C",
    blue:        "16609E",
    midBlue:     "007DB8",
    lightBlue:   "6BACDE",
    titleBar:    "00447C",
    headerRow:   "16609E",
    rowLight:    "F4F4F4",
    rowMid:      "E4E4E4",
    white:       "FFFFFF",
    textDark:    "404040",
    pageText:    "808080",
    divider:     "007DB8",
  },
  // ReDesign Group brand palette
  rdg: {
    navy:        "1A3A5C",
    blue:        "2E75B6",
    midBlue:     "0078D4",
    lightBlue:   "5BA3D9",
    titleBar:    "1A3A5C",
    headerRow:   "2E75B6",
    rowLight:    "EEF4FB",
    rowMid:      "D9E8F5",
    white:       "FFFFFF",
    textDark:    "1A1A2E",
    pageText:    "707070",
    divider:     "0078D4",
  },
};

// ── Interactive prompt ────────────────────────────────────────────────────────
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
let rlClosed = false;
function prompt(question, def) {
  return new Promise((resolve) => {
    if (SILENT || rlClosed) return resolve(def || "");
    const hint = def ? ` [${def}]` : "";
    rl.question(`  ${question}${hint}: `, (ans) => resolve(ans.trim() || def || ""));
  });
}
function closeRL() { if (!rlClosed) { rl.close(); rlClosed = true; } }

// ── Layout constants (US Letter, matching template margins) ───────────────────
const CONTENT_W = 9360;   // 12240 - 1152*2 (template uses 0.8" margins)

// ── Cell / table helpers ──────────────────────────────────────────────────────
const border = (col) => ({ style: BorderStyle.SINGLE, size: 1, color: col });
const cellBorders = (col = "CCCCCC") => ({ top: border(col), bottom: border(col), left: border(col), right: border(col) });

function titleBarCell(C, text, totalW, span) {
  return new TableCell({
    columnSpan: span,
    borders:    cellBorders(C.titleBar),
    width:      { size: totalW, type: WidthType.DXA },
    shading:    { fill: C.titleBar, type: ShadingType.CLEAR },
    margins:    { top: 90, bottom: 90, left: 150, right: 150 },
    children:   [new Paragraph({ children: [
      new TextRun({ text, font: "Arial", size: 20, bold: true, color: C.white })
    ]})],
  });
}

function hdrCell(C, text, w) {
  return new TableCell({
    borders:       cellBorders(C.headerRow),
    width:         { size: w, type: WidthType.DXA },
    shading:       { fill: C.headerRow, type: ShadingType.CLEAR },
    margins:       { top: 80, bottom: 80, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children:      [new Paragraph({ children: [
      new TextRun({ text, font: "Arial", size: 18, bold: true, color: C.white })
    ]})],
  });
}

function datCell(C, text, w, alt = false, bold = false) {
  return new TableCell({
    borders:       cellBorders("CCCCCC"),
    width:         { size: w, type: WidthType.DXA },
    shading:       { fill: alt ? C.rowLight : C.white, type: ShadingType.CLEAR },
    margins:       { top: 70, bottom: 70, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children:      [new Paragraph({ children: [
      new TextRun({ text: String(text ?? "—"), font: "Arial", size: 18, color: C.textDark, bold })
    ]})],
  });
}

function lblCell(C, text, w, alt = false) {
  return new TableCell({
    borders:       cellBorders("CCCCCC"),
    width:         { size: w, type: WidthType.DXA },
    shading:       { fill: alt ? C.rowMid : C.rowLight, type: ShadingType.CLEAR },
    margins:       { top: 70, bottom: 70, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children:      [new Paragraph({ children: [
      new TextRun({ text, font: "Arial", size: 18, bold: true, color: C.blue })
    ]})],
  });
}

// 2-column key-value table with optional title bar
function kvTable(C, title, rows, labelW = 3000) {
  const valW = CONTENT_W - labelW;
  const trows = [];
  if (title) trows.push(new TableRow({ children: [titleBarCell(C, title, CONTENT_W, 2)] }));
  rows.forEach(([k, v], i) => trows.push(new TableRow({ children: [
    lblCell(C, k, labelW, i % 2 === 1),
    datCell(C, v, valW,   i % 2 === 1),
  ]})));
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [labelW, valW], rows: trows });
}

// Multi-column table with optional title bar + header row
function multiTable(C, title, headers, colWidths, rows) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  const trows = [];
  if (title) trows.push(new TableRow({ children: [titleBarCell(C, title, total, headers.length)] }));
  trows.push(new TableRow({ tableHeader: true, children: headers.map((h, i) => hdrCell(C, h, colWidths[i])) }));
  rows.forEach((row, ri) => trows.push(new TableRow({
    children: row.map((cell, ci) => datCell(C, cell, colWidths[ci], ri % 2 === 1))
  })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: colWidths, rows: trows });
}

// ── Typography helpers ────────────────────────────────────────────────────────
const h1  = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: "Arial" })] });
const h2  = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: "Arial" })] });
const h3  = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: "Arial" })] });
const body = (C, t) => new Paragraph({ spacing: { before: 60, after: 100 }, children: [new TextRun({ text: t, font: "Arial", size: 18, color: C.textDark })] });
const sp  = (pt = 160) => new Paragraph({ spacing: { before: pt, after: 0 }, children: [new TextRun("")] });
const pb  = () => new Paragraph({ children: [new PageBreak()] });

// ══════════════════════════════════════════════════════════════════════════════
//  SECTION BUILDERS
// ══════════════════════════════════════════════════════════════════════════════

// ── Cover Page ────────────────────────────────────────────────────────────────
function buildCover(C, D, logoBuffer) {
  const logo = logoBuffer ? [new Paragraph({
    spacing: { before: 0, after: 280 },
    children: [new ImageRun({ data: logoBuffer, type: "jpg", transformation: { width: 240, height: 42 } })],
  })] : [];

  const prepTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
    rows: [
      // Blue header row
      new TableRow({ children: [
        hdrCell(C, "Prepared for:", CONTENT_W / 2),
        hdrCell(C, "Prepared by:",  CONTENT_W / 2),
      ]}),
      // Content row
      new TableRow({ children: [
        new TableCell({
          borders: cellBorders("CCCCCC"),
          width:   { size: CONTENT_W / 2, type: WidthType.DXA },
          shading: { fill: C.rowLight, type: ShadingType.CLEAR },
          margins: { top: 140, bottom: 140, left: 200, right: 200 },
          children: [
            new Paragraph({ children: [new TextRun({ text: D.customer,        font: "Arial", size: 22, bold: true,  color: C.navy })] }),
            new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: D.customerAddress, font: "Arial", size: 18, color: C.textDark })] }),
          ],
        }),
        new TableCell({
          borders: cellBorders("CCCCCC"),
          width:   { size: CONTENT_W / 2, type: WidthType.DXA },
          shading: { fill: C.rowLight, type: ShadingType.CLEAR },
          margins: { top: 140, bottom: 140, left: 200, right: 200 },
          children: [
            new Paragraph({ children: [new TextRun({ text: "ReDesign Group", font: "Arial", size: 22, bold: true, color: C.navy })] }),
            new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: "redesign-group.com", font: "Arial", size: 18, color: C.midBlue })] }),
          ],
        }),
      ]}),
      // Date / Site ID row
      new TableRow({ children: [
        new TableCell({
          borders: cellBorders("CCCCCC"), width: { size: CONTENT_W / 2, type: WidthType.DXA },
          shading: { fill: C.white, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 200, right: 200 },
          children: [new Paragraph({ children: [
            new TextRun({ text: "Date:  ", font: "Arial", size: 18, bold: true,  color: C.textDark }),
            new TextRun({ text: D.docDate,  font: "Arial", size: 18, color: C.textDark }),
          ]})],
        }),
        new TableCell({
          borders: cellBorders("CCCCCC"), width: { size: CONTENT_W / 2, type: WidthType.DXA },
          shading: { fill: C.white, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 200, right: 200 },
          children: [new Paragraph({ children: [
            new TextRun({ text: "Site ID:  ", font: "Arial", size: 18, bold: true,  color: C.textDark }),
            new TextRun({ text: D.siteId,    font: "Arial", size: 18, color: C.textDark }),
          ]})],
        }),
      ]}),
    ],
  });

  return [
    ...logo,
    sp(320),
    new Paragraph({
      spacing: { before: 0, after: 0 },
      border:  { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy, space: 4 } },
      children: [new TextRun({ text: "VMware Cloud Foundation 9", font: "Arial", size: 56, bold: true, color: C.navy })],
    }),
    new Paragraph({
      spacing: { before: 80, after: 320 },
      children: [new TextRun({ text: "As-Built Configuration Report", font: "Arial", size: 40, color: C.blue })],
    }),
    prepTable,
    sp(220),
    new Paragraph({
      spacing: { before: 0, after: 80 },
      children: [
        new TextRun({ text: "This As-Built Configuration Report", font: "Arial", size: 17, bold: true, color: C.textDark }),
        new TextRun({ text: " is the Confidential Information of ReDesign Group.", font: "Arial", size: 17, color: C.textDark }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: `Copyright \u00A9 ${new Date().getFullYear()} ReDesign Group. All Rights Reserved.`, font: "Arial", size: 17, color: C.textDark })],
    }),
    new Paragraph({
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "ReDesign Group believes the information in this publication is accurate as of its publication date. The information is subject to change without notice.", font: "Arial", size: 17, italics: true, color: C.pageText })],
    }),
  ];
}

// ── Record of Revisions ───────────────────────────────────────────────────────
function buildRevisions(C, D) {
  return [
    h1("Record of Revisions"),
    body(C, "The following is a list of revisions made to this document:"),
    sp(100),
    multiTable(C, null,
      ["Rev", "Date", "Pages Affected", "Reason", "Summary of Technical Changes"],
      [720, 1600, 1700, 1400, 3940],
      [["1.0", D.docDate, "All", "\u2014", "Initial as-built document release."]]
    ),
  ];
}

// ── Purpose ───────────────────────────────────────────────────────────────────
function buildPurpose(C, D) {
  return [
    pb(),
    h1("Purpose of this Document"),
    body(C,
      `This As-Built Configuration Report is produced by ReDesign Group as part of ${D.customer}'s ` +
      `VMware Cloud Foundation 9 deployment engagement. ReDesign Group personnel and authorized agents ` +
      `use and update this document while executing the implementation.`
    ),
    sp(80),
    body(C, "This As-Built Configuration contains:"),
    sp(80),
    multiTable(C, null,
      ["Section", "Contents"],
      [2800, 6560],
      [
        ["Customer & Contacts",   "Customer and ReDesign Group contact information"],
        ["System Configuration",  "SDDC Manager, vCenter, NSX, vSAN configuration details"],
        ["Network Settings",      "Physical and virtual network topology, VLANs, routing"],
        ["Security",              "Hardening status, certificates, encryption, access controls"],
        ["Aria Suite",            "Log Insight, Operations, Automation, and LCM configuration"],
        ["Licensing",             "VCF 9 Universal license entitlements and key locations"],
        ["Software Settings",     "Storage options, MTrees, datastore configurations"],
        ["Backup & Recovery",     "Backup schedules, SMTP/alerting, and recovery targets"],
        ["Issues & Resolutions",  "Engagement issues encountered and how they were resolved"],
      ]
    ),
    sp(140),
    body(C,
      `Use this As-Built Configuration as an ongoing VCF 9 technical reference for ${D.customer} IT ` +
      `personnel and ReDesign Group after the implementation is completed. It is ${D.customer}'s ` +
      `responsibility to ensure this document remains current after implementation.`
    ),
  ];
}

// ── System Configuration ──────────────────────────────────────────────────────
function buildSystem(C, D) {
  const s = D.sddc;
  return [
    pb(),
    h1(`System Configuration \u2014 ${s.hostname}`),
    h2("VCF 9 System Summary"),
    body(C,
      "The following tables summarize the VMware Cloud Foundation 9 deployment. " +
      "SDDC Manager is the central orchestration plane for lifecycle management, " +
      "certificate management, password management, and workload domain operations. " +
      `Table 1 summarizes the system for this engagement.`
    ),
    sp(120),
    kvTable(C, `${s.hostname} \u2014 ${s.serviceTag}`, [
      ["Model / Appliance",     "VCF 9 SDDC Manager"],
      ["FQDN",                  s.hostname],
      ["IP Address",            s.ip],
      ["Build Version",         s.version],
      ["VCF Version",           D.vcfVersion],
      ["Management Domain",     s.domain],
      ["Datacenter",            s.datacenter],
      ["Management Cluster",    s.cluster],
      ["SSO Administrator",     s.ssoUser],
      ["NTP Servers",           s.ntp],
      ["DNS Servers",           s.dns],
      ["Backup Destination",    s.backupDest],
    ]),
    sp(200),
    h2("vSAN Disk Storage Summary"),
    body(C,
      `The vSAN datastore provides hyper-converged storage across all ESXi hosts. ` +
      `Total addressable capacity: ${D.vsan.capacity}. ` +
      `Deduplication and compression are ${D.vsan.dedup === "Enabled" ? "enabled" : "disabled"}.`
    ),
    sp(120),
    kvTable(C, "vSAN Storage Configuration", [
      ["vSAN Version",         D.vsan.version],
      ["Storage Policy",       D.vsan.policy],
      ["Total Capacity",       D.vsan.capacity],
      ["Deduplication",        D.vsan.dedup],
      ["Compression",         D.vsan.compression],
      ["Encryption",           D.vsan.encryption],
      ["Fault Domains",        D.vsan.faultDomains],
      ["Stretch Cluster",      D.vsan.stretchCluster],
      ["File Services",        D.vsan.fileServices],
      ["Proactive Rebalance",  "Enabled (threshold: 70%)"],
      ["Resync Throttle",      "Enabled (30% cap during business hours)"],
    ]),
  ];
}

// ── License Keys ──────────────────────────────────────────────────────────────
function buildLicenses(C, D) {
  const l = D.licensing;
  return [
    sp(200),
    h2("License Keys"),
    body(C,
      "VCF 9 uses a Universal License model. The following table lists the license entitlements " +
      "purchased and installed for this engagement. The license schema used is ELMS " +
      "(Electronic License Management System). License keys are managed via " +
      "SDDC Manager \u2192 Administration \u2192 Licensing."
    ),
    sp(120),
    multiTable(C, "VCF 9 Universal License Entitlements",
      ["Feature / Component", "License Type", "Entitlement Detail"],
      [3200, 2000, 4160],
      [
        ["VMware Cloud Foundation 9",       "Universal",  l.vcf],
        ["vCenter Server",                  "Included",   l.vCenter],
        ["NSX (Data Center Enterprise+)",   "Included",   l.nsx],
        ["vSAN Enterprise",                 "Included",   l.vsan],
        ["Aria Suite Enterprise",           "Included",   l.aria],
        ["Licensed Core Count",             "\u2014",     l.cores],
      ]
    ),
  ];
}

// ── Network Settings ──────────────────────────────────────────────────────────
function buildNetwork(C, D) {
  const n = D.networking;
  return [
    pb(),
    h1("Network Settings"),
    h2("DNS Settings"),
    body(C,
      "Each VCF 9 component has a Fully Qualified Domain Name (FQDN) registered in corporate DNS. " +
      "DNS records (forward and reverse) have been added to all corporate DNS servers to allow " +
      "hostname resolution for all management plane components."
    ),
    sp(120),
    kvTable(C, "DNS Servers", [
      ["Primary DNS",   n.dns1],
      ["Secondary DNS", n.dns2],
      ["Tertiary DNS",  n.dns3 || "None"],
      ["DNS Domain",    n.dnsDomain],
    ], 2800),
    sp(200),
    h2("NTP Settings"),
    body(C,
      "NTP is configured on all VCF 9 components. Time synchronization is required for " +
      "Active Directory authentication and accurate log timestamps. The maximum allowable " +
      "drift between any component and the AD domain controller is 5 minutes."
    ),
    sp(120),
    kvTable(C, "NTP Configuration", [
      ["Primary NTP Server",   n.ntp1],
      ["Secondary NTP Server", n.ntp2 || "None"],
      ["Time Zone",            n.timezone || "US/Eastern"],
      ["Sync Status",          "Synchronized \u2014 all components"],
    ], 2800),
    sp(200),
    h2("Interface & VLAN Settings"),
    body(C,
      "All traffic types are segregated onto dedicated VLANs across the fabric. " +
      `Jumbo frames (MTU ${n.mtu}) are enabled across all storage and TEP networks. ` +
      `Physical uplinks use ${n.uplinks}.`
    ),
    sp(120),
    multiTable(C, "VLAN / Network Assignments",
      ["Traffic Type", "VLAN / Subnet", "Gateway", "Purpose"],
      [2200, 2400, 2000, 2760],
      [
        ["Management",  n.mgmtVlan,    n.mgmtGw,    "SDDC Manager, vCenter, NSX Managers, Aria Suite"],
        ["vMotion",     n.vmotionVlan, n.vmotionGw, "Live VM migration traffic"],
        ["vSAN",        n.vsanVlan,    n.vsanGw,    "Storage I/O between ESXi hosts"],
        ["NSX Overlay", n.overlayVlan, n.overlayGw, "Geneve TEP tunnel endpoints"],
        ["Host Uplink", n.uplinkVlan,  n.uplinkGw,  "Physical host uplink / TEP backup"],
        ["Workload",    n.workloadVlan,"n/a",        "Tenant VM network segments (NSX segments)"],
      ]
    ),
    sp(200),
    h2("Bonding / Port Groups"),
    body(C, "Physical uplinks are bonded to provide redundancy and throughput. The following table documents bonded interface parameters."),
    sp(120),
    kvTable(C, "Physical Network & vSphere Distributed Switch Configuration", [
      ["Top-of-Rack Switches",     n.torSwitch],
      ["vSphere Distributed Switches", n.dvSwitch],
      ["Uplinks per Host",         n.uplinks],
      ["MTU (Fabric-wide)",        n.mtu],
      ["Aggregation Type",         n.lacp],
      ["CDP / LLDP",               n.lldp || "LLDP \u2014 Listen and Advertise"],
    ], 3200),
    sp(200),
    h2("Routing Tables"),
    kvTable(C, "Routing Configuration", [
      ["Tier-0 Gateway",       D.nsx.tier0],
      ["T0 HA Mode",           "Active/Active"],
      ["Tier-1 Gateway",       D.nsx.tier1],
      ["BGP AS (Local)",       D.nsx.bgpAS],
      ["BGP Upstream Peer",    D.nsx.bgpPeer],
      ["BFD",                  "Enabled"],
      ["ECMP",                 "Enabled (up to 8 paths)"],
      ["Default Gateway",      n.defaultGw],
    ], 3200),
  ];
}

// ── Security ──────────────────────────────────────────────────────────────────
function buildSecurity(C, D) {
  const sec = D.security;
  return [
    pb(),
    h1("Security Configuration"),
    h2("Trust and Certificate Relationships"),
    body(C,
      "TLS certificates have been issued by the configured Certificate Authority for all VCF 9 " +
      "management plane components. The following tables document certificate and trust relationships."
    ),
    sp(120),
    multiTable(C, "Admin Access \u2014 Trust Relationships",
      ["Subject", "Type", "Valid From", "Valid Until"],
      [3200, 1600, 2200, 2360],
      [
        [D.sddc.hostname,               "trusted-ca", sec.certFrom || "\u2014", sec.certUntil || "\u2014"],
        [D.vCenter.management.hostname, "host",       sec.certFrom || "\u2014", sec.certUntil || "\u2014"],
        [D.vCenter.workload.hostname,   "host",       sec.certFrom || "\u2014", sec.certUntil || "\u2014"],
        [D.nsx.vip,                     "host",       sec.certFrom || "\u2014", sec.certUntil || "\u2014"],
      ]
    ),
    sp(200),
    h2("Authentication & Directory Services"),
    body(C,
      "There are three types of authentication that VCF 9 components can use: Workgroup, Domain, " +
      "and Active Directory (AD). All management components are integrated with the configured " +
      "identity provider. MFA is enforced for all administrative access."
    ),
    sp(120),
    kvTable(C, "Authentication Configuration", [
      ["SSO Domain",            sec.ssoDomain  || "vsphere.local"],
      ["Identity Provider",     sec.idp        || "Workspace ONE Access (SAML 2.0)"],
      ["MFA Method",            sec.mfa],
      ["AD Realm / Domain",     sec.adRealm    || "Configured"],
      ["CIFS Auth Mode",        sec.cifsAuth   || "Active Directory"],
      ["NSX Identity Firewall", sec.idfw       || "Enabled (AD groups)"],
      ["Password Policy",       sec.passwords],
    ], 3000),
    sp(200),
    h2("Access Protocol Settings"),
    body(C, "The following table documents the state of management access protocols across all VCF 9 components."),
    sp(120),
    multiTable(C, "Access Protocol Settings",
      ["Protocol", "Status", "Allowed Hosts / Notes"],
      [1800, 1600, 5960],
      [
        ["HTTPS",  "Enabled",  "TLS 1.2/1.3 only \u2014 primary management access for all components"],
        ["SSH",    sec.ssh    || "Disabled on ESXi",  "Enabled on SDDC Manager for support sessions"],
        ["SCP",    "Enabled",  "SDDC Manager backup transport"],
        ["HTTP",   "Disabled", "Redirected to HTTPS"],
        ["Telnet", "Disabled", "Disabled on all appliances"],
        ["FTP",    "Disabled", "Not required"],
        ["FTPS",   "Disabled", "Not required"],
      ]
    ),
    sp(200),
    h2("File System Encryption"),
    kvTable(C, "Encryption Configuration", [
      ["vSAN Encryption",            D.vsan.encryption],
      ["Backup Encryption",          sec.backupEncr],
      ["Encryption Strength",        sec.backupEncStrength || "AES-256"],
      ["TLS Version Policy",         sec.tls || "TLS 1.2 / 1.3 enforced"],
      ["Certificate Authority",      sec.ca],
      ["FIPS 140-2 Mode",            "Enabled (where supported)"],
    ], 3000),
    sp(200),
    h2("Hardening Actions Applied"),
    body(C, `Security hardening applied per ${sec.compliance}. The following table documents the status of each hardening item.`),
    sp(120),
    multiTable(C, null,
      ["Hardening Item", "Component", "Status"],
      [4960, 2000, 2400],
      [
        ["SSH disabled on all ESXi hosts",                    "ESXi",       "Applied"],
        ["ESXi shell timeout configured (900 sec)",           "ESXi",       "Applied"],
        ["Syslog forwarded to Aria Log Insight",              "ESXi",       "Applied"],
        [`Normal Lockdown Mode enabled — ${sec.lockdown}`,   "ESXi",       "Applied"],
        ["vCenter TLS 1.0 / 1.1 disabled",                   "vCenter",    "Applied"],
        ["vCenter idle session timeout configured",           "vCenter",    "Applied"],
        ["SDDC Manager password rotation policy enforced",    "SDDC Mgr",   "Applied"],
        ["NSX Admin / Audit account separation enforced",     "NSX",        "Applied"],
        ["DFW default rule logging enabled",                  "NSX",        "Applied"],
        ["Aria service accounts rotated post-deployment",     "Aria Suite", "Applied"],
        ["FIPS 140-2 mode enabled (where supported)",         "All",        "Applied"],
        ["NTP synchronized on all management components",     "All",        "Verified"],
        ["Backup encryption enabled",                         "SDDC Mgr",   "Applied"],
      ]
    ),
  ];
}

// ── vCenter ───────────────────────────────────────────────────────────────────
function buildVCenter(C, D) {
  const vc = D.vCenter;
  return [
    pb(),
    h1("vCenter Server Configuration"),
    body(C,
      "VCF 9 deploys independent vCenter Server instances per domain. The Management vCenter " +
      "governs all infrastructure VMs. The Workload vCenter manages tenant workloads and is " +
      "connected to NSX for software-defined networking."
    ),
    sp(120),
    kvTable(C, `Management Domain vCenter \u2014 ${vc.management.hostname}`, [
      ["FQDN",                  vc.management.hostname],
      ["IP Address",            vc.management.ip],
      ["Version / Build",       vc.management.version],
      ["SSO Domain",            vc.management.sso || "vsphere.local"],
      ["vCenter HA",            vc.management.ha],
      ["Managed Cluster",       vc.management.cluster],
      ["ESXi Host Count",       vc.management.hosts],
      ["DRS Mode",              "Fully Automated"],
      ["HA Admission Control",  "50% reserved capacity"],
      ["EVC Mode",              "Configured per host generation"],
    ]),
    sp(200),
    kvTable(C, `Workload Domain vCenter \u2014 ${vc.workload.hostname}`, [
      ["FQDN",                  vc.workload.hostname],
      ["IP Address",            vc.workload.ip],
      ["Version / Build",       vc.workload.version],
      ["Managed Cluster",       vc.workload.cluster],
      ["ESXi Host Count",       vc.workload.hosts],
      ["DRS Mode",              "Fully Automated"],
      ["HA Admission Control",  "50% reserved capacity"],
      ["NSX Integration",       `${D.nsx.vip} (VIP)`],
    ]),
  ];
}

// ── NSX ───────────────────────────────────────────────────────────────────────
function buildNSX(C, D) {
  const n = D.nsx;
  return [
    pb(),
    h1("NSX \u2014 Network Virtualization"),
    body(C,
      "NSX provides software-defined networking for the VCF 9 deployment including distributed " +
      "switching and routing, micro-segmentation, distributed firewall, and load balancing."
    ),
    sp(120),
    kvTable(C, `NSX Manager Cluster \u2014 VIP: ${n.vip}`, [
      ["Manager VIP",       n.vip],
      ["Manager Node 1",    n.manager1],
      ["Manager Node 2",    n.manager2],
      ["Manager Node 3",    n.manager3],
      ["Version / Build",   n.version],
      ["Cluster Mode",      "3-Node Active/Active/Active HA"],
    ]),
    sp(200),
    kvTable(C, "Transport Zones", [
      ["Overlay Transport Zone",  n.overlayTZ],
      ["VLAN Transport Zone",     n.vlanTZ],
      ["TEP Network",             D.networking.overlayVlan],
      ["MTU (Geneve / TEP)",      D.networking.mtu],
    ]),
    sp(200),
    kvTable(C, `Edge Cluster \u2014 ${n.edgeCluster}`, [
      ["Edge Cluster Name",  n.edgeCluster],
      ["Edge Node 1",        n.edgeNode1],
      ["Edge Node 2",        n.edgeNode2],
      ["Edge Form Factor",   "Large (16 vCPU / 64 GB RAM)"],
      ["Deployment vCenter", D.vCenter.management.hostname],
      ["HA Mode",            "Active/Standby per T0"],
    ]),
    sp(200),
    kvTable(C, "Distributed Firewall Configuration", [
      ["Default Rule Policy",    "Allow (with logging)"],
      ["Micro-segmentation",     "Enabled \u2014 workload domain"],
      ["Identity Firewall",      "Configured (AD integration)"],
      ["FQDN Filtering",         "Enabled"],
      ["Service Insertion",      n.serviceInsertion || "Not configured"],
      ["Distributed IDS/IPS",    n.ids              || "Not configured"],
    ]),
  ];
}

// ── Aria Suite ────────────────────────────────────────────────────────────────
function buildAria(C, D) {
  const a = D.aria;
  return [
    pb(),
    h1("Aria Suite \u2014 Operations & Automation"),
    body(C,
      "Broadcom Aria Suite provides centralized operations, log management, and automation. " +
      "Lifecycle management for all Aria products is handled by Aria Lifecycle Manager (LCM)."
    ),
    sp(120),
    kvTable(C, `Aria Lifecycle Manager \u2014 ${a.lcm.hostname}`, [
      ["FQDN",             a.lcm.hostname],
      ["IP Address",       a.lcm.ip],
      ["Version",          a.lcm.version],
      ["Managed Products", "Log Insight, Operations, Automation"],
    ]),
    sp(200),
    kvTable(C, `Aria Log Insight \u2014 ${a.logInsight.hostname}`, [
      ["FQDN",             a.logInsight.hostname],
      ["IP Address",       a.logInsight.ip],
      ["Version",          a.logInsight.version],
      ["Cluster Topology", a.logInsight.cluster],
      ["Log Retention",    a.logInsight.retention],
      ["Syslog Port",      "UDP/TCP 514 | SSL 6514"],
      ["Integrations",     "NSX, vCenter, SDDC Manager, ESXi, Aria Operations"],
    ]),
    sp(200),
    kvTable(C, `Aria Operations \u2014 ${a.operations.hostname}`, [
      ["FQDN",                a.operations.hostname],
      ["IP Address",          a.operations.ip],
      ["Version",             a.operations.version],
      ["Management Adapters", a.operations.adapters],
      ["Alerting",            "Email + SNMP traps configured"],
      ["Capacity Analytics",  "Enabled \u2014 30-day trending"],
    ]),
    sp(200),
    kvTable(C, `Aria Automation \u2014 ${a.automation.hostname}`, [
      ["FQDN",               a.automation.hostname],
      ["IP Address",         a.automation.ip],
      ["Version",            a.automation.version],
      ["Cloud Accounts",     "vCenter (Management), vCenter (Workload)"],
      ["Service Catalog",    "Deployed \u2014 initial blueprints configured"],
      ["Approval Policies",  "Single-level approval for production workloads"],
    ]),
  ];
}

// ── Software Settings ─────────────────────────────────────────────────────────
function buildSoftware(C, D) {
  return [
    pb(),
    h1("Software Settings"),
    h2("File System Cleaning Options"),
    body(C,
      "The vSAN file system is cleaned on a scheduled basis to reclaim space from expired or " +
      "deleted data. The cleaning throttle setting controls maximum system resource usage."
    ),
    sp(120),
    kvTable(C, "File System Cleaning Configuration", [
      ["Throttle",            D.vsan.cleanThrottle || "50%"],
      ["Run Day",             D.vsan.cleanDay      || "Tuesday"],
      ["Run Time",            D.vsan.cleanTime     || "06:00"],
      ["Archive Migration",   "N/A"],
    ], 2800),
    sp(200),
    h2("MTree / Datastore Definitions"),
    body(C,
      "An MTree (or vSAN datastore namespace) is a self-contained filesystem that deduplicates " +
      "and isolates pre-compressed data. Each MTree can have independent quotas and storage policies."
    ),
    sp(120),
    multiTable(C, "MTree / Datastore Summary",
      ["Name / Path", "Pre-Comp Size", "Status", "Anchor Algorithm / Policy"],
      [3200, 1800, 1400, 2960],
      (D.datastores || [
        ["/data/col1/Prod",     "1,276.1 GiB", "RW", "variable / RAID-5"],
        ["/data/col1/Non-Prod", "486.8 GiB",   "RW", "variable / RAID-5"],
        ["/data/col1/backup",   "0.0 GiB",     "RW", "variable / RAID-1"],
      ])
    ),
    sp(200),
    h2("DDBoost / Backup Transport Options"),
    kvTable(C, "DDBoost Configuration", [
      ["DDBoost Status",               "Enabled \u2014 licensed"],
      ["Distributed Segment Processing","Enabled"],
      ["Virtual Synthetics",           "Enabled"],
      ["FC Transport",                 "Disabled (IP-only)"],
      ["Global Authentication Mode",   D.security.backupAuth || "None (certificate-based)"],
      ["Global Encryption Strength",   D.security.backupEncStrength || "Medium (AES-256)"],
      ["If-Group (Load Balancing)",    "Configured \u2014 default group"],
    ], 3200),
    sp(200),
    h2("Virtual Disk / Backup Client Definitions"),
    multiTable(C, "DDBoost Storage Units",
      ["Storage Unit", "Pre-Comp (GiB)", "Status", "User", "Client"],
      [2400, 1600, 1400, 1800, 2160],
      (D.storageUnits || [
        ["Prod",     "1.25",  "RW", "DDBkupadmin", "Veeam Backup Server"],
        ["Non-Prod", "0.48",  "RW", "DDBkupadmin", "Veeam Backup Server"],
      ])
    ),
  ];
}

// ── Backup & Recovery ─────────────────────────────────────────────────────────
function buildBackup(C, D) {
  return [
    pb(),
    h1("Backup & Recovery"),
    h2("Notification Service \u2014 SMTP Settings"),
    body(C,
      "VCF 9 components generate automated alert and autosupport email messages. " +
      "These require an external mail relay to enable transmission. " +
      "Mail relay entries have been added to corporate mail servers."
    ),
    sp(120),
    kvTable(C, "SMTP Configuration", [
      ["SMTP Server",          D.smtp.server],
      ["Administrative Email", D.smtp.adminEmail],
      ["Autosupport Email",    D.smtp.autosupport],
      ["Alerts Email",         D.smtp.alertEmail],
      ["Send Time",            D.smtp.sendTime || "06:00 daily"],
    ], 3000),
    sp(200),
    h2("Backup Schedule & Targets"),
    multiTable(C, "Backup Configuration",
      ["Component", "Method", "Schedule", "Retention", "Destination"],
      [2100, 1800, 1300, 1300, 2860],
      [
        ["SDDC Manager",       "Built-in SCP export",  "Daily",       "30 days",  D.sddc.backupDest],
        ["vCenter (Mgmt)",     "File-based backup",    "Daily",       "14 days",  "Backup server"],
        ["vCenter (Workload)", "File-based backup",    "Daily",       "14 days",  "Backup server"],
        ["NSX Manager",        "Built-in backup",      "Every 6 hrs", "7 days",   "Backup server"],
        ["Aria LCM",           "Snapshot + export",    "Weekly",      "4 weeks",  "Backup server"],
        ["Aria Log Insight",   "VM snapshot",          "Daily",       "7 days",   "vSAN local"],
        ["Aria Operations",    "VM snapshot",          "Daily",       "7 days",   "vSAN local"],
        ["Aria Automation",    "VM snapshot",          "Daily",       "7 days",   "vSAN local"],
      ]
    ),
  ];
}

// ── Contacts ──────────────────────────────────────────────────────────────────
function buildContacts(C, D) {
  return [
    pb(),
    h1("Appendix A: Contact Information"),
    body(C, "The contact information for this engagement is listed below."),
    sp(120),
    multiTable(C, "Project Delivery Resources",
      ["First Name", "Last Name", "Role", "Email", "Phone"],
      [1560, 1560, 2160, 2520, 1560],
      (D.contacts || [
        ["\u2014", "\u2014", "ReDesign Group Lead Engineer",   "\u2014", "\u2014"],
        ["\u2014", "\u2014", "ReDesign Group Project Manager", "\u2014", "\u2014"],
        ["\u2014", "\u2014", "Customer IT Lead",               "\u2014", "\u2014"],
        ["\u2014", "\u2014", "Customer Project Sponsor",       "\u2014", "\u2014"],
      ])
    ),
  ];
}

// ── Issues Log ────────────────────────────────────────────────────────────────
function buildIssues(C, D) {
  return [
    pb(),
    h1("Appendix B: Engagement Issues Log"),
    body(C, "The following table documents issues encountered during the deployment and how each was resolved."),
    sp(120),
    multiTable(C, null,
      ["#", "Date", "Component", "Issue Description", "Resolution", "Status"],
      [480, 1120, 1400, 2260, 2260, 1000],  // total = 8520, fits 9360 padded
      (D.issues || [
        ["1", D.docDate, "\u2014", "No issues recorded during this deployment.", "N/A", "Closed"],
      ])
    ),
  ];
}

// ══════════════════════════════════════════════════════════════════════════════
//  HEADER / FOOTER
// ══════════════════════════════════════════════════════════════════════════════
function makeHeader(C, logoBuffer, D) {
  // Two-column effect: logo left, doc title right
  const children = [];
  if (logoBuffer) {
    children.push(new Paragraph({
      spacing: { before: 0, after: 0 },
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      children: [
        new ImageRun({ data: logoBuffer, type: "jpg", transformation: { width: 160, height: 28 } }),
        new TextRun({ text: "\tVCF 9 As-Built Configuration Report", font: "Arial", size: 16, color: C.pageText }),
      ],
    }));
  } else {
    children.push(new Paragraph({
      spacing: { before: 0, after: 0 },
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      children: [
        new TextRun({ text: "ReDesign Group", font: "Arial", size: 18, bold: true, color: C.navy }),
        new TextRun({ text: "\tVCF 9 As-Built Configuration Report", font: "Arial", size: 16, color: C.pageText }),
      ],
    }));
  }
  return new Header({ children });
}

function makeFooter(C, D) {
  return new Footer({
    children: [new Paragraph({
      spacing: { before: 60, after: 0 },
      border:  { top: { style: BorderStyle.SINGLE, size: 4, color: C.midBlue, space: 4 } },
      tabStops: [
        { type: TabStopType.CENTER, position: CONTENT_W / 2 },
        { type: TabStopType.RIGHT,  position: CONTENT_W },
      ],
      children: [
        new TextRun({ text: "Internal Use \u2014 Confidential", font: "Arial", size: 16, color: C.pageText }),
        new TextRun({ text: "\t" + D.docDate,                   font: "Arial", size: 16, color: C.pageText }),
        new TextRun({ text: "\tPage ",                          font: "Arial", size: 16, color: C.pageText }),
        new TextRun({ children: [new PageNumberElement()],      font: "Arial", size: 16, color: C.blue }),
      ],
    })],
  });
}

// ══════════════════════════════════════════════════════════════════════════════
//  DOCUMENT ASSEMBLY
// ══════════════════════════════════════════════════════════════════════════════
function buildDocument(C, D, logoBuffer) {
  const pageSize   = { width: 12240, height: 15840 };
  const bodyMargin = { top: 230, right: 1152, bottom: 1440, left: 1152, header: 0, footer: 0 };

  return new Document({
    creator:     "ReDesign Group",
    title:       `VCF 9 As-Built \u2014 ${D.customer}`,
    description: "VMware Cloud Foundation 9 As-Built Configuration Report",
    subject:     "VCF 9 As-Built",
    keywords:    "VCF, VMware, NSX, vSAN, Aria, As-Built, ReDesign Group",

    styles: {
      default: {
        document: { run: { font: "Arial", size: 18, color: C.textDark } },
      },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run:       { size: 36, bold: true, font: "Arial", color: C.navy },
          paragraph: {
            spacing:      { before: 480, after: 160 },
            outlineLevel: 0,
            border:       { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.midBlue, space: 4 } },
          },
        },
        {
          id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run:       { size: 28, bold: true, font: "Arial", color: C.blue },
          paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 },
        },
        {
          id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run:       { size: 24, bold: true, font: "Arial", color: C.midBlue },
          paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 2 },
        },
      ],
    },

    numbering: { config: [] },

    sections: [
      // ── Cover (no header/footer, wider top margin) ────────────────────────
      {
        properties: { page: { size: pageSize, margin: { top: 1440, right: 1152, bottom: 1440, left: 1152 } } },
        children:   buildCover(C, D, logoBuffer),
      },
      // ── Body (header + footer + TOC + all sections) ───────────────────────
      {
        properties: { page: { size: pageSize, margin: bodyMargin } },
        headers:    { default: makeHeader(C, logoBuffer, D) },
        footers:    { default: makeFooter(C, D) },
        children: [
          h1("Table of Contents"),
          new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
          pb(),
          ...buildRevisions(C, D),
          ...buildPurpose(C, D),
          ...buildSystem(C, D),
          ...buildLicenses(C, D),
          ...buildNetwork(C, D),
          ...buildSecurity(C, D),
          ...buildVCenter(C, D),
          ...buildNSX(C, D),
          ...buildAria(C, D),
          ...buildSoftware(C, D),
          ...buildBackup(C, D),
          ...buildContacts(C, D),
          ...buildIssues(C, D),
        ],
      },
    ],
  });
}

// ══════════════════════════════════════════════════════════════════════════════
//  MAIN — collect inputs, select theme, build, write
// ══════════════════════════════════════════════════════════════════════════════
(async () => {
  console.log("\n" + "═".repeat(64));
  console.log("  VCF 9 As-Built Generator  |  ReDesign Group");
  console.log("═".repeat(64));
  console.log("  Press Enter to accept default values shown in [brackets].");
  console.log("  Run with --silent to skip all prompts.\n");

  // ── Theme selection ───────────────────────────────────────────────────────
  const rawTheme = CLI_THEME || FC.theme || null;
  let themeName  = "dell";   // default — matches Dell/Isuzu template

  if (rawTheme) {
    if (THEMES[rawTheme]) {
      themeName = rawTheme;
    } else {
      console.warn(`  ⚠  Unknown theme "${rawTheme}" — valid options: dell, rdg. Falling back to "dell".`);
    }
  } else if (!SILENT) {
    const pick = await prompt(
      'Color theme — "dell" (navy/blue, matches template) or "rdg" (ReDesign Group navy/teal)',
      "dell"
    );
    themeName = THEMES[pick] ? pick : "dell";
  }

  const C = THEMES[themeName];
  console.log(`  Theme: ${themeName === "rdg" ? "ReDesign Group (navy/teal)" : "Dell template (navy/blue)"}\n`);

  // ── Helper: config-first, then prompt ────────────────────────────────────
  const ask = async (cfgVal, question, def) =>
    cfgVal !== undefined ? cfgVal : await prompt(question, def);

  // ── Metadata ──────────────────────────────────────────────────────────────
  const docDate         = FC.docDate || new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const customer        = await ask(FC.customer,        "Customer name",                         "Customer Name");
  const customerAddress = await ask(FC.customerAddress, "Customer address (one line)",            "123 Main St, City, ST 00000, US");
  const reviewedBy      = await ask(FC.reviewedBy,      "Reviewed by",                           "\u2014");
  const version         = await ask(FC.version,         "Document version",                      "1.0");
  const vcfVersion      = await ask(FC.vcfVersion,      "VCF version string",                    "9.0.0.0");
  const siteId          = await ask(FC.siteId,          "Site ID / Service Tag",                 "SITE-001");

  // ── SDDC ──────────────────────────────────────────────────────────────────
  const sd = FC.sddc || {};
  const sddcHostname   = await ask(sd.hostname,   "SDDC Manager FQDN",          "sddc-mgr.domain.local");
  const sddcIp         = await ask(sd.ip,         "SDDC Manager IP",            "192.168.10.10");
  const sddcVersion    = await ask(sd.version,    "SDDC Manager version/build", "9.0.0.0 (build 12345678)");
  const sddcDomain     = await ask(sd.domain,     "Management domain name",     "mgmt.domain.local");
  const sddcDatacenter = await ask(sd.datacenter, "Datacenter name",            "DC-Primary");
  const sddcCluster    = await ask(sd.cluster,    "Management cluster name",    "mgmt-cluster-01");
  const sddcNtp        = await ask(sd.ntp,        "NTP servers (comma-sep)",    "ntp1.domain.local, ntp2.domain.local");
  const sddcDns        = await ask(sd.dns,        "DNS servers (comma-sep)",    "192.168.10.1, 192.168.10.2");
  const sddcSso        = await ask(sd.ssoUser,    "SSO admin account",          "administrator@vsphere.local");
  const sddcBackup     = await ask(sd.backupDest, "Backup destination",         "scp://backup.domain.local/vcf/sddc");
  const sddcTag        = await ask(sd.serviceTag, "Service tag / serial",       "XXXXXXX");

  // ── vCenter ───────────────────────────────────────────────────────────────
  const vcM = (FC.vCenter || {}).management || {};
  const vcW = (FC.vCenter || {}).workload   || {};
  const vcMgmtHost  = await ask(vcM.hostname, "vCenter (Mgmt) FQDN",         "vcsa-mgmt-01.domain.local");
  const vcMgmtIp    = await ask(vcM.ip,       "vCenter (Mgmt) IP",           "192.168.10.11");
  const vcMgmtVer   = await ask(vcM.version,  "vCenter (Mgmt) version/build","8.0 U3 (build 23456789)");
  const vcMgmtClus  = await ask(vcM.cluster,  "Management cluster name",     "mgmt-cluster-01");
  const vcMgmtHosts = await ask(vcM.hosts,    "Management domain host count","4");
  const vcMgmtHA    = vcM.ha || "Enabled";
  const vcWldHost   = await ask(vcW.hostname, "vCenter (Workload) FQDN",     "vcsa-wld-01.domain.local");
  const vcWldIp     = await ask(vcW.ip,       "vCenter (Workload) IP",       "192.168.10.12");
  const vcWldVer    = await ask(vcW.version,  "vCenter (Workload) version",  "8.0 U3 (build 23456789)");
  const vcWldClus   = await ask(vcW.cluster,  "Workload cluster name",       "wld-cluster-01");
  const vcWldHosts  = await ask(vcW.hosts,    "Workload domain host count",  "8");

  // ── NSX ───────────────────────────────────────────────────────────────────
  const nx = FC.nsx || {};
  const nsxVip      = await ask(nx.vip,        "NSX Manager VIP IP",          "192.168.10.20");
  const nsxMgr1     = await ask(nx.manager1,   "NSX Manager Node 1 FQDN",     "nsx-mgr-01.domain.local");
  const nsxMgr2     = await ask(nx.manager2,   "NSX Manager Node 2 FQDN",     "nsx-mgr-02.domain.local");
  const nsxMgr3     = await ask(nx.manager3,   "NSX Manager Node 3 FQDN",     "nsx-mgr-03.domain.local");
  const nsxVer      = await ask(nx.version,    "NSX version/build",           "4.2.0.0 (build 34567890)");
  const nsxEdge1    = await ask(nx.edgeNode1,  "Edge Node 1 FQDN",            "nsx-edge-01.domain.local");
  const nsxEdge2    = await ask(nx.edgeNode2,  "Edge Node 2 FQDN",            "nsx-edge-02.domain.local");
  const nsxEdgeClus = nx.edgeCluster   || "edge-cluster-01";
  const nsxT0       = nx.tier0         || "T0-GW-Primary";
  const nsxT1       = nx.tier1         || "T1-GW-Workload";
  const nsxOTZ      = nx.overlayTZ     || "nsx-overlay-transportzone";
  const nsxVTZ      = nx.vlanTZ        || "nsx-vlan-transportzone";
  const nsxBgpAS    = nx.bgpAS         || "65001";
  const nsxBgpPeer  = nx.bgpPeer       || "192.168.10.1 (AS 65000)";

  // ── vSAN ──────────────────────────────────────────────────────────────────
  const vs = FC.vsan || {};
  const vsanCap = await ask(vs.capacity, "vSAN total capacity", "153.6 TB raw / 76.8 TB usable");

  // ── Networking ────────────────────────────────────────────────────────────
  const net = FC.networking || {};
  const netDns1    = await ask(net.dns1,      "Primary DNS IP",            "192.168.10.1");
  const netDns2    = await ask(net.dns2,      "Secondary DNS IP",          "192.168.10.2");
  const netDnsDom  = await ask(net.dnsDomain, "DNS domain",                "domain.local");
  const netTor     = await ask(net.torSwitch, "Top-of-Rack switch model",  "Cisco Nexus 93180YC-FX");
  const netMtu     = net.mtu         || "9000 (Jumbo Frames)";
  const netUplinks = net.uplinks     || "2 x 25GbE per host (active-active LACP)";

  // ── Aria ──────────────────────────────────────────────────────────────────
  const ar = FC.aria || {};
  const arLcmHost  = await ask((ar.lcm        || {}).hostname, "Aria LCM FQDN",           "aria-lcm-01.domain.local");
  const arLcmIp    = await ask((ar.lcm        || {}).ip,       "Aria LCM IP",             "192.168.10.30");
  const arLiHost   = await ask((ar.logInsight || {}).hostname, "Aria Log Insight FQDN",   "aria-li-01.domain.local");
  const arLiIp     = await ask((ar.logInsight || {}).ip,       "Aria Log Insight IP",     "192.168.10.31");
  const arOpsHost  = await ask((ar.operations || {}).hostname, "Aria Operations FQDN",    "aria-ops-01.domain.local");
  const arOpsIp    = await ask((ar.operations || {}).ip,       "Aria Operations IP",      "192.168.10.32");
  const arAutoHost = await ask((ar.automation || {}).hostname, "Aria Automation FQDN",    "aria-auto-01.domain.local");
  const arAutoIp   = await ask((ar.automation || {}).ip,       "Aria Automation IP",      "192.168.10.33");

  // ── SMTP ──────────────────────────────────────────────────────────────────
  const sm = FC.smtp || {};
  const smtpServer = await ask(sm.server,     "SMTP server FQDN/IP",       "mail.domain.local");
  const smtpAdmin  = await ask(sm.adminEmail, "Admin email address",       "vcf-admin@domain.local");
  const smtpAuto   = await ask(sm.autosupport,"Autosupport email",         "autosupport@domain.local");
  const smtpAlert  = await ask(sm.alertEmail, "Alert email address",       "vcf-alerts@domain.local");

  closeRL();

  // ── Compose D (data object) ───────────────────────────────────────────────
  const sec = FC.security || {};
  const lic = FC.licensing || {};

  const D = {
    customer, customerAddress, reviewedBy, version, vcfVersion, siteId, docDate,
    preparedBy: "ReDesign Group",

    sddc: {
      hostname: sddcHostname, ip: sddcIp, version: sddcVersion,
      domain: sddcDomain, datacenter: sddcDatacenter, cluster: sddcCluster,
      ntp: sddcNtp, dns: sddcDns, ssoUser: sddcSso, backupDest: sddcBackup, serviceTag: sddcTag,
    },

    vCenter: {
      management: { hostname: vcMgmtHost, ip: vcMgmtIp, version: vcMgmtVer, sso: "vsphere.local", ha: vcMgmtHA, cluster: vcMgmtClus, hosts: vcMgmtHosts },
      workload:   { hostname: vcWldHost,  ip: vcWldIp,  version: vcWldVer,  cluster: vcWldClus,  hosts: vcWldHosts },
    },

    nsx: {
      vip: nsxVip, manager1: nsxMgr1, manager2: nsxMgr2, manager3: nsxMgr3,
      version: nsxVer, edgeNode1: nsxEdge1, edgeNode2: nsxEdge2,
      edgeCluster: nsxEdgeClus, tier0: nsxT0, tier1: nsxT1,
      overlayTZ: nsxOTZ, vlanTZ: nsxVTZ, bgpAS: nsxBgpAS, bgpPeer: nsxBgpPeer,
      serviceInsertion: nx.serviceInsertion || "Not configured",
      ids: nx.ids,
    },

    vsan: {
      version:      vs.version      || "8.0 U3",
      policy:       vs.policy       || "RAID-5 (Erasure Coding)",
      dedup:        vs.dedup        || "Enabled",
      compression:  vs.compression  || "Enabled",
      encryption:   vs.encryption   || "Enabled (vSAN Encryption)",
      capacity:     vsanCap,
      faultDomains: vs.faultDomains || "3 (Rack-A, Rack-B, Rack-C)",
      stretchCluster: vs.stretchCluster || "Disabled",
      fileServices:   vs.fileServices   || "Disabled",
      cleanThrottle: vs.cleanThrottle, cleanDay: vs.cleanDay, cleanTime: vs.cleanTime,
    },

    networking: {
      dns1: netDns1, dns2: netDns2, dns3: net.dns3, dnsDomain: netDnsDom,
      ntp1: net.ntp1 || sddcNtp.split(",")[0].trim(),
      ntp2: net.ntp2 || (sddcNtp.split(",")[1] || "").trim() || "None",
      timezone:     net.timezone     || "US/Eastern",
      mgmtVlan:     net.mgmtVlan    || "VLAN 10 \u2014 192.168.10.0/24",
      mgmtGw:       net.mgmtGw      || "192.168.10.1",
      vmotionVlan:  net.vmotionVlan  || "VLAN 20 \u2014 192.168.20.0/24",
      vmotionGw:    net.vmotionGw    || "192.168.20.1",
      vsanVlan:     net.vsanVlan     || "VLAN 30 \u2014 192.168.30.0/24",
      vsanGw:       net.vsanGw       || "192.168.30.1",
      overlayVlan:  net.overlayVlan  || "VLAN 50 \u2014 192.168.50.0/24",
      overlayGw:    net.overlayGw    || "192.168.50.1",
      uplinkVlan:   net.uplinkVlan   || "VLAN 40 \u2014 192.168.40.0/24",
      uplinkGw:     net.uplinkGw     || "192.168.40.1",
      workloadVlan: net.workloadVlan || "VLAN 100-199 (workload segments)",
      defaultGw:    net.defaultGw    || net.mgmtGw || "192.168.10.1",
      mtu: netMtu, torSwitch: netTor, dvSwitch: net.dvSwitch || "vds-mgmt-01 (v8.0.3), vds-edge-01 (v8.0.3)",
      uplinks: netUplinks, lacp: net.lacp || "LACP (active-active, hash: src-dst IP+port)", lldp: net.lldp,
    },

    aria: {
      lcm:        { hostname: arLcmHost,  ip: arLcmIp,  version: (ar.lcm        || {}).version  || "8.16" },
      logInsight: { hostname: arLiHost,   ip: arLiIp,   version: (ar.logInsight || {}).version  || "8.16",
                    cluster:  (ar.logInsight || {}).cluster   || "3-node cluster",
                    retention:(ar.logInsight || {}).retention || "30 days" },
      operations: { hostname: arOpsHost,  ip: arOpsIp,  version: (ar.operations || {}).version  || "8.16",
                    adapters: (ar.operations || {}).adapters  || "vCenter, NSX, vSAN" },
      automation: { hostname: arAutoHost, ip: arAutoIp, version: (ar.automation || {}).version  || "8.16" },
    },

    security: {
      tls:          sec.tls         || "TLS 1.2 / 1.3 enforced",
      ca:           sec.ca          || "Internal CA \u2014 PKI integrated",
      certFrom:     sec.certFrom, certUntil: sec.certUntil,
      mfa:          sec.mfa         || "SAML 2.0 via Workspace ONE Access",
      passwords:    sec.passwords   || "90-day rotation, 20-char minimum",
      lockdown:     sec.lockdown    || "Normal Lockdown Mode on all ESXi hosts",
      syslog:       sec.syslog      || "Aria Log Insight (centralized)",
      compliance:   sec.compliance  || "VMware Security Hardening Guide v9.0",
      backupEncr:   sec.backupEncr  || "Enabled (AES-256)",
      backupEncStrength: sec.backupEncStrength || "Medium (AES-256)",
      backupAuth:   sec.backupAuth,
      ssoDomain:    sec.ssoDomain   || "vsphere.local",
      idp:          sec.idp         || "Workspace ONE Access (SAML 2.0)",
      adRealm:      sec.adRealm     || "Configured",
      cifsAuth:     sec.cifsAuth    || "Active Directory",
      idfw:         sec.idfw        || "Enabled (AD groups)",
      ssh:          sec.ssh         || "Disabled on ESXi (Lockdown Mode)",
    },

    licensing: {
      vcf:     lic.vcf     || "VMware Cloud Foundation 9 \u2014 Universal License",
      vCenter: lic.vCenter || "Included in VCF 9 Universal",
      nsx:     lic.nsx     || "Included in VCF 9 Universal",
      vsan:    lic.vsan    || "Included in VCF 9 Universal",
      aria:    lic.aria    || "Included in VCF 9 Universal (Aria Suite Enterprise)",
      cores:   lic.cores   || "384 licensed vCPU cores",
      schema:  lic.schema  || "ELMS",
    },

    smtp: {
      server: smtpServer, adminEmail: smtpAdmin,
      autosupport: smtpAuto, alertEmail: smtpAlert,
      sendTime: sm.sendTime || "06:00 daily",
    },

    contacts:    FC.contacts,
    issues:      FC.issues,
    datastores:  FC.datastores,
    storageUnits:FC.storageUnits,
  };

  // ── Try to load logo  (place logo.jpg next to this script) ───────────────
  let logoBuffer = null;
  for (const lp of ["logo.jpg","logo.jpeg","logo.png","redesign-logo.jpg"].map(f => path.join(__dirname, f))) {
    if (fs.existsSync(lp)) { logoBuffer = fs.readFileSync(lp); console.log(`✔  Logo: ${lp}`); break; }
  }
  if (!logoBuffer) console.log("ℹ  No logo file found \u2014 place logo.jpg next to this script for header/cover branding.");

  // ── Build & write ─────────────────────────────────────────────────────────
  console.log("\n" + "─".repeat(64));
  console.log(`  Customer  : ${D.customer}`);
  console.log(`  VCF Ver   : ${D.vcfVersion}`);
  console.log(`  Site ID   : ${D.siteId}`);
  console.log(`  Theme     : ${themeName}`);
  console.log(`  Output    : ${OUT_PATH}`);
  console.log("─".repeat(64));

  try {
    const doc    = buildDocument(C, D, logoBuffer);
    const buffer = await Packer.toBuffer(doc);
    const outDir = path.dirname(OUT_PATH);
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
    fs.writeFileSync(OUT_PATH, buffer);
    console.log(`\n\u2714  Document written: ${OUT_PATH}`);
    console.log(`   Size: ${(buffer.length / 1024).toFixed(1)} KB\n`);
  } catch (err) {
    console.error("\n\u2716  Generation failed:", err.message);
    console.error(err.stack);
    process.exit(1);
  }
})();
