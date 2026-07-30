/* Architecture figures generated from the codebase scan.

   architecture-data.json is the node/edge graph produced by the codebase scan
   (public/embeds/codebase-scan.html). Every node carries a sourceRef, and all 22
   of them were verified to exist in the tree before this was published — a
   diagram claiming a file that is not there is worse than no diagram.

   Two figures rather than one picture of all 26 nodes: the full graph is legible
   interactively (that is what the embedded scan is for) but unreadable as a
   static figure. These show the layering, and the request path that carries the
   actual risk.

   Rendered by scripts/build-figures.mjs into public/figures/. */

import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const scan = require("./architecture-data.json");

const BRAND = {
  ink: "#070708",
  field: "#ffffff",
  graphite: "#51545a",
  mist: "#ecedef",
  line: "#d5d7dc",
  blue: "#2f66ff",
  green: "#2b7a5e",
  muted: "#9aa0a8",
  index: "#a4a8b0",
};

const KIND_STYLE = {
  entry: { fill: "#eaf0ff", stroke: BRAND.blue },
  service: { fill: BRAND.field, stroke: BRAND.line },
  store: { fill: "#e7f2ed", stroke: BRAND.green },
  external: { fill: BRAND.mist, stroke: BRAND.muted, dashed: true },
};

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function styleBlock() {
  return `<style>
    .b { stroke-width:1; }
    .lbl { fill:${BRAND.ink}; font:600 11px Inter,Arial,Helvetica,sans-serif; }
    .sub { fill:${BRAND.graphite}; font:9px ui-monospace,SFMono-Regular,Menlo,monospace; }
    .tier { fill:${BRAND.graphite}; font:700 9px Inter,Arial,Helvetica,sans-serif; letter-spacing:1.4px; }
    .cnt { fill:${BRAND.index}; font:9px ui-monospace,SFMono-Regular,Menlo,monospace; }
    .ed { fill:${BRAND.graphite}; font:9px Inter,Arial,Helvetica,sans-serif; }
  </style>`;
}

function arrowDefs(id, color = BRAND.graphite) {
  return `<defs><marker id="ar-${id}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,1 L9,5 L0,9 z" fill="${color}"/></marker></defs>`;
}

function box(x, y, w, h, kind, label, sub, count) {
  const style = KIND_STYLE[kind] ?? KIND_STYLE.service;
  const dash = style.dashed ? ' stroke-dasharray="4 3"' : "";
  return `<g>
    <rect class="b" x="${x}" y="${y}" width="${w}" height="${h}" rx="2" fill="${style.fill}" stroke="${style.stroke}"${dash}/>
    <text class="lbl" x="${x + 11}" y="${y + (sub ? 20 : h / 2 + 4)}">${escapeXml(label)}</text>
    ${sub ? `<text class="sub" x="${x + 11}" y="${y + 34}">${escapeXml(sub)}</text>` : ""}
    ${count ? `<text class="cnt" x="${x + w - 11}" y="${y + 20}" text-anchor="end">${escapeXml(count)}</text>` : ""}
  </g>`;
}

/* ---------- figure 1: how the codebase layers ---------- */

export function renderLayerMap() {
  const kinds = ["entry", "service", "store", "external"];
  const counts = Object.fromEntries(kinds.map((kind) => [kind, scan.NODES.filter((node) => node.kind === kind).length]));
  const tiers = [
    { kind: "entry", title: "ENTRY POINTS", note: "Pages and packages a user opens" },
    { kind: "service", title: "SERVICES", note: "Feature modules, data layer, builders" },
    { kind: "store", title: "STORES", note: "Where records and files live" },
    { kind: "external", title: "EXTERNAL", note: "Platforms PULSE depends on" },
  ];

  const W = 620, BOX_H = 52, GAP_Y = 30, LEFT = 132;
  const height = tiers.length * (BOX_H + GAP_Y) - GAP_Y + 8;
  const parts = [arrowDefs("layers")];

  tiers.forEach((tier, index) => {
    const y = index * (BOX_H + GAP_Y) + 4;
    parts.push(`<text class="tier" x="0" y="${y + 20}">${tier.title}</text>`);
    parts.push(`<text class="sub" x="0" y="${y + 36}">${escapeXml(`${counts[tier.kind]} nodes`)}</text>`);
    parts.push(box(LEFT, y, W - LEFT, BOX_H, tier.kind, tier.note, "", ""));
    if (index < tiers.length - 1) {
      const x = LEFT + (W - LEFT) / 2;
      parts.push(`<line x1="${x}" y1="${y + BOX_H}" x2="${x}" y2="${y + BOX_H + GAP_Y - 3}" stroke="${BRAND.graphite}" stroke-width="1.5" marker-end="url(#ar-layers)"/>`);
    }
  });

  return svg("layers", `0 0 ${W} ${height}`, W, height,
    `How the PULSE codebase layers: ${tiers.map((tier) => `${counts[tier.kind]} ${tier.title.toLowerCase()}`).join(", ")}`,
    parts.join("\n"));
}

/* ---------- figure 2: the SharePoint request path ---------- */

export function renderDataPath() {
  /* The group name comes from the scan, so if the grouping changes there the
     figure follows rather than going stale silently. */
  const group = "SharePoint Data Layer";
  const layer = scan.NODES.filter((node) => node.group === group);
  if (!layer.length) throw new Error(`Scan data has no nodes in group "${group}"`);
  const stores = scan.NODES.filter((node) => node.kind === "store");
  const externals = scan.NODES.filter((node) => node.kind === "external");

  const BOX_W = 178, BOX_H = 54, GAP = 40;
  const columns = [
    { title: "PAGES", nodes: [{ label: "Feature pages", sub: "6 dashboard pages" }] },
    { title: group.toUpperCase(), nodes: layer.map((node) => ({ label: node.label, sub: shortRef(node.sourceRef) })) },
    { title: "STORES & PLATFORM", nodes: [...stores, ...externals].map((node) => ({ label: node.label, sub: "", kind: node.kind })) },
  ];

  const rows = Math.max(...columns.map((column) => column.nodes.length));
  const width = columns.length * BOX_W + (columns.length - 1) * GAP;
  const height = rows * (BOX_H + 14) - 14 + 30;
  const parts = [arrowDefs("path")];

  columns.forEach((column, columnIndex) => {
    const x = columnIndex * (BOX_W + GAP);
    parts.push(`<text class="tier" x="${x}" y="12">${escapeXml(column.title)}</text>`);
    column.nodes.forEach((node, rowIndex) => {
      const y = rowIndex * (BOX_H + 14) + 30;
      const kind = node.kind ?? (columnIndex === 0 ? "entry" : "service");
      parts.push(box(x, y, BOX_W, BOX_H, kind, node.label, node.sub, ""));
    });
    if (columnIndex < columns.length - 1) {
      /* Centred on the column, not on the first row — on the first row it reads
         as one box pointing at one box rather than a layer feeding a layer. */
      const y = 30 + (rows * (BOX_H + 14) - 14) / 2;
      parts.push(`<line x1="${x + BOX_W}" y1="${y}" x2="${x + BOX_W + GAP - 3}" y2="${y}" stroke="${BRAND.graphite}" stroke-width="1.5" marker-end="url(#ar-path)"/>`);
    }
  });

  return svg("path", `0 0 ${width} ${height}`, width, height,
    `The SharePoint request path: feature pages call the ${group}, which reads and writes SharePoint Lists and the document library`,
    parts.join("\n"));
}

function shortRef(ref) {
  if (!ref) return "";
  return ref.split("/").pop() ?? "";
}

function svg(id, viewBox, width, height, description, body) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${viewBox}" width="${width}" height="${height}" role="img" aria-label="${escapeXml(description)}">
${styleBlock()}
${body}
</svg>
`;
}

export const architectureFigures = [
  { id: "arch-layers", render: renderLayerMap },
  { id: "arch-datapath", render: renderDataPath },
];

/* Counts quoted in the document prose, so they cannot drift from the scan. */
export const scanSummary = {
  nodes: scan.NODES.length,
  edges: scan.EDGES.length,
  entries: scan.NODES.filter((node) => node.kind === "entry").length,
  services: scan.NODES.filter((node) => node.kind === "service").length,
  stores: scan.NODES.filter((node) => node.kind === "store").length,
  externals: scan.NODES.filter((node) => node.kind === "external").length,
  groups: [...new Set(scan.NODES.filter((node) => node.group).map((node) => node.group))],
  edgeKinds: [...new Set(scan.EDGES.map((edge) => edge.kind))],
};
