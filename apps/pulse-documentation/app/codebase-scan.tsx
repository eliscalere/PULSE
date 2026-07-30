"use client";

import { useEffect, useRef } from "react";

/* This used to be a standalone HTML file (public/embeds/codebase-scan.html)
   rendered through a sandboxed iframe, so the package carried the same map
   twice: once as this app's code, once as a second document referenced by
   URL. Inlining it here means the codebase section is just part of the page
   like everything else — one component, no second file, no iframe boundary.
   The interaction model below is a direct port of that file's script, scoped
   to this component's own container instead of `document`/`window`. */

type Kind = "entry" | "service" | "store" | "external";

type Node = {
  id: string;
  label: string;
  kind: Kind;
  sub?: string;
  sourceRef?: string;
  detail?: string;
  x: number;
  y: number;
  group?: string;
};

type Edge = { from: string; to: string; kind: string; label?: string };

const CARD_W = 190;

const NODES: Node[] = [
  { id: "pulse-code", label: "PULSE CODE", kind: "entry", sub: "Web IDE · index.html", sourceRef: "apps/PULSE-CODE/index.html", detail: "Monaco-based in-browser IDE that reads/writes files in a SharePoint document library over the user's existing session.", x: 40, y: 40 },
  { id: "svc-code-sp-client", label: "SP file client", kind: "service", sub: "read/write files in IDE", sourceRef: "apps/PULSE-CODE/index.html", x: 270, y: 40 },

  { id: "svc-build-wfc", label: "WFC package builder", kind: "service", sub: "→ releases/*.html", sourceRef: "apps/PULSE/scripts/build-sharepoint-package.js", detail: "Inlines every local script/stylesheet/asset into one self-contained HTML file with a WFC-MANIFEST header — the file uploaded to SharePoint.", x: 40, y: 450, group: "Build & Packaging" },
  { id: "svc-build-forge", label: "Forge web part builder", kind: "service", sub: "srcdoc iframe bundle", sourceRef: "apps/PULSE/scripts/build-forge.js", detail: "Base64-encodes the packaged app and injects it into a Forge platform template's srcdoc iframe for embedding on a SharePoint page.", x: 40, y: 534, group: "Build & Packaging" },
  { id: "svc-build-docs", label: "Docs site builder", kind: "service", sub: "esbuild + inline dist/", sourceRef: "apps/pulse-documentation/scripts/build-sharepoint-package.mjs", x: 40, y: 618, group: "Build & Packaging" },

  { id: "pulse-docs-site", label: "PULSE Documentation", kind: "entry", sub: "Next.js handoff site", sourceRef: "apps/pulse-documentation/app/page.tsx", detail: "Interactive technical handoff / SOP / user-guide site; also packaged as a WFC HTML upload for SharePoint.", x: 500, y: 124 },

  { id: "pulse-dashboard", label: "PULSE Dashboard", kind: "entry", sub: "SPA · index.html", sourceRef: "apps/PULSE/index.html", detail: "Main IPT dashboard: single-file SPA with a hash router, deployed as a self-contained HTML upload or a Forge web part iframe.", x: 730, y: 292 },
  { id: "svc-export", label: "PPTX export", kind: "service", sub: "overview + project decks", sourceRef: "apps/PULSE/assets/js/export.js", x: 500, y: 450 },
  { id: "svc-issue-reporting", label: "Issue reporting", kind: "service", sub: "in-app bug capture", sourceRef: "apps/PULSE/assets/js/issue-reporting.js", x: 500, y: 534 },
  { id: "store-sp-doclib", label: "SP Document Library", kind: "store", sub: "SiteAssets + files", x: 500, y: 230 },

  { id: "page-weekly", label: "Weekly Meetings", kind: "entry", sub: "/weekly", sourceRef: "apps/PULSE/assets/js/pages/weekly.js", x: 960, y: 208, group: "PULSE Dashboard Pages" },
  { id: "page-travel", label: "Travel & Debriefs", kind: "entry", sub: "/travel", sourceRef: "apps/PULSE/assets/js/pages/travel.js", x: 960, y: 292, group: "PULSE Dashboard Pages" },
  { id: "page-projects", label: "Projects & Tracker", kind: "entry", sub: "/projects", sourceRef: "apps/PULSE/assets/js/pages/projects.js", x: 960, y: 376, group: "PULSE Dashboard Pages" },
  { id: "page-docreview", label: "Document Review", kind: "entry", sub: "/docreview", sourceRef: "apps/PULSE/assets/js/pages/docreview.js", x: 960, y: 460, group: "PULSE Dashboard Pages" },
  { id: "page-admin", label: "Admin & Roles", kind: "entry", sub: "/admin", sourceRef: "apps/PULSE/assets/js/pages/admin.js", x: 960, y: 544, group: "PULSE Dashboard Pages" },
  { id: "page-overview", label: "Overview & Reporting", kind: "entry", sub: "/overview", sourceRef: "apps/PULSE/assets/js/pages/overview.js", x: 960, y: 628, group: "PULSE Dashboard Pages" },

  { id: "svc-sp-repo", label: "SharePoint repo", kind: "service", sub: "business read/write layer", sourceRef: "apps/PULSE/assets/js/sharepoint-repo.js", detail: "Business-logic read/write layer over 10 SP lists: projects, travel requests/debriefs, meetings, documents, issues, roles, audit log, notification and location config.", x: 1230, y: 376, group: "SharePoint Data Layer" },
  { id: "svc-sp-adapter", label: "SharePoint adapter", kind: "service", sub: "REST + digest calls", sourceRef: "apps/PULSE/assets/js/sharepoint-adapter.js", detail: "Low-level SP REST calls: fetches a contextinfo digest for writes, uses odata=nometadata for lean reads, and calls Microsoft Graph for people/group search.", x: 1230, y: 460, group: "SharePoint Data Layer" },
  { id: "svc-sp-schema", label: "SharePoint schema", kind: "service", sub: "list defs + seeding", sourceRef: "apps/PULSE/assets/js/sharepoint-schema.js", x: 1460, y: 376, group: "SharePoint Data Layer" },
  { id: "svc-audit-log", label: "Audit log", kind: "service", sub: "rotates every 1000 items", sourceRef: "apps/PULSE/assets/js/audit-log.js", x: 1460, y: 460, group: "SharePoint Data Layer" },

  { id: "svc-notify", label: "Notifications", kind: "service", sub: "email via SharePoint", sourceRef: "apps/PULSE/assets/js/notify.js", x: 1460, y: 544 },
  { id: "ext-sharepoint", label: "SharePoint Online", kind: "external", sub: "Flank Speed tenant", x: 1690, y: 544 },
  { id: "ext-msgraph", label: "Microsoft Graph", kind: "external", sub: "dod-graph.microsoft.us", x: 1690, y: 460 },
  { id: "store-sp-lists", label: "SharePoint Lists", kind: "store", sub: "10 PULSE lists", detail: "Projects, Travel Requests, Travel Debriefs, Meetings, Documents, Issues, App Roles, Audit Log, Notification Config, Location Config.", x: 1690, y: 292 },

  { id: "pulse-calendar", label: "PULSE Calendar", kind: "entry", sub: "Travel calendar", sourceRef: "apps/PULSE-TRAVEL-CALENDAR/index.html", x: 1690, y: 124 },
  { id: "svc-calendar", label: "Calendar logic", kind: "service", sub: "FullCalendar travel view", sourceRef: "apps/PULSE-TRAVEL-CALENDAR/assets/js/calendar.js", x: 1690, y: 208 },
];

const EDGES: Edge[] = [
  { from: "pulse-dashboard", to: "page-projects", kind: "triggers" },
  { from: "pulse-dashboard", to: "page-weekly", kind: "triggers" },
  { from: "pulse-dashboard", to: "page-travel", kind: "triggers" },
  { from: "pulse-dashboard", to: "page-docreview", kind: "triggers" },
  { from: "pulse-dashboard", to: "page-admin", kind: "triggers" },
  { from: "pulse-dashboard", to: "page-overview", kind: "triggers" },

  { from: "page-projects", to: "svc-sp-repo", kind: "calls" },
  { from: "page-weekly", to: "svc-sp-repo", kind: "calls" },
  { from: "page-travel", to: "svc-sp-repo", kind: "calls" },
  { from: "page-docreview", to: "svc-sp-repo", kind: "calls" },
  { from: "page-admin", to: "svc-sp-repo", kind: "calls" },
  { from: "page-overview", to: "svc-sp-repo", kind: "calls" },

  { from: "page-overview", to: "svc-export", kind: "calls", label: "overview deck" },
  { from: "page-projects", to: "svc-export", kind: "calls", label: "project deck" },
  { from: "page-admin", to: "svc-sp-schema", kind: "calls", label: "seeds lists" },
  { from: "svc-sp-repo", to: "svc-audit-log", kind: "calls", label: "logs every write" },

  { from: "pulse-dashboard", to: "svc-issue-reporting", kind: "triggers", label: "bug report widget" },
  { from: "svc-issue-reporting", to: "store-sp-lists", kind: "writes", label: "PULSE Issues" },
  { from: "page-docreview", to: "store-sp-lists", kind: "writes", label: "AiSuggestionsJson" },

  { from: "svc-notify", to: "ext-sharepoint", kind: "calls", label: "SP SendEmail API" },

  { from: "svc-sp-repo", to: "svc-sp-adapter", kind: "calls" },
  { from: "svc-sp-schema", to: "svc-sp-adapter", kind: "calls" },
  { from: "svc-audit-log", to: "svc-sp-adapter", kind: "calls" },
  { from: "svc-sp-adapter", to: "store-sp-lists", kind: "reads" },
  { from: "svc-sp-adapter", to: "store-sp-lists", kind: "writes", label: "contextinfo digest" },
  { from: "svc-sp-adapter", to: "ext-msgraph", kind: "calls", label: "people/group search" },
  { from: "svc-sp-adapter", to: "ext-sharepoint", kind: "calls", label: "session auth" },

  { from: "svc-build-wfc", to: "pulse-dashboard", kind: "reads", label: "inlines assets" },
  { from: "svc-build-forge", to: "pulse-dashboard", kind: "reads", label: "base64s into Forge" },
  { from: "svc-build-docs", to: "pulse-docs-site", kind: "reads", label: "inline dist/" },

  { from: "pulse-code", to: "svc-code-sp-client", kind: "triggers" },
  { from: "svc-code-sp-client", to: "store-sp-doclib", kind: "reads", label: "list/open files" },
  { from: "svc-code-sp-client", to: "store-sp-doclib", kind: "writes", label: "save edits" },

  { from: "pulse-calendar", to: "svc-calendar", kind: "triggers" },
  { from: "svc-calendar", to: "store-sp-lists", kind: "reads", label: "Travel lists" },
];

const TOP_INTEGRATIONS = [
  { id: "ext-sharepoint", label: "SharePoint Online", initial: "S", color: "blue" },
  { id: "ext-msgraph", label: "Microsoft Graph", initial: "M", color: "blue" },
];

const FILTERS: { key: Kind; label: string }[] = [
  { key: "entry", label: "Entries" },
  { key: "service", label: "Services" },
  { key: "store", label: "Stores" },
  { key: "external", label: "External" },
];

const KIND_LABEL: Record<Kind, string> = { entry: "Entry", service: "Service", store: "Store", external: "External" };

const SCAN_CSS = `
.codebase-scan-embed{
  --cs-bg:#ECEDEF; --cs-surface:#FFFFFF; --cs-surface-2:#F7F7F8; --cs-surface-3:#E4E5E8;
  --cs-text:#070708; --cs-muted:#51545A; --cs-border:#D9DADD; --cs-border-strong:#B9BBC0;
  --cs-blue:#2F66FF; --cs-blue-tint:#E8EEFF;
  --cs-green:#2B7A5E; --cs-green-tint:#E7F2ED;
  --cs-amber:#C77800; --cs-amber-tint:#FFF2DF;
  --cs-red:#B42318; --cs-red-tint:#FDECEA;
  --cs-violet:#6E3FD1; --cs-violet-tint:#F1EBFF;
  --cs-group-fill: rgba(7,7,8,.035);
  --cs-radius-lg:12px; --cs-radius-md:8px; --cs-radius-sm:4px;
  --cs-shadow-1:0 1px 2px rgba(7,7,8,.08); --cs-shadow-2:0 3px 8px rgba(7,7,8,.12); --cs-shadow-3:0 10px 30px rgba(7,7,8,.20);
  --cs-overlay:rgba(7,7,8,.58);
  color-scheme: light;
  position:relative; width:100%; height:100%; overflow:hidden;
  background:var(--cs-bg); color:var(--cs-text);
  font-family:"Inter","Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif;
  font-size:14px; line-height:1.5;
  -webkit-font-smoothing:antialiased;
}
.codebase-scan-embed[data-theme="dark"]{
  --cs-bg:#0E1116; --cs-surface:#14171D; --cs-surface-2:#1A1E25; --cs-surface-3:#22262E;
  --cs-text:#DCE1E8; --cs-muted:#8891A0; --cs-border:#262B33; --cs-border-strong:#363C46;
  --cs-blue:#6E9CFF; --cs-blue-tint:#1A2436; --cs-green:#4FBE8E; --cs-green-tint:#16261F;
  --cs-amber:#D6A24A; --cs-amber-tint:#2E2617; --cs-red:#E8776A; --cs-red-tint:#2E1D1B;
  --cs-violet:#B48CFF; --cs-violet-tint:#241B3D; --cs-group-fill: rgba(255,255,255,.035);
  --cs-shadow-1:0 1px 2px rgba(0,0,0,.5); --cs-shadow-2:0 8px 20px rgba(0,0,0,.55); --cs-shadow-3:0 16px 44px rgba(0,0,0,.6);
  --cs-overlay:rgba(0,0,0,.65);
  color-scheme: dark;
}
.codebase-scan-embed *{box-sizing:border-box;}
.codebase-scan-embed .mono{font-family:ui-monospace,SFMono-Regular,"SF Mono",Menlo,Consolas,monospace;}
.codebase-scan-embed ::selection{background:var(--cs-blue);color:#fff;}
.codebase-scan-embed button{font:inherit;color:inherit;}

.codebase-scan-embed .cs-viewport{position:absolute;inset:0;overflow:hidden;background:var(--cs-bg);cursor:grab;touch-action:none;}
.codebase-scan-embed .cs-viewport.dragging{cursor:grabbing;}
.codebase-scan-embed .cs-world{position:absolute;top:0;left:0;transform-origin:0 0;will-change:transform;}
.codebase-scan-embed svg.cs-edges{position:absolute;top:0;left:0;pointer-events:none;overflow:visible;}
.codebase-scan-embed .cs-edge-path{fill:none;stroke:var(--cs-border-strong);stroke-width:1.4;opacity:.42;}
.codebase-scan-embed .cs-edge-path.is-active{stroke-width:2;opacity:.95;}
.codebase-scan-embed .cs-edge-path.is-dim{opacity:.07;}
.codebase-scan-embed .cs-edge-label{font-size:10.5px;font-weight:700;fill:var(--cs-muted);paint-order:stroke;stroke:var(--cs-bg);stroke-width:4px;stroke-linejoin:round;}

.codebase-scan-embed .cs-group-box{position:absolute;border:1px dashed var(--cs-border-strong);border-radius:16px;background:var(--cs-group-fill);}
.codebase-scan-embed .cs-group-label{
  position:absolute;top:-10px;left:16px;background:var(--cs-bg);padding:0 8px;
  font-size:10px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:var(--cs-muted);
}

.codebase-scan-embed .cs-node{
  position:absolute;
  background:var(--cs-surface);
  border:1px solid var(--cs-border);
  border-left:3px solid var(--cs-border-strong);
  border-radius:var(--cs-radius-md);
  padding:8px 10px 9px;
  cursor:pointer;
  text-align:left;width:190px;
  box-shadow:var(--cs-shadow-1);
  user-select:none;
}
.codebase-scan-embed .cs-node:hover{border-color:var(--cs-border-strong);box-shadow:var(--cs-shadow-2);}
.codebase-scan-embed .cs-node:focus-visible{outline:2px solid var(--cs-blue);outline-offset:2px;}
.codebase-scan-embed .cs-node .kind-tag{display:block;font-size:9.5px;font-weight:800;letter-spacing:.08em;color:var(--cs-muted);text-transform:uppercase;margin-bottom:3px;}
.codebase-scan-embed .cs-node .label{display:block;font-size:12.5px;font-weight:700;letter-spacing:-.005em;line-height:1.25;}
.codebase-scan-embed .cs-node .sub{display:block;margin-top:2px;font-size:11px;color:var(--cs-muted);line-height:1.3;}

.codebase-scan-embed .cs-node[data-kind="entry"]{border-left-color:var(--cs-blue);}
.codebase-scan-embed .cs-node[data-kind="entry"] .kind-tag{color:var(--cs-blue);}
.codebase-scan-embed .cs-node[data-kind="external"]{border-left-style:dashed;border-left-color:var(--cs-border-strong);}
.codebase-scan-embed .cs-node[data-kind="store"]{border-left-color:var(--cs-blue);}

.codebase-scan-embed .cs-node.is-dim{opacity:.32;}
.codebase-scan-embed .cs-node.is-active{box-shadow:var(--cs-shadow-3);}
.codebase-scan-embed .cs-node.is-active[data-kind="entry"],.codebase-scan-embed .cs-node.is-active[data-kind="store"]{border-left-color:var(--cs-blue);background:var(--cs-blue-tint);}
.codebase-scan-embed .cs-node.is-active[data-kind="service"]{border-left-color:var(--cs-border-strong);background:var(--cs-surface-3);}
.codebase-scan-embed .cs-node.is-active[data-kind="external"]{background:var(--cs-surface-3);}
.codebase-scan-embed .cs-node.is-filtered-out{opacity:.14;}

.codebase-scan-embed .cs-brand-badge{
  position:absolute;top:16px;left:16px;z-index:20;
  display:flex;align-items:center;gap:8px;
  background:var(--cs-surface);border:1px solid var(--cs-border);border-radius:999px;
  padding:7px 14px 7px 8px;box-shadow:var(--cs-shadow-2);
}
.codebase-scan-embed .cs-brand-badge .bt{font-size:12px;color:var(--cs-muted);}
.codebase-scan-embed .cs-brand-badge .bt b{color:var(--cs-text);font-weight:700;}

.codebase-scan-embed .cs-top-right{position:absolute;top:16px;right:16px;z-index:20;display:flex;align-items:center;gap:8px;}
.codebase-scan-embed .cs-pill-btn{
  display:inline-flex;align-items:center;gap:7px;height:34px;padding:0 13px;
  background:var(--cs-surface);border:1px solid var(--cs-border);border-radius:999px;
  font-size:12px;font-weight:600;color:var(--cs-muted);box-shadow:var(--cs-shadow-1);cursor:default;
}
.codebase-scan-embed .cs-icon-btn{
  display:flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:999px;
  background:var(--cs-surface);border:1px solid var(--cs-border);box-shadow:var(--cs-shadow-1);
  cursor:pointer;font-size:14px;
}
.codebase-scan-embed .cs-icon-btn:hover{background:var(--cs-surface-2);}

.codebase-scan-embed .cs-side-panel{
  position:absolute;top:64px;left:16px;z-index:20;width:220px;max-height:calc(100% - 100px);overflow:auto;
  background:var(--cs-surface);border:1px solid var(--cs-border);border-radius:var(--cs-radius-lg);
  box-shadow:var(--cs-shadow-2);padding:16px;
}
.codebase-scan-embed .cs-side-panel h4{
  margin:0 0 9px;font-size:10.5px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:var(--cs-muted);
  display:flex;align-items:center;gap:6px;
}
.codebase-scan-embed .cs-side-row{
  display:flex;align-items:center;gap:9px;padding:7px 8px;border-radius:var(--cs-radius-sm);
  cursor:pointer;font-size:12.5px;font-weight:600;
}
.codebase-scan-embed .cs-side-row:hover{background:var(--cs-surface-2);}
.codebase-scan-embed .cs-side-row .avatar{
  width:22px;height:22px;border-radius:50%;display:flex;align-items:center;justify-content:center;
  font-size:10px;font-weight:800;color:#fff;flex:none;
}

.codebase-scan-embed .cs-filter-bar{
  position:absolute;left:50%;bottom:18px;transform:translateX(-50%);z-index:20;
  display:flex;align-items:center;gap:4px;
  background:var(--cs-surface);border:1px solid var(--cs-border);border-radius:999px;
  padding:5px;box-shadow:var(--cs-shadow-2);
}
.codebase-scan-embed .cs-filter-chip{
  display:inline-flex;align-items:center;gap:6px;height:32px;padding:0 13px;border-radius:999px;
  border:1px solid transparent;background:none;cursor:pointer;
  font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:var(--cs-muted);
}
.codebase-scan-embed .cs-filter-chip .dot{width:7px;height:7px;border-radius:50%;background:var(--cs-border-strong);}
.codebase-scan-embed .cs-filter-chip.on{background:var(--cs-surface-2);color:var(--cs-text);border-color:var(--cs-border);}
.codebase-scan-embed .cs-filter-chip:hover{color:var(--cs-text);}

.codebase-scan-embed .cs-zoom-ctl{position:absolute;right:16px;bottom:18px;z-index:20;display:flex;flex-direction:column;gap:6px;}

.codebase-scan-embed .cs-modal-backdrop{position:absolute;inset:0;background:var(--cs-overlay);display:none;align-items:center;justify-content:center;padding:20px;z-index:50;}
.codebase-scan-embed .cs-modal-backdrop.open{display:flex;}
.codebase-scan-embed .cs-modal{background:var(--cs-surface);border:1px solid var(--cs-border);border-radius:var(--cs-radius-lg);box-shadow:var(--cs-shadow-3);max-width:460px;width:100%;max-height:80%;overflow:auto;padding:20px 22px;}
.codebase-scan-embed .cs-modal-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;}
.codebase-scan-embed .cs-modal-head .kind-tag{display:inline-block;font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;padding:3px 8px;border-radius:999px;margin-bottom:8px;}
.codebase-scan-embed .cs-modal-head h3{margin:0;font-size:17px;font-weight:750;letter-spacing:-.01em;}
.codebase-scan-embed .cs-modal-head .sub{margin-top:3px;font-size:12.5px;color:var(--cs-muted);}
.codebase-scan-embed .cs-modal-close{flex:none;width:28px;height:28px;border-radius:7px;border:1px solid var(--cs-border);background:var(--cs-surface-2);cursor:pointer;font-size:14px;line-height:1;}
.codebase-scan-embed .cs-modal-close:hover{background:var(--cs-surface-3);}
.codebase-scan-embed .cs-modal p.detail{margin:14px 0 0;font-size:13px;line-height:1.6;}
.codebase-scan-embed .cs-modal .source-ref{margin-top:12px;display:inline-block;font-size:11.5px;padding:5px 9px;background:var(--cs-surface-2);border:1px solid var(--cs-border);border-radius:var(--cs-radius-sm);color:var(--cs-muted);}
.codebase-scan-embed .cs-modal .conn-title{margin:18px 0 8px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--cs-muted);}
.codebase-scan-embed .cs-conn-list{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:6px;}
.codebase-scan-embed .cs-conn-list li{display:flex;align-items:center;gap:8px;font-size:12.5px;padding:7px 9px;background:var(--cs-surface-2);border:1px solid var(--cs-border);border-radius:var(--cs-radius-sm);}
.codebase-scan-embed .cs-conn-list .arrow{color:var(--cs-muted);flex:none;}
.codebase-scan-embed .cs-conn-list .kind-pill{margin-left:auto;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.03em;color:var(--cs-muted);}

@media (prefers-reduced-motion: no-preference){
  .codebase-scan-embed .cs-node,.codebase-scan-embed .cs-edge-path,.codebase-scan-embed .cs-modal-backdrop,.codebase-scan-embed .cs-icon-btn,.codebase-scan-embed .cs-filter-chip{transition:opacity 140ms ease,box-shadow 140ms ease,background 140ms ease,border-color 140ms ease;}
}
@media (max-width:760px){
  .codebase-scan-embed .cs-side-panel{display:none;}
}
`;

function avatarColor(name: string) {
  if (name === "blue") return "var(--cs-blue)";
  return "var(--cs-border-strong)";
}

/* A direct port of the standalone scan page's script, scoped to this
   component's own container ref instead of `document`/`window` so it can
   run alongside the rest of the app without a sandboxed frame. */
function mountScan(root: HTMLDivElement) {
  const svgNS = "http://www.w3.org/2000/svg";
  const nodeById: Record<string, Node> = {};
  NODES.forEach((n) => { nodeById[n.id] = n; });
  const incoming: Record<string, Edge[]> = {};
  const outgoing: Record<string, Edge[]> = {};
  NODES.forEach((n) => { incoming[n.id] = []; outgoing[n.id] = []; });
  EDGES.forEach((e) => { outgoing[e.from].push(e); incoming[e.to].push(e); });

  const groups: Record<string, string[]> = {};
  NODES.forEach((n) => { if (n.group) (groups[n.group] = groups[n.group] || []).push(n.id); });

  const viewport = root.querySelector<HTMLDivElement>(".cs-viewport")!;
  const world = root.querySelector<HTMLDivElement>(".cs-world")!;
  const svg = root.querySelector<SVGSVGElement>(".cs-edges")!;

  const WORLD_W = 1980, WORLD_H = 760;
  world.style.width = `${WORLD_W}px`;
  world.style.height = `${WORLD_H}px`;
  svg.setAttribute("width", String(WORLD_W));
  svg.setAttribute("height", String(WORLD_H));

  const defs = document.createElementNS(svgNS, "defs");
  ([["muted", "var(--cs-border-strong)"], ["blue", "var(--cs-blue)"]] as const).forEach(([key, color]) => {
    const marker = document.createElementNS(svgNS, "marker");
    marker.setAttribute("id", `cs-arrow-${key}`);
    marker.setAttribute("markerWidth", "8"); marker.setAttribute("markerHeight", "8");
    marker.setAttribute("refX", "6"); marker.setAttribute("refY", "3");
    marker.setAttribute("orient", "auto"); marker.setAttribute("markerUnits", "userSpaceOnUse");
    const path = document.createElementNS(svgNS, "path");
    path.setAttribute("d", "M0,0 L6,3 L0,6 Z"); path.setAttribute("fill", color);
    marker.appendChild(path); defs.appendChild(marker);
  });
  svg.appendChild(defs);

  Object.keys(groups).forEach((name) => {
    const ids = groups[name];
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    ids.forEach((id) => {
      const n = nodeById[id];
      minX = Math.min(minX, n.x); minY = Math.min(minY, n.y);
      maxX = Math.max(maxX, n.x + CARD_W); maxY = Math.max(maxY, n.y + 64);
    });
    const pad = 18;
    const box = document.createElement("div");
    box.className = "cs-group-box";
    box.style.left = `${minX - pad}px`; box.style.top = `${minY - pad - 4}px`;
    box.style.width = `${maxX - minX + pad * 2}px`; box.style.height = `${maxY - minY + pad * 2 + 4}px`;
    const label = document.createElement("span");
    label.className = "cs-group-label";
    label.textContent = name;
    box.appendChild(label);
    world.appendChild(box);
  });

  const nodeEls: Record<string, HTMLButtonElement> = {};
  NODES.forEach((n) => {
    const el = document.createElement("button");
    el.type = "button";
    el.className = "cs-node";
    el.dataset.id = n.id; el.dataset.kind = n.kind;
    el.style.left = `${n.x}px`; el.style.top = `${n.y}px`;
    let html = `<span class="kind-tag">${KIND_LABEL[n.kind]}</span><span class="label">${n.label}</span>`;
    if (n.sub) html += `<span class="sub">${n.sub}</span>`;
    el.innerHTML = html;
    el.addEventListener("click", () => { if (dragMoved) return; selectNode(n.id); });
    world.appendChild(el);
    nodeEls[n.id] = el;
  });

  function edgeColorKind(e: Edge) { return "muted"; }

  type LaidOutEdge = { path: SVGPathElement; from: string; to: string; label?: string; midX: number; midY: number; colorKind: string };
  let edgePaths: LaidOutEdge[] = [];
  function layoutEdges() {
    svg.querySelectorAll("path.cs-edge-path, text.cs-edge-label").forEach((n) => n.remove());
    edgePaths = [];
    EDGES.forEach((e) => {
      const s = nodeById[e.from], d = nodeById[e.to];
      if (!s || !d) return;
      const sEl = nodeEls[e.from], dEl = nodeEls[e.to];
      const sH = sEl.offsetHeight, dH = dEl.offsetHeight;
      const sMidY = s.y + sH / 2, dMidY = d.y + dH / 2;
      const sameCol = Math.abs(s.x - d.x) < 4;

      let startX: number, startY: number, endX: number, endY: number;
      if (sameCol) {
        startX = s.x + CARD_W / 2; endX = d.x + CARD_W / 2;
        startY = s.y < d.y ? s.y + sH : s.y;
        endY = s.y < d.y ? d.y : d.y + dH;
      } else if (d.x >= s.x) {
        startX = s.x + CARD_W; startY = sMidY;
        endX = d.x; endY = dMidY;
      } else {
        startX = s.x; startY = sMidY;
        endX = d.x + CARD_W; endY = dMidY;
      }

      const dxAbs = Math.abs(endX - startX) || 1;
      const c = Math.min(90, Math.max(30, dxAbs * 0.4));
      let pathD: string;
      if (sameCol) {
        pathD = `M${startX},${startY} C${startX},${startY + c * 0.6} ${endX},${endY - c * 0.6} ${endX},${endY}`;
      } else {
        const sign = endX >= startX ? 1 : -1;
        pathD = `M${startX},${startY} C${startX + sign * c},${startY} ${endX - sign * c},${endY} ${endX},${endY}`;
      }

      const colorKind = edgeColorKind(e);
      const path = document.createElementNS(svgNS, "path");
      path.setAttribute("d", pathD);
      path.setAttribute("class", "cs-edge-path");
      path.setAttribute("marker-end", `url(#cs-arrow-${colorKind})`);
      svg.appendChild(path);
      edgePaths.push({ path, from: e.from, to: e.to, label: e.label, midX: (startX + endX) / 2, midY: (startY + endY) / 2, colorKind });
    });
  }
  layoutEdges();

  /* ---------- modal ---------- */
  const backdrop = root.querySelector<HTMLDivElement>(".cs-modal-backdrop")!;
  const modalKind = root.querySelector<HTMLSpanElement>(".cs-modal-kind")!;
  const modalTitle = root.querySelector<HTMLHeadingElement>(".cs-modal-title")!;
  const modalSub = root.querySelector<HTMLDivElement>(".cs-modal-sub")!;
  const modalDetail = root.querySelector<HTMLParagraphElement>(".cs-modal-detail")!;
  const modalSrc = root.querySelector<HTMLSpanElement>(".cs-modal-src")!;
  const modalConnWrap = root.querySelector<HTMLDivElement>(".cs-modal-conn-wrap")!;

  function kindTint(kind: Kind) { return kind === "entry" || kind === "store" ? "blue" : "muted"; }
  function kindColorVar(kind: Kind) { return kindTint(kind) === "blue" ? "var(--cs-blue)" : "var(--cs-muted)"; }
  function kindBgVar(kind: Kind) { return kindTint(kind) === "blue" ? "var(--cs-blue-tint)" : "var(--cs-surface-2)"; }

  function openModal(n: Node) {
    modalKind.textContent = KIND_LABEL[n.kind];
    modalKind.style.color = kindColorVar(n.kind);
    modalKind.style.background = kindBgVar(n.kind);
    modalTitle.textContent = n.label;
    modalSub.textContent = n.sub || "";
    modalDetail.textContent = n.detail || "";
    modalDetail.style.display = n.detail ? "block" : "none";
    if (n.sourceRef) { modalSrc.textContent = n.sourceRef; modalSrc.style.display = "inline-block"; }
    else modalSrc.style.display = "none";

    type ConnRow = { dir: "in" | "out"; other: Node; kind: string; label?: string };
    const rows: ConnRow[] = incoming[n.id].map((e): ConnRow => ({ dir: "in", other: nodeById[e.from], kind: e.kind, label: e.label }))
      .concat(outgoing[n.id].map((e): ConnRow => ({ dir: "out", other: nodeById[e.to], kind: e.kind, label: e.label })));

    if (rows.length) {
      let html = `<div class="conn-title">Connections (${rows.length})</div><ul class="cs-conn-list">`;
      rows.forEach((r) => {
        const arrow = r.dir === "in" ? "←" : "→";
        html += `<li><span class="arrow">${arrow}</span>${r.other.label}${r.label ? ` <span class="mono" style="color:var(--cs-muted)">· ${r.label}</span>` : ""}<span class="kind-pill">${r.kind}</span></li>`;
      });
      html += "</ul>";
      modalConnWrap.innerHTML = html;
    } else modalConnWrap.innerHTML = "";

    backdrop.classList.add("open");
  }
  function closeModal(skipClearSelection?: boolean) {
    backdrop.classList.remove("open");
    if (!skipClearSelection) clearSelection();
  }

  /* ---------- selection ---------- */
  let activeId: string | null = null;
  function clearSelection() {
    activeId = null;
    Object.keys(nodeEls).forEach((id) => nodeEls[id].classList.remove("is-active", "is-dim"));
    edgePaths.forEach((ep) => {
      ep.path.classList.remove("is-active", "is-dim");
      ep.path.setAttribute("stroke", "var(--cs-border-strong)");
      ep.path.setAttribute("marker-end", "url(#cs-arrow-muted)");
    });
    svg.querySelectorAll("text.cs-edge-label").forEach((n) => n.remove());
  }
  function selectNode(id: string) {
    if (activeId === id) { clearSelection(); closeModal(true); return; }
    clearSelection();
    activeId = id;
    const related = incoming[id].concat(outgoing[id]);
    const relatedIds: Record<string, boolean> = { [id]: true };
    related.forEach((e) => { relatedIds[e.from] = true; relatedIds[e.to] = true; });

    Object.keys(nodeEls).forEach((nid) => {
      if (relatedIds[nid]) { if (nid === id) nodeEls[nid].classList.add("is-active"); }
      else nodeEls[nid].classList.add("is-dim");
    });

    edgePaths.forEach((ep) => {
      if (ep.from === id || ep.to === id) {
        ep.path.classList.add("is-active");
        ep.path.setAttribute("stroke", "var(--cs-blue)");
        ep.path.setAttribute("marker-end", "url(#cs-arrow-blue)");
        if (ep.label) {
          const t = document.createElementNS(svgNS, "text");
          t.setAttribute("x", String(ep.midX)); t.setAttribute("y", String(ep.midY - 4));
          t.setAttribute("text-anchor", "middle"); t.setAttribute("class", "cs-edge-label");
          t.textContent = ep.label;
          svg.appendChild(t);
        }
      } else {
        ep.path.classList.add("is-dim");
      }
    });
    openModal(nodeById[id]);
  }

  function focusNode(id: string) {
    const n = nodeById[id];
    if (!n) return;
    const vw = viewport.clientWidth, vh = viewport.clientHeight;
    scale = 1;
    panX = vw / 2 - (n.x + CARD_W / 2) * scale - 60;
    panY = vh / 2 - (n.y + 32) * scale;
    applyTransform();
    selectNode(id);
  }

  root.querySelector(".cs-modal-close")!.addEventListener("click", () => closeModal());
  backdrop.addEventListener("click", (ev) => { if (ev.target === backdrop) closeModal(); });
  const onKeyDown = (ev: KeyboardEvent) => { if (ev.key === "Escape") closeModal(); };
  document.addEventListener("keydown", onKeyDown);

  /* ---------- side panel ---------- */
  const sideIntegrations = root.querySelector<HTMLDivElement>(".cs-side-integrations")!;
  sideIntegrations.innerHTML = TOP_INTEGRATIONS.map((it) =>
    `<div class="cs-side-row" data-id="${it.id}"><span class="avatar" style="background:${avatarColor(it.color)}">${it.initial}</span>${it.label}</div>`,
  ).join("");
  sideIntegrations.querySelectorAll<HTMLDivElement>(".cs-side-row").forEach((row) => {
    row.addEventListener("click", () => focusNode(row.dataset.id!));
  });

  /* ---------- filter bar ---------- */
  const activeFilters: Record<string, boolean> = {};
  const filterBar = root.querySelector<HTMLDivElement>(".cs-filter-bar")!;
  FILTERS.forEach((f) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "cs-filter-chip";
    chip.innerHTML = `<span class="dot"></span>${f.label}`;
    chip.addEventListener("click", () => {
      if (activeFilters[f.key]) delete activeFilters[f.key]; else activeFilters[f.key] = true;
      chip.classList.toggle("on");
      applyFilters();
    });
    filterBar.appendChild(chip);
  });
  function applyFilters() {
    const anyOn = Object.keys(activeFilters).length > 0;
    NODES.forEach((n) => {
      const el = nodeEls[n.id];
      const match = !anyOn || activeFilters[n.kind];
      el.classList.toggle("is-filtered-out", !match);
    });
  }

  /* ---------- pan / zoom ---------- */
  let panX = 340, panY = 90, scale = 0.85;
  let isDragging = false, dragMoved = false, lastX = 0, lastY = 0;

  function applyTransform() {
    world.style.transform = `translate(${panX}px,${panY}px) scale(${scale})`;
  }
  function resetView() { panX = 340; panY = 90; scale = 0.85; applyTransform(); }
  applyTransform();

  const onPointerDown = (ev: PointerEvent) => {
    isDragging = true; dragMoved = false; lastX = ev.clientX; lastY = ev.clientY;
    viewport.classList.add("dragging");
    viewport.setPointerCapture(ev.pointerId);
  };
  const onPointerMove = (ev: PointerEvent) => {
    if (!isDragging) return;
    const dx = ev.clientX - lastX, dy = ev.clientY - lastY;
    if (Math.abs(dx) + Math.abs(dy) > 2) dragMoved = true;
    panX += dx; panY += dy; lastX = ev.clientX; lastY = ev.clientY;
    applyTransform();
  };
  const endDrag = () => { isDragging = false; viewport.classList.remove("dragging"); };
  const onWheel = (ev: WheelEvent) => {
    ev.preventDefault();
    const rect = viewport.getBoundingClientRect();
    const cx = ev.clientX - rect.left, cy = ev.clientY - rect.top;
    const prevScale = scale;
    const factor = ev.deltaY < 0 ? 1.1 : 0.9;
    scale = Math.min(1.8, Math.max(0.35, scale * factor));
    panX = cx - (cx - panX) * (scale / prevScale);
    panY = cy - (cy - panY) * (scale / prevScale);
    applyTransform();
  };
  viewport.addEventListener("pointerdown", onPointerDown);
  viewport.addEventListener("pointermove", onPointerMove);
  viewport.addEventListener("pointerup", endDrag);
  viewport.addEventListener("pointercancel", endDrag);
  viewport.addEventListener("wheel", onWheel, { passive: false });

  const zoomIn = root.querySelector(".cs-zoom-in")!;
  const zoomOut = root.querySelector(".cs-zoom-out")!;
  const zoomReset = root.querySelector(".cs-zoom-reset")!;
  const onZoomIn = () => {
    const prevScale = scale; scale = Math.min(1.8, scale * 1.2);
    const vw = viewport.clientWidth / 2, vh = viewport.clientHeight / 2;
    panX = vw - (vw - panX) * (scale / prevScale); panY = vh - (vh - panY) * (scale / prevScale);
    applyTransform();
  };
  const onZoomOut = () => {
    const prevScale = scale; scale = Math.max(0.35, scale * 0.8);
    const vw = viewport.clientWidth / 2, vh = viewport.clientHeight / 2;
    panX = vw - (vw - panX) * (scale / prevScale); panY = vh - (vh - panY) * (scale / prevScale);
    applyTransform();
  };
  zoomIn.addEventListener("click", onZoomIn);
  zoomOut.addEventListener("click", onZoomOut);
  zoomReset.addEventListener("click", resetView);

  /* ---------- theme toggle ---------- */
  const toggle = root.querySelector(".cs-theme-toggle")!;
  function currentTheme() {
    const stored = root.getAttribute("data-theme");
    if (stored) return stored;
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }
  const onToggleTheme = () => {
    root.setAttribute("data-theme", currentTheme() === "dark" ? "light" : "dark");
  };
  toggle.addEventListener("click", onToggleTheme);

  return function cleanup() {
    document.removeEventListener("keydown", onKeyDown);
    viewport.removeEventListener("pointerdown", onPointerDown);
    viewport.removeEventListener("pointermove", onPointerMove);
    viewport.removeEventListener("pointerup", endDrag);
    viewport.removeEventListener("pointercancel", endDrag);
    viewport.removeEventListener("wheel", onWheel);
    zoomIn.removeEventListener("click", onZoomIn);
    zoomOut.removeEventListener("click", onZoomOut);
    zoomReset.removeEventListener("click", resetView);
    toggle.removeEventListener("click", onToggleTheme);
  };
}

export function CodebaseScanEmbed() {
  const rootRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!rootRef.current) return;
    return mountScan(rootRef.current);
  }, []);

  return <div className="codebase-scan-embed" ref={rootRef}>
    <style>{SCAN_CSS}</style>
    <div className="cs-viewport">
      <div className="cs-world">
        <svg className="cs-edges" />
      </div>
    </div>
    <div className="cs-brand-badge">
      <span className="bt"><b>PULSE</b> · Codebase Scan</span>
    </div>
    <div className="cs-top-right">
      <span className="cs-pill-btn">Generated from apps/</span>
      <button className="cs-icon-btn cs-theme-toggle" type="button" aria-label="Toggle color theme" title="Toggle theme">◐</button>
    </div>
    <div className="cs-side-panel">
      <div className="sec">
        <h4>Integrations</h4>
        <div className="cs-side-integrations" />
      </div>
    </div>
    <div className="cs-filter-bar" />
    <div className="cs-zoom-ctl">
      <button className="cs-icon-btn cs-zoom-in" type="button" aria-label="Zoom in">+</button>
      <button className="cs-icon-btn cs-zoom-out" type="button" aria-label="Zoom out">−</button>
      <button className="cs-icon-btn cs-zoom-reset" type="button" aria-label="Reset view" title="Reset view" style={{ fontSize: 11 }}>⤢</button>
    </div>
    <div className="cs-modal-backdrop">
      <div className="cs-modal" role="dialog" aria-modal="true">
        <div className="cs-modal-head">
          <div>
            <span className="kind-tag cs-modal-kind" />
            <h3 className="cs-modal-title" />
            <div className="sub cs-modal-sub" />
          </div>
          <button className="cs-modal-close" type="button" aria-label="Close">✕</button>
        </div>
        <p className="detail cs-modal-detail" />
        <span className="source-ref mono cs-modal-src" style={{ display: "none" }} />
        <div className="cs-modal-conn-wrap" />
      </div>
    </div>
  </div>;
}
