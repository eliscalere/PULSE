# AEWTTR PULSE — IPT Dashboard Suite

**Version:** 1.0.0  
**Program:** AEWTTR IPT DB46200 · NAVAIR  
**Developer:** Eli Scalere (Contractor) — elijah.t.scalere.ctr@us.navy.mil  
**Out-of-service contact:** eli.scalere@scaleredesign.com · (516) 265-2636

---

PULSE is a suite of single-file web apps deployed directly into **SharePoint Online (Flank Speed)**. No server, no build pipeline to deploy — edit source, run the build script, upload the `.html` file. SharePoint handles authentication and data storage through its native REST API using the user's existing browser session (CAC/PIV/SSO).

---

## Apps

| App | Release file | Purpose |
|---|---|---|
| **PULSE** | `releases/PULSE-v1.0.0.html` | Main IPT dashboard — projects, travel, meetings, docs, admin |
| **PULSE Tickets** | `releases/PULSE-Tickets-v1.0.0.html` | Issue tracker |
| **PULSE Calendar** | `releases/PULSE-Calendar-v1.0.0.html` | Team travel and leave calendar |
| **PULSE CODE** | `releases/PULSE-CODE-v1.0.0.html` | In-browser code editor for SharePoint-hosted files |

---

## Project Map

```
AEWTTR PAS/
│
├── releases/                               ← DEPLOY THESE TO SHAREPOINT
│   ├── PULSE-v1.0.0.html                  (6.8 MB — main dashboard)
│   ├── PULSE-Calendar-v1.0.0.html         (314 KB — travel calendar)
│   ├── PULSE-Tickets-v1.0.0.html          ( 64 KB — issue tracker)
│   └── PULSE-CODE-v1.0.0.html             ( 92 KB — code editor)
│
├── apps/
│   ├── PULSE/                              Primary app source (start here)
│   │   ├── index.html                      Entry point
│   │   ├── assets/js/
│   │   │   ├── app-config.js               App-level config + APP_VERSION
│   │   │   ├── app.js                      Core shell, router, shared utils
│   │   │   ├── data.js                     In-memory data model + normalizers
│   │   │   ├── sharepoint-repo.js          SharePoint data layer (read/write)
│   │   │   ├── sharepoint-schema.js        SP list definitions + seeding
│   │   │   ├── sharepoint-adapter.js       Low-level SP REST calls
│   │   │   ├── audit-log.js                Audit trail (multi-list rotation)
│   │   │   ├── export.js                   PowerPoint export (pptxgenjs)
│   │   │   ├── notify.js                   Email notifications via SP
│   │   │   └── pages/                      One file per route
│   │   │       ├── dashboard.js
│   │   │       ├── projects.js
│   │   │       ├── overview.js
│   │   │       ├── weekly.js
│   │   │       ├── travel.js
│   │   │       ├── docreview.js
│   │   │       ├── admin.js
│   │   │       ├── logs.js
│   │   │       └── users.js
│   │   ├── assets/css/
│   │   │   ├── style.css
│   │   │   └── dropdowns.css
│   │   ├── assets/images/                  Seals, wordmarks, logos
│   │   ├── vendor/                         Vendored libraries (offline-safe)
│   │   │   ├── boxicons/
│   │   │   ├── flatpickr/
│   │   │   ├── fullcalendar/
│   │   │   └── pptxgenjs/
│   │   └── scripts/
│   │       ├── build-sharepoint-package.js → releases/PULSE-v1.0.0.html
│   │       └── build-forge.js              → Forge web part bundle (see below)
│   │
│   ├── PULSE-TICKETS/
│   │   ├── index.html
│   │   ├── assets/js/tickets.js
│   │   ├── assets/css/tickets.css
│   │   └── scripts/build-sharepoint-package.js
│   │
│   ├── PULSE-TRAVEL-CALENDAR/
│   │   ├── index.html
│   │   ├── assets/js/calendar.js
│   │   ├── assets/css/calendar.css
│   │   └── scripts/build-sharepoint-package.js
│   │
│   └── PULSE-CODE/                         In-browser code editor
│       ├── index.html                       Single self-contained file
│       └── scripts/build-sharepoint-package.js
│
├── artifacts/generated/                    SharePoint list schema CSVs + Excel template
├── assets/branding/                        Logos, seals (AEWTTR, NAVAIR, TTSD)
├── deliverables/                           Final AoA and presentation files
└── tools/sharepoint/                       SP list schema build scripts
```

---

## How the Stack Works

Every app follows the same pattern:

```
Browser session (CAC/PIV login to Flank Speed)
        │
        ▼
HTML file served from SharePoint SiteAssets or a Forge web part
        │
        ├─► SharePoint REST API  (/_api/web/…)
        │     • Same-origin fetch with credentials:'include'
        │     • No Azure AD app registration
        │     • No separate auth tokens — uses the existing session
        │
        └─► (PULSE CODE only) AskSage AI
              POST https://api.asksage.ai/server/openai/v1/chat/completions
              Authorization: Bearer {api_key}
```

**SharePoint lists** store all data. `sharepoint-adapter.js` handles the low-level REST calls (contextinfo digest for writes, `odata=nometadata` for lean responses). `sharepoint-repo.js` provides the business-logic read/write layer on top.

**No npm, no bundler, no framework.** Vendor libraries are checked in under `apps/PULSE/vendor/`. The build scripts are plain Node.js `require`-style scripts with zero dependencies.

---

## Building a Release (WFC Package)

Each app has a `scripts/build-sharepoint-package.js` that inlines all local assets into a single self-contained HTML file with a `WFC-MANIFEST` header.

```bash
# Run from the repo root

# PULSE (main dashboard)
node apps/PULSE/scripts/build-sharepoint-package.js
# → releases/PULSE-v1.0.0.html

# PULSE Tickets
node apps/PULSE-TICKETS/scripts/build-sharepoint-package.js
# → releases/PULSE-Tickets-v1.0.0.html

# PULSE Travel Calendar
node apps/PULSE-TRAVEL-CALENDAR/scripts/build-sharepoint-package.js
# → releases/PULSE-Calendar-v1.0.0.html

# PULSE CODE
node apps/PULSE-CODE/scripts/build-sharepoint-package.js
# → releases/PULSE-CODE-v1.0.0.html
```

No `npm install` needed — PULSE, Tickets, and Calendar vendor all their dependencies locally. PULSE CODE loads Monaco Editor from the jsdelivr CDN at runtime.

**What the build does:**
1. Reads `index.html`
2. Inlines every local `<link rel="stylesheet">` and `<script src="...">` (replaces with `<style>` and `<script>` blocks)
3. Inlines CSS `url()` asset references as base64 data URIs
4. Writes a `<!--WFC-MANIFEST:...-->` comment at the top (base64 JSON inventory of bundled files)
5. Outputs a single `.html` file to `releases/`

The output file is what you upload to SharePoint. Once uploaded, it is completely self-contained — no other files needed.

---

## Packaging with Forge

**Forge** is the Flank Speed web part host that wraps an app in an iframe using an `srcdoc` attribute. Use this path when you need to embed the app inside an existing SharePoint page (rather than opening it as a standalone HTML document).

### What you need

- A **Forge template file** — a Forge `.html` stub that contains the placeholder string `const CHILD_HTML_B64 = "";`. This file is provided by the Forge platform administrator and is **not** in this repository.
- Node.js (any recent version)

### How to run it

```bash
# From the repo root:
node apps/PULSE/scripts/build-forge.js <path-to-forge-template.html> [output-path]
```

Example:

```bash
node apps/PULSE/scripts/build-forge.js ~/forge-template.html releases/forge-builds/PULSE-Forge.html
```

The script outputs two files to the output path you specify (defaulting to `apps/PULSE/releases/forge-builds/`):
- `Forge.html` — the packed file for upload (base64-encoded child app)
- `Forge-readable.html` — a human-readable version using a template literal instead of base64 (useful for debugging)

### What it does internally

```
index.html + assets/
      │
      ▼ build-forge.js
  1. Inlines all local CSS → <style> blocks
  2. Inlines all local JS  → <script> blocks
  3. Base64-encodes the assembled child HTML
  4. Replaces `const CHILD_HTML_B64 = "";` in the Forge template
      │
      ▼
Forge.html  (Forge template + base64-encoded PULSE inside)
```

The Forge template then decodes and writes the child HTML into an iframe's `srcdoc` at runtime, isolating PULSE from the SharePoint page chrome.

> **Note:** Only `apps/PULSE/` has a Forge build script. PULSE Tickets, PULSE Calendar, and PULSE CODE are deployed as standalone HTML files (WFC packages), not as Forge web parts.

---

## PULSE CODE — Quick Start

PULSE CODE is a Monaco-based code editor that reads and writes files directly to a SharePoint document library using the same session auth as the other apps.

1. Upload `releases/PULSE-CODE-v1.0.0.html` to `SiteAssets` on your SharePoint site
2. Open it from SharePoint (not from a local file — session auth requires the SharePoint origin)
3. Click **◈ No project — click to add** in the sidebar
4. Paste any SharePoint folder URL from your browser's address bar
5. The editor connects, lists the folder's files, and auto-syncs your Ask Sage API key from `SiteAssets/pulse-code-config.json`

**Settings → Ask Sage AI:** Enter your API key once — it's saved to `SiteAssets/pulse-code-config.json` and shared automatically with anyone else running PULSE CODE on the same site.

---

## Versioning

| App | Current version |
|---|---|
| PULSE | `1.0.0` |
| PULSE Calendar | `1.0.0` |
| PULSE Tickets | `1.0.0` |
| PULSE CODE | `1.0.0` |

To cut a new release: update `VERSION` in the relevant `scripts/build-sharepoint-package.js`, update `APP_VERSION` in `apps/PULSE/assets/js/app-config.js` (for PULSE), then run the build.

---

## SharePoint Lists (PULSE)

| List | Purpose |
|---|---|
| PULSE Projects | Project records |
| PULSE Travel Requests | Travel submissions + approvals |
| PULSE Travel Debriefs | Post-trip debrief records |
| PULSE Meetings | Weekly and project meeting records |
| PULSE Documents | Document review queue |
| PULSE Issues | Issue / ticket records (Tickets app) |
| PULSE App Roles | User role assignments |
| PULSE Audit Log | Activity audit trail (rotates to Audit Log 2, 3 … at 1000 items) |
| PULSE Notification Config | Per-user notification preferences |
| PULSE Location Config | Office / site location records |

Schema CSVs and an Excel template for initial list setup are in `artifacts/generated/`.

---

## Contact

> **Developer:** Eli Scalere (Contractor)  
> **Program contact:** elijah.t.scalere.ctr@us.navy.mil  
> **Out-of-service / personal:** eli.scalere@scaleredesign.com · (516) 265-2636  
> **Organization:** AEWTTR IPT · DB46200 · NAVAIR
