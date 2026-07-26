# AEWTTR PULSE — IPT Dashboard Suite

**Version:** 1.0.0  
**Program:** AEWTTR IPT DB46200 · NAVAIR  
**Developer:** Eli Scalere (Contractor) — elijah.t.scalere.ctr@us.navy.mil  
**Out-of-service contact:** eli.scalere@scaleredesign.com · (516) 265-2636

---

PULSE is a single-file web dashboard deployed inside a **Forge web part on Flank Speed (SharePoint Online)**. No server, no build pipeline to deploy — edit source, run the build script, upload the `.html` file. SharePoint handles authentication and data storage via its native REST API.

---

## Project Map

```
AEWTTR PAS/
│
├── releases/                          ← DEPLOY THESE TO SHAREPOINT
│   ├── PULSE-v1.0.0.html             (6.8 MB — main dashboard)
│   ├── PULSE-Calendar-v1.0.0.html    (314 KB — travel calendar)
│   └── PULSE-Tickets-v1.0.0.html     ( 64 KB — issue tracker)
│
├── apps/
│   ├── PULSE/                         ← Primary app source (start here)
│   │   ├── index.html                 Entry point
│   │   ├── assets/
│   │   │   ├── js/
│   │   │   │   ├── app-config.js      App-level config + APP_VERSION
│   │   │   │   ├── app.js             Core shell, router, shared utils
│   │   │   │   ├── data.js            In-memory data model + normalizers
│   │   │   │   ├── sharepoint-repo.js SharePoint data layer (read/write)
│   │   │   │   ├── sharepoint-schema.js  SP list definitions + seeding
│   │   │   │   ├── sharepoint-adapter.js Low-level SP REST calls
│   │   │   │   ├── audit-log.js       Audit trail (multi-list rotation)
│   │   │   │   ├── export.js          PowerPoint export (pptxgenjs)
│   │   │   │   ├── notify.js          Email notifications via SP
│   │   │   │   ├── project-pptx-export.js  Per-project slide generation
│   │   │   │   └── pages/             One file per route
│   │   │   │       ├── dashboard.js
│   │   │   │       ├── projects.js
│   │   │   │       ├── overview.js
│   │   │   │       ├── weekly.js
│   │   │   │       ├── travel.js
│   │   │   │       ├── docreview.js
│   │   │   │       ├── admin.js
│   │   │   │       ├── logs.js
│   │   │   │       └── users.js
│   │   │   ├── css/
│   │   │   │   ├── style.css          Main stylesheet
│   │   │   │   └── dropdowns.css      Dropdown component styles
│   │   │   └── images/                Seals, wordmarks, logos
│   │   ├── vendor/                    Vendored libraries (offline-safe)
│   │   │   ├── boxicons/              Icon font
│   │   │   ├── flatpickr/             Date picker
│   │   │   ├── fullcalendar/          Calendar widget
│   │   │   └── pptxgenjs/            PowerPoint generation
│   │   ├── docs/                      Developer + user documentation
│   │   │   ├── current/               Canonical current-version docs
│   │   │   │   ├── PULSE-User-Guide.md
│   │   │   │   ├── PULSE-Technical-Handoff.md
│   │   │   │   └── PULSE-Operations-SOP.md
│   │   │   ├── handoff/
│   │   │   │   ├── FORGE-DEPLOYMENT-GUIDE.md   How to deploy to Flank Speed
│   │   │   │   └── FS-FORGE-STEPS.md            Step-by-step upload checklist
│   │   │   └── architecture/          Technical deep-dives
│   │   ├── scripts/
│   │   │   ├── build-sharepoint-package.js  → outputs releases/PULSE-v1.0.0.html
│   │   │   └── maintenance/           Diagram + screenshot utilities
│   │   └── validation/pdf-samples/    Engineering travel PDF test output
│   │
│   ├── PULSE-TICKETS/                 Issue tracker companion app
│   │   ├── index.html
│   │   ├── assets/js/tickets.js
│   │   ├── assets/css/tickets.css
│   │   └── scripts/build-sharepoint-package.js  → releases/PULSE-Tickets-v1.0.0.html
│   │
│   ├── PULSE-TRAVEL-CALENDAR/         Travel calendar companion app
│   │   ├── index.html
│   │   ├── assets/js/calendar.js
│   │   ├── assets/css/calendar.css
│   │   └── scripts/build-sharepoint-package.js  → releases/PULSE-Calendar-v1.0.0.html
│   │
│   └── spfx-pulse-workspace/          Legacy SPFx typed implementation (reference)
│
├── artifacts/
│   └── generated/
│       ├── spo-column-schemas-*/      SharePoint column schema CSVs
│       └── spo-list-templates-*/      SP list template CSVs + Excel workbook
│
├── assets/
│   ├── branding/                      Logos, seals (AEWTTR, NAVAIR, TTSD)
│   └── presentation/screenshots/      UI screenshots for presentations
│
├── archive/backups/                   Point-in-time source backups
├── deliverables/                      Final deliverable files (AoA, presentations)
├── docs/                              Project-level documentation
└── tools/sharepoint/                  SP list schema + template build scripts
```

---

## Building a Release

Each app has its own build script that inlines all assets into a single self-contained HTML file.

```bash
# Build PULSE (main dashboard)
cd apps/PULSE
node scripts/build-sharepoint-package.js
# → releases/PULSE-v1.0.0.html

# Build Calendar companion
cd apps/PULSE-TRAVEL-CALENDAR
node scripts/build-sharepoint-package.js
# → releases/PULSE-Calendar-v1.0.0.html

# Build Tickets companion
cd apps/PULSE-TICKETS
node scripts/build-sharepoint-package.js
# → releases/PULSE-Tickets-v1.0.0.html
```

No `npm install` needed — all dependencies are vendored in `apps/PULSE/vendor/`.

---

## Versioning

This project uses **Semantic Versioning (SemVer — MAJOR.MINOR.PATCH)**:

| Component | Current | Notes |
|---|---|---|
| PULSE | `1.0.0` | Core IPT dashboard — increment MINOR for new features, PATCH for fixes |
| PULSE-Calendar | `1.0.0` | Travel calendar companion |
| PULSE-Tickets | `1.0.0` | Issue tracker companion |

Update `VERSION` in each app's `scripts/build-sharepoint-package.js` before a release. Also update `APP_VERSION` in `apps/PULSE/assets/js/app-config.js`.

---

## Deployment (Flank Speed / Forge)

1. Run the build script for the app you are updating.
2. Open your Forge web part on SharePoint.
3. Upload the `.html` file from `releases/` using the Forge upload interface.
4. The app reads site context from `window._spPageContextInfo` automatically.

See `apps/PULSE/docs/handoff/FORGE-DEPLOYMENT-GUIDE.md` for the full deployment walkthrough.

---

## SharePoint Lists

PULSE reads from and writes to these SharePoint lists on the host site:

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
