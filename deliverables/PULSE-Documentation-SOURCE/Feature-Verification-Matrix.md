# Feature Verification Matrix

Local route capture reflects the 2026-07-29 set described in [Screenshot Manifest](Screenshot-Manifest.md). That set was taken against seeded records, so most rows now show populated state rather than an empty route.

| Feature | Source evidence | Local route capture | Live SharePoint verification |
|---|---|---|---|
| Shell, Dashboard, notifications, account settings | `app.js`, `dashboard.js`, `notification-settings.js` | Captured populated (01, 17) | [VERIFY] |
| Overview and workload views | `pages/overview.js`, `pages/workload.js` | Captured overview populated (02); workload not captured | [VERIFY] |
| Projects, portfolios, programs, end-item configurations | `pages/projects.js`, repository mapper | Captured populated (03, 04) | [VERIFY] create, edit, filtering, access |
| Project People and Documents | `pages/people.js`, `pages/project-documents.js` | Captured populated (07, 08) | [VERIFY] directory and library operations |
| Tracker, Gantt, Notes, risks, settings | `pages/projects.js`, `data.js` | Captured tracker with tasks (05); Gantt/risks tabs not captured | [VERIFY] create, update, persistence |
| Boards and checklists | `pages/checklists.js`, project boards mapper | Captured checklist (06) | [VERIFY] board creation and persistence |
| Reporting and PowerPoint output | export modules and bundled PptxGenJS | Source reviewed; not captured | [VERIFY] generated file and SharePoint upload |
| Weekly and project meetings | `pages/weekly.js` | Captured live session with attendance and minutes (09) | [VERIFY] session and action-item persistence |
| Travel requests, leave, and debriefs | `pages/travel.js` | Captured form and debrief (10, 13) | [VERIFY] approvals, finance, charge object, debriefs |
| My Travel | `pages/travel.js`, travel-request mapper | Captured populated with state tabs (11) | [VERIFY] per-user filtering against live identity |
| Travel Calendar and team events | `pages/travel.js`, team-event mapper | Captured (12) | [VERIFY] event persistence and cross-user visibility |
| Formal Document Review | `pages/docreview.js` | Captured with reviewer decisions (14) | [VERIFY] file, revision, reviewer, signer workflows |
| Tickets | `pages/tickets.js`, `issue-reporting.js` | Captured populated (15) | [VERIFY] submission and update persistence in **PULSE Issues** |
| Admin, users, roles, locations, logs, schema setup | admin, users, logs, schema, adapter modules | Captured admin entry and activity log (16, 18) | [VERIFY] administrator actions |
| Focused single-file packages | `scripts/build-travel-packages.js`, `PULSE_PORT_CONFIG` entry files | Build verified for all eight packages | [VERIFY] site resolution and read/write from a web-part page |
| Documentation package | `apps/pulse-documentation` | Build verified; boot and asset integrity verified top-level and in a sandboxed iframe | [VERIFY] rendering on the hosted page |
| PULSE CODE | `apps/PULSE-CODE` | Build verified; not functionally exercised | [VERIFY] authentication, file read/write, and AI features |
| REST and role resolution | adapter, repository, schema modules | Not applicable | [VERIFY] on approved SharePoint site |
| Behaviour at scale | paging guard, server-side sorts, full-workspace refresh | Not applicable; no large data set exists | [VERIFY] on a loaded development site per SOP 10 |
| Development environment | none present in source or configuration | Not applicable | Not yet established; see SOP 10 and document 10 |

`[VERIFY]` means the repository supports the feature but the current workspace has no authorized live SharePoint endpoint to validate it.

Local capture confirms routes, labels, and rendered state only. It is not evidence of a completed SharePoint transaction, of administrator-only control, or of notification delivery.
