# Feature Verification Matrix

| Feature | Source evidence | Local route capture | Live SharePoint verification |
|---|---|---|---|
| Shell, Dashboard, notifications, account settings | `app.js`, `dashboard.js`, `notification-settings.js` | Captured | [VERIFY] |
| Overview and workload views | `pages/overview.js`, `pages/workload.js` | Captured overview | [VERIFY] |
| Projects, portfolios, programs, end-item configurations | `pages/projects.js`, repository mapper | Captured projects | [VERIFY] create, edit, filtering, access |
| Project People and Documents | `pages/people.js`, `pages/project-documents.js` | Not populated | [VERIFY] directory and library operations |
| Tracker, Gantt, Notes, risks, settings | `pages/projects.js`, `data.js` | Not populated | [VERIFY] create, update, persistence |
| Boards | `pages/checklists.js`, project boards mapper | Not populated | [VERIFY] board creation and persistence |
| Reporting and PowerPoint output | export modules and bundled PptxGenJS | Source reviewed | [VERIFY] generated file and SharePoint upload |
| Weekly and project meetings | `pages/weekly.js` | Captured weekly route | [VERIFY] session and action-item persistence |
| Travel and debriefs | `pages/travel.js` | Captured travel route | [VERIFY] approvals, finance, debriefs |
| Formal Document Review | `pages/docreview.js` | Captured review route | [VERIFY] file, revision, reviewer, signer workflows |
| Tickets | `pages/tickets.js`, `issue-reporting.js` | Source reviewed | [VERIFY] submission and update persistence |
| Admin, users, roles, locations, logs, schema setup | admin, users, logs, schema, adapter modules | Captured restricted entry state | [VERIFY] administrator actions |
| REST and role resolution | adapter, repository, schema modules | Not applicable | [VERIFY] on approved SharePoint site |

`[VERIFY]` means the repository supports the feature but the current workspace has no authorized live SharePoint endpoint to validate it.
