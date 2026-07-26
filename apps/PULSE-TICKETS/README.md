# PULSE Tickets — Standalone Issue Tracker

A standalone SharePoint-hosted tool for viewing and managing issues reported by PULSE users. Reads from the same **"PULSE Issues"** SharePoint list that the built-in PULSE tickets page writes to.

---

## What it is

PULSE Tickets is a lightweight companion app — a single HTML file deployed to a Forge web part or a SharePoint page — that gives admins a full-page, tab-based issue tracker without needing to open the main PULSE dashboard. It shares no code with PULSE at runtime, but reads `../PULSE/assets/js/app-config.js` for site URL configuration.

---

## SharePoint list

**List name:** `PULSE Issues`

| Field | Internal name | Type | Notes |
|---|---|---|---|
| Title | `Title` | Single line | Issue headline |
| Issue code | `IssueCode` | Single line | e.g. `ISS-000042` |
| Status | `IssueStatus` | Choice | New · In Progress · Resolved · Closed |
| Type | `IssueType` | Choice | Bug · Feature Request · Other |
| Reporter name | `ReportedByName` | Single line | |
| Reporter email | `ReporterEmail` | Single line | |
| Occurred at | `OccurredAt` | DateTime | |
| Description | `Description` | Multi-line | |
| Expected behavior | `ExpectedBehavior` | Multi-line | |
| Additional context | `AdditionalContext` | Multi-line | Append-only update log |
| Page title | `PageTitle` | Single line | PULSE page where issue occurred |
| Route | `Route` | Single line | Hash route, e.g. `#/projects` |
| Page URL | `PageUrl` | Single line | Full URL |
| Error codes | `ErrorCodesJson` | Multi-line | JSON array of strings |
| Browser logs | `LogsJson` | Multi-line | JSON array `[{level, message, ts}]` |
| Diagnostics | `DiagnosticsJson` | Multi-line | JSON object — browser, viewport, `sharePointDebug[]` |
| Screenshot data URL | `ScreenshotDataUrl` | Multi-line | Base64 data URI (inline) |
| Screenshot file URL | `ScreenshotFileUrl` | Single line | SP-hosted file URL |
| Screenshot server-relative URL | `ScreenshotServerRelativeUrl` | Single line | |
| Screenshot filename | `ScreenshotFileName` | Single line | |
| Resolution note | `ResolutionNote` | Multi-line | |

---

## Detail view tabs

Opening a ticket shows a tabbed detail page:

| Tab | Content |
|---|---|
| Overview | Description, expected result, additional context/updates, add-update composer |
| Screenshot | Captured screenshot with lightbox; thumbnail also in sidebar |
| Browser Logs | All `LogsJson` entries — errors/warnings highlighted, info/debug below |
| App Logs | `DiagnosticsJson.sharePointDebug` entries captured by PULSE |
| Error Codes | `ErrorCodesJson` chips + full client diagnostics grid |

Tabs only appear if that data is present on the ticket.

---

## Building

```bash
cd apps/PULSE-TICKETS
node scripts/build-sharepoint-package.js
# → releases/PULSE-Tickets-v1.0.0.html
```

Upload `releases/PULSE-Tickets-v1.0.0.html` to a SharePoint Forge web part.

---

## Hash routing / deep-links

Navigating to `PULSE-Tickets.aspx#/ISS-000042` opens that ticket's detail view directly. The list view is the default when no hash is present.

---

## Site URL detection

The tool detects the SharePoint site URL via this priority chain:

1. `window._spPageContextInfo.webAbsoluteUrl` — injected by SharePoint when hosted in a Forge web part
2. `APP_CONFIG.manualSharePointSiteUrl` — set in `apps/PULSE/assets/js/app-config.js`
3. `window.location.origin + _spPageContextInfo.webServerRelativeUrl`
4. `window.location.origin` — last resort

If tickets fail to load in development, set `manualSharePointSiteUrl` in `app-config.js`.

---

## Version bump convention

After editing source files, bump the query-string version on the changed file in `index.html` before rebuilding:

```html
<link rel="stylesheet" href="assets/css/tickets.css?v=20260726c">
<script src="assets/js/tickets.js?v=20260726b"></script>
```

Use the date (`YYYYMMDD`) plus a letter suffix for same-day changes (`a`, `b`, `c`…).
