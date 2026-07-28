# PULSE — Technical Reference

<img src="../assets/images/aewttr-seal.png" alt="AEWTTR Seal" width="80" align="right">

| | |
|---|---|
| **Document** | PULSE Technical Reference |
| **Audience** | Developers/maintainers of PULSE |
| **Scope** | System architecture, Firepit/Forge hosting model, SharePoint integration, and the code-update workflow |
| **Companion docs** | This document summarizes and organizes material from `AI-HANDOFF.md`, `FIREPIT-DEVELOPER-GUIDE.md`, `FS-FORGE-STEPS.md`, and `SHAREPOINT-LISTS-GUIDE.md` at the repository root. Those files carry more raw detail (postmortems, exact CAML XML, etc.) and should be treated as the primary source if this document and one of them ever disagree. |

---

## Table of Contents

1. [System Architecture at a Glance](#1-system-architecture-at-a-glance)
2. [Repository Layout](#2-repository-layout)
3. [Firepit — the Hosting Platform](#3-firepit--the-hosting-platform)
4. [Forge — the Packaging Tool](#4-forge--the-packaging-tool)
5. [The Deployment Pipeline](#5-the-deployment-pipeline)
6. [SharePoint Integration](#6-sharepoint-integration)
7. [Application Architecture](#7-application-architecture)
8. [How to Update the Code](#8-how-to-update-the-code)
9. [Known Gotchas](#9-known-gotchas)
10. [File Map](#10-file-map)

---

## 1. System Architecture at a Glance

PULSE is a **single-page, browser-only application**. There is no application server, no database server, and no separate authentication system.

```
 Browser (user's SharePoint session)
   │
   │  loads one HTML page from SharePoint / Firepit
   ▼
 PULSE (vanilla HTML/CSS/JS, no framework)
   │
   │  fetch() with credentials: "same-origin"
   ▼
 SharePoint Online REST API  (/_api/web/...)
   │
   ▼
 SharePoint Lists  =  the database
```

Everything the app needs to run — the UI, the logic, and the "database" — lives inside SharePoint itself:

- **Compute**: runs entirely in the visiting user's browser.
- **Identity**: the user's existing SharePoint/CAC session. No separate login, no Azure app registration, no OAuth flow.
- **Storage**: SharePoint Lists (structured records) and a SharePoint document library (files). No SQL database, no external API.
- **Hosting**: the app is delivered to the browser as a single HTML file, either through a **Firepit** web part or directly from a SharePoint document library.

This is a deliberate constraint, not an oversight — see `AI-HANDOFF.md`'s "Hard Requirements": no Microsoft Graph as a dependency, no Azure app registration, no backend server, no external APIs for app data, no CDN/runtime internet dependencies in the shipped build. Every one of those would require infrastructure and approvals outside what a SharePoint site owner can control alone.

---

## 2. Repository Layout

```
apps/PULSE/
├── index.html                    Entry point (dev/unbundled — references separate files)
├── assets/
│   ├── css/style.css              All application styling
│   ├── images/                    Logo, icons
│   └── js/
│       ├── app-config.js          APP_CONFIG — site overrides, list prefix, boot diagnostics
│       ├── app.js                 Shell, router, state, boot sequence (see §7)
│       ├── data.js                Data model / store shape (see §7.4)
│       ├── audit-log.js           Activity Log subsystem
│       ├── notify.js              Notification/email/Teams templating
│       ├── sharepoint-adapter.js  All raw SharePoint REST/Graph plumbing (see §6)
│       ├── sharepoint-repo.js     Repo.save/remove, list↔object mappers (see §6)
│       ├── sharepoint-schema.js   List/column schema + setup/provisioning (see §6)
│       └── pages/*.js             One file per feature/page (dashboard, projects, travel, etc.)
├── vendor/                        Third-party libraries (Bootstrap, FullCalendar, Flatpickr, Boxicons)
├── scripts/
│   ├── build-sharepoint-package.js  Preferred packaging script (see §4)
│   └── build-forge.js               Secondary/offline-wrapper packaging script (see §4)
├── dist/                          Output of build-forge.js (not the preferred ship artifact)
├── FS packages/                   Output of build-sharepoint-package.js — THE SHIP ARTIFACT
├── releases/reference-packages/                  Known-good reference copy of a real Forge-tool output
├── AI-HANDOFF.md                  Maintainer handoff notes — hard requirements, editing priorities
├── FIREPIT-DEVELOPER-GUIDE.md     Firepit/Forge deep-dive, postmortems
├── FS-FORGE-STEPS.md              Step-by-step Forge packaging notes
├── SHAREPOINT-LISTS-GUIDE.md      SharePoint REST/schema deep-dive
└── docs/                        This document and its companions
```

---

## 3. Firepit — the Hosting Platform

**Firepit** is the SharePoint Online web part that hosts and executes PULSE. It is not a build system — it takes whatever HTML it is given and renders it directly. Inline `<script>` tags execute; inline `<style>` tags apply. There is no server-side dependency resolution.

That has real consequences for how the app must be packaged:

| Constraint | Why it matters |
|---|---|
| **Must be one self-contained HTML file** | No `<script src="...">` / `<link href="...">` pointing at separate files — Firepit has nowhere to fetch them from. Everything (JS, CSS, fonts, images) must already be inlined. |
| **Same-origin only** | The page runs inside the SharePoint origin, so `fetch()` calls carry the user's session cookie automatically. This is *why* no separate login is needed — but it also means the app cannot call anything off-origin without breaking that assumption. |
| **Variable web part width** | The web part can be placed next to other web parts at less than full page width. The UI must stay usable at a narrow viewport. |
| **No cache-busting query strings on asset tags** | SharePoint's static-file serving mishandles `?v=N`-style query strings on script/link URLs — this causes a **silent** 404 (blank page, no console errors). See §9. |
| **No CDN / runtime internet dependencies** | The shipped build cannot rely on any `https://` script or stylesheet reference. Both packaging scripts described in §4 hard-fail if they find one. |

Firepit consumes exactly **one artifact**: a flat `.html` file, typically uploaded to (or read from) a SharePoint document library that backs the web part.

---

## 4. Forge — the Packaging Tool

**"Forge"** is the internal name for the packaging step that turns the multi-file `index.html` + `assets/` source tree into the single self-contained HTML file Firepit needs. Two scripts exist in this repository for this purpose — they are **not interchangeable**, and picking the wrong one was a real, documented mistake early in this project (`FIREPIT-DEVELOPER-GUIDE.md`, Postmortem #8).

### 4.1 `scripts/build-sharepoint-package.js` — **the preferred script**

This is the correct tool for shipping to Firepit/SharePoint. Run it with:

```bash
node "scripts/build-sharepoint-package.js" "FS packages/PULSE-v1.0.0.html"
```

What it does, mechanically:

1. Reads `index.html`.
2. Inlines every `<link rel="stylesheet">` as a `<style>` block, rewriting any `url(...)` references inside that CSS (fonts, background images) into `data:` URIs.
3. Inlines every `<script src="...">` as a `<script>` block, escaping any literal `</script` inside the source so it can't terminate the tag early.
4. Replaces known image references inside JS with `data:` URIs: `aewttr-seal.png`, `aewttr-seal-template.png`, `ttsd-seal-template.png`, `navair-seal-template.png`, `navair-wordmark-template.png`. SVG variants (`ttsd-seal.svg`, `navair-seal.svg`) are only replaced if the file exists on disk — an `fs.existsSync()` guard prevents a crash when those assets are absent.
5. Wraps PptxGenJS with `wrapUmdAsBrowserGlobal()` — sets `var define = undefined` and `var module = undefined` around the UMD bundle so SharePoint's AMD `define()` cannot intercept PptxGenJS's module registration, which would otherwise leave `window.PptxGenJS` unset.
6. Strips the `?devcache=N` query string that `index.html` uses for local dev-only cache-busting — this must never survive into the shipped file (see the cache-busting gotcha in §3 and §9).
7. Builds a `WFC-MANIFEST` — a JSON object (version, project name, ISO generation timestamp, and a hash-fingerprinted list of every inlined file) — base64-encodes it, and prepends it to the output as an HTML comment. This is provenance metadata for humans/tooling; Firepit itself does not parse it.
8. **Hard-fails** (throws) if it finds any `https://` script or stylesheet reference, rather than silently skipping it.
9. Writes the final single-file HTML to the path given as the second CLI argument.

Output lands in `FS packages/`. This is the file you upload to SharePoint.

### 4.2 `scripts/build-forge.js` — secondary / offline-wrapper script

```bash
node "scripts/build-forge.js" "<path to a real Forge-tool template HTML file>" "dist/Forge.html"
```

This produces a **different, heavier artifact**: it inlines CSS/JS the same way, but then base64-encodes the entire resulting child page and injects it into a `const CHILD_HTML_B64 = "...";` placeholder inside a real Forge-tool-exported template, ultimately rendering the child app inside an `<iframe srcdoc>` sandbox at runtime. It also writes a `-readable.html` twin with the child HTML left as a literal (non-encoded) string, purely for debugging.

> **This is not the format Firepit expects for a direct SharePoint ship.** It exists for offline/portable distribution scenarios. Use it only if you specifically need a Forge-format artifact for a reason other than "put this in Firepit."

### 4.3 Why two scripts exist at all

The project initially used the offline/iframe-wrapper approach for what was actually a Firepit hosting target, and it "behaved noticeably worse" there — extra layers of encoding/decoding, a heavier payload, and an iframe sandbox that didn't match how Firepit actually runs pages. `build-sharepoint-package.js` was written afterward to reproduce the real Forge tool's flat, no-wrapper output format directly, without depending on having the actual Forge tool on hand. The reference file at `releases/reference-packages/PULSE - Only Secure in FS Sharepoint-current.html` is a known-good sample of genuine Forge-tool output, kept specifically to validate that `build-sharepoint-package.js`'s output matches the real platform's expected format.

---

## 5. The Deployment Pipeline

```
 1. Edit source
    assets/js/*.js, assets/css/style.css, index.html, vendor/*
        │
 2. Rebuild
    node "scripts/build-sharepoint-package.js" "FS packages/<name>.html"
        │
 3. Ship
    Upload the resulting FS packages/<name>.html to the SharePoint
    document library backing the Firepit web part
        │
 4. Verify
    Open the real Firepit web part (not just a local browser tab)
    and confirm the change works end-to-end
```

Rebuild **after every change, before every upload** — there is no watch mode or auto-deploy. A few practical notes:

- **Naming convention**: release filenames use semantic versioning in `MAJOR.MINOR.PATCH` form (for example, `PULSE-v1.0.0.html`). Increment MINOR for backward-compatible features and PATCH for backward-compatible fixes; use the `WFC-MANIFEST` `generated` timestamp for the build time rather than encoding a calendar date in the filename.
- **Rollback**: there is no automated build history/archive. `releases/reference-packages/PULSE - Only Secure in FS Sharepoint-current.html` functions as the last-known-good reference copy. Keeping a prior build around before shipping a risky change is a manual discipline, not something the tooling does for you.
- **Verification is not optional.** A local dev server or a plain browser tab cannot fully substitute for opening the real Firepit web part — some failure modes (sandbox restrictions, cache-busting breakage, SharePoint-only globals like `_spPageContextInfo`) only appear there.

### 5.1 Local development loop

Before packaging, you can run the unbundled source directly:

```bash
python3 -m http.server 8743
```

(matches the checked-in `.claude/launch.json` config) and open `index.html` in a browser. Outside a real SharePoint site, the app auto-detects this and falls back to **local mode** — an empty, non-persisted in-memory store — so you can exercise the UI without a live tenant. `index.html` uses a `?devcache=N` query parameter on asset tags for local-only cache-busting during development; `build-sharepoint-package.js` strips this automatically, but never add it back manually to a shipped file (§3, §9).

---

## 6. SharePoint Integration

SharePoint Lists **are** the database. There is no separate backend. All reads/writes are plain `fetch()` calls to SharePoint's native REST API (`/_api/web/...`), authenticated implicitly by the browser's existing SharePoint session (`credentials: "same-origin"`). Microsoft Graph is used only as an optional, best-effort extra for tenant-wide people search — if it's unreachable or misconfigured, the app falls back to native SharePoint People Picker/Search with no loss of core functionality.

### 6.1 Site Detection & Authentication

`getSiteUrl()` (`assets/js/sharepoint-adapter.js`) resolves the current SharePoint site in this fallback order:

1. `window._spPageContextInfo.webAbsoluteUrl` — SharePoint's own injected context global. Most reliable when present.
2. `APP_CONFIG.manualSharePointSiteUrl` (`assets/js/app-config.js`) — checked deliberately *before* the origin-guess below, because a partially-populated Firepit context can otherwise produce a wrong-origin guess that fails with an opaque error.
3. `window.location.origin + window._spPageContextInfo.webServerRelativeUrl` — an "unverified guess," used only when `_spPageContextInfo` is partially present.
4. Otherwise, returns `""` (undetected), and the app falls back to local mode.

The current user is resolved from **two sources and merged** (`bootSharePointMode`, `assets/js/app.js`): a `GET /_api/web/currentuser` REST call, and whatever `window._spPageContextInfo` already exposes. REST wins on any conflict. If the REST call fails outright, boot fails immediately to a minimal, unconfigured store rather than partially rendering with bad identity data.

**Role resolution**: after identity is established, the adapter looks up a matching row in the `PULSE App Roles` list by email → login name → SharePoint user ID → display name → Title, in that priority order. A SharePoint **site admin** (`IsSiteAdmin`) can still reach setup/recovery screens even with no matching role row, specifically so first-time setup can never lock everyone out — but ordinary app permissions are **not** derived from `IsSiteAdmin` (see §9).

### 6.2 The List Schema

Defined once, centrally, in `SHAREPOINT_SCHEMA` (`assets/js/sharepoint-schema.js`). Every list is created with the display-name prefix `"PULSE "` (`APP_CONFIG.listPrefix`) so they group together in Site Contents (SharePoint lists cannot be organized into folders).

| List | Backs feature | Key columns |
|---|---|---|
| PULSE App Roles | Users & permissions | `UserEmail`, `Role`, `IsActive`, `SharePointUserId`, `HideFromMeetings` |
| PULSE Projects | Project portfolio | `ProjectCode`, `Rag`, `HistoryJson`, `RisksJson`, `PeopleJson`, `BoardsJson`, `PortfoliosJson` |
| PULSE Action Items | Tracker tasks + Board/checklist items (one list, discriminated) | `ProjectId`, `Source` (Tracker/Checklist), `SubtasksJson`, `SortOrder` |
| PULSE Meetings | Weekly Meeting sessions (global or per-project) | `ProjectId` (blank = global), `Notes`, `ActivityJson`, `SessionStatus` |
| PULSE Rocks | Quarterly Rocks | `OwnerId`, `RockStatus`, `CheckupsJson` |
| PULSE Travel Requests / PULSE Travel Debriefs | Travel workflow | `RequestCode` (e.g. `TR-0041`), joined by that business key |
| PULSE Document Review | Document approval workflow | `ReviewColumn`, `ReviewersJson`, `RevisionHistoryJson` |
| PULSE RAG Config | RAG thresholds (single row) | `OverdueAmber`, `OverdueRed`, `BehindAmber`, `BehindRed`, `CompletionAmberBelow`, `CompletionRedBelow` |
| PULSE AI Config | AI review integration settings (single row) | `ApiKey` *(admin-only screen — treat as sensitive)* |
| PULSE Audit Log | Activity Log | `Action`, `Area`, `DetailJson` |
| PULSE Notifications | Outbound Teams/email queue for Power Automate | `ToEmails`, `BodyHtml`, `DeliveryStatus` |
| PULSE Location Config / PULSE Notification Config | Admin single-row settings | — |
| PULSE Decisions | Reserved — schema exists, not yet wired into the UI | `ProjectId`, `DecisionText`, `DecidedBy` |

Nested or array-shaped data (subtasks, activity feeds, reviewer lists, board column definitions) is deliberately **not** modeled as extra lists. It's serialized JSON stored in a `*Json`-suffixed Note column and decoded/re-encoded on load/save — see `HistoryJson`, `SubtasksJson`, etc. above.

### 6.3 The Repo Pattern (`sharepoint-repo.js`)

**Load model**: read-all-then-cache, not real-time. On boot, every list is read in parallel (`Promise.all`) into one in-memory `db` object that every page module reads from. There is no live subscription or webhook. Freshness comes from:

- A full load on boot.
- A short-lived `sessionStorage` cache so a page reload can render instantly, with a background re-fetch behind it.
- Polling-style background refresh on a timer, on tab focus, and opportunistically on navigation — but skipped entirely while a save is in flight, a modal is open, or a write landed in roughly the last 8 seconds, specifically to avoid clobbering in-progress edits.

**`SP_LISTS`** is the single place that maps an app "kind" (`project`, `actionItem`, `travelRequest`, …) to its SharePoint list display name. Each kind has a matched pair of mapper functions — e.g. `projectToSpItem` / `spItemToProject` — that translate between the app's in-memory object shape and the SharePoint list-item shape. Loaded objects carry hidden bookkeeping fields (`_spId`, `_projectCode`, `_source`, `_sortOrder`) that page code should never need to touch directly.

**`Repo.save(kind, obj, extra)`**:
1. If not running in SharePoint mode, delegates straight to the local-mode store save (`aewttrSaveStore()`) — this is the entire "local vs. SharePoint mode" branch.
2. Stamps routing context from `extra` (e.g. which project/column a task belongs to).
3. **Debounces per-object at 500ms** — rapid repeated saves on the same object (e.g. dragging a Gantt bar) collapse into one write of the final state.
4. **Chains writes per object**, so a create always finishes and stamps `_spId` before a later update on the same object fires, preventing duplicate-create races.
5. Issues the actual `POST` (create) or `MERGE` (update) call, decided by whether `_spId` is already present.

**`Repo.remove(kind, obj)`** cancels any pending debounce and deletes the item; deleting a project also cascades to delete its child tracker/checklist items so they don't reappear as orphans.

**Conflict handling**: because every record is its own list item, two users editing *different* records never collide. Writes use `IF-MATCH: "*"` (last-write-wins per item) — there is no field-level merge or optimistic-concurrency check on the list-item path.

### 6.4 Setup & Provisioning (`sharepoint-schema.js`)

**Run SharePoint Setup** (Admin → SharePoint Setup) calls `runSharePointSetup(siteUrl)`, which:

1. Creates any missing list from `SHAREPOINT_SCHEMA`, then ensures every defined field/column exists on it.
2. Seeds the single-row RAG Config and AI Config lists if empty.
3. Syncs SharePoint site users/groups into `PULSE App Roles` (deactivating — not deleting — role records for users no longer on the site).
4. Repairs any column left `Required` from an older schema version that no longer matches the current field list, so a stale Required column can't silently block every future write.

The whole routine is **idempotent** — safe to re-run at any time; every step is a "create/repair only what's missing" check.

**Column creation is the fragile part.** Fields are created via CAML XML through the `createfieldasxml` endpoint with `Options: 28`. This specific bitmask matters:

| Bit | Meaning |
|---|---|
| `4` | Add to all content types |
| `8` | **Honor the internal-name hint** — omitting this is what causes SharePoint to invent internal names like `field_7` instead of the name you specified |
| `16` | Add to the list's default view |

> **CAUTION** — Never pass `Options: 0` when creating a field. This is documented at length in `SHAREPOINT-LISTS-GUIDE.md` as the single most damaging mistake possible here: it produces mangled internal column names that silently break every future read/write against that column.

**Check Setup** (`getSetupStatus`) reports, per list, missing fields, type mismatches, and "suspicious" (mangled `field_N`-style) columns left over from a bad prior run. Suspicious columns are *detected*, not auto-repaired — they have to be removed by hand in SharePoint's List Settings.

### 6.5 Adding or Modifying a SharePoint-Backed Field

Following the pattern the codebase already uses, adding a new field (e.g. `RiskLevel` on `PULSE Projects`) touches four places:

1. **`sharepoint-schema.js`** — add the field definition (name, type, choices if applicable) to that list's entry in `SHAREPOINT_SCHEMA`.
2. **`sharepoint-repo.js`** — add it to *both* directions of the relevant mapper pair: e.g. `RiskLevel: proj.riskLevel || ""` in `projectToSpItem`, and `riskLevel: item.RiskLevel || ""` in `spItemToProject`.
3. **The relevant page module** — read/write `proj.riskLevel` like any other in-memory field. No direct SharePoint call is needed; `Repo.save("project", proj)` picks up the new mapper output automatically.
4. **Run SharePoint Setup** once against the live site so the column actually exists before the mapper writes to it.

> **CAUTION** — If step 4 is skipped, nothing errors. `filterItemPayloadForExistingFields()` in the adapter silently drops any payload key that doesn't match a real, existing column. The save appears to succeed; the value simply never persists. This is the most common cause of "I saved it but it's gone" bugs — always run Setup after a schema change.

Other rules that matter here: DateTime fields must be sent as `${value}T00:00:00Z` or `null` — never `""`. Choice-type fields reject any value not in their defined choice list with no partial fill-in, so adding a new status option to a dropdown in the UI requires a matching schema change and a re-run of Setup, or the new value will silently fail to save. Prefer a new or existing `*Json` Note column over a new list for any nested/array data.

---

## 7. Application Architecture

### 7.1 The `window.AEWTTR` Global

Defined at the top of `app.js`:

```js
window.AEWTTR = { db: null, state: {}, debugLog: [] };
```

- **`db`** — the entire in-memory "database" (see §7.4). Replaced wholesale on a SharePoint sync, not mutated field-by-field at the top level.
- **`state`** — transient, non-persisted UI state (active filters, active sub-tab, wizard step), namespaced per feature.
- Runtime-attached properties include `mode` (`"sharepoint"` or `"local"`), `siteUrl`, `currentSpUser`, `bootComplete`/`bootFailed`, `setupStatus`, and cache/refresh bookkeeping.

### 7.2 Routing

`PAGE_RENDERERS` is a plain object populated as a side effect of each `assets/js/pages/*.js` file loading — e.g. `PAGE_RENDERERS.dashboard = function(parts){...}`. Navigation is **hash-only** (`location.hash = "#/route/sub/parts"`) — not `pushState` — because the app can be embedded inside a SharePoint `.aspx` page/web part where History API navigation doesn't behave reliably. `navigate(path)` writes the hash; `renderPage()` (hooked to `hashchange`/`popstate`) parses it, looks up the matching renderer (falling back to the dashboard), gates `admin`/`logs` behind an access check, and wraps the whole render in a try/catch so one broken page shows an inline error panel instead of crashing the app.

### 7.2a Overview Architecture

The Overview module (`pages/overview.js`) has two independently-rendered views, selected by a My/Team tab in the page UI:

- **`renderMyOverview()`** — computes personal stats via `getMyData()`, renders a year-in-review hero with six metric cards (Tasks Completed, Docs Reviewed, Docs Signed, Meetings, Travel Requests, Projects), a My Tasks table with filter pills (All/Open/No Date/Completed), and a My Projects spotlight list. Tracks filter state in `window.AEWTTR.state.overviewMyTasksFilter`.
- **`renderTeamOverview(projects)`** — computes portfolio-level counts via `buildTeamMetrics(projects)`, renders a Team Snapshot panel (six headline numbers) followed by four inner tabs (Portfolio/Workload/Resources/Operations) driven by `window.AEWTTR.state.teamOverviewTab`. Blocked Tasks count is highlighted red when non-zero.

The Project Meeting tab (`renderProjectMeetingApp`, `pages/weekly.js`) receives the current project object and sets `scope = { type: "project", project: proj }`, which routes `meetingData()` and `meetingProjects()` to return only that project's tasks and roster rather than the full portfolio.

### 7.3 The App Shell — `renderShell()`

Renders the persistent chrome once per boot: the top nav bar (built from a `NAV_ITEMS` list, with Admin hidden unless the current user can access it), help-mode toggle, theme toggle, notification bell, and the user badge. Nav items can show an alert badge when something in that module needs attention. Below that sits the page-header row (title/subtitle/actions) and the `#page-content` mount point every page renderer writes into.

### 7.4 The Data Model (`data.js`)

`buildEmptyStore(currentUser)` defines the canonical shape of "the database": `projects`, `projectExtra` (status/history/risks/handoff notes, keyed by project id), `projectPeople`, `ganttTasks` (keyed by project id), `travelRequests`, `debriefs`, `docs` (keyed by review-column name), `checklistBoards` / `checklistTasks`, `tickets`, `rocks`, `weeklyMeeting` (roster/rocks/sessions/currentSession/meetingStatus/projectMeetings), `members`, `ragConfig`, `aiConfig`, `locationConfig`, and `user`.

`normalizeStoreShape(db)` runs on every load and is the defensive migration layer: it fills in any top-level key missing from an older store shape, back-fills newer project fields, normalizes Gantt tasks/subtasks, dedupes weekly-meeting rocks/sessions, and derives `meetingStatus` from session state. This is what lets the schema evolve over time without a formal migration step — new fields just default sensibly the first time an older record is loaded.

### 7.5 State Changes & Re-rendering

There is no virtual DOM — every page renderer builds an HTML string and sets `innerHTML`. After a local mutation, the calling page notifies the shell (`notifyLocalDataChanged`), which triggers a **soft re-render of the current route** (re-invoking that page's renderer while preserving scroll position) plus a refresh of the nav/notification badges. In SharePoint mode, writes go through `Repo.save`/`Repo.remove` (§6.3); `audit-log.js` monkey-patches those functions (and `navigate`) so almost every create/update/delete/navigation is captured in the Activity Log automatically, without individual pages having to call the logger by hand.

### 7.6 Boot Sequence

1. Show a boot screen with a progress indicator and an expandable diagnostic log.
2. Call `sharePointAdapter.detectSharePointSite()`. If a site is found → `bootSharePointMode()`; otherwise → local fallback mode with an empty store.
3. `bootSharePointMode()` resolves the current user, ensures a few self-healing lists exist (App Roles, AI Config, Document Review), resolves the user's role, and hydrates `db` — from a short-lived cache if available, otherwise a full load.
4. Render the shell, navigate to the current hash (or the dashboard), mark boot complete, log a Login audit entry, and start background refresh.
5. A top-level `try/catch` around the entire sequence shows a plain "PULSE failed to start" screen with the raw error if anything above throws, rather than a blank page.

### 7.6 Reporting & PPTX Export

The Reporting tab and its associated PowerPoint export are implemented across two files:

**`assets/js/pages/projects.js` — `drawProjectReporting(body, proj)`**

This function renders the entire Reporting tab UI each time the tab is opened. It builds:

- A topbar with a milestone-count checkbox, **Download .pptx**, and **Save to SharePoint** buttons.
- A live 4-quadrant **slide preview** (`renderSlidePreview()`) — a `div.rep-sld-mock` container with `aspect-ratio: 4/3` matching the 10 × 7.5-inch PPTX slide. The quadrant grid is `grid-template-columns: 47.3% 1.5px 1fr; grid-template-rows: 54% 1.5px 1fr` — the divider positions mirror the exact PPTX layout constants.
- A **Tech Bullets** editing card (`renderTechBullets()`) — inline add/edit/delete for `cfg.techBullets`, the string array stored in `projectExtra[pid].reportConfig.techBullets`.
- A **Risk Status** section — three badge + notes rows for Schedule / Budget / Technical, stored as `cfg.slideContent.riskRows[{cat, rag, notes}]`.
- A **Milestone Tracker** table — one row per task, with an **In Report** checkbox and date inputs for the six MC event types.
- A **Timeline Preview** — a mini horizontal Gantt showing in-report tasks.

**`buildGanttHtml()`** (inside `drawProjectReporting`'s closure) generates the HTML string for the mini Gantt chart shown in the slide preview's bottom-right quadrant. It:

1. Filters tasks to those with `inReport: true`.
2. Computes the month-aligned date range across all milestone events.
3. Builds a header row with year groups, then a second row with single-letter month abbreviations (`J F M A M J J A S O N D`).
4. Renders one bar row per task, placing milestone symbols at their percentage-offset positions.
5. Inserts a `div.rep-sld-gantt-today` absolutely positioned at the today percentage (`color: #f97316` — orange).

**`assets/js/project-pptx-export.js` — `exportProjectStatusPptx(proj)`**

Top-level export entry point. Calls `buildStatusSlide(pptx, proj)` which assembles the full 4-quadrant 10 × 7.5-inch slide. Key sub-functions:

| Function | What it draws |
|---|---|
| `drawMcMilestoneChart(slide, shapes, proj, x, y, w, h)` | Milestone Gantt in the bottom-right quadrant. Dual-row axis: year row (fill `#DCE7F3`, bold) above month-letter row (fill `#EEF4F9`, 5.5pt). Today-line: solid orange `#F97316`, width 1.5pt. Milestone symbols placed as text shapes at percentage offsets. |
| `riskSummaryRows(risks, proj)` | Returns the 3-row risk table content, checking `reportConfig.slideContent.riskRows` overrides first before computing from the live risk register. |
| `drawTechBullets(slide, proj, x, y, w)` | Renders `reportConfig.techBullets` as a bulleted text box in the top-right quadrant. |

**PPTX Layout constants (`L` object):**

| Key | Value (inches) | Role |
|---|---|---|
| `vDiv` | x: 4.731, y: 0.942, h: 6.05 | Vertical center divider |
| `hDiv` | x: 0.083, y: 4.053, w: 9.65 | Horizontal center divider |
| `photo` | x: 0.28, y: 1.12, w: 4.2, h: 2.7 | Photos area (top-left) |
| `techPanel` | x: 4.91, y: 1.08 | Tech bullets origin (top-right) |
| `riskTable` | x: 4.91, y: 2.28 | Risk table origin (top-right) |
| `descBody` | x: 0.15, y: 4.38, w: 4.4, h: 2.65 | Project description (bottom-left) |
| `milesBody` | x: 4.85, y: 4.38, w: 5.0, h: 2.65 | Milestones Gantt (bottom-right) |

**MC (Milestone Category) array** — defined in both `projects.js` and `project-pptx-export.js` and must be kept in sync if new milestone types are added:

| Key | Label | Symbol | Hex |
|---|---|---|---|
| `contractAwarded` | Contract Award | ☆ | `#2563EB` |
| `fat` | First Article Test | △ | `#6366F1` |
| `sat` | Site Acceptance Test | △ | `#3B82F6` |
| `add` | ADD | △ | `#F97316` |
| `fielding` | Fielding | △ | `#8B5CF6` |
| `complete` | Complete | ★ | `#DC2626` |

**PptxGenJS AMD conflict** — PptxGenJS ships as a UMD bundle. On SharePoint pages where an AMD `define()` is already present, the UMD detection code registers JSZip via `define()` instead of `window.JSZip`, which causes `JSZip is not defined` at slide-generation time. The build script's `wrapUmdAsBrowserGlobal()` wrapper (§4.1, step 5) prevents this by shadowing `define` and `module` with `undefined` inside the bundle's scope.

### 7.7 Supporting Subsystems

- **`audit-log.js`** — the Activity Log's write path. Caps the in-memory log, persists to SharePoint (or `localStorage` in local mode), and auto-instruments `Repo.save`/`Repo.remove`/`navigate`.
- **`notify.js`** — builds notifications in three parallel formats (HTML email, Teams markdown, Teams Adaptive Card) from one payload, respecting each user's per-area and per-channel preferences (see User Guide §10). Every feature module sends through one shared entry point rather than building its own message.

---

## 8. How to Update the Code

1. **Understand the editing priority order** (from `AI-HANDOFF.md`): keep SharePoint REST calls working → keep role resolution working → keep setup/schema checks working → keep the Logs/diagnostics screens usable → *then* touch UI/features. If you're not sure whether a change is safe, protecting these in order is the tie-breaker.
2. **Route all SharePoint calls through `sharepoint-adapter.js`.** Don't scatter new raw `fetch()` calls elsewhere in the app if the adapter already has (or could reasonably gain) a method for the operation you need.
3. **Make the change locally first** (§5.1) — run the unbundled source with a local static server and confirm the change behaves before packaging.
4. **Never make the boot handler `async`-block on a network call before first render.** A slow or hung SharePoint call (common on constrained networks) would otherwise leave the page blank indefinitely with no feedback to the user.
5. **Never swallow errors in a broad `try/catch` without surfacing them somewhere** (the boot log, a toast, the Activity Log). Every historical "my data just disappeared" bug traced back to a throw that a catch-all silently absorbed.
6. **`await` on `Repo.save()` really does mean "written."** Code that depends on a record's `_spId` existing must run after the `await`, not assume a fire-and-forget save has already landed.
7. **Rebuild before every ship**: `node "scripts/build-sharepoint-package.js" "FS packages/<name>.html"` (§4.1, §5).
8. **Verify inside a real Firepit web part**, not just a local browser tab, before considering the change done (§5).
9. For SharePoint-specific issues, diagnose in this order before assuming an architecture change is needed: does `/_api/web/currentuser` succeed (Admin → SharePoint Setup diagnostics)? Does the detected site URL look right? Does the Activity Log show live rows? Is a "missing" field actually missing from the schema, or just hidden from the SharePoint list UI?

---

## 9. Known Gotchas

| Symptom | Root cause | Where it's guarded against |
|---|---|---|
| Blank page, zero console errors | `?v=N`-style cache-busting query strings on a script/link tag — SharePoint's static file serving mishandles them | Never add cache-busting query strings to shipped asset tags; the packaging script strips the dev-only `?devcache=N` automatically |
| "I saved it but it's gone" | A field was written to the app object but the matching SharePoint column doesn't exist yet — the payload key is silently dropped | Always run **Run SharePoint Setup** after any schema change (§6.5) |
| Column shows as `field_7` / list only shows Title | A field was created with `Options: 0` instead of `28` | Always use `Options: 28` when creating fields (§6.4) |
| "Connection: Error" stuck in Admin diagnostics even though everything is working | A single transient failure (digest race, a momentary 429) got stored and never cleared | The adapter clears the last-error state on every *successful* request, not just on demand |
| A choice field silently fails to save a value the UI clearly offers | The Choice column's defined option list is out of sync with the schema/UI | Adding a new status/option requires a schema update in `sharepoint-schema.js` **and** a re-run of Setup |
| A background refresh clobbers an in-progress edit | Polling refresh landed mid-edit | Background refresh is gated off while a modal is open, a save is pending, or a write landed within the last ~8 seconds |
| Two rapid saves on the same object create two SharePoint items instead of one create + one update | A second `Repo.save()` fired before the first create's `_spId` was stamped | Per-object write chaining in `Repo.save` (§6.3) ensures a create always resolves before a later update on the same object |
| Tenant-wide people search returns nothing even though the person exists | Wrong Graph API cloud endpoint for the tenant (e.g. a DoD IL4/IL5 `*.sharepoint-mil.us` site needs `dod-graph.microsoft.us`, not the commercial/GCC-High endpoint) | Check `APP_CONFIG.graphApiBase` in `app-config.js` against the tenant's actual cloud; native SharePoint People Picker/Search still works even if Graph is misconfigured |
| A file uploads but is corrupted (especially `.docx` or images) | A text-encoding upload path was used for binary content | Binary uploads must use raw bytes, never a `TextEncoder`-based text path |
| Roles behave unpredictably across environments | App permissions were derived from `IsSiteAdmin` instead of the `PULSE App Roles` list | `IsSiteAdmin` is a SharePoint permission concept, not an app role — never gate app features on it directly (site admins get *setup/recovery* access as a safety net only) |

---

## 10. File Map

| File | Role |
|---|---|
| `index.html` | Unbundled dev entry point |
| `assets/js/app-config.js` | `APP_CONFIG` — site URL override, list prefix, Graph endpoint, boot diagnostics |
| `assets/js/app.js` | Shell, router, boot sequence, background refresh |
| `assets/js/data.js` | Store shape, `normalizeStoreShape` migration layer |
| `assets/js/audit-log.js` | Activity Log write path |
| `assets/js/notify.js` | Notification/email/Teams message building |
| `assets/js/sharepoint-adapter.js` | Raw SharePoint REST/Graph plumbing, CRUD, diagnostics |
| `assets/js/sharepoint-schema.js` | `SHAREPOINT_SCHEMA`, setup/provisioning, field repair |
| `assets/js/sharepoint-repo.js` | `SP_LISTS`, object↔list-item mappers, `Repo.save`/`Repo.remove` |
| `assets/js/pages/overview.js` | `renderMyOverview()` — personal year-in-review dashboard; `renderTeamOverview()` — leadership snapshot with Portfolio/Workload/Resources/Operations tabs |
| `assets/js/pages/weekly.js` | Weekly Meeting engine; `renderProjectMeetingApp(body, proj)` — project-scoped meeting view |
| `assets/js/pages/projects.js` | Project workspace pages including `drawProjectReporting()` (Reporting tab), `renderSlidePreview()`, `buildGanttHtml()`, `renderTechBullets()` |
| `assets/js/project-pptx-export.js` | PPTX export — `exportProjectStatusPptx(proj)`, `buildStatusSlide()`, `drawMcMilestoneChart()`, `riskSummaryRows()` |
| `assets/js/pages/*.js` | One file per feature (see User Guide for the full feature list) |
| `scripts/build-sharepoint-package.js` | **Preferred** packaging script — produces the Firepit ship artifact |
| `scripts/build-forge.js` | Secondary/offline-wrapper packaging script |
| `FS packages/` | Shipped SharePoint artifacts (output of the preferred script) |
| `releases/reference-packages/` | Known-good reference copy of real Forge-tool output |
| `dist/` | Output of `build-forge.js` (not the preferred ship path) |
