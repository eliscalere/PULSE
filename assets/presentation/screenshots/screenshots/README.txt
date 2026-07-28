PULSE — Presentation Screenshots
=====================================
Captured: 2026-07-09
Source: PULSE/ (static HTML/CSS/JS single-page app)

HOW THESE WERE CAPTURED
------------------------
PULSE has no build step — it is plain HTML/CSS/JS designed to run inside
a SharePoint site (via a Firepit/SPFx web part) using SharePoint lists as its
backend. It was run locally with the project's existing dev launch config
(python3 -m http.server, see PULSE/.claude/launch.json) and driven with
Playwright (Chromium) at a fixed 1600x1000 viewport, screenshotted directly
to PNG (no browser chrome).

Outside of a real SharePoint site, the app automatically falls back to a
"local mode" with a completely empty database (this is real app behavior,
in assets/js/data.js). To produce screens that actually demonstrate the
app's features instead of empty tables, two things were done, both using
the app's own real code paths — nothing was mocked up or drawn by hand:

  1. The app ships an actual built-in sample-data generator, "Demo Portfolio
     Seed" (assets/js/pages/admin.js, seedDemoPortfolio()), intended for
     exactly this kind of demo. It was invoked to create 5 sample projects
     with tracker tasks. A few tasks were then marked complete (a normal
     user action) so the portfolio shows a realistic mix of Red/Amber/Green
     status from the app's real, computed RAG logic (computeProjectRag() in
     assets/js/app.js) rather than every project defaulting to the same
     color.
  2. Five generic, non-identifying demo members (e.g. "Program Manager",
     "Systems Engineer", placeholder @example.mil addresses — not real
     people or domains) were added so the Users/roles screen, which is
     normally populated live from a SharePoint site, wasn't empty.

The local browser session was also flagged as an app admin (window.AEWTTR.
db.user.isAdmin) purely to reach admin-gated screens for this screenshot
session — in real deployments, admin status comes from SharePoint group
membership, not a local flag.

No screens were invented, and no placeholder/lorem-ipsum content was used.
Every screenshot below is a real, currently-implemented page of the app,
rendered by its real rendering code.

SCREENSHOTS
------------

01-dashboard.png
  Main Dashboard — quick-launch tiles for every module (Overview, Projects,
  Weekly Meeting, Travel, Document Review, Admin) plus a live "My Tasks" and
  "My Rocks" sidebar pulled from real project/tracker data.
  Status: Fully implemented.

02-executive-overview.png
  Overview > Executive Summary — portfolio leadership view: Red/Amber/Green
  rollup counts, mission snapshot (active projects, docs in review, travel
  requests, support tickets), and a "Needs Attention" queue. Demonstrates
  cross-module data aggregation.
  Status: Fully implemented.

03-project-portfolio.png
  Projects > All Projects — the project portfolio table (ID, name, priority,
  computed RAG status, last updated), with My Projects / All Projects and
  priority filters.
  Status: Fully implemented.

04-project-detail.png
  Project workspace, Home tab (PULSE-01 "Radar Integration") — shows the
  per-project side navigation (Home, People, Documents, Meeting, Tracker,
  Boards, Settings), open task/subitem/people/board counts, and project
  detail fields. Demonstrates the app's modular per-project structure.
  Status: Fully implemented.

05-project-tracker.png
  Project workspace, Tracker tab — the task list for a project (owner,
  start/end dates, health status, progress), the same tracker view used
  inside the Weekly Meeting flow.
  Status: Fully implemented.

06-weekly-meeting.png
  Weekly Meeting, active session, Project View — a live cross-project
  meeting queue showing every project's RAG status and open/at-risk task
  counts down the left rail, with the selected project's tracker and a
  Rocks panel on the right.
  Status: Fully implemented.

07-travel-request.png
  Travel > New Request — step 1 of the 5-step travel request submission
  wizard (Type > Trip > Purpose > Budget > Review), showing the four
  supported request types (Standard, Leave, Contractor Travel, Engineering
  TTSD/CL export). Captured mid-flow, not submitted, to show the real
  workflow without inventing a fake approved/denied request record.
  Status: Fully implemented.

08-document-review.png
  Document Review — the review-status board (Not Started / In Review /
  Concurrence / Final-Signed / Archive) with scope and status filters and a
  "Submit Document" action. Shown in its real empty state, since no
  documents were submitted during this local demo session — no fabricated
  document records were added.
  Status: Fully implemented (shown with no data).

09-sharepoint-setup.png
  Admin > SharePoint Setup — the backend connection/configuration screen:
  detected SharePoint site, connection status, list/column setup and
  permission-check tools, and a diagnostics panel. This capture was taken
  outside a live SharePoint site, so it correctly displays "Local fallback
  (dev)" mode rather than an active connection — it demonstrates the
  screen and its SharePoint-integration tooling, not a live data sync.
  Status: Fully implemented; shown outside a live SharePoint session.

10-users-roles.png
  Admin > Users — the site users / app roles table. In production this
  list is synced live from SharePoint site membership; for this local
  capture it is populated with generic placeholder members (no real names,
  emails, or identifiers) purely to demonstrate the table layout and role
  badges.
  Status: Fully implemented; populated with placeholder demo data.

NOT INCLUDED
------------
- Notifications panel: exists in the app shell (bell icon, top right) but
  had no content to display in this local/no-SharePoint session, so it was
  left out rather than showing a misleadingly empty state as a dedicated
  screenshot.
- A separate "navigation shell" screenshot was not produced as its own
  file — the persistent top navigation bar (Dashboard / Overview / Projects
  / Weekly Meeting / Travel / Document Review / Admin) is already visible
  at the top of every screenshot above.
- Teams chat-log image: excluded per instructions.
- Checklists/Boards, Tickets, and AI Review Config screens exist in the
  codebase but were not captured — they were empty in this demo session and
  populating them would have required fabricating content beyond the
  built-in demo-seed feature.
