# PULSE — User Guide

<img src="../assets/images/aewttr-seal.png" alt="AEWTTR Seal" width="80" align="right">

| | |
|---|---|
| **Document** | PULSE User Guide |
| **Audience** | AEWTTR IPT personnel (all roles) |
| **Scope** | Every screen and feature currently implemented in the application |
| **Prepared for** | AEWTTR IPT |

---

## How to read this manual

This guide is organized the same way the application's top navigation bar is organized — left to right, top to bottom. Each section covers one module: what it's for, how to reach it, what's on the screen, and how to use it. Screenshots are pulled directly from a running instance of the application, so what you see in this document is what you will see on screen.

Callouts are used throughout:

> **NOTE** — Additional context that helps you understand *why* something works the way it does.

> **CAUTION** — An action that is easy to get wrong or has a consequence you should know about before proceeding.

> **ADMIN ONLY** — This screen or control is only visible to users holding the Admin (or a specifically named) role.

---

## Table of Contents

1. [Getting Started](#1-getting-started)
2. [Dashboard](#2-dashboard)
3. [Overview](#3-overview)
4. [Projects](#4-projects)
5. [Reporting](#5-reporting)
6. [Weekly Meeting](#6-weekly-meeting)
7. [Travel](#7-travel)
8. [Document Review](#8-document-review)
9. [Support Tickets](#9-support-tickets)
10. [Admin Tools](#10-admin-tools)
11. [Notification Settings](#11-notification-settings)
12. [Roles & Permissions Reference](#12-roles--permissions-reference)
13. [Glossary](#13-glossary)

---

## 1. Getting Started

PULSE runs **inside your SharePoint site** — there is nothing to install. You open it the same way you'd open any other page or web part on your team's SharePoint site, using your existing SharePoint/CAC login. There is no separate PULSE username or password.

### 1.1 The App Shell

Every screen in PULSE shares the same top bar:

![Dashboard and app shell](screenshots/01-dashboard.png)

| Element | Location | What it does |
|---|---|---|
| PULSE logo | Top-left | Click to jump to the Dashboard from anywhere |
| Main navigation | Top-center | Dashboard, Overview, Projects, Weekly Meeting, Travel, Document Review, Admin (Admin only shows if you hold an admin-level role) |
| Help toggle (`?`) | Top-right | Turns on inline hint text next to fields and controls across the app |
| Theme toggle (moon/sun) | Top-right | Switches between light and dark mode. Your preference is remembered. |
| Notification bell | Top-right | Shows an unread badge and a dropdown of recent notifications |
| Your name / role badge | Top-right | Click it to open your [Notification Settings](#10-notification-settings) |

> **NOTE** — A small red badge can appear on a nav item (e.g. Travel, Document Review, Admin) when something in that module needs your attention — a pending approval, a document waiting on your review decision, etc.

### 1.2 Local / Offline Fallback

If PULSE is ever opened somewhere it cannot detect a SharePoint site (for example, a developer testing on a laptop), it automatically starts in a **local fallback mode** with an empty, non-persisted workspace instead of failing to load. In normal day-to-day use inside SharePoint, you will never see this — it exists purely so the app never shows a blank white screen.

---

## 2. Dashboard

**Reach it:** Nav bar → **Dashboard**, or click the PULSE logo.

The Dashboard is your landing page and quick-launch pad.

![Dashboard](screenshots/01-dashboard.png)

- **Quick-launch tiles** — one tile per module (Overview, Projects, Weekly Meeting, Travel, Document Review, Admin). Click **Open →** to jump straight in.
- **My Tasks** (right rail) — every open tracker task, across every project, that is assigned to you. Click a task to jump to that project's Tracker.
- **My Rocks** (right rail) — your active quarterly Rocks from the Weekly Meeting module, with a shortcut to open the meeting.

---

## 3. Overview

**Reach it:** Nav bar → **Overview**.

Overview has two views, switched by the **My / Team** tab at the top of the page.

### 3.1 My Overview

![My Overview](screenshots/ss-my-overview.png)

The personal year-in-review view, scoped entirely to you. Switch to it by clicking the **My** tab.

**Hero section** — shows your name, total open task count, number of assigned projects, and the current date. Quick-action buttons link directly to My Projects, Document Review, and Weekly Meeting.

**Year metrics** — six cards summarizing your activity since January 1 of the current year:

| Card | Shows |
|---|---|
| Tasks Completed | Tasks you completed this calendar year |
| Docs Reviewed | Document Review decisions you made this year |
| Docs Signed | Documents reaching Final/Signed where you were a reviewer |
| Meetings | Weekly meeting sessions you participated in |
| Travel Requests | Travel requests you submitted this year |
| Projects | Total projects you are currently assigned to |

**Needs Attention** — appears only when applicable; highlights documents awaiting your review decision or pending travel items.

**My Tasks** — a full table of your tasks, filtered with pills (All / Open / No Date / Completed). Click any row to jump to that project.

**My Projects** — one row per project you are assigned to. Click to open.

### 3.2 Team Overview

![Team Overview](screenshots/ss-team-overview.png)

The leadership-level portfolio snapshot. Switch to it by clicking the **Team** tab.

**Team Snapshot panel** — six headline metrics updated in real time:

| Metric | What it counts |
|---|---|
| Active Projects | Projects with status Active |
| Open Tasks | Tracker tasks not yet marked Done across all projects |
| Blocked Tasks | Tasks explicitly marked Blocked — shown in red when non-zero |
| Open Risks | Risks recorded in any project that are still open |
| Docs in Review | Documents in the formal review pipeline |
| Pending Travel | Travel requests awaiting approval |

**Inner tabs** — four tabs drill down into the data:

| Tab | Shows |
|---|---|
| Portfolio | RAG status rollup for every active project with completion % and blocked task count |
| Workload | Open task count per person across the portfolio |
| Resources | Team roster with roles, functional areas, and project assignment count |
| Operations | Pending travel requests and documents currently in review |

Use the **All Projects** and **Team Members** quick-action buttons to jump to those modules.

> **NOTE** — Team Overview is read-only. Make changes inside individual project workspaces; the Team Overview reflects those changes automatically on the next view render.

---

## 4. Projects

**Reach it:** Nav bar → **Projects**.

### 4.1 Portfolio View

![Project Portfolio](screenshots/05-project-portfolio.png)

The full list of projects, toggle between **My Projects** (projects you're attached to) and **All Projects**. Filter by priority or search by name. Columns: ID, Name, Priority, RAG status, Last Updated. Click **+ New Project** to stand up a new project, or click any row to open it.

### 4.2 Inside a Project

Every project opens into its own workspace with a left-hand rail of tabs. This is the "front door" of PULSE — from here you can reach every tool relevant to that specific project.

#### Home

![Project Home](screenshots/06-project-home.png)

The project's cover page: name, RAG badge, priority, lifecycle stage, key dates, and quick counts (open tasks, subitems, people, boards). The **Project details** panel shows role assignments (PM, Engineer, ISSO, Range POC, Contractor), tools in use, stakeholders, funding, and change-request status. The **Boards** panel gives one-click shortcuts into that project's Tracker, People, and Meeting.

#### People

![Project People](screenshots/07-project-people.png)

The project's roster. Click **Add Person** to attach an internal team member (picked from the org directory), a named contractor, or a company. This roster automatically becomes the participant list for the project's own Weekly Meeting sessions.

#### Documents

![Project Documents](screenshots/08-project-documents.png)

A file browser (folders, upload, search, drag-and-drop) scoped to that one project's SharePoint document library. This is raw file storage — a shared drive for the project.

> **NOTE** — This is **not** the same thing as the top-level [Document Review](#7-document-review) module. Documents here have no approval workflow attached; Document Review is where a document goes through a formal reviewer sign-off process.

#### Tracker

![Project Tracker](screenshots/09-project-tracker.png)

The project's task list / Gantt chart — the working schedule. Switch between **Main table** and **Gantt** view. Each task has an owner, start/end dates, a health flag (On Track / At Risk / Off Track), and can carry subtasks with their own assignee, dates, and related documents. Click **+ New task** to add one, drag a Gantt bar to reschedule it, or drag a row to reorder.

> **NOTE** — Task health here directly drives the project's RAG status. See [RAG Configuration](#93-rag-configuration) for how the thresholds work.

#### Boards

![Project Boards](screenshots/10-project-boards.png)

A separate, ad-hoc Kanban/checklist tool for workflows that don't belong on the formal Tracker — intake queues, review checklists, anything with its own set of custom stages. Click **+ New board** to create one as either a **table** board (spreadsheet-style rows with expandable subitems) or a **kanban** board (drag cards across columns). Columns are fully customizable — add, rename, recolor, or remove them, and add custom fields that show as tags on each card.

#### Settings

![Project Settings](screenshots/11-project-settings.png)

The project's configuration form. Covers:

- **Identity** — name, description
- **Status & Schedule** — priority, lifecycle, dates, change-request flag
- **Locations** — range/location assignments
- **Role Assignments** — PM, Engineer, ISSO, Range POC, Contractor
- **Program & Funding** — multi-select tag pickers for:
  - **Fiscal Year** — one or more applicable fiscal years (e.g. FY25, FY26)
  - **Funding Type** — RDTE, PROC, O&M, etc.
  - **Funding Status** — Funded, Unfunded, TBD, etc.
  - **Task Order** — the contract task order(s) associated with this project
  - **Program / End Item** — the parent program or configuration end item
- **Handoff Notes** — freeform transition notes

Click **Save Settings** to apply changes. **Delete project** is in the top-right corner.

> **CAUTION** — Deleting a project also deletes its Tracker tasks and checklist items. There is no undo.

#### Notes

A freeform project journal/notes panel — separate from Weekly Meeting notes — for anything worth writing down that doesn't belong in a specific task.

#### Meeting

A task review view **scoped to this project only**. It shows the current project's tasks — not the full portfolio — organized by assignee, with status, priority, and due date columns. Use this during the weekly meeting to walk through a specific project's open work in real-time.

> **NOTE** — The project Meeting tab is intentionally limited to the current project. For a cross-project or around-the-room view, use the top-level [Weekly Meeting](#5-weekly-meeting) module.

---

## 5. Reporting

**Reach it:** Open a project → **Reporting** tab.

The Reporting tab is where you compose, review, and export the project's official DIADS-style status slide. Everything on this tab feeds directly into the PowerPoint export — what you see in the live preview is what you get in the .pptx file.

### 5.1 Slide Preview

At the top of the page is a live, to-scale mockup of the 4-quadrant status slide:

| Quadrant | Contents |
|---|---|
| **Top-left** | Project photos (populated from the Photos tab) |
| **Top-right** | Technical Status bullet list + Schedule/Budget/Technical risk table |
| **Bottom-left** | Project description (from Project Settings) |
| **Bottom-right** | Milestone Gantt chart with month-letter axis and orange today-line |

The preview updates in real time as you edit any section below it. The slide header shows the project name, POC, and CUI markings; the footer shows NAVAIR and the As-of date.

### 5.2 Technical Status Bullets

Below the preview, the **Technical Status** card lets you manage the bullet points that appear in the top-right quadrant.

- Click **+ Add bullet** to add a new entry.
- Click an existing bullet text to edit it inline.
- Click **×** to delete a bullet.

Changes reflect in the preview immediately and save automatically.

### 5.3 Risk Status

The **Risk Status** section has three rows — Schedule, Budget, and Technical. For each:

1. Click the colored badge to cycle through **Green → Yellow → Red**.
2. Type a short comment in the notes field (e.g. "No change", "Funding TBD").

These values override the computed risk-register summary and are what appear in the slide's risk table. If you leave a row at its default, the computed value from the Risk Register is used instead.

### 5.4 Milestone Tracker

The **Milestone Tracker** table lists every task in the project's Tracker. Use it to:

1. Check **In Report** on the tasks you want to show on the slide's Gantt chart.
2. Enter dates for any combination of the six milestone event types:

| Symbol | Event | Color on slide |
|---|---|---|
| ☆ | Contract Awarded | Blue |
| △ | First Article Test (FAT) | Indigo |
| △ | Site Acceptance Test (SAT) | Sky blue |
| △ | ADD | Orange |
| △ | Fielding | Purple |
| ★ | Complete | Red |

Only tasks with **In Report** checked and at least one milestone date entered will appear on the Gantt.

> **NOTE** — When a task's status is set to **Complete** in the Tracker, its completion date auto-seeds the Complete (★) milestone field if it's blank.

### 5.5 Timeline Preview

A condensed timeline below the Milestone Tracker shows all in-report tasks as bars across the calendar year, with the same symbol markers. This is a quick sanity check before export — use it to confirm date alignment and sequencing look right.

### 5.6 Exporting to PowerPoint

Click **Download .pptx** (top-right of the Reporting tab) to generate and download the slide as a PowerPoint file. The export runs entirely in the browser — no server call is needed.

The generated slide is a 10 × 7.5-inch, 4:3 presentation with:

- Dual-row axis on the Gantt chart: year labels in the first row, single-letter month abbreviations (J F M A M J J A S O N D) below
- Orange vertical today-line on the Gantt
- Risk table matching the RAG overrides set in §5.3
- Tech bullets from §5.2
- Project description from Settings

To save the file to SharePoint instead (SharePoint mode only), click **Save to SharePoint** — the file is uploaded to the project's document library and a confirmation toast appears.

---

## 6. Weekly Meeting

**Reach it:** Nav bar → **Weekly Meeting** (portfolio-wide), or a project's **Meeting** tab (project only).

The Weekly Meeting module runs your standing status meeting live, in the room, with everyone's tasks pulled up automatically instead of a shared slide deck.

### 6.1 Starting a Meeting

A Meeting Admin starts the session with the **Start Meeting** button. Once live, a second row of tabs appears:

### 6.2 Project View

![Weekly Meeting — Project View](screenshots/12-weekly-meeting-project-view.png)

Every project listed down the left with its live RAG status and open/at-risk task counts. Click a project to pull up its Tracker in the main panel and work through it with the room. **Rocks** for the selected project/portfolio appear on the right.

### 6.3 Around the Room

![Weekly Meeting — Around the Room](screenshots/13-weekly-meeting-around-the-room.png)

The facilitation mode for going person-by-person instead of project-by-project. The left rail lists meeting participants (drawn from the roster); click a name and their tasks/Rocks come up one at a time — built for literally going "around the room."

### 6.4 Meeting Notes

![Weekly Meeting — Meeting Notes](screenshots/14-weekly-meeting-notes.png)

A shared, chat-style notes feed for the session. Type an update, decision, or follow-up and press **Post Note** (Enter posts, Shift+Enter adds a line break). Everyone in the meeting sees the same feed — nobody's note overwrites anyone else's.

### 6.5 History

![Weekly Meeting — History](screenshots/15-weekly-meeting-history.png)

Every past session, searchable by date, people, notes, or activity. Use this to answer "what did we say about this two weeks ago" without digging through email.

### 6.6 Rocks

Rocks are quarterly-goal-style objectives — a title, an owner, an optional linked project, a due date, and a status (On Track / Off Track / Done). Each Rock has a "checkups" drawer for posting dated progress notes without changing the Rock itself. Rocks can be tied to a specific project or kept portfolio-wide, and reassigning ownership sends the new owner a notification.

---

## 7. Travel

**Reach it:** Nav bar → **Travel**.

### 7.1 Submitting a Request

![Travel — New Request](screenshots/16-travel-new-request.png)

Click **New request** and pick a type:

| Type | For |
|---|---|
| **Standard** | Quick TDY, conference, or training request |
| **Leave** | Personal leave / time away from the project (shorter, single-step form) |
| **Contractor Travel** | Travel for contractors or external team members |

The wizard then walks you through **Trip → Purpose → Budget → Review**, with a step tracker across the top and a "jump back and edit" link on every section of the final Review step before you submit.

> **NOTE** — Editing a request that's already Approved switches to an expanded 5-step flow (Flights → Transport → Lodging → Details → Review) to capture final booking details.

### 7.2 Tracking Requests

![Travel — My Travel / All Travel](screenshots/17-travel-my-travel.png)

Toggle between **My Travel** (just yours) and **All Travel**. Filter by status: All, Upcoming, Pending, Pending Finance, Approved, Denied, Revoked, Cancelled. If you hold approval rights, Approve / Deny / Assign-C/O actions appear directly in the table — there's no separate Approvals screen to hunt for.

### 7.3 Debrief

After an approved (non-Leave) trip, file a debrief: systems and locations visited, trip classification, a technical summary, issues/follow-ups, and file attachments. Debriefs link back to the original request, and a **View Debrief** button appears on that request once filed.

### 7.4 Calendar

![Travel — Calendar](screenshots/18-travel-calendar.png)

A monthly calendar of all non-Denied/Cancelled/Revoked travel, color-coded by type (Leave, TDY, Conference, Training, Contractor, Other, Pending approval). Toggle contractor travel on/off with the checkbox, top-right.

---

## 8. Document Review

**Reach it:** Nav bar → **Document Review**.

This is the **formal document approval workflow** — distinct from a project's own Documents file browser (Section 4.2).

![Document Review Board](screenshots/19-document-review-board.png)

- **Board columns**: Not Started → In Review → Concurrence → Final / Signed, plus an Archive bucket.
- **Scope toggle**: My Documents vs. All Documents.
- **Status filter**: Active / Concurrence / Signed / Archive / All.
- Each card shows the linked project, submitter, current revision, and a reviewer-decision fraction (e.g. "2/3 approved").

Click **+ Submit Document** to start a new review: give it a title, type, project, deadline, upload the first revision, and add reviewers (you're added automatically as the submitter).

**How the review works:** each reviewer records a decision — Pending, Approved, or Requested Changes — with an optional note. Uploading a new revision resets every reviewer back to Pending, since the document has changed. A document with no outstanding review action for 7+ days auto-archives. Reviewers get reminder notifications on a scheduled sweep if their decision has been pending too long.

---

## 9. Support Tickets

**Reach it:** Nav bar link (or Overview → Operations Queue → **Open Tickets**).

![Support Tickets](screenshots/20-tickets.png)

A lightweight internal helpdesk. Filter by status (Open / In Progress / Resolved) or type (Blocker / Bug / Access / Platform / Question), and search by keyword. Click **+ New Ticket** to file one — optionally locked to a specific project. Opening a ticket shows its description, workaround notes, an escalation-path field, and a running updates timeline where you can post progress and change status.

> **NOTE** — Filing a new ticket notifies all admin users; status changes and updates notify the original reporter (you won't be notified of your own updates).

---

## 10. Admin Tools

**Reach it:** Nav bar → **Admin** (only visible if you hold an admin-level role).

![Admin launcher](screenshots/21-admin-launcher.png)

Six tools, each described below.

### 10.1 RAG Configuration

![RAG Configuration](screenshots/22-admin-rag-config.png)

Sets the thresholds that automatically compute every project's Red/Amber/Green status — no code changes needed to retune them:

| Group | Controls |
|---|---|
| Overdue Tasks | Amber if ≥ N overdue tasks, Red if ≥ N |
| Tasks Marked Behind | Amber/Red thresholds on tasks flagged At Risk or Off Track |
| Completion | Amber if completion % is below X, Red if below Y |

If a project trips any Red threshold, it shows Red regardless of the others ("Red wins"). A project with no tracker tasks yet keeps its seeded baseline color until tasks are added.

### 10.2 Locations

![Locations](screenshots/23-admin-locations.png)

Manages the list of ranges/locations selectable on every project's Settings tab. Add, rename, or remove entries; duplicate names are blocked and deletion asks for confirmation.

### 10.3 SharePoint Setup

![SharePoint Setup](screenshots/24-admin-sharepoint-setup.png)

Backend connection status and provisioning tools — see the [Technical Reference](02-TECHNICAL-REFERENCE.md) for what each button does under the hood. In everyday admin use: **Check Setup** tells you if any SharePoint lists/columns are missing, **Run SharePoint Setup** creates or repairs them, and the Diagnostics panel on the right helps answer "why am I not showing as Admin?" without opening a support ticket.

### 10.4 Users

![Users & App Roles](screenshots/25-admin-users-roles.png)

Every site user and their PULSE app role (Admin, Lead, Member, etc.), synced from SharePoint site membership.

### 10.5 Activity Log

![Activity Log](screenshots/26-admin-activity-log.png)

An append-only record of every create/update/delete/navigation action taken in the app — who, what, when, and where. Filter by area or action type, search by actor/summary/record, and expand **Details** on any row to see the raw change payload. **Export CSV** downloads the currently filtered rows (and is itself logged as an audit entry).

### 10.6 Notification Settings

Shortcut to the same screen described in [Section 11](#11-notification-settings) — included here because it's also a site-wide admin concern (channel overrides).

---

## 11. Notification Settings

**Reach it:** Click your name/role badge, top-right of any screen.

![Notification Settings](screenshots/27-notification-settings.png)

Self-service and **autosaving** — there's no Save button, changes take effect as you make them.

- **Delivery Channels** — turn Email and Teams notifications on or off independently. (A site-wide admin toggle can override this if disabled tenant-wide.)
- **Apps & Areas** — choose which parts of PULSE are allowed to notify you at all (Travel & Leave, Document Review, Projects, Weekly Meeting, Support Tickets, Admin, general PULSE notices).
- **Tone** — pick how PULSE "talks" to you in its messages (e.g. Robotic vs. Friendly). Only the greeting/sign-off changes — the facts in the message never change.

Use **Send Test Notification** to confirm delivery is working before you rely on it.

---

## 12. Roles & Permissions Reference

| Role | Can do |
|---|---|
| **Admin** | Everything, including all Admin Tools, RAG/Locations configuration, and SharePoint Setup |
| **Meeting Admin** | Start/end Weekly Meeting sessions, manage the roster and guests |
| **Finance Admin** | Travel finance-stage actions |
| **Manager / Lead** | Elevated visibility and approval actions within their scope |
| **Member** | Standard day-to-day access: projects, tasks, travel, documents, tickets |
| **Viewer** | Read-only |

> **NOTE** — Your role comes from your SharePoint site membership, synced automatically — there is nothing to request separately inside the app. If a screen looks locked and you believe it shouldn't be, ask your site's PULSE Admin to check **Admin → Users**.

---

## 13. Glossary

| Term | Meaning |
|---|---|
| **IPT** | Integrated Product/Project Team |
| **RAG** | Red / Amber / Green status indicator |
| **Rock** | A quarterly-goal-style objective tracked in Weekly Meeting |
| **Tracker** | A project's task list / Gantt schedule |
| **Board** | A custom Kanban or table workflow inside a project, separate from the Tracker |
| **Debrief** | The after-action report filed following an approved trip |
| **TDY** | Temporary Duty (a travel request type) |
| **Concurrence** | A Document Review status meaning the document has stakeholder sign-off pending |
