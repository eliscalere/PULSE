# PULSE-IPT — General Overview

<img src="../assets/images/aewttr-seal.png" alt="AEWTTR Seal" width="80" align="right">

| | |
|---|---|
| **Document** | PULSE-IPT General Overview |
| **Audience** | Leadership, stakeholders, and anyone getting oriented before reading the User Guide or Technical Reference |
| **Companion docs** | [User Guide](01-USER-GUIDE.md) · [Technical Reference](02-TECHNICAL-REFERENCE.md) |

---

## 1. What PULSE-IPT Is

PULSE-IPT is an internal, SharePoint-backed operations workspace built for AEWTTR IPT (Integrated Product/Project Team) personnel. It gives a project team one place to see status, manage tasks, run the weekly meeting, move travel requests and document reviews through approval, and track support issues — instead of stitching that picture together across email, spreadsheets, and a dozen disconnected SharePoint lists.

The design premise, carried through every screen, is simple: **a project is the front door.** Open a project and every relevant tool — its schedule, its people, its files, its meeting, its own checklists — is one click away, in context, instead of behaving like isolated apps that happen to share a login.

## 2. Who Uses It

AEWTTR team members across Program Management, Engineering, Cyber, Contracts, Support, and leadership-adjacent review roles, during day-to-day program execution — checking status, moving work forward, and finding the right file or form without hunting across disconnected screens.

## 3. What It Actually Does

| Capability | In practice |
|---|---|
| **Portfolio visibility** | An executive rollup of every project's health (Red/Amber/Green), completion, and blockers — computed automatically from real task data, not self-reported |
| **Personal Overview** | A year-in-review dashboard for each user: tasks completed, documents reviewed and signed, meetings attended, travel requests, and assigned projects — all filtered to that individual |
| **Team Overview** | Leadership-level snapshot: active projects, open tasks, blocked tasks, risks, docs in review, and pending travel — with drill-down tabs for Portfolio, Workload, Resources, and Operations |
| **Project workspaces** | Each project gets its own home page, roster, document library, tracker/Gantt schedule, custom boards, settings, and a project-scoped meeting view |
| **Status Reporting & PPTX Export** | A Reporting tab inside every project: a live 4-quadrant slide preview (photos · tech status + risk table · project description · milestone Gantt), editable tech bullets and risk RAG overrides, per-task milestone date tracking, and one-click PowerPoint export with dual-row month/year axis and orange today-line |
| **Weekly Meeting** | Runs the standing status meeting live — project-by-project or person-by-person — with shared notes and quarterly Rocks tracked in the same place |
| **Travel** | A full request-to-approval-to-debrief workflow, plus a shared travel calendar |
| **Document Review** | A formal, reviewer-driven approval workflow (Not Started → In Review → Concurrence → Final/Signed), separate from ordinary file storage |
| **Support Tickets** | A lightweight internal helpdesk for blockers, bugs, and access requests |
| **Admin controls** | Configurable RAG thresholds, location lists, role management, and a full activity audit trail |

Every one of these is a real, working feature today — not a mockup or a planned capability. The [User Guide](01-USER-GUIDE.md) documents each one screen-by-screen.

## 4. How It's Built and Hosted

PULSE-IPT is deliberately **infrastructure-light**. It's a browser-only application — no application server, no database server, and no separate login system to stand up or maintain. It runs inside the team's existing SharePoint site, using SharePoint Lists as its database and the user's own SharePoint session as authentication. There is nothing for a user to install, and nothing beyond standard SharePoint site administration for an admin to operate.

This matters operationally: the app inherits SharePoint's existing security boundary, backup posture, and access model rather than introducing a new system that has to be independently secured and approved. The tradeoff is packaging discipline — because the hosting environment (an internal tool called Firepit) only accepts a single self-contained file, every change goes through a build step before it ships. That process, and the reasoning behind it, is covered in full in the [Technical Reference](02-TECHNICAL-REFERENCE.md).

## 5. What Makes It Different from a Shared SharePoint List

A plain SharePoint list can hold data. PULSE-IPT adds the layer on top that makes the data usable day-to-day:

- Status (RAG) is **computed** from real task health and completion, not typed in by hand and left stale.
- Every module is **cross-linked** — a task assignee, a travel request, a document reviewer, and a support ticket reporter all resolve back to the same person and project.
- Every meaningful action is **audited** automatically, with no extra effort from the user.
- The interface is built for **fast execution during working hours** — explicit controls, visible state, no hidden gestures — not for browsing at leisure.

## 6. Where to Go Next

- **New user, want to learn the screens?** → [User Guide](01-USER-GUIDE.md)
- **Developer, need to change or ship code?** → [Technical Reference](02-TECHNICAL-REFERENCE.md)
- **Need the raw historical build notes and postmortems?** → `AI-HANDOFF.md`, `FIREPIT-DEVELOPER-GUIDE.md`, `FS-FORGE-STEPS.md`, `SHAREPOINT-LISTS-GUIDE.md` at the repository root
