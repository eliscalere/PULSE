# Screenshot Manifest

The current capture set was produced on 2026-07-29 from a clean local browser instance by `apps/PULSE/scripts/capture-documentation-screenshots.cjs`, at a 1600 x 1000 viewport and 2x device scale. The capture identity is the neutral display name **Program Lead**, supported by **Program Manager**, **Systems Engineer**, and **Operations Lead**. Every record is synthetic and every address is under `example.mil`; no real project, personal, or organizational data is present.

Unlike the previous set, these captures are taken against seeded records rather than an empty workspace, so each screen shows populated state instead of an empty placeholder.

Captures are written to `apps/PULSE/docs/current/screenshots/` at 2x for document use, and a 1600px-wide copy is published to `apps/pulse-documentation/public/screenshots/` for the documentation site's Interface reference view.

Local capture verifies visible routes, labels, and rendered state only. SharePoint identity, list persistence, permissions, and notification delivery require live-site verification — see [Feature Verification Matrix](Feature-Verification-Matrix.md). These screenshots are intentionally not evidence of a completed SharePoint transaction.

| Filename | Screen | Route | Demonstrates |
|---|---|---|---|
| `01-dashboard.png` | Dashboard | `#/dashboard` | Counts, quick access, assigned work across projects |
| `02-overview-portfolio.png` | Overview — portfolio | `#/overview` | Team and portfolio roll-up for status, workload, approvals |
| `03-projects-workspaces.png` | All workspaces | `#/projects` | Project list with health and workspace entry points |
| `04-project-workspace.png` | Workspace home | `#/projects/<code>` | Project landing view and workspace section rail |
| `05-project-tracker.png` | Tracker | `#/projects/<code>/tracker` | Tasks and milestones with owner, dates, health; timeline and risk tabs |
| `06-project-checklist.png` | Checklist | `#/projects/<code>/checklist` | Column-based checklist for repeatable delivery steps |
| `07-project-documents.png` | Documents | `#/projects/<code>/documents` | Project document surface backed by the SharePoint library |
| `08-project-people.png` | People | `#/projects/<code>/people` | Assigned members and roles |
| `09-weekly-meeting.png` | Weekly meeting | `#/weekly` | Live session: attendance, minutes, project updates, around-the-room |
| `10-travel-request-form.png` | Travel request form | `#/travel/submit` | TDY, conference, training, and leave intake |
| `11-travel-my-travel.png` | My travel | `#/travel/mytravel` | A traveler's requests by state, with concurrence and charge-object status |
| `12-travel-calendar.png` | Travel calendar | `#/travel/calendar` | Team travel and events on a shared calendar |
| `13-travel-debrief.png` | Travel debrief | `#/travel/debrief` | Post-trip debrief against an approved request |
| `14-document-review.png` | Document review | `#/docreview` | Review packages in concurrence and signature |
| `15-tickets.png` | Tickets | `#/tickets` | Blockers, bugs, access needs, questions tracked to resolution |
| `16-admin.png` | Admin | `#/admin` | SharePoint setup, users and roles, configuration, diagnostics |
| `17-notification-settings.png` | Notification settings | `#/notification-settings` | Per-user areas, tone, and channel preferences |
| `18-logs.png` | Activity log | `#/logs` | Audit record by actor, area, and time |

## Regenerating the set

1. Serve the app on port 8743 with any static server rooted at `apps/PULSE`.
2. Run `node apps/PULSE/scripts/capture-documentation-screenshots.cjs`.
3. Copy the output into `apps/pulse-documentation/public/screenshots/` downscaled to 1600px wide, then rebuild the documentation package.

The script exits non-zero if a route's expected selector never appears, so a blank capture is not silently written as a success. Update `SCREEN_CAPTURE_DATE` in `apps/pulse-documentation/app/page.tsx` whenever the set is regenerated.
