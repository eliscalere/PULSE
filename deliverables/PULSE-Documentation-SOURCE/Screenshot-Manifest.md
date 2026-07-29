# Screenshot Manifest

The current capture set was produced on 2026-07-29 from a clean local browser instance by `apps/PULSE/scripts/capture-documentation-screenshots.cjs`, at a 1600 x 1000 viewport and 2x device scale. The capture identity is the neutral display name **Program Lead**, supported by **Program Manager**, **Systems Engineer**, and **Operations Lead**. Every record is synthetic and every address is under `example.mil`; no real project, personal, or organizational data is present.

Unlike the previous set, these captures are taken against seeded records rather than an empty workspace, so each screen shows populated state instead of an empty placeholder.

Captures are written to `apps/PULSE/docs/current/screenshots/` at 2x for document use. A 1600px-wide WebP copy is published to `apps/pulse-documentation/public/screenshots/` and appears in documentation library document 11, *Interface Reference*, in both the PDF and the reader. WebP keeps the delivered package loadable in SharePoint: the same set as PNG was 2.9 MB against 668 KB.

Local capture verifies visible routes, labels, and rendered state only. SharePoint identity, list persistence, permissions, and notification delivery require live-site verification — see [Feature Verification Matrix](Feature-Verification-Matrix.md). These screenshots are intentionally not evidence of a completed SharePoint transaction.

| Filename | Screen | Route | Demonstrates |
|---|---|---|---|
| `01-dashboard.webp` | Dashboard | `#/dashboard` | Counts, quick access, assigned work across projects |
| `02-overview-portfolio.webp` | Overview — portfolio | `#/overview` | Team and portfolio roll-up for status, workload, approvals |
| `03-projects-workspaces.webp` | All workspaces | `#/projects` | Project list with health and workspace entry points |
| `04-project-workspace.webp` | Workspace home | `#/projects/<code>` | Project landing view and workspace section rail |
| `05-project-tracker.webp` | Tracker | `#/projects/<code>/tracker` | Tasks and milestones with owner, dates, health; timeline and risk tabs |
| `06-project-checklist.webp` | Checklist | `#/projects/<code>/checklist` | Column-based checklist for repeatable delivery steps |
| `07-project-documents.webp` | Documents | `#/projects/<code>/documents` | Project document surface backed by the SharePoint library |
| `08-project-people.webp` | People | `#/projects/<code>/people` | Assigned members and roles |
| `09-weekly-meeting.webp` | Weekly meeting | `#/weekly` | Live session: attendance, minutes, project updates, around-the-room |
| `10-travel-request-form.webp` | Travel request form | `#/travel/submit` | TDY, conference, training, and leave intake |
| `11-travel-my-travel.webp` | My travel | `#/travel/mytravel` | A traveler's requests by state, with concurrence and charge-object status |
| `12-travel-calendar.webp` | Travel calendar | `#/travel/calendar` | Team travel and events on a shared calendar |
| `13-travel-debrief.webp` | Travel debrief | `#/travel/debrief` | Post-trip debrief against an approved request |
| `14-document-review.webp` | Document review | `#/docreview` | Review packages in concurrence and signature |
| `15-tickets.webp` | Tickets | `#/tickets` | Blockers, bugs, access needs, questions tracked to resolution |
| `16-admin.webp` | Admin | `#/admin` | SharePoint setup, users and roles, configuration, diagnostics |
| `17-notification-settings.webp` | Notification settings | `#/notification-settings` | Per-user areas, tone, and channel preferences |
| `18-logs.webp` | Activity log | `#/logs` | Audit record by actor, area, and time |

## Regenerating the set

1. Serve the app on port 8743 with any static server rooted at `apps/PULSE`.
2. Run `node apps/PULSE/scripts/capture-documentation-screenshots.cjs`.
3. Downscale to 1600px wide and convert to WebP into `apps/pulse-documentation/public/screenshots/`:
   `for f in *.png; do cwebp -q 80 "$f" -o "${f%.png}.webp"; done`
4. Regenerate document 11 and rebuild the package:
   `node apps/pulse-documentation/scripts/build-source-document.mjs 11`

The script exits non-zero if a route's expected selector never appears, so a blank capture is not silently written as a success. Regenerating document 11 rewrites its PDF, its searchable text, and the figure manifest the reader uses, so all three stay in agreement; update the capture date in `documents/11-interface-reference.mjs` at the same time.
