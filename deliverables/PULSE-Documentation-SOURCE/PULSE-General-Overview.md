# PULSE General Overview

PULSE is the SharePoint-hosted operational workspace for connecting project delivery, updates, logistics, status, and execution. It creates one current operating picture so teams can identify the next action, its owner, and its supporting record without assembling the answer from disconnected lists, spreadsheets, and email.

## Who uses PULSE

PULSE supports project leads, delivery teams, reviewers, travelers, meeting participants, support staff, and administrators. Access is established from the existing SharePoint session and the active PULSE role record. Administrators additionally maintain users, configuration, diagnostics, activity records, and SharePoint setup.

## Operational problem addressed

Project information commonly spans trackers, document libraries, meeting notes, travel requests, review correspondence, and support messages. PULSE joins these operational activities through a project-centered workspace, a shared navigation model, notification preferences, and an audit trail. It is designed to make current ownership, timing, conditions, decisions, and dependencies easier to find and act on.

## Current capabilities

- Dashboard and Overview views for personal and team-level situational awareness.
- Project, Portfolio, Program, and End Item Configuration organization.
- Project workspaces for people, project documents, tracker, Gantt, boards, reporting, meetings, notes, tickets, and settings.
- Weekly Meeting, travel requests and debriefs, formal Document Review, support tickets, notifications, administration, and activity records.
- Project and group PowerPoint reporting generated in the browser, with SharePoint upload when the app is running in SharePoint.

## SharePoint use

PULSE is a browser application served from SharePoint. The existing SharePoint session provides identity. The application uses same-site SharePoint REST calls for PULSE Lists and project document-library operations; it does not provide a separate PULSE sign-in or application database. The source schema defines the required lists and columns, and the **SharePoint Setup** tool validates or provisions them.

Project **Documents** are project-associated document-library files and references. **Document Review** is a separate formal workflow that records reviewers, revisions, approvals, final-pack handling, signatures, and workflow state. Do not treat a file stored under a project as evidence of formal review completion.

## Why not disconnected tools

PULSE preserves the relationship between a project, its people, work items, risks, files, meetings, travel, review activity, and reporting. It supplies consistent labels, current-state routing, linked operational records, role-aware controls, notifications, and activity logging that a loose collection of independently maintained lists cannot reliably provide.

## Read next

- [User Guide](PULSE-User-Guide.md) for navigation and role-aware operating instructions.
- [Technical Reference](PULSE-Technical-Reference.md) for source, configuration, SharePoint integration, packaging, and maintenance.
- [Standard Operating Procedures](PULSE-Standard-Operating-Procedures.md) for repeatable operating workflows and evidence.
- [Source Audit](Source-Audit.md) for verified boundaries and items requiring live-site confirmation.
