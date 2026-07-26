# Legacy to SPFx Migration Map

## Source reference

The original prototype lives in `../PULSE` and currently demonstrates these user-facing areas:

- dashboard
- projects and project workspace
- weekly meeting
- document review
- admin
- travel, checklists, tickets, and related operational workflows

## Target mapping

- `dashboard.js` -> `src/pages/Dashboard`
- `projects.js` -> `src/pages/Projects` and `src/pages/Project`
- `weekly.js` -> `src/pages/Meetings`
- `docreview.js` -> `src/pages/Documents`
- `admin.js` -> `src/pages/Admin`
- checklist-style boards -> `src/pages/Tasks`
- travel and timeline flows -> `src/pages/Calendar`
- reporting and export flows -> `src/pages/Reports`

## Recommended next implementation slices

1. Port project portfolio and project detail first.
2. Port weekly meeting orchestration second.
3. Port document review and task boards third.
4. Add travel and reporting workflows after the core schema stabilizes.
