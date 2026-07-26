# PULSE Overview PPTX Data Contract

Use this contract when creating a reference PowerPoint for the **My Overview**
and **Team Overview** exports. Dates use ISO `YYYY-MM-DD` format. Counts are
whole numbers greater than or equal to zero.

Files:

- `samples/overview-pptx-data.schema.json` — machine-readable JSON Schema
- `samples/overview-pptx-sample-data.json` — filled examples for both views

## Shared presentation fields

| Field | Type | Required | Purpose |
|---|---|---:|---|
| `schemaVersion` | string | Yes | Contract version; use `1.0`. |
| `viewType` | enum | Yes | `myOverview` or `teamOverview`. |
| `title` | string | Yes | Presentation title. |
| `subtitle` | string | Yes | Scope and reporting date. |
| `generatedAt` | date-time | Yes | When the snapshot was generated. |
| `generatedBy` | string | Yes | Display name of the person generating it. |
| `reportingAsOf` | date | Yes | Date through which the data is current. |
| `summaryNarrative` | string array | No | Up to five concise executive summary bullets. |

## My Overview

The personal view should contain:

1. A cover/title slide.
2. Personal workload and activity metrics.
3. Assigned projects.
4. Active task detail.

### Metrics

| Field | Meaning |
|---|---|
| `totalProjects` | Projects where the user has an assigned role or task. |
| `totalTasks` | All tasks assigned to the user. |
| `activeTasks` | Assigned tasks not complete or closed. |
| `overdueTasks` | Active assigned tasks past their due date. |
| `completedTasks` | Assigned tasks completed during the selected reporting scope. |
| `documentsReviewed` | Documents reviewed or concurred on. |
| `documentsSigned` | Documents signed or approved. |
| `meetingsAttended` | Weekly meeting sessions attended. |
| `travelRequests` | Travel requests submitted. |

### Project fields

`projectId`, `projectName`, `team`, `role`, `status`, `myActiveTaskCount`, and
`dueDate`.

### Task fields

`taskId`, `title`, `projectId`, `projectName`, `owner`, `status`, `health`,
`startDate`, `dueDate`, `percentComplete`, `blockedReason`, and `nextAction`.

## Team Overview

The team view should contain:

1. A cover/title slide.
2. Portfolio metrics and status distribution.
3. Work grouped by end item.
4. Project inventory.
5. One project-detail slide per project when space permits.
6. Team workload.

### Metrics

| Field | Meaning |
|---|---|
| `totalProjects` | All projects visible in the selected scope. |
| `activeProjects` | Projects with active work. |
| `activeTasks` | Tasks not complete or closed. |
| `overdueTasks` | Active tasks past their due date. |
| `blockedTasks` | Active tasks in a blocked state. |
| `openRisks` | Risks not closed, resolved, or accepted. |
| `teamSize` | People with active assigned work. |
| `documentsInReview` | Document-review records not complete. |
| `pendingTravel` | Travel requests awaiting action or approval. |

### End-item fields

`name`, `projectCount`, `activeTaskCount`, `openRiskCount`, and `teams`.

### Project fields

`projectId`, `projectName`, `endItem`, `projectManager`, `team`, `fundingType`,
`fiscalYear`, `taskOrder`, `lifecycleStatus`, `dueDate`, `derivedStatus`,
`activeTaskCount`, `overdueTaskCount`, `openRiskCount`, `tasks`, and `risks`.

### Risk fields

`riskId`, `title`, `description`, `owner`, `likelihood`, `impact`, `status`,
`mitigation`, and `dueDate`.

### Workload fields

`person`, `activeTaskCount`, `blockedTaskCount`, `projectCount`, and
`nextDueDate`.

## Reference PPTX notes

- Use 4:3 slides.
- Keep the same section order listed above.
- Use real-looking but sanitized values.
- Use no more than 10 project rows or 12 task rows on a single slide; continue
  onto another slide when necessary.
- Keep status, health, likelihood, and impact values exactly as written in the
  sample data so the export can map them consistently.
