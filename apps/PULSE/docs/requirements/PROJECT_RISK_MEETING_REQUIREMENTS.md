# PULSE Project, Risk, Meeting, and Export Requirements

This document captures the next data and workflow requirements for AEWTTR-PULSE.
It aligns the app with project tracker needs where related efforts can be
grouped, briefed, reviewed in meetings, and exported without losing project
context.

## Project Management

Project records must include the full tracker metadata needed for execution:
project manager, engineer, ISSO, range POC, contract, contractor, task order,
funding type, funding year, funded status, change request requirement,
configuration end item, location, start date, completion date, scope, and
objectives.

Task labels must be distinct:

- At Risk means the work can continue, but schedule, cost, technical, staffing,
  or external indicators show likely trouble without intervention.
- Off Track means the work is already outside the approved plan or expected
  timeline.
- Blocked means the work cannot continue because of a specific dependency,
  access issue, decision, input, funding action, document, or problem.
- On Hold means work is intentionally paused by direction or timing, not
  because of an unresolved blocker.
- Complete means the work is finished and no further action is expected.

The hierarchy must support projects, subprojects, tasks, and subtasks. A real
subproject is stored as its own project record with a parent project link. It
must be able to carry its own funding, dates, task order, location, risks,
documents, and team members.

## Portfolio Views

Users need grouped project views that can pivot by portfolio, program,
configuration end item, product, location, contract, task order, and funding
year. This supports related efforts such as TRES, DIADS, and Fury, where
multiple projects may share a broader operational relationship but differ in
funding, schedules, and personnel.

Related project navigation should let users move from one grouped project or
subproject to another without returning to the full project list.

## Risk Management

Risks are first-class records, not just notes on a project. Each risk must
include risk name, description, owner, likelihood, impact, category, mitigation
plan, response strategy, due date, and status.

Supported response strategies are Accept, Avoid, Mitigate, Transfer, and
Escalate.

Risk rating is calculated from likelihood and impact:

- Red for high combined likelihood and impact.
- Amber for moderate combined likelihood and impact.
- Green for low combined likelihood and impact.

Risk views must show highest risks first, red and amber risks, risks by
project, risks by portfolio, and overall project risk posture. Meeting pages
must allow risk review and updates during the meeting.

## Meetings

Meetings must support an agenda before and during the meeting. Team members
must be able to add future agenda items.

Meeting pages must include agenda, projects, tasks, risks, around the room,
general notes, decisions, action items, and attendance.

Meeting history must preserve notes, decisions, task changes, risk changes,
project status changes, action items, and attendees.

Recurring meeting series must group and search related meetings such as Weekly
PM Sync, Technical Exchange, and Government Biweekly Meeting. Starting a new
meeting should show previous notes, unfinished agenda items, and carry-forward
actions from the prior meeting in that series.

## Project Documents

Each project must have a direct link to its main SharePoint folder. Each
project must also expose a project-specific document review view showing review
status, revision, deadline, pending reviewers, signature status, owner, and
related file information.

## Project Images And PowerPoint Export

Projects must support multiple images. Each image should allow a caption, main
image selection, export selection, and ordering.

Project slide export must use the existing PowerPoint quad template and include
project title, scope, objectives, schedule, status, accomplishments, upcoming
work, funding, key risks, key personnel, and selected images.

Before export, the app must show missing required information so the user can
complete the record before generating briefing slides.
