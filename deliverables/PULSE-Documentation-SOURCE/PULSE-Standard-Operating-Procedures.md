# PULSE Standard Operating Procedures

## SOP 1: Maintain a project and tracker

**Purpose:** keep project delivery records current. **Applies to:** authorized project members and leads. **When to use:** when creating a project, changing ownership, timing, health, scope, or work items.

**Prerequisites:** appropriate PULSE role, project access, and current source information.

1. Open **Projects**, locate or create the project, and confirm its identity.
2. In **Settings**, maintain lifecycle, technical health, ownership, dates, portfolios, program, and end-item configuration as applicable.
3. In **People**, maintain the project roster and responsibilities.
4. In **Tracker**, create or update work items with owner, start date, due date, workflow value, health, notes, and related documents.
5. Use **Blocked** only when a dependency prevents progress and state the dependency and recovery action. Use **On Hold** only with the reason and restart condition.
6. Use **Gantt** to confirm schedule placement. Use **Boards** only for the board workflow; do not assume board activity updates tracker items.
7. Confirm the save indication, refresh the affected record, and check the Activity Log when required.

**Expected result and verification:** the project summary and tracker show the updated owner, timing, condition, and work. Verify after a refresh in SharePoint mode. **Exceptions and escalation:** report failed saves, inaccessible roles, or schema errors to an administrator. **Records:** project, action-item, risk, and activity records. **Related procedures:** reporting, weekly meeting, and troubleshooting.

## SOP 2: Plan and report project work

**Purpose:** produce an accurate project or group report. **Applies to:** authorized project leads and reporting users. **When to use:** before status reviews or management briefings.

1. Review current project settings, tracker, Gantt, risks, meetings, and document-review records.
2. Open **Reporting** and verify technical bullets, milestones, risks, timeline, and export selections.
3. Generate the project or group PowerPoint output.
4. Inspect the file for current dates, names, condition statements, and intended audience before distribution.
5. If running in SharePoint mode, confirm the upload destination and file access; otherwise record the approved storage location.

**Verification:** generated content agrees with source records. **Caution:** generated output does not supersede source records. **Records:** report configuration, generated file, and release evidence.

## SOP 3: Run a weekly or project meeting

**Purpose:** convert meeting discussion into owned work and decisions. **Applies to:** meeting participants and facilitators. **When to use:** scheduled weekly sessions or project meetings.

1. Open **Weekly Meeting** or project **Meeting** and select the right scope.
2. Confirm date, roster, agenda, attendance, rocks, carry-forward items, and active project context.
3. Capture decisions separately from general notes.
4. Create or update a tracker item for each follow-up requiring an owner or timing.
5. Review action ownership and dates before closing the session.
6. Use meeting history to verify the session record and follow up on carry-forward items.

**Expected result:** decisions and action items remain traceable to the meeting. **Escalation:** report inability to save, an incorrect roster, or access problems. **Records:** meeting session, notes, decisions, action items, and activity record.

## SOP 4: Submit, approve, and debrief travel

**Purpose:** manage travel and leave records through their applicable workflow. **Applies to:** requesters, approvers, finance users, and travelers. **Prerequisites:** approved business need and required trip information.

1. Open **Travel** and select **New Request**.
2. Enter required request, traveler, destination, dates, purpose, project, estimate, and supporting information.
3. Submit and monitor the request in the applicable travel view.
4. Approvers use the role-restricted decision actions and provide required denial or approval information.
5. Finance users record required finance information where enabled.
6. After travel, each traveler creates the required debrief with systems or subjects, classification, summary, and follow-up.

**Verification:** request and debrief records show their current values after refresh. **Exceptions:** correct incomplete information before approval; escalate role or persistence faults. **Records:** travel request, approval history, debrief, notification, and activity record.

## SOP 5: Conduct formal Document Review

**Purpose:** maintain a controlled review, revision, final-pack, and signature record. **Applies to:** document owners, reviewers, signers, and administrators. **Prerequisites:** current file, related project, reviewers, and due date.

1. Open **Document Review** and select **New Review**.
2. Record the project, type, title, due date, revision label, current file, reviewers, and any required signers.
3. Reviewers record their decision and actionable comments in the formal review record.
4. When changes are requested, upload a new revision and run the review again; preserve prior revision history.
5. When all reviewers approve, follow the displayed formal-review path for review completion or final-pack signing.
6. Archive only through the workflow after required activity is complete.

**Verification:** reviewer decisions, revision history, and workflow value are visible on the review record. **Cautions:** project Documents are not formal review evidence; never overwrite a reviewed revision. **Records:** review record, revision files, comments, signature activity, notifications, and audit trail.

## SOP 6: Manage tickets and notifications

**Purpose:** create actionable support records and maintain notification preferences. **Applies to:** all users for their own tickets and preferences; administrators for platform configuration.

1. Use **Report issue** or **Tickets** to open a ticket.
2. State the issue, affected route, project or record identifier, expected result, actual result, priority, and approved screenshot.
3. Monitor the ticket workflow and add factual updates.
4. Use **Notification Settings** to manage your own enabled areas and channels.
5. Administrators review notification configuration only through authorized Admin tools.

**Verification:** ticket updates and preferences persist after refresh in SharePoint mode. **Records:** issue, notification, and audit record.

## SOP 7: Administer users, roles, configuration, logs, and setup

**Purpose:** maintain PULSE access and schema without uncontrolled changes. **Applies to:** PULSE administrators only.

1. Open **Admin** and confirm current-user and site diagnostics.
2. In **Users**, locate the approved person, associate them with PULSE, apply the least-privileged appropriate role, and retain required role metadata.
3. Deactivate a departing user record rather than removing history.
4. Use **SharePoint Setup** to validate or provision the defined lists and fields. Do not create ad hoc schema changes.
5. Maintain locations and application configuration only through the provided administrative workflows.
6. Use **Activity Log** and diagnostics to investigate operational problems.

**Verification:** role resolution, setup results, and audit evidence reflect the approved action. **Escalation:** obtain organizational authorization before broad access, schema, retention, or production-host changes. **Records:** role, configuration, setup result, and audit records.

## SOP 8: Export, troubleshoot, and preserve evidence

**Purpose:** safely produce exports and respond to an application fault. **Applies to:** all users for issue reporting; maintainers for technical response.

1. Before export, verify source records and required access.
2. Generate the output and verify content, filename, destination, and access.
3. For a fault, capture time, route, user role, record identifier, action, expected result, actual result, browser/device, and approved screenshot.
4. Check the application diagnostics and Activity Log when authorized.
5. Maintainers validate locally, then on a safe SharePoint site, then on the actual hosted page.
6. For a deployment regression, restore the prior known-good package, preserve evidence, and isolate the source change.

**Records:** issue report, diagnostics, log evidence, validation evidence, and rollback artifact. **Related procedures:** every SOP above.
