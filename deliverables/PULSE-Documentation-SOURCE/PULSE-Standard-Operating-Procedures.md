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

1. Open **Travel** and select **New Request**. Leave requests are filed from the same form; there is no separate leave route.
2. Enter required request, traveler, destination, dates, purpose, project, estimate, and supporting information.
3. Submit, then monitor the request under **Travel → Travel → My Travel**, using the state tabs (**Upcoming**, **Submitted**, **Withdrawn**, **Cancelled**, **Completed**) to confirm where it stands. The row reports request status, concurrence state, charge-object state, and any outstanding debrief.
4. Approvers use the role-restricted decision actions and provide required denial or approval information. Before approving, check **Travel → Calendar** for overlapping trips and team events that would leave a coverage gap.
5. Finance users record required finance information where enabled, including charge-object assignment.
6. After travel, each traveler creates the required debrief from **Travel → Debrief** with systems or subjects, classification, summary, and follow-up. File one debrief per traveler; a trip with an outstanding debrief stays flagged in **My Travel**.

**Verification:** request and debrief records show their current values after refresh, and the trip appears on the travel calendar for the approved dates. **Exceptions:** correct incomplete information before approval; escalate role or persistence faults. **Records:** travel request, approval history, charge-object assignment, concurrence, debrief, notification, and activity record.

This procedure is identical when performed from the standalone **Travel Request Forms**, **My Travel**, or **Travel Calendar** packages; those tools write the same records. See SOP 9.

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

1. Use **Report issue**, project **Tickets**, or the standalone Tickets package to open a ticket.
2. State the issue, affected route, project or record identifier, expected result, actual result, priority, and approved screenshot.
3. Monitor the ticket workflow and add factual updates on the ticket record itself, not by email, so the history stays complete. Ticket values are **Open**, **In Progress**, and **Resolved**.
4. Use **Notification Settings** to manage your own enabled areas and channels.
5. Administrators review notification configuration only through authorized Admin tools.

**Verification:** ticket updates and preferences persist after refresh in SharePoint mode. **Note:** tickets and issue reports share the **PULSE Issues** list, so a ticket opened in the standalone Tickets package is the same record the full application shows; do not open a second ticket for the same problem in the other surface. **Records:** issue, notification, and audit record.

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

## SOP 9: Publish and validate a focused tool page

**Purpose:** place a single-area PULSE package on a SharePoint page and confirm it is operating against the site rather than local browser state. **Applies to:** site owners and release owners. **Prerequisites:** the provisioned PULSE schema on the target site and the approved package artifact.

Available packages are **PULSE** (full application), **Travel Request Forms**, **My Travel**, **Travel Calendar**, **Tickets**, **PULSE Calendar**, **PULSE CODE**, and **PULSE Documentation**.

1. Build the artifacts with `node apps/PULSE/scripts/build-travel-packages.js`, which writes every package to `apps/PULSE/fs-packages/` and `releases/`. Retain the prior known-good artifact before replacing anything.
2. Upload the package for the intended area and add it to the SharePoint page through the organization's current hosted-page process.
3. Open the page and confirm the tool resolved the site: a record created in the tool must appear in the full PULSE application on the same site after refresh. A tool that shows only local behavior did not resolve the site and is not a system of record.
4. If site resolution fails, open the full PULSE page once on the same site, then reload the tool; the full application caches the resolved site address for the focused packages. Persistent failure is a ticket with the page address attached.
5. Confirm navigation is limited to the intended area and that no unintended area is reachable from the page.
6. Record which package version is published on which page, so a later regression can be traced to an artifact.

**Verification:** a controlled read and write for the tool's record type is visible in the full application, and the Activity Log shows the action under the expected actor. **Cautions:** a focused package carries the full script set and is configured only by `PULSE_PORT_CONFIG`; do not hand-trim scripts from a package to make it smaller, which has previously broken boot. **Records:** published artifact, page address, validation evidence, and activity record.

## SOP 10: Stand up a development site and validate at scale

**Purpose:** create a disposable SharePoint environment for PULSE development, establish what the application does under realistic data volumes, and evaluate whether Dataverse is a better store. **Applies to:** developers and maintainers. **Prerequisites:** authority to request a separate site collection, and agreement that no production records will be copied into it.

**Status: not yet performed.** PULSE currently has no development environment, no automated tests, and no measured scale evidence. Development runs against browser-local fallback state or against a site holding real records. Until this procedure is completed, PULSE should make no claim about the data volumes it supports.

1. Request a separate site collection for PULSE development, outside any retention or records scope that would make deletion difficult.
2. Publish the current package to a page on that site, run **SharePoint Setup** against the empty site, and record every list and column provisioned and every failure. This is the first observation of provisioning from a clean start.
3. Confirm identity and role resolution for one administrator and one ordinary member.
4. Seed synthetic records only, using neutral role display names and `example.mil` addresses. Do not copy production records.
5. Load the site past the ceilings in the code — roughly 5,000 items for lists loaded without an explicit page size and roughly 25,000 for those paged at 500 — using a generator held in the repository that requires its target site to be passed explicitly.
6. Measure boot time and request count, refresh duration and overlap, render time for the heaviest views, memory after idling, and save duration under load.
7. Provoke the known failure modes: silent truncation at the request guard, list-view-threshold failures on server-sorted date columns, render stalls on large trackers, and save behaviour while refreshes run.
8. Evaluate Dataverse against the same loaded data set, comparing query behaviour at volume, paging, relationships, identity, licensing, migration, and whether single-file packaging survives.
9. Record a written recommendation on the refresh and paging model, and on Dataverse: adopt, reject, or revisit at a stated threshold.

**Verification:** counts shown in the application match counts in the list on every list, at volumes above the guard. A page that renders is not evidence of a complete dataset. **Cautions:** point any data generator only at the development site; an empty or stale site value is a realistic way to load synthetic records into the wrong place. **Records:** provisioning log, rebuild procedure, generator and its volumes, measured results stamped with package version and site state, confirmed failure points with the volume at which each appears, and the two written recommendations.

Full detail, including the specific ceilings, suggested volumes, and the Dataverse comparison, is in documentation library document 10, *Development Environment & Scale Validation*.
