# Procedure Verification Matrix

Rows correspond to the procedures in [Standard Operating Procedures](PULSE-Standard-Operating-Procedures.md).

| Procedure | Source-supported steps | Evidence needed before operational use |
|---|---|---|
| Project and tracker maintenance | Project, action-item, risk, and settings UI plus repository mapping | Live create, update, refresh, and activity evidence |
| Gantt and boards | Tracker/Gantt and board modules | Live task scheduling and board persistence |
| Reporting and PowerPoint | Browser export modules | Output review, upload destination, and access confirmation |
| Weekly and project meetings | Meeting module | Session save, action-item linkage, and history retrieval |
| Travel, leave, and debriefs | Travel module and schema | Request, role-restricted decisions, finance and charge-object assignment, and debrief transaction |
| My Travel and travel calendar | Travel module, travel-request and team-event mappers | Per-user filtering under live identity; trip and event visibility across users |
| Formal Document Review | Review module and schema | Revision, decision, signing, and archive workflow |
| Tickets and notifications | Ticket and notification modules | Ticket updates persisting to **PULSE Issues**, and preference persistence |
| Users, roles, setup, and logs | Admin, users, schema, adapter, audit modules | Administrator operation on approved site |
| Export and troubleshooting | Build and diagnostics code | Hosted-page release, rollback, and evidence retention |
| Development site and scale validation | Existing site-detection hook and provisioning tooling; no environment, generator, or tests yet | A disposable site provisioned from empty, a repository-held data generator, measured results stamped with package version, confirmed failure volumes, and written recommendations on the refresh model and on Dataverse |
| Publishing a focused tool page | Unified package builder, `PULSE_PORT_CONFIG` entry files, site-detection chain | Site resolution from a web-part page, a controlled read/write visible in the full application, and a recorded artifact-to-page mapping |
