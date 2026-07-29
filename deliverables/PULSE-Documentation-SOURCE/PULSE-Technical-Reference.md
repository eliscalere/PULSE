# PULSE Technical Reference

## Architecture

PULSE is a browser-only single-page application. The current delivered SharePoint Framework web part embeds the browser implementation from `apps/PULSE` through `apps/spfx-pulse-workspace/src/legacy/LegacyWorkspaceFrame.tsx`; the typed SPFx page modules are not the active delivered feature surface. The browser app is hosted as a self-contained HTML package on SharePoint or a compatible host and uses the existing SharePoint session plus same-site REST calls.

```text
Browser + SharePoint session -> PULSE package -> shell and feature modules
-> SharePoint REST adapter + repository -> PULSE Lists and document libraries
```

There is no separate PULSE server, database, or sign-in. The bundled app includes browser-side libraries for calendar, date selection, PDF generation, and PowerPoint generation.

## Repository layout

| Path | Responsibility |
|---|---|
| `apps/PULSE/index.html` | Browser entry point and dependency order. |
| `apps/PULSE/assets/js/app.js` | Boot, shell, routing, shared UI helpers, and diagnostics. |
| `apps/PULSE/assets/js/data.js` | Browser store, normalization, and local fallback. |
| `apps/PULSE/assets/js/pages/` | Module-specific UI and workflows. |
| `apps/PULSE/assets/js/sharepoint-adapter.js` | Same-site REST transport, site detection, identity, and role operations. |
| `apps/PULSE/assets/js/sharepoint-repo.js` | Repository mapping between browser objects and list items. |
| `apps/PULSE/assets/js/sharepoint-schema.js` | List definitions, setup, schema validation, and provisioning. |
| `apps/PULSE/assets/js/project-pptx-export.js` and `overview-pptx-export.js` | Browser-side PowerPoint output. |
| `apps/PULSE/scripts/build-sharepoint-package.js` | Preferred self-contained package builder. |
| `apps/spfx-pulse-workspace/` | SPFx host solution and embedded browser-app integration. |

## Configuration and security boundaries

`apps/PULSE/assets/js/app-config.js` holds non-secret runtime configuration such as list prefix, site-detection fallback, host-page naming, group fallbacks, and the government Graph base URL. Do not place credentials, tokens, private site addresses, or personal information in source or documentation. The AI-review configuration contains an empty key field and an endpoint comment requiring organizational confirmation; it is not treated as verified operational capability in this package.

SharePoint authentication relies on the page's existing session. The REST adapter obtains a request digest for write requests and uses same-origin credentials. It detects the site from SharePoint page context, configured manual site value, or a controlled fallback. A live SharePoint site is required to verify identity, permissions, list access, and persistence.

## Build, package, deploy, and rollback

For the browser implementation:

```bash
cd apps/PULSE
node scripts/build-sharepoint-package.js "releases/PULSE-v1.0.0.html"
```

The preferred builder reads `index.html`, inlines local styles and scripts, rewrites local CSS asset URLs, embeds selected images, protects browser-global library registration, removes local cache query strings, and writes a provenance manifest comment. It rejects external runtime style and script dependencies.

To build every delivered package in one pass:

```bash
node apps/PULSE/scripts/build-travel-packages.js
```

That script builds all eight packages — the full application, Travel Request Forms, My Travel, Travel Calendar, Tickets, PULSE Calendar, PULSE CODE, and PULSE Documentation — writing each to both `apps/PULSE/fs-packages/` and `releases/`. The five PULSE-native packages share the builder above and differ only by entry file; the other three delegate to their own build scripts. A failure in any package aborts the run with a non-zero exit code.

The focused packages are ordinary PULSE entry files that set `window.PULSE_PORT_CONFIG` before the application loads, restricting navigation and the sidebar to one area. They carry the same script set as `index.html`; trimming scripts from a focused entry file has broken boot in the past and should be avoided. Document 09 in the documentation library, *Focused Tools & Package Delivery*, covers the package set, each tool's purpose, and the publishing procedure in full.

The documentation package is built from a Vite/RSC application, so its packager additionally rewrites the rendered payload: the route stylesheet is re-pointed at an inlined `data:` URI, the RSC preload hint is dropped, and the dynamic-import chunk table is emptied. Without those steps React suspends its first commit waiting on a stylesheet that does not exist beside a single-file package, which inside a SharePoint iframe leaves the page blank until the user clicks it.

For the SPFx host solution, verified package scripts are:

```bash
cd apps/spfx-pulse-workspace
./use-node22.sh npm install
./use-node22.sh gulp serve
./use-node22.sh npm run typecheck
./use-node22.sh npm run lint
./use-node22.sh npm run build
```

Deploy the approved browser package through the organization’s current hosted-page process. Before deployment, keep the prior known-good artifact. After upload, open the real hosted page, validate identity, role resolution, navigation, and a controlled read/write for each changed record type. If acceptance fails, restore the prior known-good artifact first, preserve evidence, and investigate from source.

## Boot sequence and routing

The browser entry point loads vendor scripts, configuration, SharePoint integration, store normalization, reporting exports, shell code, support subsystems, and page modules. `app.js` determines whether SharePoint context is available; SharePoint mode loads identity, role association, setup state, and repository data. Local fallback loads only browser-local state and is not an official system of record. Routing is hash-based and dispatches to feature renderers for dashboard, overview, projects and project tabs, weekly meeting, travel, Document Review, tickets, notification settings, users, workload, logs, and administration.

## SharePoint integration and data model

Current schema list titles include **PULSE App Roles**, **PULSE Projects**, **PULSE Risks**, **PULSE Action Items**, **PULSE Meetings**, **PULSE Travel Requests**, **PULSE Travel Debriefs**, **PULSE Team Events**, **PULSE Document Review**, **PULSE Doc Reviewer Groups**, **PULSE Audit Log**, **PULSE Issues**, **PULSE Notifications**, **PULSE Notification Config**, **PULSE App Settings**, and **PULSE Location Config**.

Support tickets are stored in **PULSE Issues** rather than a separate ticket list. The repository maps the `ticket` record kind onto that list and translates the type and status vocabularies in both directions, so a ticket raised in the Tickets package and an issue report raised from the full application land in the same list and are visible to both.

### Site detection inside web-part iframes

A focused package hosted as a SharePoint web part runs in a sandboxed iframe that does not receive `_spPageContextInfo`; only the host page has it. Site detection therefore resolves in order from the frame's own context, explicit configuration, the parent window, the top window, a server-relative path combined with the current origin, and finally a site URL cached by the full application on a previous successful boot. If every step fails the package runs in local-only mode, which is the intended signal that it was loaded outside a SharePoint page and is not a system of record.

Repository mappers serialize browser objects to list fields and hydrate list items back into normalized store records. Several composite record areas use JSON fields, including project people, boards, risks, reporting configuration, notes, and other nested values. A field change must be implemented coherently across schema, mapper write path, mapper read path, normalization, UI validation, and live SharePoint read/write verification.

Project document operations use SharePoint document libraries through the same application integration. Formal Document Review record data is held in its dedicated list and workflow; do not combine its lifecycle with project document storage.

## Roles, notifications, reporting, and audit

PULSE role association is resolved through active **PULSE App Roles** records. The role service distinguishes administrative rights from ordinary membership and individual project access. User and role administration is restricted to administrators.

Notification preferences are held per user record and configuration. Notifications may link to routes and record context; users must confirm the source record. Audit logging records supported actions with actor, role, route, object type, and relevant context, subject to successful SharePoint persistence.

Project and group PowerPoint output is generated client-side with the bundled PptxGenJS runtime. In SharePoint mode, the project export flow can upload a status package to project files and attempt to open it in the desktop PowerPoint application. Verify generated content, destination, and file access before use.

## Testing and troubleshooting

Run browser UI checks locally, then validate on a safe SharePoint site. Local tests cannot prove session identity, role association, list permissions, schema readiness, REST write behavior, document-library access, or hosted-page restrictions.

| Symptom | First inspection | Safe response |
|---|---|---|
| Blank or partial page | Hosted page, console, generated package, asset paths | Rebuild with the preferred builder and check package provenance. |
| Incorrect role | Current user, active role record, site diagnostics | Correct the approved role record; do not infer authorization from site membership. |
| Save failure | REST response, permissions, schema, mapper, payload | Validate the live response and full mapping before changing UI. |
| Data inconsistency | List item, JSON field, mapper pair, local cache | Repair the data boundary coherently and repeat read/write validation. |
| Release-only defect | Hosted page and prior artifact | Roll back, preserve evidence, then isolate the change. |

## Known limitations and maintainer handoff

- [VERIFY] The current production hosting arrangement is not available in this workspace; confirm whether the deployed surface is the browser package, SPFx host, or another approved host.
- [VERIFY] Confirm the approved live SharePoint site, document-library names, retention controls, and deployment authority outside this package.
- [VERIFY] Confirm whether AI grammar review has an approved endpoint, key-management path, and authorization to operate before enabling it.
- The typed SPFx structure is incomplete as a standalone feature implementation because the web part currently renders the embedded browser application.

Maintainers should trace any change from feature UI to store, repository mapper, schema, and live validation; use the preferred package builder; retain rollback artifacts; avoid documenting secrets; and update the verification matrices with test evidence.
