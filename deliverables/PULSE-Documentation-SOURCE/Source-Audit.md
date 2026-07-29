# Source Audit

## Verified source of truth

- The active SharePoint Framework web part renders `LegacyWorkspaceFrame`, which embeds generated browser HTML; the feature-complete source is `apps/PULSE`.
- `apps/PULSE/index.html` defines the active browser module order.
- Current browser routes and UI labels were checked from the clean local running application.
- The current source schema is `apps/PULSE/assets/js/sharepoint-schema.js`.
- Brand terminology, logo treatment, color vocabulary, and voice were checked against the supplied PULSE brand guide.
- The delivered package set is defined by `apps/PULSE/scripts/build-travel-packages.js`: PULSE, Travel Request Forms, My Travel, Travel Calendar, Tickets, PULSE Calendar, PULSE CODE, and PULSE Documentation. All eight build from source in this workspace.
- Focused packages are ordinary PULSE entry files differing only by `PULSE_PORT_CONFIG`; they carry the same script set as `index.html`.
- Support tickets resolve to the **PULSE Issues** list. There is no separate ticket list in the schema, and the repository translates ticket type and status onto the issue vocabulary.
- The documentation package boot path was exercised locally, top-level and inside a sandboxed iframe, with no failed asset requests and no page errors.
- The delivered documentation package omits the controlled PDFs to stay loadable in SharePoint (2.3 MB against 14.1 MB with them). `INCLUDE_PDFS=1` produces the archival package that carries them; PDF affordances are hidden when the assets are absent.
- There is no development environment: no isolated site in configuration or documentation, no test directory, and no test script. `manualSharePointSiteUrl` in `app-config.js` is the intended hook for aiming at a specific site and is empty.
- The REST item fetch pages on SharePoint's continuation link under a guard of 50 requests and returns silently on reaching it, giving ceilings of roughly 5,000 items for lists loaded without an explicit page size and roughly 25,000 for those paged at 500.
- Issues, tickets, and the audit log are sorted server-side on date columns and grow without bound, so they are the lists most exposed to SharePoint's 5,000-item view threshold.
- Background refresh reloads the entire workspace rather than fetching deltas, gated to no more than one successful refresh every five seconds while the tab is visible, and also on navigation and tab focus.
- No Dataverse reference exists anywhere in the source; evaluating it would be new work with no existing commitment.

## Items requiring confirmation

- [VERIFY] Production host type, approved deployment location, and release authority cannot be determined from the local repository.
- [VERIFY] Live SharePoint list permissions, document-library permissions, current role assignments, and REST write behavior require an authorized site session.
- [VERIFY] AI grammar review endpoint, authorization, credential management, and approved model must be confirmed before it is enabled in an operational environment.
- [VERIFY] Organizational retention, records-management, privacy, and classification controls are not specified in code.
- [VERIFY] Site resolution for focused packages hosted as web parts depends on the host page exposing SharePoint context to the frame's parent window; this can only be confirmed on a real hosted page.
- [VERIFY] PULSE CODE authentication, file read/write behavior, and AI features were not functionally exercised in this workspace; only its build was verified.
- [VERIFY] Behaviour at realistic data volumes is unmeasured. The ceilings above are read from source, not observed; the volume at which each first causes a visible failure is unknown until SOP 10 is performed.

## Documentation library composition

Documents 01 through 08 under `apps/pulse-documentation/public/source-pdfs/` are externally authored portfolio documents with no generator in this repository; the matching `.txt` files are text extractions of them. They were checked for claims invalidated by current behaviour and none were found — they are incomplete with respect to the newer surfaces, not inaccurate. They are therefore unchanged, and their PDF-to-extraction agreement is preserved.

Document 09, **Focused Tools & Package Delivery**, closes the coverage gap for My Travel, the travel calendar, Tickets and the PULSE Issues binding, PULSE CODE, site resolution in web parts, and publishing a focused tool page. It is generated from a single content source, so its PDF and its searchable extraction cannot drift:

- Content source: `apps/pulse-documentation/documents/09-focused-tools.mjs`
- Generator: `apps/pulse-documentation/scripts/build-source-document.mjs`
- Command: `node apps/pulse-documentation/scripts/build-source-document.mjs 09`

Documents 11 (**Interface Reference**) and 12 (**Process Flows**) are generated the same way and carry figures. Screenshots and process diagrams now live inside the documents rather than in separate views of the site, so they appear in the PDF, in the searchable text as captions, and in the reader as the same image files:

- Diagram source: `apps/pulse-documentation/documents/flows.mjs`, rendered to `public/figures/flow-<id>.svg` by `scripts/build-figures.mjs`. The PDF and the reader display that one file, so a diagram cannot be drawn twice and disagree with itself.
- Figure placement: the generator emits a per-section manifest to `app/generated/figures/`, which the reader uses to place the same figures the PDF shows.

Document 10, **Development Environment & Scale Validation**, is generated the same way from `apps/pulse-documentation/documents/10-development-environment.mjs`. It records that PULSE has no development environment, states the scale ceilings read from source, and specifies the work needed to stand up a disposable development site, load it past those ceilings, and evaluate Dataverse. It describes planned work; nothing in it is a current capability.

Any future document should be authored the same way rather than as a PDF with a separate extraction. Reissuing documents 01 through 08 on that basis remains an open ownership question.

## Deliberate exclusions

- Older source, filenames, and reference documents use legacy product naming. This package uses **PULSE** throughout.
- Older documents describe a superseded color-coded project scheme. This package uses the current displayed lifecycle, technical-health, task, travel, and formal-review values instead.
- Browser-local fallback is documented only as a development interface mode, never as an authoritative operational record.
