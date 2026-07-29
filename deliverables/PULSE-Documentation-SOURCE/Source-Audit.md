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

## Items requiring confirmation

- [VERIFY] Production host type, approved deployment location, and release authority cannot be determined from the local repository.
- [VERIFY] Live SharePoint list permissions, document-library permissions, current role assignments, and REST write behavior require an authorized site session.
- [VERIFY] AI grammar review endpoint, authorization, credential management, and approved model must be confirmed before it is enabled in an operational environment.
- [VERIFY] Organizational retention, records-management, privacy, and classification controls are not specified in code.
- [VERIFY] Site resolution for focused packages hosted as web parts depends on the host page exposing SharePoint context to the frame's parent window; this can only be confirmed on a real hosted page.
- [VERIFY] PULSE CODE authentication, file read/write behavior, and AI features were not functionally exercised in this workspace; only its build was verified.

## Known documentation gap

The eight portfolio PDFs under `apps/pulse-documentation/public/source-pdfs/` are externally authored and have no generator in this repository. The `.txt` files the documentation site searches are text extractions of those PDFs. Both were therefore left unchanged in this pass, so the searchable library does not yet describe My Travel, Travel Calendar, PULSE CODE, or the focused package model — that coverage currently exists in these Markdown sources, the site's Interface reference view, and the SOP set. Closing the gap requires regenerating the PDFs and re-extracting the text so the two stay in agreement.

## Deliberate exclusions

- Older source, filenames, and reference documents use legacy product naming. This package uses **PULSE** throughout.
- Older documents describe a superseded color-coded project scheme. This package uses the current displayed lifecycle, technical-health, task, travel, and formal-review values instead.
- Browser-local fallback is documented only as a development interface mode, never as an authoritative operational record.
