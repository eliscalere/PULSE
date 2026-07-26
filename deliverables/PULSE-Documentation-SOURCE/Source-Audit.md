# Source Audit

## Verified source of truth

- The active SharePoint Framework web part renders `LegacyWorkspaceFrame`, which embeds generated browser HTML; the feature-complete source is `apps/PULSE`.
- `apps/PULSE/index.html` defines the active browser module order.
- Current browser routes and UI labels were checked from the clean local running application.
- The current source schema is `apps/PULSE/assets/js/sharepoint-schema.js`.
- Brand terminology, logo treatment, color vocabulary, and voice were checked against the supplied PULSE brand guide.

## Items requiring confirmation

- [VERIFY] Production host type, approved deployment location, and release authority cannot be determined from the local repository.
- [VERIFY] Live SharePoint list permissions, document-library permissions, current role assignments, and REST write behavior require an authorized site session.
- [VERIFY] AI grammar review endpoint, authorization, credential management, and approved model must be confirmed before it is enabled in an operational environment.
- [VERIFY] Organizational retention, records-management, privacy, and classification controls are not specified in code.

## Deliberate exclusions

- Older source, filenames, and reference documents use legacy product naming. This package uses **PULSE** throughout.
- Older documents describe a superseded color-coded project scheme. This package uses the current displayed lifecycle, technical-health, task, travel, and formal-review values instead.
- Browser-local fallback is documented only as a development interface mode, never as an authoritative operational record.
