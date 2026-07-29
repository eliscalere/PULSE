# PULSE Technical Handoff

## Operating model

PULSE is a browser-only single-page application hosted as a self-contained HTML file in SharePoint. The existing SharePoint session supplies identity; SharePoint Lists and libraries supply shared records. There is no application server, separate database, separate login, required runtime CDN, or required external API.

```text
Browser + SharePoint session → one-file PULSE package → shell + feature modules
→ SharePoint REST adapter + repository → SharePoint lists and libraries
```

## Source structure

| Area | Primary paths | Responsibility |
|---|---|---|
| Entry | `index.html` | Development dependency order. |
| Shell/store | `assets/js/app.js`, `data.js` | Boot, routing, rendering shell, shared state. |
| Features | `assets/js/pages/` | Workspace-specific UI and interactions. |
| SharePoint | `sharepoint-adapter.js`, `sharepoint-repo.js`, `sharepoint-schema.js` | REST, mapping, persistence, schema. |
| Preferred packaging | `scripts/build-sharepoint-package.js` | One-file release artifact. |

## Change workflow

1. Trace page module → store → repository mapper → schema before editing.
2. Change schema, object/list mappers, normalization, UI, and live validation together when adding data.
3. Test UI locally; use a safe SharePoint site for identity, role, schema, and persistence checks.
4. Build the one-file package, then validate in the actual hosted page.
5. Retain the prior known-good package and release evidence.

## Packaging

Use:

```bash
node scripts/build-sharepoint-package.js "fs-packages/PULSE-v1.0.0.html"
```

The preferred builder reads `index.html`, inlines styles and scripts, rewrites CSS URLs and selected images to data URIs, strips development cache queries, protects the presentation-library UMD global, writes a provenance manifest, and fails external runtime dependencies. The secondary wrapper builder is for portable/iframe distribution, not the normal hosted release.

## Persistence and troubleshooting

Feature modules change in-memory objects and call the repository. The repository maps objects to list items, sends same-origin REST calls, and refreshes/reconciles local state. JSON fields are a compatibility contract; change both mapper directions and validate a live read/write. Saves are debounced and chained per object, with effective last-write-wins behavior.

For blank pages, inspect the real hosted page, console, package, and asset queries. For save failures, inspect the live REST response, permissions, list schema, mapper, and payload before changing UI. For a release-only defect, roll back to the prior package, capture evidence, and isolate the source change.
