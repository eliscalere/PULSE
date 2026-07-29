# PULSE — Developer Handoff

PULSE is a browser-based SharePoint/Firepit application. Start here when maintaining the app.

## Where to work

- `index.html` — application entry point.
- `assets/js/app.js` — application shell, hash routing, and shared UI behavior.
- `assets/js/pages/` — feature pages (projects, weekly meetings, travel, document review, tickets, admin, and more).
- `assets/js/sharepoint-schema.js` — PULSE list definitions.
- `assets/js/sharepoint-adapter.js` and `assets/js/sharepoint-repo.js` — SharePoint integration and data access.
- `assets/css/style.css` — application styling.
- `scripts/` — build and capture helpers. Use `scripts/build-sharepoint-package.js` to generate a SharePoint-ready file.

## Keep these boundaries

- `assets/` and `vendor/` are runtime source dependencies; keep their relative paths intact.
- `docs/` is the maintenance reference: begin with `docs/current/README.md` (paired Markdown/Word documentation set — user guide, technical handoff, operations SOP).
- `fs-packages/` contains the preferred SharePoint/Firepit upload packages.
- `releases/` contains Forge builds and reference deployment packages, not editable application source.
- `validation/` contains verification samples, not runtime code.

## Shipping

From this folder, create a SharePoint-ready package with:

```sh
node scripts/build-sharepoint-package.js "fs-packages/PULSE-v1.0.0.html"
```

Use semantic versioning (`MAJOR.MINOR.PATCH`) in release filenames. Increment MINOR for backward-compatible features and PATCH for backward-compatible fixes. See `docs/handoff/FS-FORGE-STEPS.md` for the fuller packaging procedure.

The secondary Forge path, only when specifically required, is:

```sh
node scripts/build-forge.js <path-to-forge-template.html> releases/forge-builds/Forge.html
```

The Forge template is a user-supplied file (e.g. exported from the Forge/Firepit platform), not something checked into this repo.
