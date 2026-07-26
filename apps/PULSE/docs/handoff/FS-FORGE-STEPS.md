# AEWTTR-PULSE Steps And Forge Usage

This file is a plain-English handoff for the AEWTTR-PULSE app.

It covers:
- what was tried so far
- what changed in the app
- how the SharePoint backend works now
- how to build/package the app
- how to use the Forge-related files in this project

## 1. Project location

- App root:
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE`

- Main source entry:
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/index.html`

- Main JavaScript files:
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/assets/js/app-config.js`
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/assets/js/sharepoint-adapter.js`
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/assets/js/sharepoint-schema.js`
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/assets/js/sharepoint-repo.js`
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/assets/js/app.js`

## 2. What was tried so far

The original plan was:
- use SharePoint REST
- use SharePoint Lists as the backend
- create lists and columns automatically
- sync SharePoint site users into a `PULSE App Roles` list

What went wrong:
- lists could be created, but columns were unreliable
- SharePoint created columns with generic internal names like `field_1`, `field_2`, `field_7`
- some columns were created with the wrong type
- duplicate/hidden/internal columns caused writes to appear successful while the visible SharePoint form stayed empty

That made the list/column model too fragile for this environment.

## 3. Current backend design

The app was changed to use a SharePoint document library instead of wide SharePoint lists.

Current design:
- document library: `PULSE App Data`
- app folder inside library: `AEWTTR-PULSE`
- JSON files:
  - `db.json`
  - `roles.json`

Purpose:
- `db.json` stores shared app data
- `roles.json` stores app roles and synced site users

Why this is better:
- avoids column creation problems
- avoids bad SharePoint internal field names
- avoids hidden/duplicate list columns
- still uses SharePoint REST only

## 4. Important code changes

### `sharepoint-adapter.js`

This now handles:
- SharePoint site detection
- current user loading
- diagnostics
- document library creation
- folder creation
- JSON file existence checks
- JSON file reads/writes
- roles loading/saving from `roles.json`

Key behavior:
- it now resolves the library's real `RootFolder.ServerRelativeUrl`
- it creates the folder before trying to create files
- missing JSON files fall back cleanly instead of crashing boot immediately

### `sharepoint-repo.js`

This was simplified so the app saves shared data to:
- `db.json`

instead of writing each object into separate SharePoint lists.

### `sharepoint-schema.js`

This was repurposed from "list/field setup" to "file-store setup".

It now checks:
- data library exists
- app folder exists
- `db.json` exists
- `roles.json` exists

### `dashboard.js`

The home page now includes repair/debug buttons:
- `Run SharePoint Setup`
- `Check Setup`
- `Create Data Store`
- `Create roles.json`
- `Create db.json`
- `Open Logs`

## 5. Current SharePoint setup flow

When the app runs in SharePoint:

1. Detect SharePoint site URL
2. Call `/_api/web/currentuser`
3. Prepare SharePoint data store
4. Sync SharePoint site users into `roles.json`
5. Resolve current app role from `roles.json`
6. Load app data from `db.json`

If there are no app admins yet:
- a SharePoint site admin is temporarily treated as app admin
- this is only to let setup/recovery happen

## 6. Current app storage files

Configured in:
- `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/assets/js/app-config.js`

Current config values:
- `dataLibraryTitle: "PULSE App Data"`
- `dataFolderName: "AEWTTR-PULSE"`
- `dataFiles.db: "db.json"`
- `dataFiles.roles: "roles.json"`

## 7. How to repair from the home page

Recommended click order:

1. `Create Data Store`
2. `Create roles.json`
3. `Create db.json`
4. `Check Setup`
5. `Run SharePoint Setup`
6. `Open Logs` if anything still fails

What to capture if it still breaks:
- `Last REST URL`
- `Last REST status code`
- `Last REST error text`
- newest debug log entry

## 8. Packaging/build tools in this repo

There are two packaging scripts in:
- `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/scripts`

### A. Recommended for Firepit/SharePoint

File:
- `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/scripts/build-sharepoint-package.js`

Purpose:
- creates one flat self-contained HTML package
- matches the real shipped SharePoint app style
- embeds CSS/JS into one file
- adds the `WFC-MANIFEST` header

Use this for:
- SharePoint / Firepit uploads
- the final single-file app deliverable

Command:

```bash
node "/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/scripts/build-sharepoint-package.js" \
  "/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/FS packages/PULSE-v1.0.0.html"
```

Output:
- `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/FS packages/PULSE-v1.0.0.html`

### B. Forge wrapper builder

File:
- `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/scripts/build-forge.js`

Purpose:
- generates a Forge-oriented wrapper/output from a Forge template HTML
- useful when working with the Forge tool itself
- not the preferred final SharePoint shipping format

Use this only if you specifically need a Forge-format artifact.

Command shape:

```bash
node "/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/scripts/build-forge.js" \
  "<path to source Forge.html>" \
  "/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/dist/Forge.html"
```

Example:

```bash
node "/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/scripts/build-forge.js" \
  "/Users/eliscalere/Downloads/DOD SAFE-IJM0O4RwLMJQdEby/apps/PULSE/Website-20260701T182117Z-3-001/Website/vendor/boxicons/fonts/extra/Forge.html" \
  "/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/dist/Forge.html"
```

## 9. Forge-related files

Relevant files:

- Existing generated Forge-format outputs:
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/dist/Forge.html`
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/dist/Forge-readable.html`
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/dist/AEWTTR-PULSE-Forge.html`

- Decoded/reference Forge internals:
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/dist/forge-child-decoded.html`

What we learned:
- the real SharePoint-shipped app is a flat HTML file with `WFC-MANIFEST`
- the best reference shipped file is:
  `/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/releases/reference-packages/PULSE-IPT - Only Secure in FS Sharepoint-current.html`

That shipped file is the reason `build-sharepoint-package.js` is the preferred packaging script now.

## 10. Best current build path

If you just want the Firepit-ready app:

1. Edit source files under `assets/js`, `assets/css`, and `index.html`
2. Run:

```bash
node "/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/scripts/build-sharepoint-package.js" \
  "/Users/eliscalere/Downloads/AEWTTR PAS/apps/PULSE/FS packages/PULSE-v1.0.0.html"
```

3. Upload the output HTML into SharePoint/Firepit

## 11. If another AI is taking over

Tell it:
- this app no longer relies on SharePoint list columns for shared data
- the backend is now a SharePoint document library + folder + JSON files
- the main integration file is `sharepoint-adapter.js`
- the main persistence bridge is `sharepoint-repo.js`
- the recommended ship step is `build-sharepoint-package.js`

## 12. Current known issue to watch

If setup still fails, the next likely issue is not schema anymore, but one of:
- SharePoint blocking folder creation
- SharePoint blocking file upload to the folder
- library root path mismatch
- delayed library/folder availability after creation

When debugging, always capture:
- `Last REST URL`
- `Last REST status code`
- `Last REST error text`
- newest log entry from the Logs page
