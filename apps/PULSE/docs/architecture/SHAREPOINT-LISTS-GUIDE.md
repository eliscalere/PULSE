# AEWTTR-PULSE — SharePoint Lists: How They're Created and How the App Uses Them

This document has two halves:

- **Section 0 — the universal recipe.** Standalone, copy-pasteable code for
  creating a SharePoint list, adding columns, and reading/writing items from
  ANY browser-hosted app (Firepit web part, script editor, single-file HTML).
  No PULSE code required. Start here if you're building something new.
- **Sections 1–6 — how PULSE applies it.** The bug history, this app's schema,
  and how the app's data layer is wired.

Audience: whoever maintains this app next — a developer or an AI assistant.

---

## 0. The universal recipe — create and use any SharePoint list via REST

Everything below is plain `fetch` against SharePoint's REST API. It works from
any page served from the SharePoint site (Firepit web part, document-library
HTML file) because the browser already carries the user's CAC/SSO session —
`credentials: "same-origin"` on every request is what authenticates you.
No Graph, no Azure app registration, no npm packages.

`SITE` below is your site URL, e.g. `https://flankspeed.sharepoint-mil.us/sites/YourSite`.

### 0.1 The one rule for all WRITES: get a request digest first

Every POST/MERGE/DELETE needs a fresh `X-RequestDigest` token. GETs don't.

```js
async function getDigest(SITE) {
  const r = await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST",
    credentials: "same-origin",
    headers: { "Accept": "application/json;odata=nometadata" }
  });
  const d = await r.json();
  return d.FormDigestValue;
}
```

Digests expire (~30 min). Fetch one per burst of writes, not one per page load.

### 0.2 Create a list

```js
async function createList(SITE, title, description) {
  const r = await fetch(`${SITE}/_api/web/lists`, {
    method: "POST",
    credentials: "same-origin",
    headers: {
      "Accept": "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": await getDigest(SITE)
    },
    body: JSON.stringify({ Title: title, Description: description || "", BaseTemplate: 100 })
  });
  if (!r.ok) throw new Error(`createList ${r.status}: ${await r.text()}`);
  return r.json(); // includes the new list's Id
}
```

`BaseTemplate: 100` = generic custom list. A list that already exists → HTTP 409;
check existence first with a GET to `/_api/web/lists/getbytitle('Title')`
(404 = doesn't exist).

### 0.3 Create columns — THE PART THAT BITES EVERYONE

Create each column from CAML XML via `createfieldasxml`, and you MUST pass
`Options: 28`. If you pass 0 (or omit thinking it's optional), SharePoint
ignores your internal name and generates `field_1`-style names, your writes
will target columns that don't exist, and the list will open showing only
Title. This single flag is the difference between "it works" and every
symptom in Section 1.

```js
// type: "Text" | "Note" | "Number" | "DateTime" | "Boolean" | "Choice"
function fieldXml(name, type, choices) {
  const esc = (s) => String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
  const common = `DisplayName="${esc(name)}" Name="${esc(name)}" StaticName="${esc(name)}" Required="FALSE"`;
  if (type === "Choice") {
    const opts = (choices || []).map(c => `<CHOICE>${esc(c)}</CHOICE>`).join("");
    return `<Field Type="Choice" ${common} Format="Dropdown"><CHOICES>${opts}</CHOICES></Field>`;
  }
  if (type === "Note")     return `<Field Type="Note" ${common} NumLines="6" RichText="FALSE" />`;
  if (type === "DateTime") return `<Field Type="DateTime" ${common} Format="DateTime" />`;
  if (type === "Boolean")  return `<Field Type="Boolean" ${common} />`;
  if (type === "Number")   return `<Field Type="Number" ${common} />`;
  return `<Field Type="Text" ${common} />`;
}

async function createColumn(SITE, listTitle, name, type, choices) {
  const r = await fetch(
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/createfieldasxml`,
    {
      method: "POST",
      credentials: "same-origin",
      headers: {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": await getDigest(SITE)
      },
      body: JSON.stringify({
        parameters: {
          __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
          SchemaXml: fieldXml(name, type, choices),
          // 4 = show on the new/edit form; 8 = HONOR THE INTERNAL NAME (critical);
          // 16 = show in the default view. Never pass 0.
          Options: 28
        }
      })
    }
  );
  if (!r.ok) throw new Error(`createColumn ${name} ${r.status}: ${await r.text()}`);
}
```

Notes:
- This one endpoint wants `odata=verbose` on both headers and the
  `__metadata.type` wrapper — most other endpoints are happy with `nometadata`.
- Check a column exists before creating:
  `GET /_api/web/lists/getbytitle('X')/fields?$filter=InternalName eq 'Name'`.
- Use simple alphanumeric names (`ProjectCode`, not `Project Code`) — spaces
  get encoded into `_x0020_` messes.
- Verify after creating: re-read `/fields?$select=InternalName,TypeAsString`
  and confirm your exact name is there. If it isn't, you (or an earlier run)
  created a mangled column — delete it in List Settings and create again.
  Mangled columns can't be renamed via REST.

If a column exists but the list still opens showing only Title, the column
just isn't in the default view — repair with:

```js
await fetch(`${SITE}/_api/web/lists/getbytitle('${listTitle}')/DefaultView/ViewFields/addviewfield('${name}')`,
  { method: "POST", credentials: "same-origin",
    headers: { "Accept": "application/json;odata=nometadata", "X-RequestDigest": await getDigest(SITE) } });
```

### 0.4 Read items ("linking to" a list)

```js
async function getItems(SITE, listTitle, odataQuery) {
  // odataQuery e.g. "$select=Id,Title,Status&$filter=Status eq 'Open'&$orderby=Id&$top=500"
  const r = await fetch(
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items${odataQuery ? "?" + odataQuery : ""}`,
    { method: "GET", credentials: "same-origin", headers: { "Accept": "application/json;odata=nometadata" } }
  );
  if (!r.ok) throw new Error(`getItems ${r.status}: ${await r.text()}`);
  const data = await r.json();
  return data.value; // array of items; follow data["odata.nextLink"] for >100 items
}
```

Gotchas:
- Property names in results are the columns' **internal names**, not display
  names. If your columns were created correctly (0.3) they're the same thing.
- Results page at ~100 items; keep following `odata.nextLink` until absent.
- 404 on the list URL = list doesn't exist (or the title has a typo — the
  title in `getbytitle` is the DISPLAY title, case-insensitive but exact).

### 0.5 Create, update, delete items

```js
async function createItem(SITE, listTitle, fields) {
  const r = await fetch(`${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`, {
    method: "POST",
    credentials: "same-origin",
    headers: {
      "Accept": "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": await getDigest(SITE)
    },
    body: JSON.stringify(fields)   // e.g. { Title: "Item 1", Status: "Open", DueDate: "2026-07-15T00:00:00Z" }
  });
  if (!r.ok) throw new Error(`createItem ${r.status}: ${await r.text()}`);
  return r.json();                 // includes .Id — SAVE IT; updates/deletes need it
}

async function updateItem(SITE, listTitle, itemId, fields) {
  const r = await fetch(`${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`, {
    method: "POST",
    credentials: "same-origin",
    headers: {
      "Accept": "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": await getDigest(SITE),
      "IF-MATCH": "*",             // overwrite regardless of item version
      "X-HTTP-Method": "MERGE"     // partial update — only the fields you send
    },
    body: JSON.stringify(fields)
  });
  if (!r.ok && r.status !== 204) throw new Error(`updateItem ${r.status}: ${await r.text()}`);
}

async function deleteItem(SITE, listTitle, itemId) {
  const r = await fetch(`${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`, {
    method: "POST",
    credentials: "same-origin",
    headers: {
      "Accept": "application/json;odata=nometadata",
      "X-RequestDigest": await getDigest(SITE),
      "IF-MATCH": "*",
      "X-HTTP-Method": "DELETE"
    }
  });
  if (!r.ok && r.status !== 204) throw new Error(`deleteItem ${r.status}: ${await r.text()}`);
}
```

Payload rules that cause mysterious 400s:
- **DateTime** columns: ISO strings (`"2026-07-15T00:00:00Z"`) or `null`.
  Never `""`.
- **Choice** columns: the value must be one of the defined choices exactly.
- **Number** columns: send a number, not a numeric string.
- **Boolean** columns: send `true`/`false`.
- Sending a property that doesn't match ANY column's internal name → 400
  ("The property ___ does not exist"). This is the classic symptom of columns
  created with `Options: 0` — the internal names aren't what you think.

### 0.6 Linking lists to each other

SharePoint has lookup columns, but plain columns are simpler and more
portable: store the related item's numeric `Id` in a Number column (that's a
foreign key), or a human business key in a Text column. Resolve the join in
your own code after loading both lists. That's exactly what PULSE does —
see "ProjectCode ↔ list-item ID" in Section 4.

### 0.7 Minimal end-to-end example

```js
const SITE = "https://flankspeed.sharepoint-mil.us/sites/YourSite";

// one-time setup
await createList(SITE, "My Tasks", "demo");
await createColumn(SITE, "My Tasks", "Status", "Choice", ["Open", "Done"]);
await createColumn(SITE, "My Tasks", "DueDate", "DateTime");

// runtime
const item = await createItem(SITE, "My Tasks", { Title: "First", Status: "Open", DueDate: "2026-07-15T00:00:00Z" });
await updateItem(SITE, "My Tasks", item.Id, { Status: "Done" });
const open = await getItems(SITE, "My Tasks", "$filter=Status eq 'Open'");
```

---

## 1. The problem that broke earlier attempts

Earlier versions of setup could create lists, but the columns came out wrong:

- internal names like `field_1`, `field_7` instead of `ProjectCode`
- some columns created with the wrong type
- duplicate columns appearing after re-running setup
- writes "succeeded" but the SharePoint form stayed empty
- lists that opened showing only the Title column

**Root cause — one line.** The REST call that creates a column,
`/_api/web/lists/getbytitle('X')/fields/createfieldasxml`, takes an `Options`
bitmask alongside the field XML. The old code passed `Options: 0`. With that
value SharePoint **ignores the `Name` attribute in your field XML** and
auto-generates an internal name (`field_N`). Everything else cascaded from
that:

1. The column got a mangled internal name.
2. The app's "does this column exist?" check looked for the *correct* internal
   name, found nothing, and created the column again → duplicates.
3. The app wrote item properties using the correct names, which didn't match
   the real columns → data landed nowhere visible.

**The fix.** Pass `Options: 28`, which is three flags OR'd together:

| Flag | Value | Effect |
|---|---|---|
| `addToAllContentTypes` | 4 | Column appears on the list's new/edit form |
| `addFieldInternalNameHint` | 8 | **SharePoint honors the `Name` attribute in the XML as the internal name** — this is the critical one |
| `addFieldToDefaultView` | 16 | Column shows when the list is opened (otherwise the default view keeps showing only Title) |

Location in code: `createFieldAsXml()` in `assets/js/sharepoint-adapter.js`.

---

## 2. How a column is created (the full sequence)

All in `assets/js/sharepoint-adapter.js`:

1. **`buildFieldXml(field)`** builds CAML XML with `Name`, `StaticName`, and
   `DisplayName` all set to the same value, e.g.
   `<Field Type="Choice" DisplayName="Rag" Name="Rag" StaticName="Rag" ...>`.
   Field types used: `Text`, `Note`, `Number`, `DateTime`, `Boolean`, `Choice`.

2. **`createFieldAsXml(siteUrl, listTitle, xml)`** POSTs it with `Options: 28`
   (see above). If the XML route fails, **`createFieldByType`** is a fallback
   that creates the field through the typed REST endpoint — safe here because
   all our field names are plain alphanumeric, so `Title` becomes the internal
   name unchanged.

3. **Verification (`ensureField`)** — after creating, the code re-reads the
   list's fields and requires an **exact** `InternalName` match. If the column
   came back with a mangled name it throws a loud error telling you which
   column to delete in list settings, instead of letting writes silently
   target an invisible column.

4. **View repair (`ensureDefaultViewFields`)** — after all columns are
   ensured, setup reads the list's default view
   (`/DefaultView/ViewFields`) and calls `addviewfield('Name')` for any schema
   column missing from it. This is what fixed lists created by earlier runs:
   their columns existed but were never added to the view, so opening the list
   showed only Title.

Requests use the standard SharePoint REST plumbing: `X-RequestDigest` from
`/_api/contextinfo` for writes, `credentials: "same-origin"` so the user's
existing CAC/SSO session authenticates every call. No Graph, no Azure app
registration — this is what makes it work inside Firepit on Flank Speed.

---

## 3. The schema — what lists exist

Defined once in `SHAREPOINT_SCHEMA` in `assets/js/sharepoint-schema.js`.
12 lists, all prefixed **"PULSE "** so they group together in Site Contents
(lists cannot live inside folders — the prefix is the SharePoint convention
for grouping an app's lists):

| List | Holds | Notable columns |
|---|---|---|
| PULSE App Roles | one item per user | `UserEmail`, `Role` (Admin/Manager/Member/Viewer), `IsActive` |
| PULSE Projects | one item per project | `ProjectCode` (the app's "P05"-style key), `Rag`, `HistoryJson`, `PeopleJson` |
| PULSE Action Items | tracker tasks AND project checklist items | `ProjectId` (number FK), `Source` (Tracker/Checklist), `SubtasksJson`, `SortOrder` |
| PULSE Risks / PULSE Updates / PULSE Decisions | reserved per spec | `ProjectId` |
| PULSE Meetings | one item per meeting session | `ProjectId` (0 = global weekly meeting), `ActivityJson` |
| PULSE Rocks | EOS-style rocks | `RockStatus`, `Pct`, `CheckupsJson` |
| PULSE Travel Requests / PULSE Travel Debriefs | travel workflow | joined by `RequestCode` (e.g. "TR-0041") |
| PULSE Document Review | doc review kanban | `ReviewColumn` (the kanban column), `ReviewersJson` |
| PULSE RAG Config | single settings row | RAG thresholds |

Nested data the schema didn't want as separate lists (subtasks, activity logs,
reviewers) is stored as JSON text in `*Json` Note columns.

**Run SharePoint Setup** (Dashboard button) walks this schema: creates missing
lists (`BaseTemplate: 100`), ensures every column, repairs default views,
seeds the RAG Config row, and syncs SharePoint site users into PULSE App
Roles. It's idempotent — safe to run repeatedly; it only creates what's
missing. **Check Setup** reports per-list status including missing columns,
type mismatches, and "suspicious" columns (leftover `field_N` ones from old
runs — those must be deleted by hand; setup detects but can't repair them).

---

## 4. How the app links to the lists at runtime

### Boot (in `assets/js/app.js` → `bootSharePointMode`)

1. Detect the site URL (`_spPageContextInfo`, URL heuristics, or probe
   `/_api/web/currentuser`).
2. Load the current user from REST; look up their role in PULSE App Roles.
   If no app Admin exists yet, a SharePoint **site admin** is treated as app
   admin so setup/recovery is always possible.
3. `getSetupStatus()` — if lists/columns are missing, the app renders with an
   empty database and points the admin at Run SharePoint Setup.
4. `loadAllFromSharePoint()` — queries **all lists in parallel** and assembles
   the in-memory `db` object every page module reads.

### The key translation: ProjectCode ↔ list-item ID

Pages use `"P05"`-style project codes everywhere. Child lists store the
project's numeric **list-item ID** in their `ProjectId` column. The repo
(`assets/js/sharepoint-repo.js`) builds a bidirectional map right after
loading PULSE Projects (`registerProjectId`), and every mapper translates in
both directions. Page code never sees a raw SharePoint ID.

### Mappers

For each kind there's a pair in `sharepoint-repo.js`:
`projectToSpItem`/`spItemToProject`, `actionItemToSpItem`/`spItemToActionItem`,
etc. They translate between the app's object shape and list-item property
names. Loaded objects carry hidden bookkeeping fields:

- `_spId` — the list-item ID (present ⇒ save updates; absent ⇒ save creates,
  then stamps it)
- `_projectCode`, `_source`, `_column` — routing context (which project,
  Tracker vs Checklist, which kanban column)
- `_sortOrder` — last saved `SortOrder` value

### Saving — `Repo.save(kind, obj, extra)`

Page code mutates local objects then calls `Repo.save`. The repo:

1. Stamps routing context from `extra` onto the object.
2. **Debounces per object (500 ms)** — a gantt-bar drag fires a save per
   pixel of mouse movement; these collapse into ONE REST call carrying the
   final state.
3. **Chains writes per object** — a create always finishes (and stamps
   `_spId`) before a follow-up update runs, so rapid edits can't create
   duplicates.
4. Creates (`POST .../items`) or updates (`MERGE .../items(id)`) the one item.

Because each record is its own list item, two people editing different things
can never overwrite each other — the whole-file clobbering problem of the old
db.json backend is gone.

### Task ordering

Lists don't preserve array order, so PULSE Action Items has a `SortOrder`
number column. Loads sort by it; when a drag-reorder saves a task, the repo
re-indexes any siblings whose stored `SortOrder` no longer matches their array
position (writes are proportional to how many items actually moved).

### Deletes

`Repo.remove` deletes the one item. Deleting a **project** also deletes its
tracker/checklist items so they don't return as orphans on the next load.

### Background refresh

When the tab regains focus (throttled to 30 s, skipped if local unsaved
changes are pending), the app silently re-runs `loadAllFromSharePoint` and
re-renders — teammates' changes appear without a manual reload.

---

## 5. File map

| File | Role |
|---|---|
| `assets/js/app-config.js` | site-URL override, list prefix, default role, boot error catcher |
| `assets/js/sharepoint-adapter.js` | all raw REST: digest, CRUD, field/list/view creation, roles, diagnostics |
| `assets/js/sharepoint-schema.js` | `SHAREPOINT_SCHEMA` + `runSharePointSetup` / `getSetupStatus` |
| `assets/js/sharepoint-repo.js` | mappers, ProjectCode↔ID map, `Repo.save`/`Repo.remove`, per-object debounce |
| `assets/js/app.js` | boot sequence, mode detection, background refresh |

Ship step: `node scripts/build-sharepoint-package.js "FS packages/PULSE-v1.0.0.html"`
produces the single-file Firepit upload.

## 6. Gotchas for the next maintainer

- **Never** create columns with `Options: 0` — that's the original bug.
- DateTime item properties must be ISO strings or `null`; SharePoint rejects `""`.
- Choice columns reject values not in their choice list (no fill-in). If a new
  `Source`/status value is needed, add it to the schema and re-run setup.
- A column with a mangled internal name cannot be renamed via REST — delete it
  in list settings and re-run setup.
- The legacy `PULSE App Data` document library (db.json/roles.json) is no
  longer read by the app. It can be deleted once you're confident nothing in
  it is needed.
