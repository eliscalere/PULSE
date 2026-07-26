# Building Apps for Firepit + SharePoint Online — A Developer's Guide

This document explains, from first principles, how to build a static
HTML/CSS/JavaScript application that runs inside a **Firepit web part** on
**SharePoint Online / Flank Speed**, using **SharePoint itself** as both the
identity provider and the data backend — no server, no Microsoft Graph, no
Azure app registration, no CDN dependencies.

Everything in this document was learned by building and shipping a real
production app (AEWTTR-PULSE) this way. Where something is inferred from
observed behavior rather than documented platform source, that's called out
explicitly — don't take anything here as gospel about Firepit/Forge internals
we haven't directly verified.

Audience: a developer or AI assistant starting a **new** Firepit-hosted app.

---

## 1. The mental model

A Firepit web part hosts a single blob of HTML. Whatever you give it, it
renders and executes — inline `<script>` tags run, inline `<style>` tags
apply. There is no build step, no bundler, no npm install step at runtime.
This has three consequences that shape everything else in this guide:

1. **Ship one self-contained HTML file.** No external `<script src>`/`<link
   href>` to files sitting elsewhere — everything (JS, CSS, images, fonts)
   must be inlined into that one file before it goes to Firepit.
2. **SharePoint is your only backend.** The page is served from/embedded in
   a SharePoint site, so the browser already carries the user's session
   cookie. SharePoint's REST API (`/_api/web/...`) is reachable with
   `credentials: "same-origin"` and no separate login — this is what makes
   "no Azure app registration" possible.
3. **The web part's available width varies.** A Firepit web part can sit
   narrower than a full page (next to other web parts), so responsive layout
   isn't optional polish — design for it from the start.

## 2. Forge — the packaging/ship tool

"Forge" is the internal tool used to compile and ship an app into
SharePoint/Firepit. Two things are worth knowing, one **verified by
inspecting a real shipped artifact**, one **inferred from an earlier
engineering handoff note** (flagged as such):

### 2.1 Verified: the shipped file format

A real app shipped through Forge's SharePoint path is a **single flat HTML
file** starting with an HTML comment:

```html
<!--WFC-MANIFEST:eyJ2ZXJzaW9uIjoxLCJwcm9qZWN0Ijoi...-->
<!DOCTYPE html>
<html lang="en">
...
```

That comment is a base64-encoded JSON manifest:

```json
{
  "version": 1,
  "project": "YOUR-APP-NAME",
  "generated": "2026-07-01T18:59:12.176Z",
  "index": "index.html",
  "files": [
    { "kind": "css", "path": "assets/css/style.css", "external": false, "hash": "e711f1b7" },
    { "kind": "js",  "path": "assets/js/app.js",      "external": false, "hash": "65f300a8" }
  ]
}
```

It's provenance metadata — a manifest of every source file that got inlined
and a content hash of each, so a human or tool can tell what shipped and
whether a given source file changed since the last ship. Firepit itself
doesn't need this comment to render the page; it's there for traceability.

Below the manifest comment, the file is a completely ordinary HTML document:
all CSS inlined into `<style>` blocks, all JS inlined into `<script>` blocks,
all images/fonts as `data:` URIs. Nothing references an external path.

**This flat, no-iframe, no-wrapper format is the one that actually works
reliably in Firepit.** Build your own version of this packaging step (see
§3) rather than depending on a specific tool's UI — the format itself is
simple enough to reproduce with a ~150-line Node script.

### 2.2 Inferred from handoff notes: Forge's two ship paths

An earlier engineering handoff described Forge (the authoring tool, as
opposed to the shipped-file format above) as having two distinct compile
targets:

- **SharePoint/Flank Speed ship** — calls something like
  `startCompilation({ forceNoSecurityHeaders: true, ... })` and produces the
  flat manifest-tagged HTML described in §2.1, typically saved with a suffix
  like `" - Only Secure in FS Sharepoint"`. **This is the correct path for
  Firepit.**
- **Offline wrapper ship** — produces a heavier artifact: a parent shell page
  that base64-encodes the entire child app and injects it into an `<iframe
  srcdoc>` sandbox at runtime. This is meant for offline/portable
  distribution, **not** for direct Firepit hosting, and behaved noticeably
  worse there in practice (see §7's "the wrong build tool" postmortem).

If you only have access to the offline-wrapper-style tool, don't use it as
your Firepit ship path — replicate the flat-manifest format directly instead
(§3). If your Forge deployment exposes a "SharePoint" or "Flank Speed"
target distinct from a generic/offline export, prefer that one.

## 3. Build your own packaging script

Because the shipped format is just "inline everything into one HTML file,"
you don't need to depend on any specific external tool to produce it. A
Node script under 200 lines does the whole job:

```js
// scripts/build-sharepoint-package.js (abbreviated)
const fs = require("fs");
const path = require("path");

function toDataUri(filePath) {
  const mime = { ".css":"text/css", ".js":"application/javascript", ".png":"image/png",
    ".woff2":"font/woff2", ".ttf":"font/ttf" /* ...etc */ }[path.extname(filePath).toLowerCase()]
    || "application/octet-stream";
  return `data:${mime};base64,${fs.readFileSync(filePath).toString("base64")}`;
}

function inlineCssUrls(cssText, cssPath) {
  // rewrite url(...) references to local fonts/images as data: URIs
  return cssText.replace(/url\(([^)]+)\)/g, (match, raw) => {
    const cleaned = raw.trim().replace(/^['"]|['"]$/g, "");
    if (!cleaned || cleaned.startsWith("data:") || /^https?:/i.test(cleaned)) return match;
    const target = path.resolve(path.dirname(cssPath), cleaned);
    return fs.existsSync(target) ? `url("${toDataUri(target)}")` : match;
  });
}

// 1. Read index.html
// 2. For every <link rel="stylesheet" href="...">, inline the CSS as <style>
// 3. For every <script src="...">, inline the JS as <script>...</script>
//    (escape any literal `</script` inside the source so it doesn't
//    terminate the tag early)
// 4. Replace any local image paths referenced directly in JS/CSS with data: URIs
// 5. Prepend the WFC-MANIFEST comment with a file list + content hash per file
// 6. Write out the single resulting HTML file
```

Key rules this script enforces:
- **Throw, don't silently skip, on any external (`https://`) script or
  stylesheet reference.** Firepit ships need to work with zero runtime
  network dependencies; catching this at build time beats discovering it in
  production.
- **Escape `</script` inside inlined JS source** (e.g. `s/<\/script/<\\/script/gi`)
  — a literal closing tag inside a string constant would otherwise terminate
  the `<script>` block early and corrupt the page.
- Inline images referenced by absolute path in JS/CSS as `data:` URIs too, not
  just ones referenced via `<img src>` in the HTML.

Run it after every change, before every upload:

```bash
node scripts/build-sharepoint-package.js "output/YourApp_v1.html"
```

Upload the resulting single file wherever your Firepit web part reads from
(typically a SharePoint document library backing the web part).

### 3.1 The cache-busting trap

**Do not add `?v=N` query-string cache-busting parameters to script/link
tags** if you also use this flat-file packaging approach, and be cautious
even in unbundled multi-file dev setups hosted from SharePoint. We hit this
directly: SharePoint's static-file serving mishandled query strings on
script URLs in a way that caused every `<script src="app.js?v=10">` to
silently 404, which — because the failure was silent — looked exactly like a
blank white page with no console errors. Removing the query strings fixed it
immediately. If you need to force a fresh reload during development, use
your browser's dev tools "disable cache" option or a hard refresh, not a
version query string baked into the shipped HTML.

## 4. Local development without a real SharePoint site

You won't have a live tenant to test against during development. Structure
your app to auto-detect its environment and boot in one of two modes:

- **SharePoint mode** — `window._spPageContextInfo` exists, or a probe GET to
  `/_api/web/currentuser` succeeds. Real REST calls, real Lists.
- **Local-fallback mode** — neither is present (e.g. `python -m
  http.server` during dev). Falls back to an in-memory/localStorage seed so
  the app is still usable and testable without a tenant.

```js
async function detectSharePointSite() {
  if (window._spPageContextInfo) return window._spPageContextInfo.webAbsoluteUrl;
  for (const candidate of siteUrlCandidates()) {
    try {
      const res = await fetchWithTimeout(`${candidate}/_api/web/currentuser`,
        { headers: { Accept: "application/json;odata=nometadata" }, credentials: "same-origin" }, 6000);
      if (res.ok) return candidate;
    } catch (e) { /* try next candidate */ }
  }
  return null; // local-fallback mode
}
```

Boot **synchronously with local data first**, then attempt the SharePoint
upgrade in the background — never block the very first paint on a network
call that might hang (see §7's "async boot" postmortem for why this matters).

## 5. SharePoint integration — the universal recipe

Everything below is plain `fetch` against SharePoint's REST API. It works
from any page served from the SharePoint site (Firepit web part, or a file
in a document library) because the browser already carries the user's
session — `credentials: "same-origin"` on every request is what
authenticates you. No Graph, no Azure app registration, no npm packages.

`SITE` below is your site URL, e.g. `https://yourtenant.sharepoint.com/sites/YourSite`.

### 5.1 The one rule for all writes: get a request digest first

Every POST/MERGE/DELETE needs a fresh `X-RequestDigest` token. GETs don't.

```js
async function getDigest(SITE) {
  const r = await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST",
    credentials: "same-origin",
    headers: { Accept: "application/json;odata=nometadata" }
  });
  return (await r.json()).FormDigestValue;
}
```

Digests expire (~30 min). Fetch one per burst of writes, not one per page load.

### 5.2 Current user and identity

```js
async function getCurrentUser(SITE) {
  const r = await fetch(`${SITE}/_api/web/currentuser`, {
    credentials: "same-origin",
    headers: { Accept: "application/json;odata=nometadata" }
  });
  const u = await r.json();
  return {
    displayName: u.Title, email: u.Email, loginName: u.LoginName,
    spUserId: u.Id, isSiteAdmin: !!u.IsSiteAdmin, principalType: u.PrincipalType
  };
}
```

**Roles/permissions should NOT be derived from `IsSiteAdmin` alone.** A
site admin is a SharePoint permission concept, not necessarily your app's
"Admin" role. Build your own roles list (see §5.6) and treat SharePoint
site-admin as a bootstrap fallback only (e.g., "if no app Admin exists yet,
let a site admin in so they can run first-time setup").

### 5.3 Create a list

```js
async function createList(SITE, title, description) {
  const r = await fetch(`${SITE}/_api/web/lists`, {
    method: "POST",
    credentials: "same-origin",
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": await getDigest(SITE)
    },
    body: JSON.stringify({ Title: title, Description: description || "", BaseTemplate: 100 })
  });
  if (!r.ok) throw new Error(`createList ${r.status}: ${await r.text()}`);
  return r.json();
}
```

`BaseTemplate: 100` = generic custom list (use `101` for a document
library). Check existence first with a GET to
`/_api/web/lists/getbytitle('Title')` (404 = doesn't exist) — a list with a
duplicate exact title can't be created twice on one site anyway, so this is
mainly to avoid a 409 and to short-circuit idempotent setup runs.

### 5.4 Create columns — THE PART THAT BREAKS EVERYONE

Create each column via CAML XML through `createfieldasxml`, and pass
**`Options: 28`**. Get this wrong and every symptom below cascades from it:

| Passing `Options: 0` (or omitting it) causes | Fix with `Options: 28` |
|---|---|
| SharePoint ignores your `Name` attribute and invents an internal name like `field_7` | flag `8` (`addFieldInternalNameHint`) makes it honor your name |
| Column doesn't appear on the add/edit form | flag `4` (`addToAllContentTypes`) |
| List opens showing only the `Title` column | flag `16` (`addFieldToDefaultView`) |

```js
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
        Accept: "application/json;odata=verbose",       // this endpoint wants verbose, not nometadata
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": await getDigest(SITE)
      },
      body: JSON.stringify({
        parameters: {
          __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
          SchemaXml: fieldXml(name, type, choices),
          Options: 28   // NEVER pass 0 — see table above
        }
      })
    }
  );
  if (!r.ok) throw new Error(`createColumn ${name} ${r.status}: ${await r.text()}`);
}
```

Other rules:
- Use simple alphanumeric column names (`ProjectCode`, not `Project Code`) —
  spaces get encoded into `_x0020_` messes in the internal name.
- **Verify after creating.** Re-read `/fields?$select=InternalName,TypeAsString`
  and confirm your exact name is present. If it isn't, an earlier run (or a
  manually-created column) got mangled — it can't be renamed via REST, only
  deleted and recreated.
- If a column exists but the list still shows only `Title` when opened, the
  column just isn't in the default view (happens to columns created by an
  older version of your setup code, before you added `Options: 28`). Repair
  it directly:

```js
await fetch(
  `${SITE}/_api/web/lists/getbytitle('${listTitle}')/DefaultView/ViewFields/addviewfield('${name}')`,
  { method: "POST", credentials: "same-origin",
    headers: { Accept: "application/json;odata=nometadata", "X-RequestDigest": await getDigest(SITE) } }
);
```

Build this repair into your setup routine so it runs every time (idempotent —
skip a column that's already in the view), not just once by hand.

### 5.5 Read, create, update, delete items

```js
async function getItems(SITE, listTitle, odataQuery) {
  const all = [];
  let endpoint = `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items${odataQuery ? "?" + odataQuery : ""}`;
  while (endpoint) {
    const r = await fetch(endpoint, { credentials: "same-origin", headers: { Accept: "application/json;odata=nometadata" } });
    if (!r.ok) throw new Error(`getItems ${r.status}: ${await r.text()}`);
    const data = await r.json();
    all.push(...data.value);
    endpoint = data["odata.nextLink"] || null;   // pages at ~100 items — must follow this
  }
  return all;
}

async function createItem(SITE, listTitle, fields) {
  const r = await fetch(`${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`, {
    method: "POST", credentials: "same-origin",
    headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": await getDigest(SITE) },
    body: JSON.stringify(fields)
  });
  if (!r.ok) throw new Error(`createItem ${r.status}: ${await r.text()}`);
  return r.json(); // .Id — save it, updates/deletes need it
}

async function updateItem(SITE, listTitle, itemId, fields) {
  const r = await fetch(`${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`, {
    method: "POST", credentials: "same-origin",
    headers: { "Content-Type": "application/json;odata=nometadata", "X-RequestDigest": await getDigest(SITE),
      "IF-MATCH": "*", "X-HTTP-Method": "MERGE" },
    body: JSON.stringify(fields)
  });
  if (!r.ok) throw new Error(`updateItem ${r.status}: ${await r.text()}`);
}

async function deleteItem(SITE, listTitle, itemId) {
  const r = await fetch(`${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`, {
    method: "POST", credentials: "same-origin",
    headers: { "X-RequestDigest": await getDigest(SITE), "IF-MATCH": "*", "X-HTTP-Method": "DELETE" }
  });
  if (!r.ok && r.status !== 404) throw new Error(`deleteItem ${r.status}: ${await r.text()}`);
}
```

Property names in results are **internal names**, not display names — if
your columns were created correctly (§5.4) they're the same thing.

### 5.6 Roles and identity — one item per user

Don't reinvent per-user identity plumbing. Create a list (e.g. "App Roles")
with columns `UserEmail`, `UserDisplayName`, `SharePointUserId`, `LoginName`,
`Role` (Choice), `IsActive` (Boolean). On boot:

1. Get the current user (§5.2).
2. Look up their row by email (fall back to login name, then SharePoint user
   ID, then display name — a manually-added row is often missing one of
   these fields, so match on whichever is present).
3. No matching row → assign a default role (e.g. "Member"). Don't write a
   row until an explicit "sync users" action runs — otherwise every viewer
   creates a row on every visit.
4. Bootstrap: if literally no "Admin" role row exists anywhere, treat the
   current SharePoint **site admin** as an app Admin so first-time setup is
   always reachable, without hardcoding anyone's email.

A "sync site users" action (admin-triggered, not automatic on every boot —
see §7's "slow boot" postmortem) can batch-create rows for
`/_api/web/siteusers` entries that don't have one yet.

### 5.7 File uploads and document libraries

Uploading a binary file (not JSON/text) needs the raw bytes as the fetch
body — a `TextEncoder`-based text upload path will corrupt non-UTF8 content
like a `.docx` or image.

```js
async function ensureFolderPath(SITE, serverRelativeUrl) {
  // recursively create nested folders — check existence, create if missing,
  // recursing on the parent path first
}

async function uploadFile(SITE, folderServerRelativeUrl, fileName, arrayBuffer) {
  await ensureFolderPath(SITE, folderServerRelativeUrl);
  const r = await fetch(
    `${SITE}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')` +
    `/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`,
    {
      method: "POST", credentials: "same-origin",
      headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/octet-stream",
        "X-RequestDigest": await getDigest(SITE) },
      body: arrayBuffer   // NOT JSON.stringify'd, NOT TextEncoder'd — raw ArrayBuffer/Blob
    }
  );
  if (!r.ok) throw new Error(`upload ${r.status}: ${await r.text()}`);
  return r.json();
}
```

### 5.8 Email notifications, no server required

SharePoint's own `SP.Utilities.Utility.SendEmail` REST endpoint sends mail
without Graph, Power Automate, or any separate mail infrastructure:

```js
async function sendEmail(SITE, { to, subject, body }) {
  const r = await fetch(`${SITE}/_api/SP.Utilities.Utility.SendEmail`, {
    method: "POST", credentials: "same-origin",
    headers: { Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": await getDigest(SITE) },
    body: JSON.stringify({
      properties: {
        __metadata: { type: "SP.Utilities.EmailProperties" },
        To: { results: Array.isArray(to) ? to : [to] },
        Subject: subject, Body: body
      }
    })
  });
  if (!r.ok) throw new Error(`sendEmail ${r.status}: ${await r.text()}`);
}
```

Some tenants disable outbound mail via this endpoint by policy. Treat
notification failures as non-fatal — the underlying data save should already
have succeeded before you attempt to notify anyone, and a missed email
shouldn't roll back or block real work.

### 5.9 Optimistic concurrency (ETags)

For a scenario where the whole document might be overwritten wholesale
(rare if you're doing per-record CRUD per §5.5, but relevant for a
single-file JSON blob approach), use `If-Match` on writes:

```js
async function getFileETag(SITE, serverRelativeUrl) {
  const r = await fetch(`${SITE}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')?$select=ETag`,
    { credentials: "same-origin", headers: { Accept: "application/json;odata=nometadata" } });
  return (await r.json()).ETag;
}
// Pass that ETag as an If-Match header on the write; a 412/409 means
// someone else saved first — re-read, merge, retry.
```

Prefer **per-record CRUD** (§5.5) over "read the whole thing, mutate, write
the whole thing back" wherever practical — it avoids most concurrency
problems structurally instead of needing conflict-merge logic at all.

## 6. Recommended app architecture

A clean separation that scales well for this kind of app:

```
app-config.js       — the ONE place secrets/config/endpoints live
sharepoint-adapter.js — every raw REST call (digest, CRUD, field/list/view
                        creation, roles, email, file upload, diagnostics).
                        Nothing else in the app calls fetch() directly.
sharepoint-schema.js — your list/column schema as data + setup/check routines
                        that create what's missing and repair what's wrong
sharepoint-repo.js   — maps between your app's in-memory object shapes and
                        SharePoint list-item property names; exposes a
                        generic save(kind, obj) / remove(kind, obj) that
                        pages call instead of touching SharePoint directly
app.js               — boot sequence, mode detection (SharePoint vs. local
                        fallback), shared UI shell/router
pages/*.js           — one file per feature area, calling Repo.save/remove
```

### 6.1 Boot sequence

```
detect mode (SharePoint vs. local fallback)
  → resolve current user identity
  → resolve their role (cheap: read whatever's already in the roles list)
  → check/repair schema status
  → load data
  → render
  → (in the background, NOT blocking the above): sync all site users into
    the roles list, refresh in place if it changes anything for this user
```

The "sync all site users" step is the slowest thing in boot (a full
site-users fetch plus a create/update per changed user) and almost nobody
needs it done before they can see their actual data — see §7.

### 6.2 Per-record save with debounce + serialization

If a UI interaction can fire many rapid saves (dragging a task on a
timeline, resizing a date range), debounce them per-object and chain writes
so a create always finishes — and stamps the new item's ID — before a
follow-up update for the same object runs:

```js
const saveTimers = new Map();      // obj -> {timer, kind}
const saveChains = new WeakMap();  // obj -> promise chain (serializes writes per object)
const pendingSettlers = new Map(); // obj -> shared promise, resolved when the write actually lands

function repoSave(kind, obj) {
  const pending = saveTimers.get(obj);
  if (pending) clearTimeout(pending.timer);
  let settler = pendingSettlers.get(obj);
  if (!settler) { let resolve; settler = { promise: new Promise(r => resolve = r), resolve }; pendingSettlers.set(obj, settler); }
  const timer = setTimeout(() => {
    saveTimers.delete(obj);
    const prev = saveChains.get(obj) || Promise.resolve();
    const next = prev.then(() => writeRecord(kind, obj)).finally(() => { pendingSettlers.delete(obj); settler.resolve(); });
    saveChains.set(obj, next);
  }, 500);
  saveTimers.set(obj, { timer, kind });
  return settler.promise; // resolves once the ACTUAL write lands, not just when debounced
}
```

Callers that need the new object's server-assigned ID before doing
something dependent (like creating child records with a foreign key)
**must** `await Repo.save(...)` and rely on it resolving only once the write
is real — see §7's "fire-and-forget save" postmortem for what goes wrong if
it doesn't.

## 7. Postmortems — real bugs, in the order we hit them

Each of these cost real debugging time. Read them before you build the
equivalent piece yourself.

1. **Mangled column names (`Options: 0`).** Covered in §5.4. Root cause of
   nearly every "columns are missing/duplicated/blank form" symptom this
   project hit. Fix once, verify with a post-create name check, and never
   regress it.

2. **Cache-busting query strings silently 404 on SharePoint.** Covered in
   §3.1. Symptom: totally blank page, zero console errors, because the
   `<script src="app.js?v=10">` request itself failed and nothing downstream
   ever ran.

3. **Async boot blocking first paint.** Making the boot handler `async` and
   awaiting SharePoint calls before the first render meant a hung or slow
   network call (common on government/CAC-gated networks) left the page
   blank indefinitely. Fix: render synchronously with whatever local data
   you have immediately, then upgrade to SharePoint data in the background.

4. **A crash in the shared load path silently falling back to empty state.**
   A data-loading function called a normalization step that read a field
   (`db.user.id`) which wasn't set yet at that point in the call chain — it
   threw on literally every load, and the surrounding `try/catch` silently
   swallowed it and substituted an empty database. Every symptom that looked
   like "my data disappeared on refresh" traced back to this one throw.
   Lesson: audit exactly what a shared function assumes is already set
   before it runs, and don't let a broad catch-and-fallback block hide a
   deterministic bug behind an intermittent-looking symptom.

5. **Fire-and-forget saves.** A `Repo.save()` that returned before the
   network write actually completed meant `await Repo.save(parent)` gave
   false confidence — code immediately after it, which created child
   records referencing the parent's new ID, ran before that ID existed,
   silently writing a null/zero foreign key. Fix: make save() return a
   promise that resolves only when the real write lands (§6.2).

6. **A debug aid that clobbered the live app.** A boot-logging function
   unconditionally rebuilt a "startup console" screen by wiping the app
   root element — harmless during boot, but it had no guard against being
   called again *after* boot completed. Once anything post-boot (a manual
   "reload data" action, a background refresh) logged a message through the
   same function, it wiped the live, already-rendered app back to the boot
   screen. Fix: any diagnostic/logging utility that touches the DOM needs an
   explicit "has boot finished?" guard baked in from the start, not added
   after the fact.

7. **Slow, blocking user-directory sync.** Syncing every SharePoint site
   user into an app-roles list on every single boot — before the user could
   see any of their actual data — made the app feel slow for a step almost
   nobody needed done synchronously. Fix: run it in the background after the
   real data has rendered, and only apply its result to the live UI if it
   actually changes something for the current user (e.g., their own role).

8. **The wrong build tool for the target.** An offline/portable packaging
   path (heavy iframe-`srcdoc` wrapper, base64-embedding the entire child
   app) was initially used for what was actually a SharePoint/Firepit ship.
   It technically produced a loadable file, but the flat single-file format
   with inlined assets (§2.1, §3) is what the platform's own SharePoint ship
   path produces and is the one that behaves predictably in a Firepit web
   part. If your tooling offers a distinctly-named SharePoint/Flank-Speed
   export option, use that one — don't default to a generic/offline export
   just because it's the first one you find.

## 8. New-app checklist

- [ ] `index.html` with no external script/link tags — everything local
- [ ] `app-config.js` — site URL override, feature flags, any API keys
      (clearly marked where to put them), nothing else duplicates these values
- [ ] `sharepoint-adapter.js` — digest, CRUD, list/field/view creation with
      `Options: 28`, roles, email, file upload, all diagnostics in one place
- [ ] `sharepoint-schema.js` — your schema as data + idempotent setup/check
- [ ] `sharepoint-repo.js` — object↔item mappers, `Repo.save`/`Repo.remove`
      with real (not fire-and-forget) promises
- [ ] Boot: synchronous local render first, SharePoint upgrade in background,
      slow steps (user sync) deferred further still
- [ ] A packaging script producing the flat single-file WFC-manifest format
- [ ] No `?v=N` cache-busting on script tags in the shipped file
- [ ] Verified in a real Firepit web part before considering it done — local
      dev / a plain browser tab cannot fully substitute for this

## 9. What's out of scope here

This guide covers the plumbing: packaging, identity, CRUD, files, email.
It does not cover SharePoint permission/governance topics (who can create
lists on a site, tenant-level REST restrictions, DLP policies) — those are
organization-specific and should be confirmed with whoever administers your
SharePoint tenant before you assume a given REST call will succeed there.
