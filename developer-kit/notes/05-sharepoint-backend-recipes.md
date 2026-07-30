<!-- UNCLASSIFIED -->

# 05 — SharePoint as your backend

**UNCLASSIFIED**

Plumbing: identity, storage, files, mail. Working code for all of it is in
[connectors/](../connectors/); this document is the decisions and the rules, not a
second copy of the code.

---

## 1. The one rule for reads and the one rule for writes

**Reads:** `credentials: "same-origin"` and you are authenticated. That is all.

**Writes:** get a request digest first.

```js
const r = await fetch(`${SITE}/_api/contextinfo`, {
  method: "POST",
  credentials: "same-origin",
  headers: { Accept: "application/json;odata=nometadata" },
});
const digest = (await r.json()).FormDigestValue;   // → X-RequestDigest header
```

Digests live about 30 minutes. Fetch one per burst of writes, not one per write.
A 403 on a write is *usually a stale digest, not a permission problem* — retry
once with a fresh one before you surface an error, and only once, so a real
permission failure still fails fast. [sp-rest-core.js](../connectors/sp-rest-core.js)
does both.

**`odata=nometadata` everywhere**, for smaller responses — with two exceptions
that require `odata=verbose`:

- `fields/createfieldasxml`
- `SP.Utilities.Utility.SendEmail`

Get that backwards and you get a 400 that reads like a malformed payload.

---

## 2. Choosing your storage shape

This is the decision to make deliberately, because changing it later is a
migration. PULSE made it twice.

| | **Lists** | **JSON in a document library** |
|---|---|---|
| Schema | Columns you create via REST | Whatever your objects are |
| Query | Server-side `$filter`/`$orderby` | Load all, filter in memory |
| Scale | Large; paging built in | Thousands of records, not millions |
| SharePoint UI | Users can view, sort, export | Nothing; your app is the only reader |
| Power Automate | Triggers on item change | No triggers |
| Concurrency | Per-item, mostly free | Yours to handle, with ETags |
| Setup risk | **High** — see §3 | Low |

**What happened here.** PULSE started on lists. Column creation in this tenant
kept producing generic internal names (`field_1`, `field_7`), occasional wrong
types, and hidden duplicates. The visible symptom was the worst kind: writes
returned 201 and the SharePoint form stayed empty, because the app was writing to
a property name SharePoint had silently replaced. The team judged the list/column
model too fragile and moved to a document library holding `db.json` and
`roles.json`.

**The honest retrospective:** that failure was `Options: 0` on column creation
(§3), and it was fixable. The JSON store is a legitimate architecture and PULSE
ships on it — but if you are starting fresh, know that "lists were too fragile"
was really "we did not know about `Options: 28`". Choose on the tradeoffs in the
table, not on that history.

---

## 3. Creating columns — the part that breaks everyone

Pass **`Options: 28`** to `createfieldasxml`. Always.

```
28 = 4 | 8 | 16
     4  addToAllContentTypes      → appears on the add/edit form
     8  addFieldInternalNameHint  → SharePoint honours YOUR Name attribute
    16  addFieldToDefaultView     → appears when someone opens the list
```

| Symptom | Missing flag |
|---|---|
| Internal name came out `field_7` | 8 |
| Column absent from the new/edit form | 4 |
| List opens showing only `Title` | 16 |
| Writes return 201 but the form is blank | 8 |

Then three habits:

1. **Alphanumeric names only.** `ProjectCode`, not `Project Code` — a space
   becomes `_x0020_` in the internal name and every query against it is an
   escaping puzzle forever.
2. **Verify after creating.** Re-read `/fields?$select=InternalName,TypeAsString`
   and confirm your exact names are present. **A column cannot be renamed through
   REST** — if it came out wrong the only fix is delete and recreate, and you want
   to know that now, not after three weeks of data.
3. **Repair the default view every run.** Columns created by an older build that
   predates `Options: 28` exist but are invisible. `addviewfield` is idempotent;
   put it in your setup routine rather than doing it once by hand.

All of this, done properly: [sp-columns.js](../connectors/sp-columns.js).

---

## 4. Items

`BaseTemplate: 100` is a custom list, `101` a document library.

Two things that will bite:

**Paging.** Lists page at ~100 items. Follow `odata.nextLink` until it is absent.
Code that reads `.value` once works perfectly until the list grows past a page,
then silently shows partial data — and nobody notices until a report is wrong.

**Field shapes.** Most types take a plain value. These do not:

| Type | How to write it |
|---|---|
| Person/Group | `{ AssignedToId: 42 }` — the `…Id` property, by SharePoint user id |
| Multi-person | `{ ReviewersId: { results: [42, 51] } }` |
| MultiChoice | `{ Tags: { results: ["A", "B"] } }` |
| Lookup | `{ ProjectId: 17 }` |
| URL | `{ Link: { Url: "https://…", Description: "text" } }` |
| DateTime | an explicit ISO 8601 string |

[sp-list-crud.js](../connectors/sp-list-crud.js) has a `values` helper per shape.

---

## 5. Files

**Binary must go as bytes.** Pass an `ArrayBuffer` or `Blob` as the fetch body.
A `TextEncoder`- or string-based path corrupts anything non-UTF-8 — a `.docx`
becomes unopenable, a `.png` becomes noise — and the request returns 200. This is
the single easiest way to ship a silent data-corruption bug in this stack.

**Create folders parents-first.** SharePoint will not create intermediate folders,
and the failure is a 404 on the leaf that reads like a permissions error.

**Resolve a library's real root folder** rather than guessing its URL from its
title. `RootFolder/ServerRelativeUrl` is authoritative; "PULSE App Data" may live
at any of several paths depending on how it was created.

**Use ETags when you overwrite a whole document.** `If-Match` with a real ETag
turns a blind overwrite into a checked one; a 412 or 409 means someone saved first,
so re-read, re-apply, retry. Prefer per-record CRUD where you can — it avoids the
problem structurally instead of needing merge logic.

[sp-files.js](../connectors/sp-files.js), [sp-json-store.js](../connectors/sp-json-store.js).

---

## 6. Identity and roles

`/_api/web/currentuser` gives you `Title`, `Email`, `LoginName`, `Id`,
`IsSiteAdmin`, `PrincipalType`.

**`IsSiteAdmin` is not your app's Admin role.** It is a SharePoint permission.
Build your own roles list — `UserEmail`, `UserDisplayName`, `SharePointUserId`,
`LoginName`, `Role` (Choice), `IsActive` (Boolean) — and treat site-admin purely
as a bootstrap: *if no app Admin exists anywhere yet, let a site admin in so
first-run setup is reachable.* Never hardcode an email.

Three details that matter more than they look:

- **Match a user to their row on whatever field is present** — email, then login
  name, then user id, then display name. Rows added by hand through the SharePoint
  UI are routinely missing one of these, and a single-field lookup fails to find
  someone who is plainly there. To them, the app forgot who they are.
- **Do not write a row during role resolution.** If resolution provisions, every
  viewer who opens the app creates a record and your roles list becomes a visitor
  log. Provision from an explicit, admin-triggered sync.
- **Keep capabilities in one table.** Scattering `if (role === "Admin")` through
  the UI guarantees an inconsistency, and read-only users are the ones who find
  it. Every screen has a real view-only audience; read-only states must look
  intentional, not broken.

[sp-identity-roles.js](../connectors/sp-identity-roles.js).

---

## 7. Mail, with no server

`SP.Utilities.Utility.SendEmail` sends as the site, using the user's session — no
Graph, no Power Automate, no mail infrastructure. Requires `odata=verbose`.
Recipients must already have site access.

**Treat notification failure as non-fatal.** Some tenants disable outbound mail
from this endpoint by policy. Save the data first, notify second, and never let a
missed email roll back or block real work. A user whose travel request vanished
because a confirmation email failed will not forgive the design.

[sp-email.js](../connectors/sp-email.js) returns `{sent, reason, policyBlocked}`
rather than throwing, so a caller cannot make mail fatal by forgetting a
try/catch.

---

## 8. Recommended layering

```
app-config.js          the ONE place site URLs, feature flags, and keys live
sp-rest-core.js        every raw fetch. Nothing else calls fetch() on SharePoint
sp-columns.js /        schema as data, plus idempotent setup and repair
  sp-json-store.js
repo.js                maps your in-memory objects ↔ storage property names;
                       exposes save(kind, obj) / remove(kind, obj)
app.js                 boot, mode detection, shell, router
pages/*.js             one file per feature area, calling repo.*
```

The point of the single-fetch rule: when a call starts failing in production you
want exactly one function to instrument, and one place holding the last URL,
status, and error text for your diagnostics screen. Capture those three in every
bug report and most "it doesn't work" tickets answer themselves.

### Boot order

```
detect mode (SharePoint reachable, or local fallback)
  → resolve current user
  → resolve their role from whatever is already stored
  → check/repair schema
  → load data
  → render
  → THEN, in the background: sync site users, refresh in place only if it
    changes something for this user
```

Render before you fetch, and keep slow steps off the critical path. Both halves of
that sentence are postmortems, not preferences — see `notes/06`.

---

## 9. Out of scope here

Tenant governance: who may create lists on a site, tenant-level REST
restrictions, DLP policy. Those are organization-specific and must be confirmed
with whoever administers your SharePoint tenant. Do not assume a call that works
on your dev site is permitted everywhere.

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
