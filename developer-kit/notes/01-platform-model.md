<!-- UNCLASSIFIED -->

# 01 — The platform model

**UNCLASSIFIED**

Three named things sit between your code and the user. They are easy to confuse,
and confusing them is how people waste their first week.

| Name | What it actually is |
|---|---|
| **SharePoint Online** (Flank Speed) | The origin your app is served from, the identity provider, and the database. |
| **Firepit** | An SPFx web part that renders one blob of HTML inside a sandboxed iframe on a SharePoint page. Version read here: **7.1.0**. |
| **Forge** | A browser-based IDE that compiles a folder of HTML/CSS/JS into the single self-contained file Firepit renders. |

Forge is a build tool. Firepit is a runtime host. Neither is a framework, and
neither is required to write the app — you can produce a Firepit-compatible file
with a 200-line Node script (`notes/03`) and edit it in any editor you like.

---

## The whole stack, top to bottom

```
User signs in to Flank Speed (CAC/PIV/SSO)
        │
        ▼
SharePoint page containing a Firepit web part
        │
        │  the web part fetches your .html from a document library
        │  (or reads it out of its own stored properties)
        ▼
Sandboxed iframe, loaded from a blob: URL, with an injected CSP
        │
        │  your app is now running, with the user's session cookie
        ▼
        ├──► SharePoint REST   /_api/web/...        same-origin, no tokens
        ├──► Dataverse         via postMessage to the host web part
        └──► An allow-listed API host, if the CSP permits it
```

Everything downstream of "sandboxed iframe" is your app. Everything upstream is
platform you do not control.

---

## The four consequences that shape every decision

### 1. One file, everything inlined

The host hands the browser a single HTML document. There is no second request for
`app.js`. Every stylesheet, every script, every font and image must be inlined —
`<style>` blocks, `<script>` blocks, and `data:` URIs — before the file ships.

This is not a style preference. The host injects `default-src 'none'` into your
frame (`notes/02`), so a file that references anything by URL gets a blank page
and a console full of CSP violations.

### 2. Authentication is already done, and you must not re-do it

Your page runs on the SharePoint origin with the user's session cookie. That is
the entire authentication story:

```js
const r = await fetch(`${SITE}/_api/web/currentuser`, {
  credentials: "same-origin",
  headers: { Accept: "application/json;odata=nometadata" }
});
```

No Azure AD app registration. No MSAL. No token cache, no refresh flow, no
consent screen. If you find yourself reaching for an auth library, stop — you
have misread the platform. See [connectors/sp-rest-core.js](../connectors/sp-rest-core.js).

The one thing this does *not* give you is authorization. `IsSiteAdmin` is a
SharePoint permission, not your app's "admin" role; build your own roles list
(`notes/05`).

### 3. SharePoint is the database, and it has opinions

Two viable storage shapes, and this program has shipped both:

| Shape | Use when | Cost |
|---|---|---|
| **Lists** — a list per entity, columns per field | You want data queryable and editable in SharePoint's own UI, or Power Automate needs to see it | Column creation is genuinely treacherous — see [connectors/sp-columns.js](../connectors/sp-columns.js) |
| **JSON in a document library** — one `db.json`, read/modify/write | You want to move fast and own your schema | You hand-roll concurrency; use ETags |

PULSE started with lists, hit column-naming failures severe enough to change
architecture mid-flight, and moved to a JSON store. Both connectors are in this
kit. Read the honest version of that decision in `notes/05` before choosing.

### 4. The web part's width is not the page's width

A Firepit web part can sit in a narrow column beside other web parts. Responsive
layout is a requirement, not polish. `fullScreen` is available but is
deliberately disabled while the page is in edit mode, so the author can still
reach the property pane.

---

## Where each piece lives in this kit

| Piece | Where |
|---|---|
| Firepit web part source (TypeScript, 1,750 lines) | [firepit-webpart/src/](../firepit-webpart/src/) |
| Firepit deployment package + manifests | [forge/firepit-prod.sppkg](../forge/), [firepit-webpart/sppkg-manifests/](../firepit-webpart/sppkg-manifests/) |
| Forge IDE, compiled | [forge/Forge.html](../forge/Forge.html) |
| Forge IDE, unpacked into readable modules | [forge-ide-decompiled/](../forge-ide-decompiled/) |
| A complete app you can read end to end | [examples/pulse-code-ide/](../examples/pulse-code-ide/) |
| Working integration code | [connectors/](../connectors/) |

---

## What is verified and what is not

**Verified** (read directly from the artifacts in [forge/](../forge/)):
the Firepit web part's rendering path, sandbox attributes, injected CSP, property
schema, chunking limits, and postMessage protocol; the WFC manifest format; Forge's
compiler and decompiler behaviour, including its hash function.

**Inferred, and labelled as such where it matters:** anything about Forge's or
Firepit's roadmap, server-side behaviour, or tenant configuration. Those belong to
their maintainers. Tenant policy in particular — whether a given REST call is
permitted on your site — must be confirmed with whoever administers the tenant,
not assumed from this document.

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
