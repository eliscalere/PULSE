<!-- UNCLASSIFIED -->

# 02 — What the Firepit host actually does to your HTML

**UNCLASSIFIED**

Everything in this document was read from
[firepit-webpart/src/webparts/htmlFileRenderer/HtmlFileRendererWebPart.ts](../firepit-webpart/src/webparts/htmlFileRenderer/HtmlFileRendererWebPart.ts)
(1,750 lines, Firepit **7.1.0**) and the manifests in
[firepit-webpart/sppkg-manifests/](../firepit-webpart/sppkg-manifests/). Line
references are to that file. If your Firepit version differs, re-read it — this
is the most version-sensitive document in the kit.

This is the document to read first. The host's injected CSP is the constraint
behind nearly every "why can't I just…" question in this stack.

---

## 1. Identity of the thing

| Property | Value |
|---|---|
| Solution | `firepit-client-side-solution` |
| Solution version | `7.1.0.0` |
| Product ID | `6dd7e353-52da-4020-8eaf-4ff0de1ec35a` |
| Web part class | `HtmlFileRendererWebPart` |
| Web part ID | `4b2ad7ba-a39b-49e8-90e3-10e08433578b` |
| Display name | **Firepit** (toolbox group: *Advanced*) |
| `requiresCustomScript` | `false` — installs on sites with Custom Script disabled |
| `SkipFeatureDeployment` | `true` — tenant-wide deployable |
| Supported hosts | `SharePointWebPart`, `SharePointFullPage`, `TeamsTab`, `TeamsPersonalApp` |
| Requested API permission | **Microsoft Dataverse**, `user_impersonation` |

That last row matters: the package requests a Dataverse scope at the tenant level.
Your app never sees a Dataverse token — it asks the host to make the call
(section 6).

---

## 2. Where your HTML comes from

Two sources, checked in order (`render()`, line 155):

**A file in SharePoint** — the `htmlFileUrl` property. The host fetches it with
`SPHttpClient` and renders the response body. This is the normal path: upload
your built `.html` to a document library, point the web part at it, and
redeploying is a file overwrite.

**Inline content** — the `htmlCode` property, stored in the web part's own
configuration. Because a single SPFx property cannot hold a large document, the
host splits content across numbered properties:

| Constant | Value |
|---|---|
| `CHUNK_SIZE` | 250,000 chars per chunk |
| `MAX_CHUNKS` | 120 |
| Effective ceiling | ~30 MB |

Chunks are written as `htmlCode_0 … htmlCode_N` with `htmlCodeChunkCount`
recording how many, and reassembled by concatenation on load (lines 424–470).

> **Sharp edge.** Content exceeding 120 chunks is **truncated**, and the only
> signal is a `console.warn` (line 456–457). It does not throw and the page does
> not visibly fail — you get a silently half-loaded app. If you ship inline and
> your file is in the tens of megabytes, check that warning. The file-URL path
> has no such limit; prefer it for anything large.

---

## 3. It renders through a `blob:` URL, not `srcdoc`

This is the single most surprising implementation detail, and the source explains
why in a comment worth quoting almost in full (lines 227–240):

```js
// Build container without iframe in innerHTML to avoid MCAS js-wrapper
// intercepting innerHTML and re-executing scripts inside the iframe
...
// Create iframe programmatically and load via blob URL.
// blob: URLs bypass MCAS js-wrapper interception of srcdoc content,
// which otherwise causes scripts to be double-executed ("Identifier X
// has already been declared" errors for every const/let in the app).
const iframe = document.createElement('iframe');
iframe.setAttribute('sandbox', sandboxAttr);
const blob = new Blob([this._htmlContent], { type: 'text/html' });
iframe.src = URL.createObjectURL(blob);   // src set BEFORE append to DOM
```

Read that error message carefully, because you will meet it:

> **`Identifier 'X' has already been declared`**, thrown for every top-level
> `const`/`let` in your app.

That is not a bug in your code. It is Microsoft Defender for Cloud Apps (MCAS)
session-proxying the page, wrapping DOM APIs, and executing your scripts twice.
Firepit works around it by never putting the iframe in `innerHTML` and never
using `srcdoc` — the blob URL and the pre-append `src` assignment both exist to
get ahead of a `MutationObserver`-based interception.

Two things follow for you:

- **Do not assume `srcdoc` semantics.** Your app loads from an opaque
  `blob:` origin, not the SharePoint origin.
- **If you build your own host, copy this.** It is the difference between an app
  that runs and an app that throws on every declaration.

The blob URL is revoked after load to avoid leaking.

### Sandbox attributes

```
allow-scripts allow-modals allow-popups allow-popups-to-escape-sandbox
allow-forms allow-same-origin allow-downloads
```

`allow-same-origin` is why your same-origin SharePoint REST calls work at all.
Absent from the list, notably: `allow-top-navigation`.

---

## 4. The injected Content-Security-Policy — read this twice

The host builds a `<meta http-equiv="Content-Security-Policy">` and injects it
into *your* document (lines 902–919). Verbatim directive set:

```
default-src 'none';
script-src 'unsafe-inline' 'unsafe-eval';
style-src  'unsafe-inline';
worker-src blob:;
connect-src 'self' <sharePointOrigin> <mcasOrigins…> <dataverseOrigin?>;
img-src    data: blob: 'self';
font-src   data:;
media-src  'none';
manifest-src 'none';
form-action 'none';
frame-src  <sharePointOrigin> <mcasWrappedOrigin>;   /* or 'none' */
object-src 'none';
```

Every constraint this stack seems to impose arbitrarily is right here:

| The rule you keep hitting | The directive causing it |
|---|---|
| No CDN scripts or stylesheets, ever | `default-src 'none'` with no `https:` in `script-src`/`style-src` |
| Fonts must be `data:` URIs | `font-src data:` — that is the *only* allowed source |
| Images must be inlined or same-origin | `img-src data: blob: 'self'` |
| Inline `<script>`/`<style>` do work | `'unsafe-inline'` on both |
| `eval` and `new Function` work | `'unsafe-eval'` on `script-src` |
| `fetch` to a random API host fails | `connect-src` allow-list |
| Native `<form>` submits do nothing | `form-action 'none'` — use `fetch` |
| Web workers need a blob | `worker-src blob:` |
| No Flash/applets/embeds | `object-src 'none'` |

**Practical consequence:** an app that works when you open it from your desktop
can fail completely inside Firepit, and the reverse is never true. Test in a real
web part before calling anything done.

### The `lockDown` property

Setting `lockDown` on the web part replaces two directives:

```
connect-src 'none'      /* no outbound requests AT ALL — including SharePoint REST */
frame-src   'none'
```

A locked-down Firepit web part cannot talk to SharePoint. It is for rendering
static content. If your app boots to "cannot reach SharePoint" and the code looks
right, check this toggle before you debug anything else.

### CSP nonce

If the host page runs a nonce-based CSP, Firepit reads it and forwards it:
`window.__firepitCspNonce` is set inside your frame (line 874) and the host
rewrites your `<script>` tags to carry it (`_applyNonceToScriptTags`, line 950).
You do not normally touch this — but if you generate `<script>` elements at
runtime, that global is where the nonce you need lives.

### Additional injected headers

`X-Content-Type-Options: nosniff` · `Referrer-Policy: no-referrer` ·
`X-XSS-Protection: 1; mode=block` · `x-dns-prefetch-control: off` ·
`Permissions-Policy` denying geolocation, payment, USB, magnetometer, gyroscope,
accelerometer, ambient-light, autoplay, encrypted-media, fullscreen,
picture-in-picture, and screen-wake-lock.

That `Permissions-Policy` line is worth noting: **`fullscreen=()` is denied**, so
the Fullscreen API will not work from inside the frame. The web part's own
`fullScreen` property — which resizes the container on the SharePoint page — is a
different mechanism and does work.

---

## 5. Five security layers, not one

The CSP is only the first. `_processHtmlCode()` (line 843) enumerates the set in
its own doc comment, and describes them as *"ported from the Forge compiler"* —
which is why the same guards turn up inside compiled Forge apps as well:

> 1. CSP meta tag restricts connect-src to SharePoint origin, or none in lock-down
> 2. Additional security meta tags
> 3. Link-hint guard neutralizes link hints that can trigger outbound requests
> 4. Query-param guard strips query params from navigation URLs
> 5. Network guard blocks non-SharePoint fetch, XHR, WebSocket, EventSource, and sendBeacon

### The network guard is the one that will surprise you

`_buildNetworkGuardScript()` (line 1458) injects code into your document that
**monkeypatches the browser's networking APIs** — `window.fetch`,
`XMLHttpRequest.prototype`, `WebSocket`, `EventSource`, and
`navigator.sendBeacon` — and checks every request origin against an allow-list
(SharePoint origin, the MCAS-wrapped origin, and the Dataverse origin when
configured). The comment at line 940 labels it *"Data-exfiltration prevention"*.

This is enforcement **independent of the CSP**, in your own JS context. Two
consequences:

- **An origin being CSP-allowed is not sufficient.** Both layers must permit a
  request. Getting an AI or API host added means getting it into both.
- **The failure looks different.** A CSP block logs a violation in the console. A
  network-guard block comes back as a rejection from a patched `fetch` — so a
  `TypeError`/rejected promise with no CSP violation anywhere is still a platform
  block, not your bug.

Per the architecture guide in [../forge/](../forge/), the allow-list in this
tenant includes specific government AI endpoints (`api.genai.mil`,
`api.genai.army.mil`, `api.capra.flankspeed.us.navy.mil`). Treat that as tenant
configuration to confirm with your administrator, not as a stable contract — and
see [../connectors/ai-openai-compatible.js](../connectors/ai-openai-compatible.js)
for what this means when you wire up an LLM.

### DOM guards

A `MutationObserver` with
`attributeFilter: ["href","action","method","src","data","srcdoc","content","http-equiv"]`
re-sanitizes those attributes as they change (lines ~1330–1455). Observed
policies:

- **Frames** are permitted only to `/_layouts/15/embed.aspx` (line 1340).
- **`srcdoc` on any nested iframe/frame is sanitized** (lines 1388–1423) — you
  cannot smuggle content in through a child frame.
- **`mailto:` links** are restricted to the parameters `subject`, `body`, `cc`,
  `bcc`.
- **`<link rel>`** is restricted to `stylesheet`, `icon`, `shortcut`,
  `apple-touch-icon`, `apple-touch-icon-precomposed` — this is layer 3, since
  `rel="preconnect"`/`prefetch`/`dns-prefetch` would each trigger an outbound
  request without any script running.

### You will find these guards twice

The three untracked inline scripts at the top of the decompiled Forge tree —
[`_app-inline/head-01.js`](../forge-ide-decompiled/_app-inline/), `head-02.js`,
`head-03.js` — are Forge's own copies, and their error strings say so: `Forge
child bridge bootstrap failed`, `Forge query param guard init failed`, `Forge link
hint guard init failed`.

That duplication is deliberate. Forge's compiler embeds the guards into everything
it ships; Firepit re-injects them at render time so that hand-written HTML — an app
that never went through Forge — is constrained identically. Neither layer trusts
the other to have done it.

---

## 6. Talking to the host: the postMessage protocol

Four message types, all `window.parent.postMessage`, all validated by the host
against `event.source === iframe.contentWindow` before acting:

| Message `type` | Direction | Purpose |
|---|---|---|
| `firepit:open-edit-panel` | child → host | Open the SharePoint property pane (hotkey passthrough) |
| `firepit:dataverse-request` | child → host | Ask the host to make a Dataverse call on your behalf |
| `firepit:dataverse-response` | host → child | Correlated reply, keyed by `requestId` |
| `firepit:destructive-command-approved` | child → host | Record that the user approved a destructive operation for this session |

**The Dataverse proxy** is the interesting one. Your frame cannot get an AAD
token — but the host can, through `AadHttpClient` and the tenant-approved
`user_impersonation` scope. So you post a request `{url, init:{method, headers,
body}}` with a `requestId`, and the host performs it and posts back
`{ok, status, statusText, headers, body}`. Requires the `dataverseEnvironmentUrl`
property to be set.

**Destructive-operation throttle** (lines 65–66): mutations are rate-limited to
**5 per 60-second window** per target, and there is a session approval gate the
user must pass. Build for it — a bulk import that fires more than 5 mutations in
a minute will be refused, and the refusal is a platform decision your code cannot
override.

A minimal, commented client for all of this:
[connectors/firepit-bridge.js](../connectors/firepit-bridge.js).

---

## 7. Property reference

| Property | Type | Notes |
|---|---|---|
| `htmlFileUrl` | string | Source file in SharePoint. Preferred for large apps. |
| `htmlCode` | string | Inline content; auto-chunked when large. |
| `htmlCode_0 … htmlCode_N` | string | Storage chunks. Do not hand-edit. |
| `htmlCodeChunkCount` | number | How many chunks are valid. |
| `fullScreen` | boolean | Ignored while the page is in edit mode, by design. |
| `iframeHeight` | string | CSS length. Default `600px`. |
| `lockDown` | boolean | `connect-src 'none'` + `frame-src 'none'`. See §4. |
| `dataverseEnvironmentUrl` | string | Enables the Dataverse proxy and adds that origin to `connect-src`. |

---

## 8. Checklist before you blame your own code

1. Is `lockDown` on? Then nothing outbound works, including SharePoint.
2. Console full of CSP violations? You referenced something by URL. Inline it.
   And note the reverse: a rejected `fetch` with **no** CSP violation is still
   probably a platform block — the network guard (§5) enforces separately.
3. `Identifier 'X' has already been declared`? MCAS double-execution, not you
   (§3). Confirm you are on a Firepit version that uses the blob path.
4. App loads but truncated? Inline-content chunk ceiling (§2).
5. Fonts not rendering? `font-src data:` — they must be base64 in your CSS.
6. `<form>` submit doing nothing? `form-action 'none'`. Use `fetch`.
7. Fullscreen API rejected? `Permissions-Policy: fullscreen=()`. Use the web
   part's own `fullScreen` property.
8. Bulk writes failing after the fifth? Destructive-operation throttle (§6).

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
