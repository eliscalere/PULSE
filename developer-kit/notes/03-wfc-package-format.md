<!-- UNCLASSIFIED -->

# 03 — The WFC package format, and how to build one

**UNCLASSIFIED**

A shipped app is one HTML file with a base64 JSON manifest in a leading comment.
That is the entire format. You do not need Forge to produce it, and this document
shows both how to read it and how to emit it in about 200 lines of Node.

---

## 1. The shipped file

```html
<!--WFC-MANIFEST:eyJ2ZXJzaW9uIjoxLCJwcm9qZWN0Ijoi...-->
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Your App</title>
<style>/* every stylesheet, inlined */</style>
</head>
<body>
<!-- your markup -->
<script>/* every script, inlined */</script>
</body>
</html>
```

Below the comment it is an ordinary HTML document with nothing external.
Firepit does not need the manifest to render the page — it exists for
traceability and for decompiling.

## 2. The manifest

Base64-decode the comment payload and you get:

```json
{
  "version": 1,
  "project": "PULSE",
  "generated": "2026-07-29T22:46:06.778Z",
  "index": "index.html",
  "files": [
    { "kind": "css", "path": "assets/css/style.css", "external": false, "hash": "b9624e43" },
    { "kind": "js",  "path": "assets/js/app.js",     "external": false, "hash": "22c73a00" }
  ]
}
```

| Field | Meaning |
|---|---|
| `kind` | `css` · `js` · `img` · `wasm` · `js-cdn` |
| `path` | Source path as written in `index.html`, or a full URL when `external` |
| `external` | Was the source a remote URL. **Does not tell you whether it was inlined** — see §4 |
| `hash` | Truncated content hash. Four different conventions exist — see §5 |

Entries appear in **emit order**, which is the order the corresponding blocks
appear in the shipped document. That ordering is the only thing tying a block to
its filename, and it is what makes decompiling possible at all (`notes/04`).

## 3. Two shipped shapes

**Flat** (what you want). Manifest comment, then a plain document. Every packager
in this repo produces this. It is what Firepit renders predictably.

**Wrapper** (avoid for Firepit). A parent shell containing
`const CHILD_HTML_B64 = "…"`, which it decodes into an `<iframe srcdoc>` at
runtime. Forge ships *itself* this way — [forge/Forge.html](../forge/Forge.html)
is 8.2 MB of which 8.14 MB is one base64 string. It loads, but you are then
running an iframe inside Firepit's iframe, and the host sanitizes nested
`srcdoc` attributes (`notes/02` §5). Ship flat.

> An earlier handoff in this repo described these as Forge's two compile targets —
> a SharePoint/Flank Speed ship producing the flat file, and an offline wrapper
> ship producing the nested one. The two output *shapes* are verified from real
> artifacts. The tool's internal option names quoted in that handoff are not, and
> are marked inferred there.

## 4. `external: true` does not mean "left as a URL"

The packagers disagree here, and assuming either behaviour will burn you:

| Packager | On encountering `<script src="https://…">` |
|---|---|
| `build-sharepoint-package.js` (PULSE) | **Throws.** The build fails until you vendor the file locally. |
| Forge's `compiler.js` | **Fetches it at compile time and inlines it**, recording `external: true`. |

Both produce a runtime with zero network dependencies — which is mandatory under
`default-src 'none'` — they just differ on who does the downloading. Forge's
manifest lists seven external entries (jQuery, Bootstrap CSS+JS, FileSaver,
SheetJS, mammoth, PDF.js); all seven are present, inlined, in the shipped file.
See [forge-ide-decompiled/_vendor-cdn/README.md](../forge-ide-decompiled/_vendor-cdn/README.md).

One exception worth knowing: PULSE CODE keeps the Monaco loader as a genuine
runtime CDN reference and records it as `kind: "js-cdn"`. That works when the file
is opened directly from a document library, and would **not** work inside a
Firepit web part, where the CSP would block it.

## 5. The hash field: four conventions

Reproducing the hash is what lets you *verify* a decompile instead of guessing.
There is no single algorithm, and there is not even a single truncation length:

| Convention | Who | Computed over |
|---|---|---|
| sha256, first 8 hex | `build-sharepoint-package.js` | the bare source text |
| sha256, first 8 hex | PULSE CODE's packager | the whole emitted `<script>…</script>` block |
| sha256, first **12** hex | `pulse-documentation`'s packager | the pre-inline bundle text |
| **DJB2-xor, 32-bit** | **Forge** | the **escaped** block body, exactly as written out |

Compare by *prefix* against a full digest rather than truncating to a guessed
width — assume 8 and every entry in the 12-char package silently fails to verify.

Forge's, from `hashString()` in its own compiler
([forge-ide-decompiled/compiler.js](../forge-ide-decompiled/compiler.js), ~line 2531):

```js
const hashString = (str) => {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h) ^ str.charCodeAt(i);
  return (h >>> 0).toString(16).padStart(8, '0');
};
```

The subtlety that costs an afternoon: Forge hashes
`sanitizeScriptContent(code)` — the text *after* `</script` has been rewritten to
`<\/script>`. If you unescape a block before hashing it, all 34 entries of a real
Forge build fail to verify and you conclude the hashes are unusable. They are not.
`tools/wfc-decompile.js` tries all four conventions and reports which one landed.

## 6. Building a package yourself

Real, working, zero-dependency examples in this repo:

| Script | Output |
|---|---|
| `apps/PULSE/scripts/build-sharepoint-package.js` | `releases/PULSE-v1.0.0.html` |
| `apps/PULSE-CODE/scripts/build-sharepoint-package.js` | `releases/PULSE-CODE-v1.0.0.html` |
| `apps/PULSE/scripts/build-travel-packages.js` | all eight packages at once |

The algorithm:

1. Read `index.html`.
2. For each `<link rel="stylesheet" href="…">`, read the CSS, rewrite its
   `url(...)` references to `data:` URIs, emit a `<style>` block, and record a
   manifest entry.
3. For each `<script src="…">`, read the JS, **escape `</script`**, emit a
   `<script>` block, and record a manifest entry.
4. Rewrite local asset paths that appear inside JS/CSS string literals to `data:`
   URIs too — not just the ones in `<img src>`.
5. Prepend `<!--WFC-MANIFEST:…-->`.
6. Write one file.

### The four rules that are not optional

**Escape `</script` in every inlined script.** A literal `</script` inside a
string constant terminates the block early and corrupts everything after it. Both
packagers do `s/<\/script/<\\\/script/gi`. Forge additionally escapes `<script` →
`<\x73cript`.

**Throw on external references** rather than skipping them silently — unless you
are going to fetch and inline them like Forge does. A silent skip becomes a
CSP violation in production, which is a blank page with no obvious cause.

**Never add `?v=N` cache-busting to script or link tags.** SharePoint's static
file serving mishandles query strings on script URLs: every
`<script src="app.js?v=10">` 404s *silently*. The symptom is a blank white page
with an empty console, because the request failed before any of your code ran.
This cost real debugging time here (`notes/06`, item 2). Use a hard refresh
during development instead.

**Unwrap UMD bundles that check for AMD.** SharePoint pages define a global
`define`, so a UMD bundle registers itself with AMD instead of setting
`window.Foo`, and the next library that expects `window.JSZip` throws. The PULSE
packager wraps such bundles in an IIFE that shadows `define`, `module`, and
`exports`:

```js
(function (root) {
  var define = undefined, module = undefined, exports = undefined;
  /* …bundle… */
  if (typeof JSZip !== "undefined") root.JSZip = JSZip;
})(typeof window !== "undefined" ? window : this);
```

This rewrite is also why `jspdf.umd.min.js` and `pptxgen.browser.js` are the only
two files in `PULSE-v1.0.0.html` whose manifest hashes do not verify — the shipped
bytes are deliberately not the source bytes. Expected, and reported rather than
hidden.

## 7. Deploying

1. Build the single file.
2. Upload it to a SharePoint document library (`SiteAssets` is conventional).
3. Point a Firepit web part's `htmlFileUrl` at it, **or** paste the content into
   `htmlCode` if it is small.
4. Redeploy = overwrite the file. No package upload, no app catalog, no cache to
   purge.

Verify in a real Firepit web part before calling it done. A browser tab on your
desktop does not apply the host's CSP, so it cannot reproduce the failure mode
that matters most.

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
