<!-- UNCLASSIFIED -->

# Forge IDE — decompiled

**UNCLASSIFIED**

[`../forge/Forge.html`](../forge/Forge.html) is the Forge IDE as shipped: one
8.2 MB HTML file. This folder is that file unpacked back into its source tree.

**All 34 files the manifest describes are hash-verified**, using Forge's own hash
function — so the filenames here are facts read out of the package, not guesses.
Nine additional inline blocks had no manifest entry; those are in `_app-inline/`
and are labelled untracked.

Produced by [`../tools/wfc-decompile.js`](../tools/wfc-decompile.js). Nothing here
was hand-edited except this README.

## Reproduce it

```bash
node developer-kit/tools/prepare-forge-tree.js
```

That regenerates this folder exactly as committed. It runs the decompiler and then
drops three kinds of byte that are all derivable from `forge/Forge.html`, which is
itself in the repo:

- `_vendor-cdn/` → a README listing the seven public library releases instead of
  2.2 MB of library code
- `_child-app-as-shipped.html` (5.8 MB) → deleted; it is `Forge.html` minus one
  base64 decode
- the 8.14 MB base64 payload inside `_forge-parent-shell.html` → elided, keeping
  its 38 KB of wrapper logic

For the untrimmed output:

```bash
node developer-kit/tools/wfc-decompile.js developer-kit/forge/Forge.html /tmp/forge-full
```

## What is in here

### Forge's own modules

| File | Size | What it is |
|---|---|---|
| `athenaAgent.js` | 532 KB | "Prometheus" — the agentic coding assistant. `const aiAgent`, profiles, API keys, abort control |
| `compiler.js` | 328 KB | **The packager.** Inlines CSS/JS, fetches and inlines CDN deps, writes the WFC manifest, emits the Firepit renderer |
| `decompiler.js` | 24 KB | Forge's own unpacker — see below |
| `editor.js` | 84 KB | CodeMirror integration, saves, file-tree resilience |
| `cm6-bundle.js` | 1.1 MB | CodeMirror 6, bundled offline by Forge |
| `loadFolder.js` | 132 KB | File System Access API project loading |
| `aiResponseImporter.js` | 156 KB | Parses AI output back into file writes |
| `aiResponseParser.js` | 24 KB | Language/extension mapping for AI code blocks |
| `aiHelper.js` | 100 KB | Prompt construction and AI plumbing |
| `promptLab.js` | 44 KB | Prompt authoring surface |
| `securityReviewer.js` | 88 KB | Security review pass over project files |
| `sastTab.js` | 20 KB | Static analysis tab |
| `checkpointManager.js` | 32 KB | Snapshots into a hidden `.checkpoints` folder |
| `newProjectWalkthrough.js` | 52 KB | New-project wizard |
| `search.js` | 16 KB | Project-wide search |
| `devconsoleTab.js` | 64 KB | In-IDE dev console |
| `testRecorderTab.js` | 32 KB | Records interactions into a replayable test |
| `advancedDebugTab.js` | 64 KB | Debug tooling |
| `mathLogicTesterTab.js` | 8 KB | Expression tester |
| `sharedriveNoSqlTab.js` | 48 KB | UI for the shared-drive store |
| `sharedrive-nosql.js` | 24 KB | `fs-nosql-db.js` — a File System Access API JSON store with last-write-wins merging |
| `leafletMapTab.js` | 12 KB | Map tab |
| `athenaCompat.js` | 16 KB | Compatibility shims for the agent |
| `domPurify.js` | 24 KB | DOMPurify 3.3.0, held as a string for injection |
| `js/merge-editor.js` | 44 KB | Three-way merge UI |
| `css/styles.css` | 84 KB | Forge's stylesheet |
| `css/codicon-inline.css` | 196 KB | VS Code icon font, inlined as base64 |

### `_app-inline/` — the nine untracked blocks

Inline blocks in Forge's own `index.html` that the manifest never recorded. Worth
reading, because three of them are the child-side security layer:

| File | What it is |
|---|---|
| `head-01.js` (38 KB) | The Forge child bridge bootstrap — parent↔child postMessage plumbing |
| `head-02.js` (15 KB) | Query-parameter guard (`Forge query param guard init failed`) |
| `head-03.js` (1.7 KB) | Link-hint guard (`Forge link hint guard init failed`) |
| `head-01.css` (68 KB) | Base stylesheet; its own comment points at `css/styles.css` for the rest |
| `body-*.js` | PDF.js worker config, CodeMirror wiring shim, jQuery event bindings, an Athena panel controller |

These mirror guards the Firepit host *also* injects (see
[../notes/02-firepit-webpart-internals.md](../notes/02-firepit-webpart-internals.md) §5),
so the same protections exist in both layers. Do not be surprised to find them
twice.

### `_forge-parent-shell.html`

The wrapper Forge ships itself inside: it base64-decodes the child app into an
`<iframe srcdoc>` at runtime. Payload elided; the wrapper logic is intact. Note
this shape is **not** what you want for a Firepit ship — the host sanitizes nested
`srcdoc` ([notes/02](../notes/02-firepit-webpart-internals.md) §5,
[notes/03](../notes/03-wfc-package-format.md) §3).

### `BLOCK-INVENTORY.md`

One row per extracted block: hash, size, the evidence tier that named it, and its
opening line. This is the audit trail — every row reads `exact (manifest hash)` or
`untracked`.

---

## Two things worth reading in this tree

### `compiler.js` — how a WFC package is really built

The authority on the format documented in
[notes/03](../notes/03-wfc-package-format.md). Specifics worth finding:

- `hashString()` (~line 2531) — the DJB2-xor 32-bit hash used for every manifest
  entry. **Not** sha256. Reproducing it is what makes verification possible.
- `sanitizeScriptContent()` (~line 2715) — escapes `</script` → `<\/script>` *and*
  `<script` → `<\x73cript`. The manifest hash is computed over this escaped form,
  which is the detail that makes or breaks a verifier.
- `checkForCspIssues()` (~line 249) — Forge warns at compile time about patterns
  the host CSP will reject, and `buildAiRemediationPrompt()` turns those warnings
  into a prompt for its own AI to fix them.
- `buildFusionFirepitRendererHtml()` (~line 700) — the renderer wrapper it emits.
- `bumpShipVersion()` / `buildVersionedOutputName()` — how release naming works.

### `decompiler.js` — Forge unpacks its own output

Reading this is where the hash function came from. Also instructive:

- `_normalizeInline()` — strips exactly one leading and one trailing newline,
  tolerating CRLF. The compiler wraps every block in newlines; this undoes it.
- `_collectInlineBlocks()` — uses the naive tag regex, working around its own
  self-termination problem by writing its tag literals as `<\x73cript`. It is
  vulnerable to the shredding failure described in
  [notes/04](../notes/04-decompiling-a-shipped-app.md) §1; our tool takes the
  parser-accurate route instead.
- `_stripSecurityFeatures()` — removes the injected guards and CSP meta on the way
  back out. Useful to know if you round-trip a file through Forge and wonder where
  the security wrapper went.
- `_extractWasmAssets()`, `_collectDataUriAssets()` — how binary assets are
  recovered from `data:` URIs.

---

## Caveats

- **This is a snapshot**, compiled 2026-06-29. Forge is someone else's product and
  will have moved on.
- **Extraction is byte-exact, so minified stays minified.** Vendor bundles come
  back as they shipped.
- **Binary assets are not recoverable by filename.** They survive as `data:` URIs
  inside the CSS and HTML; the manifest records `kind: "img"` hashes, which gives
  you a way to match them up by hand.
- **Read-only reference.** Nothing here is wired to build or run; it is here to be
  read. Forge's maintainers own Forge.

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
