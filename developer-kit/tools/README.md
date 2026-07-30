<!-- UNCLASSIFIED -->

# Tools

**UNCLASSIFIED**

## `wfc-decompile.js`

Unpacks a shipped single-file app back into a source tree. Zero dependencies,
Node 14+.

```bash
node wfc-decompile.js <package.html> [output-dir] [--map <overrides.json>] [--quiet]
```

Handles both shipped shapes: the flat WFC package (manifest comment + plain
document) and the Forge wrapper (a shell holding `CHILD_HTML_B64`, unwrapped
first). Writes the extracted files, a reconstructed `index.html` with normal
`<link>`/`<script src>` tags, `WFC-MANIFEST.json`, and `BLOCK-INVENTORY.md`.

### Results on the artifacts in this repo

| Package | Files | Verified |
|---|---|---|
| `forge/Forge.html` | 43 | **34 / 34** manifest entries, exact |
| `releases/PULSE-v1.0.0.html` | 47 | 43 / 45 (two UMD bundles rewritten at build time) |
| `releases/PULSE-CODE-v1.0.0.html` | 2 | 2 / 2 |
| `releases/PULSE-Calendar-v1.0.0.html` | 5 | 4 / 4 |
| `releases/PULSE-Tickets-v1.0.0.html` | 28 | 24 / 26 |
| `releases/PULSE-Documentation-v1.0.0.html` | 11 | 1 / 2 (12-char hashes; its bundle is rewritten at build time) |

Extracted application files are byte-identical to the sources in `apps/`:

```bash
node tools/wfc-decompile.js releases/PULSE-v1.0.0.html /tmp/pulse --quiet
diff /tmp/pulse/assets/js/sharepoint-adapter.js apps/PULSE/assets/js/sharepoint-adapter.js
```

### Every filename comes with its evidence

The manifest lists source paths in emit order but never says which block is which,
so pairing is inferred — and the tool reports *how* each name was established
rather than presenting a guess as a fact:

| Tier | Meaning |
|---|---|
| `exact (manifest hash)` | Content hash matched. This is a fact. |
| `strong (sourceMappingURL)` / `strong (// File: header)` | The file names itself |
| `good (matching declaration)` / `good (library banner)` / `good (self-titling banner)` | Strong circumstantial evidence |
| `weak (filename mentioned)` | The name appears in the code |
| `inferred (position only)` | Order in the sequence, nothing else. Eyeball these. |
| `untracked (not in manifest)` | An inline block the manifest never recorded |
| `pinned (operator map)` | You told it, via `--map` |

`BLOCK-INVENTORY.md` lists every block with its hash, tier, and opening line. To
correct a name, pin it by content hash:

```json
{ "565227c7": "athenaAgent.js" }
```

```bash
node tools/wfc-decompile.js forge/Forge.html out --map overrides.json
```

Hashes are content-addressed, so a map survives reruns and reviews cleanly in a
diff. Stale entries are reported rather than ignored.

### Why it is 600 lines and not 20

Four failure modes, each of which yields a *plausible wrong answer* rather than an
error. All four were hit while writing it, and all four are documented in the
file's own header and in
[../notes/04-decompiling-a-shipped-app.md](../notes/04-decompiling-a-shipped-app.md):

1. **A naive `<script>` regex shreds the document.** These packages inline
   libraries that emit HTML as data — and, in Forge's case, an IDE whose compiler
   manipulates `"</script>"` as a string. A naive sweep read 39 script and 25 style
   blocks out of `Forge.html`; the truth is 39 and 4.
2. **`toLowerCase()` is not length-preserving.** Indexing a lowercased copy of
   `Forge.html`'s child drifts by 3 characters, which silently skips whole blocks
   and misnames everything after them.
3. **Untracked blocks are interleaved,** so the tracked ones are an in-order
   subsequence, not a contiguous run. Needs sequence alignment, not an offset.
4. **Four different hash conventions exist,** and the truncation width varies too
   (8 hex chars in most packagers here, 12 in `pulse-documentation`'s). Forge's is
   a DJB2-xor over the *escaped* block body. Unescape first, or truncate to a
   guessed width, and whole packages silently fail to verify.

### Limits

- Binary assets come back as `data:` URIs inside the CSS/HTML, not as named files —
  the manifest never recorded their paths.
- Byte-exact extraction means minified stays minified.
- Files the packager rewrote (UMD→global wraps) come back rewritten, and are
  flagged.
- A `kind: "js-cdn"` entry was never inlined, so there is nothing to extract.

## `sp-columns.js` and friends

Not here — the runnable integration code lives in
[../connectors/](../connectors/).

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
