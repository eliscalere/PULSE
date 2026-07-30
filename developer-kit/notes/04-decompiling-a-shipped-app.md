<!-- UNCLASSIFIED -->

# 04 — Decompiling a shipped app

**UNCLASSIFIED**

You have a single 8 MB `.html` file and you want the source tree back. This
document covers how that works, the four ways a naive attempt fails silently, and
how to tell a verified filename from a guess.

Tool: [tools/wfc-decompile.js](../tools/wfc-decompile.js) (zero dependencies).
Worked output: [forge-ide-decompiled/](../forge-ide-decompiled/).

```bash
node developer-kit/tools/wfc-decompile.js developer-kit/forge/Forge.html /tmp/forge
```

Results on the artifacts in this repo:

| Package | Extracted | Hash-verified | Notes |
|---|---|---|---|
| `forge/Forge.html` | 43 files | **34 / 34 manifest entries** | 9 untracked inline blocks |
| `releases/PULSE-v1.0.0.html` | 47 files | 43 / 45 | 2 UMD bundles rewritten at build time |
| `releases/PULSE-CODE-v1.0.0.html` | 2 files | 2 / 2 | both were inline in the source |
| `releases/PULSE-Calendar-v1.0.0.html` | 5 files | 4 / 4 | 1 untracked inline block |
| `releases/PULSE-Tickets-v1.0.0.html` | 28 files | 24 / 26 | same two UMD bundles |
| `releases/PULSE-Documentation-v1.0.0.html` | 11 files | 1 / 2 | 12-char hashes; its JS bundle is rewritten at build time |

Extracted application files are byte-identical to the sources in `apps/`. Spot-check:

```bash
diff /tmp/pulse/assets/js/sharepoint-adapter.js apps/PULSE/assets/js/sharepoint-adapter.js
```

---

## Why this is not a ten-line regex

Four failure modes, each of which produces a *plausible-looking wrong answer*
rather than an error. Every one was hit while building this tool.

### 1. `/<script[^>]*>([\s\S]*?)<\/script>/g` shreds the document

These packages inline libraries that generate HTML as data — SheetJS writing ODS
XML, mammoth writing docx markup — and, in Forge's case, an IDE whose own compiler
manipulates `"<style>"` and `"</script>"` as strings. A naive sweep matches those
string literals as tags.

On `Forge.html` the naive approach reported **39 script blocks and 25 style
blocks**, with "stylesheets" whose content was `ss:ID="Default"` (SheetJS's ODS
generator) and jQuery's `/<script|<style|<link/i` regex parsed as markup. The
correct count is **39 script and 4 style blocks**.

Fix: walk the document the way an HTML parser treats raw-text elements. A
`<script>` body runs to its first matching close tag; scanning resumes strictly
after it; everything inside a body is opaque.

### 2. Indexing into a lowercased copy silently skews every offset

The obvious way to match tags case-insensitively is `html.toLowerCase().indexOf("<script")`,
then slice from the original. **`toLowerCase()` is not length-preserving in
Unicode.** `Forge.html`'s child document is 5,853,651 chars; its lowercased copy
is 5,853,654 — a **3-character drift**.

The consequence is not a crash. Offsets land 3 bytes off, the character-after-tag
check fails, and whole blocks get skipped — mammoth and PDF.js vanished entirely
from one run, and the files after them were all named one position off. Every
extracted file looked syntactically fine.

Fix: match with a case-insensitive regex over the original text. Never build a
lowercased index.

### 3. Untracked blocks are interleaved, so positional pairing needs alignment

The manifest lists files in emit order, but the shipped document also contains the
app's *own* inline blocks, which get no manifest entry. In `Forge.html` there are
nine, and they are scattered through the tracked ones: a bridge bootstrap and two
guards in `<head>`, a PDF.js worker config between PDF.js and CodeMirror, a
CodeMirror wiring shim after the bundle, and event-binding blocks near the end.

So neither "the tracked run is at the front" nor "at the back" is true. Counting
one untracked block against the manifest shifts every following name by one —
`compiler.js` gets `decompiler.js`'s contents and nothing complains.

Fix: treat the tracked blocks as an in-order *subsequence*. The tool runs a small
dynamic program — at each block, either it is the next manifest entry or it is
untracked — maximizing total evidence.

### 4. The hash you know about is not the hash that was used

See `notes/03` §5. Four conventions, and not even a single truncation width —
Forge's is a DJB2-xor over the *escaped* block body, and `pulse-documentation`'s
manifest truncates sha256 to 12 hex chars where everything else uses 8. Unescape
before hashing and 34/34 Forge entries fail; truncate to a guessed 8 and every
entry in the 12-char package fails. Both look like "these hashes are decorative".
They are not — they are the difference between a verified extraction and a
plausible one. Compare by prefix against the full digest.

---

## How a filename is decided

Each candidate pairing is scored by evidence, strongest first. Every output file
is labelled with the tier that named it, so a guess is never presented as a fact.

| Tier | Signal | Score |
|---|---|---|
| pinned | an operator-supplied `--map` entry | 1000 |
| **exact** | manifest hash matches (any of four conventions) | 100 |
| strong | `//# sourceMappingURL=<name>.map` | 20 |
| strong | a `// File: <name>.js` header near the top | 16 |
| good | a top-level `const`/`let`/`class` named like the file | 12 |
| good | a vendor license banner naming the library | 10 |
| good | the file titles itself in its opening comment | 6 |
| weak | the filename appears somewhere in the code | 2 |
| inferred | position in the sequence, nothing else | 0 |

Two refinements that mattered, both discovered by getting a real file wrong:

**Comments are stripped before "mention" is counted.** A 249-byte shim in
`Forge.html` opens with `// Runtime feature flags consumed by Prometheus (set
before athenaAgent.js loads)`. On a raw substring search that shim looks more like
`athenaAgent.js` than the actual 542 KB agent module does. Discounting
comment-only mentions puts the name back on the real file.

**A self-title is distinguished from a cross-reference by the extension.** A file
titling itself writes prose — `/* ===== Forge v2 — Forge Styles ===== */` — while
a reference cites the full filename — `are in css/styles.css`. Reading those the
same way swapped Forge's `styles.css` with an inline base-stylesheet. The rule the
tool uses: a stem in the leading comment counts as a self-title only when it is
*not* followed by its own extension.

Both were later confirmed correct independently, when the DJB2 hash convention
landed and all 34 entries verified exactly. That is the useful property of a
scored approach: the evidence tiers and the hashes are independent checks on each
other.

## Reviewing and correcting a result

Every run writes `BLOCK-INVENTORY.md` — one row per block with its hash, size,
naming tier, and opening line:

| # | hash | bytes | named by | path | opens with |
|---|---|---|---|---|---|
| 15 | `ccbf026d` | 1145947 | exact (manifest hash) | `cm6-bundle.js` | `var cm6 = (() => {` |

Scan the `positional` and `untracked` rows. If a name is wrong, pin it by hash:

```json
{ "565227c7": "athenaAgent.js" }
```

```bash
node tools/wfc-decompile.js forge/Forge.html out --map overrides.json
```

Hashes are content-addressed, so a map survives reruns and reviews cleanly in a
diff. A stale hash is reported rather than ignored.

## What cannot be recovered

- **Binary assets.** Images and fonts survive as `data:` URIs inside the CSS and
  HTML, but the manifest never recorded their original paths, so you get the bytes
  without the filenames. (Forge records `kind: "img"` entries with hashes, which
  gives you a way to match them up by hand if you need to.)
- **Original formatting of minified vendor bundles.** Extraction is byte-exact,
  which means a minified file comes back minified.
- **Files the packager rewrote.** The two PULSE UMD bundles come back in their
  AMD-shimmed form. The tool flags them rather than pretending otherwise.
- **Whatever was never inlined.** A `kind: "js-cdn"` entry stays a URL.

## Forge has its own decompiler

[forge-ide-decompiled/decompiler.js](../forge-ide-decompiled/decompiler.js) — the
IDE ships one, and reading it is what revealed the DJB2 hash function
(`_hashString`), the newline normalization (`_normalizeInline`), and the fact that
`_collectInlineBlocks` uses exactly the naive regex described in failure mode 1
above, obfuscating its own tags as `<\x73cript` to avoid self-termination.

It also has `_stripSecurityFeatures`, which removes the injected guards and CSP
meta on the way back out — worth knowing exists if you round-trip a file through
Forge and wonder where the security wrapper went.

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
