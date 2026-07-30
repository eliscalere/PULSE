#!/usr/bin/env node
/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   wfc-decompile.js — unpack a shipped single-file app back into a source tree

   Give it any shipped .html package — Forge's own compiled output, a PULSE
   release, anything carrying a WFC-MANIFEST — and it writes a browsable source
   tree: every inlined stylesheet and script back under its original manifest
   path, plus a reconstructed index.html with ordinary <link>/<script src> tags.

   Handles both shipped shapes:
     1. FLAT WFC       — <!--WFC-MANIFEST:base64--> followed by a plain HTML
                         document whose <style>/<script> blocks hold the sources.
                         This is the format that works in a Firepit web part.
     2. FORGE WRAPPER  — a parent shell holding `const CHILD_HTML_B64 = "…"`
                         that it decodes into an <iframe srcdoc> at runtime.
                         Forge ships itself this way. The child is unwrapped
                         first, then treated as (1).

   ── Two things make this harder than it looks ─────────────────────────────

   1. YOU CANNOT TOKENIZE WITH A NAIVE REGEX. These packages inline libraries
      that generate HTML (SheetJS writing ODS, mammoth writing docx markup) and,
      in Forge's case, an IDE whose own compiler manipulates "<style>" and
      "</script>" as string data. A `/<script[^>]*>([\s\S]*?)<\/script>/g` sweep
      matches those string literals as if they were tags and shreds the
      document — you get a plausible-looking tree of wrong files.

      Nor can you index into a lowercased copy of the document: toLowerCase()
      is not length-preserving in Unicode (this very package drifts by 3 chars),
      so every offset silently skews and whole blocks get skipped. Scan with a
      case-insensitive regex over the ORIGINAL text and step strictly past each
      block's close tag. See scanRawTextBlocks().

   2. THE MANIFEST DOES NOT SAY WHICH BLOCK IS WHICH. It lists source paths in
      emit order, and the shipped HTML has no data-source attributes to match
      them to. Pairing must be inferred, and the packagers make simple
      approaches wrong:
        · Untracked blocks are INTERLEAVED. Forge's index.html has its own
          inline scripts (a bridge bootstrap, a PDF.js worker config, a
          CodeMirror wiring shim) sitting between the tracked file blocks — so
          neither front- nor back-alignment works.
        · `external: true` does not mean absent. Forge inlines its CDN
          dependencies and records them as external; other packagers refuse to
          build if a CDN reference exists at all.

      So this tool does an in-order sequence alignment (a small DP, allowing
      blocks to be skipped as untracked) and scores each candidate pairing by
      real evidence found in the block: a manifest hash match, a
      sourceMappingURL naming the file, a `// File: name.js` header, a
      top-level declaration matching the filename. Every extracted file is
      reported with the evidence tier that named it, so a guess never
      masquerades as a fact.

   ── Hash verification ─────────────────────────────────────────────────────

   Every manifest entry carries a truncated content hash, and matching it turns an
   inferred name into a verified one. FOUR conventions are in play across the
   packagers here, and trying only the obvious one leaves whole packages
   unverifiable:

     · sha256 of the bare source text — build-sharepoint-package.js
     · sha256 of the emitted <script>…</script> block — PULSE-CODE's packager,
       for the blocks that were inline in its index.html
     · a DJB2-xor 32-bit hash of the ESCAPED block body — Forge. Forge hashes
       what it writes out, `</script` already rewritten to `<\/script`, so the
       check must run BEFORE unescaping. Undo the escaping first and all 34
       entries of a real Forge build silently fail to verify.
     · and the TRUNCATION LENGTH itself varies — 8 hex chars in most packagers
       here, 12 in pulse-documentation's. So comparison is by prefix against the
       full digest, never by truncating to a guessed width first.

   With all four, `Forge.html` verifies 34/34 and `PULSE-v1.0.0.html` 43/45.

   A surviving mismatch is reported, never silenced: it means the packager
   rewrote that file on the way in (local asset path swapped for a data: URI, a
   UMD bundle wrapped for AMD safety — the two PULSE stragglers are exactly
   that). Expected for vendor bundles; a real finding on your own app code.

   Usage:
     node wfc-decompile.js <package.html> [output-dir] [--map <file>] [--quiet]

   Default output-dir: ./<manifest.project>-decompiled

   Zero dependencies. Node 14+.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const sha256Hex = (text) => crypto.createHash("sha256").update(text).digest("hex");
const hash8 = (text) => sha256Hex(text).slice(0, 8);
const escapeRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

function fail(message) {
  console.error(`\n  ERROR: ${message}\n`);
  process.exit(1);
}

/* ── Format detection ────────────────────────────────────────────────────── */

function readManifest(html) {
  const match = html.match(/<!--\s*WFC-MANIFEST:([A-Za-z0-9+/=]+)\s*-->/);
  if (!match) return null;
  try {
    return JSON.parse(Buffer.from(match[1], "base64").toString("utf8"));
  } catch (err) {
    fail(`found a WFC-MANIFEST comment but could not parse it: ${err.message}`);
  }
}

/* The Forge offline wrapper base64-encodes the whole child app into one const,
   then writes it into an iframe srcdoc. Unwrap before decompiling. Also accepts
   the `-readable` variant, which uses a String.raw template literal.

   Guard: Forge's own compiler.js carries this wrapper as a TEMPLATE, with the
   const left empty. Requiring a long base64 payload keeps us from "unwrapping"
   a package on the strength of its own build tooling. */
function unwrapForge(html) {
  /* Find the payload with indexOf, not a regex. A quantified character class
     spanning a multi-megabyte base64 literal overflows V8's regex stack
     ("Maximum call stack size exceeded") on a real 8 MB Forge build. */
  const MIN_PAYLOAD = 1024;

  const readQuoted = (marker, quote) => {
    const declared = html.indexOf(marker);
    if (declared === -1) return null;
    const open = html.indexOf(quote, declared + marker.length);
    if (open === -1) return null;
    const close = html.indexOf(quote, open + 1);
    if (close === -1 || close - open - 1 < MIN_PAYLOAD) return null;
    return html.slice(open + 1, close);
  };

  const b64 = readQuoted("CHILD_HTML_B64", '"');
  if (b64 && /^[A-Za-z0-9+/=\s]+$/.test(b64.slice(0, 4096))) {
    return { child: Buffer.from(b64, "base64").toString("utf8"), kind: "forge-b64" };
  }

  const raw = readQuoted("CHILD_HTML = String.raw", "`");
  if (raw) {
    const child = raw.replace(/\\`/g, "`").replace(/\\\$\{/g, "${").replace(/\\\\/g, "\\");
    return { child, kind: "forge-readable" };
  }
  return null;
}

/* ── Tokenizing ──────────────────────────────────────────────────────────── */

/* Walk the document as an HTML parser would treat raw-text elements: a <script>
   or <style> body runs to its first matching close tag, and scanning resumes
   strictly after it. Everything inside a body is opaque data, so string
   literals that look like tags can never be mistaken for markup.

   Case-insensitive matching happens in the regex, over the original text — no
   lowercased copy, no offset drift. */
function scanRawTextBlocks(html) {
  const open = /<(script|style)(?=[\s/>])([^>]*)>/gi;
  const blocks = [];
  let resumeAt = 0;
  let m;

  while ((m = open.exec(html))) {
    if (m.index < resumeAt) continue; // inside a body we already consumed
    const name = m[1].toLowerCase();
    const bodyStart = m.index + m[0].length;
    const close = new RegExp(`</${name}(?=[\\s/>])[^>]*>`, "i");
    const cm = html.slice(bodyStart).match(close);
    const bodyEnd = cm ? bodyStart + cm.index : html.length;
    const end = cm ? bodyEnd + cm[0].length : html.length;

    blocks.push({
      kind: name === "style" ? "css" : "js",
      attrs: m[2] || "",
      body: html.slice(bodyStart, bodyEnd),
      raw: html.slice(m.index, end),
      inHead: false, // filled in below
      unterminated: !cm,
    });

    resumeAt = end;
    open.lastIndex = end;
  }

  /* <head> boundary, located the same drift-free way. */
  const headClose = html.match(/<\/head(?=[\s/>])[^>]*>/i);
  const headEnd = headClose ? headClose.index : 0;
  let offset = 0;
  for (const block of blocks) {
    offset = html.indexOf(block.raw, offset);
    block.inHead = offset !== -1 && offset < headEnd;
    if (offset !== -1) offset += block.raw.length;
  }

  return blocks.filter((b) => !(b.kind === "js" && /\bsrc\s*=/i.test(b.attrs)));
}

/* Packagers escape `</script` inside inlined JS so the tag can't terminate the
   block early. Undo it. */
const unescapeScript = (js) => js.replace(/<\\\/script/gi, "</script");

/* A data-source="path" attribute (build-forge.js emits these) is an exact
   answer and beats any inference. */
const dataSource = (attrs) => {
  const m = attrs.match(/data-source=["']([^"']+)["']/);
  return m ? m[1] : null;
};

const blockText = (block) => {
  const body = block.kind === "css" ? block.body : unescapeScript(block.body);
  return body.replace(/^\r?\n/, "").replace(/\r?\n$/, "");
};

/* The block body exactly as shipped — escaping intact. Forge hashes this form. */
const rawBlockBody = (block) => block.body.replace(/^\r?\n/, "").replace(/\r?\n$/, "");

/* Forge's own hash: a DJB2-xor variant, 32-bit, zero-padded hex — not sha256.
   Reproducing it is what turns a Forge package from "names inferred from
   evidence" into "names verified exactly". Lifted from the hashString() inside
   Forge's compiler.js, which the extracted tree here documents. */
function forgeHash(str) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h) ^ str.charCodeAt(i);
  return (h >>> 0).toString(16).padStart(8, "0");
}

/* Three hashing conventions across these packagers:
     sha256 raw     — bare source text                  (build-sharepoint-package.js)
     sha256 wrapped — the emitted <style>/<script> block (PULSE-CODE, inline blocks)
     Forge DJB2     — the ESCAPED block body as shipped, i.e. after `</script` →
                      `<\/script`. Forge hashes what it writes out, so this must
                      be computed BEFORE unescaping — undo the escaping first and
                      nothing matches. That one detail is the whole difference
                      between verifying a Forge package and guessing at it. */
function hashCandidates(text, isCss, rawBody) {
  const escaped = String(text).replace(/<\/script/gi, "<\\/script");
  /* Full digests, not truncated — the truncation LENGTH also varies by packager
     (8 hex chars here, 12 in pulse-documentation's), so comparison is done by
     prefix in hashMatches(). Truncating to 8 first makes a 12-char manifest hash
     unverifiable and every file in that package fall back to positional naming. */
  const out = [
    sha256Hex(text),
    isCss ? sha256Hex(`<style>\n${text}\n</style>`) : sha256Hex(`<script>\n${escaped}\n</script>`),
  ];
  if (rawBody != null) {
    out.push(forgeHash(rawBody), forgeHash(rawBody.replace(/\r\n/g, "\n")));
  }
  return out;
}

/* A manifest hash is a prefix of some digest we can compute. Forge's DJB2 value is
   a complete 8-char hash rather than a prefix, so an exact match counts too. */
function hashMatches(expected, candidates) {
  if (!expected) return false;
  const want = String(expected).toLowerCase();
  return candidates.some((c) => c === want || c.startsWith(want));
}

/* ── Evidence ────────────────────────────────────────────────────────────── */

const EVIDENCE = {
  MAP: { score: 1000, tier: "map override" },
  FORBIDDEN: { score: -1000, tier: "forbidden" },
  HASH: { score: 100, tier: "hash" },
  SOURCEMAP: { score: 20, tier: "sourcemap" },
  HEADER: { score: 16, tier: "file-header" },
  DECLARATION: { score: 12, tier: "declaration" },
  BANNER: { score: 10, tier: "library banner" },
  TITLE: { score: 6, tier: "title comment" },
  MENTION: { score: 2, tier: "mention" },
  NONE: { score: 0, tier: "positional" },
};

/* jquery-3.7.1.min.js → jquery · xlsx.full.min.js → xlsx
   mammoth.browser.min.js → mammoth · bootstrap.bundle.min.js → bootstrap
   A versioned, minified vendor filename shares no literal substring with its
   own source, so the raw stem finds nothing. Normalize before looking. */
function libraryName(base) {
  return base
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/[-.]?\d+(\.\d+)*$/g, "")
    .replace(/\.(min|full|browser|bundle|esm|umd|prod|dev|slim)$/gi, "")
    .replace(/\.(min|full|browser|bundle|esm|umd|prod|dev|slim)$/gi, "")
    .replace(/[-.]\d+(\.\d+)*/g, "");
}

/* Strip comments so a note *referring* to another file ("set before
   athenaAgent.js loads") is not read as evidence of *being* that file. */
function stripComments(text) {
  return text.replace(/\/\*[\s\S]*?\*\//g, " ").replace(/^[ \t]*\/\/.*$/gm, " ");
}

/* The leading comment of a file, if it opens with one. */
function leadingComment(text) {
  const trimmed = text.replace(/^[\s﻿]+/, "");
  const blockEnd = trimmed.startsWith("/*") ? trimmed.indexOf("*/") : -1;
  if (blockEnd !== -1) return trimmed.slice(0, blockEnd + 2);
  const lines = [];
  for (const line of trimmed.split(/\r?\n/)) {
    if (!/^\s*\/\//.test(line)) break;
    lines.push(line);
  }
  return lines.join("\n");
}

/* Only a versioned/minified/CDN entry can be identified by a license banner.
   Applying this to a hand-written module lets a generic token like "styles" or
   "pdf" claim the wrong block. */
const looksVendored = (entry, base) => !!entry.external || /\.min\.|\d+\.\d+/.test(base);

/* A real vendor banner announces itself with a copyright or version. */
const isLicenseBanner = (comment) => /copyright|\(c\)|©|@license|@licstart|\bv\d+\.\d+/i.test(comment);

/* How strongly does this block look like it came from this manifest path? */
function evidenceFor(block, entry, pinnedPath) {
  if (!entry) return EVIDENCE.NONE;

  /* An operator-supplied map is the final word, in both directions: it pins one
     pairing and forbids every other for that block. */
  if (pinnedPath) return pinnedPath === entry.path ? EVIDENCE.MAP : EVIDENCE.FORBIDDEN;

  const text = blockText(block);
  const isCss = block.kind === "css";
  if (hashMatches(entry.hash, hashCandidates(text, isCss, rawBlockBody(block)))) return EVIDENCE.HASH;

  const base = entry.path.split("/").pop().split("?")[0]; // editor.js
  const stem = base.replace(/\.[a-z0-9]+$/i, ""); // editor
  if (!stem || stem.length < 3) return EVIDENCE.NONE;

  /* Minified bundles keep a trailing `//# sourceMappingURL=<file>.map`. */
  if (new RegExp(`sourceMappingURL=[^\\s'"]*${escapeRe(base)}`, "i").test(text)) return EVIDENCE.SOURCEMAP;

  /* Hand-written modules here often open with `// File: editor.js`. Head of the
     block only, so a cross-reference further down doesn't count. */
  const head = text.slice(0, 400);
  if (new RegExp(`(?:File|Module|Source)\\s*:\\s*${escapeRe(base)}`, "i").test(head)) return EVIDENCE.HEADER;

  /* `const decompiler = {` / `var cm6 = (() => {` — a top-level binding whose
     name matches the filename stem, compared case-insensitively because the
     repo mixes sharedriveNoSqlTab.js with `const sharedriveNosqlTab`. */
  const code = stripComments(text);
  if (!isCss && new RegExp(`\\b(?:const|let|var|class|function)\\s+${escapeRe(stem)}\\b`, "i").test(code)) {
    return EVIDENCE.DECLARATION;
  }

  /* `/*! jQuery v3.7.1` · `/*! xlsx.js (C) …` — a vendor library announcing
     itself in its own license banner. Restricted hard: only for entries that
     look vendored, only inside the file's leading comment, and only when that
     comment reads like a banner. Loosen any of the three and a generic token
     ("pdf", "styles") starts claiming the wrong block. */
  const lib = libraryName(base);
  const banner = leadingComment(text);
  if (looksVendored(entry, base) && lib.length >= 4 && isLicenseBanner(banner)) {
    if (new RegExp(`\\b${escapeRe(lib)}\\b`, "i").test(banner)) return EVIDENCE.BANNER;
  }

  /* A file titling itself in its opening comment — `/* ===== Forge Styles ===== *​/`
     — versus one pointing at a different file — `// set before athenaAgent.js
     loads`. The discriminator that holds up across this codebase: prose titles
     drop the extension, cross-references cite the full filename. Both patterns
     appear in the same package, and reading them the same way swaps two files. */
  const ext = path.extname(base).slice(1);
  if (stem.length >= 4) {
    const selfTitle = new RegExp(`\\b${escapeRe(stem)}\\b(?!\\s*\\.\\s*${escapeRe(ext)})`, "i");
    if (selfTitle.test(banner)) return EVIDENCE.TITLE;
  }

  if (code.toLowerCase().includes(stem.toLowerCase())) return EVIDENCE.MENTION;
  return EVIDENCE.NONE;
}

/* ── Alignment ───────────────────────────────────────────────────────────── */

/* The manifest lists files in emit order and the packager writes them in that
   order, so the tracked blocks are an in-order SUBSEQUENCE of all blocks —
   untracked app inline blocks may sit anywhere between them.

   Find the assignment maximizing total evidence with a monotonic DP:
     dp[i][j] = best score using the first i blocks and first j entries
              = max( dp[i-1][j]              (block i is untracked)
                   , dp[i-1][j-1] + ev(i,j)  (block i is entry j) )
   Every entry must be consumed, so the answer is dp[nBlocks][nEntries]. */
function alignBySubsequence(blocks, entries, pinnedFor) {
  const n = blocks.length;
  const k = entries.length;
  if (!k) return { pairs: blocks.map((block) => ({ block, entry: null, evidence: EVIDENCE.NONE })), score: 0 };

  const NEG = -Infinity;
  const dp = Array.from({ length: n + 1 }, () => new Float64Array(k + 1).fill(NEG));
  const from = Array.from({ length: n + 1 }, () => new Uint8Array(k + 1)); // 1 = skip block, 2 = match
  const ev = Array.from({ length: n }, (_, i) =>
    entries.map((e) => evidenceFor(blocks[i], e, pinnedFor && pinnedFor(blocks[i])))
  );

  dp[0][0] = 0;
  for (let i = 1; i <= n; i++) {
    for (let j = 0; j <= Math.min(k, i); j++) {
      if (dp[i - 1][j] > NEG) {
        dp[i][j] = dp[i - 1][j];
        from[i][j] = 1;
      }
      if (j > 0 && dp[i - 1][j - 1] > NEG) {
        /* +0.001 per pairing keeps the DP from wandering among equal-scoring
           layouts; it never outweighs real evidence. */
        const candidate = dp[i - 1][j - 1] + ev[i - 1][j - 1].score + 0.001;
        if (candidate > dp[i][j]) {
          dp[i][j] = candidate;
          from[i][j] = 2;
        }
      }
    }
  }

  if (dp[n][k] === NEG) {
    /* More manifest entries than blocks — the package is not what the manifest
       describes. Fall back to back-alignment and let the report show it. */
    const firstTracked = Math.max(0, n - k);
    return {
      pairs: blocks.map((block, i) => {
        const entry = i >= firstTracked ? entries[i - firstTracked] || null : null;
        return {
          block,
          entry,
          evidence: entry ? evidenceFor(block, entry, pinnedFor && pinnedFor(block)) : EVIDENCE.NONE,
        };
      }),
      score: 0,
      degraded: true,
    };
  }

  const pairs = new Array(n);
  let i = n;
  let j = k;
  while (i > 0) {
    if (from[i][j] === 2) {
      pairs[i - 1] = { block: blocks[i - 1], entry: entries[j - 1], evidence: ev[i - 1][j - 1] };
      i--;
      j--;
    } else {
      pairs[i - 1] = { block: blocks[i - 1], entry: null, evidence: EVIDENCE.NONE };
      i--;
    }
  }
  return { pairs, score: dp[n][k] };
}

/* Whether external entries were inlined varies by packager, so try both
   candidate entry lists and keep whichever the evidence endorses. */
function bestAlignment(blocks, manifestFiles, kind, pinnedFor) {
  const ofKind = manifestFiles.filter((f) => f.kind === kind);
  const candidates = [
    { label: "all manifest entries (incl. inlined CDN)", entries: ofKind },
    { label: "local manifest entries only", entries: ofKind.filter((f) => !f.external) },
  ];

  let best = null;
  for (const candidate of candidates) {
    const result = alignBySubsequence(blocks, candidate.entries, pinnedFor);
    if (!best || result.score > best.score) best = { ...candidate, ...result };
  }
  return best;
}

/* ── Main ────────────────────────────────────────────────────────────────── */

function decompile(pkgPath, outDir, overrides) {
  let html = fs.readFileSync(pkgPath, "utf8");
  const notes = [];
  const map = overrides || {};

  /* Forge ships itself through its own offline-wrapper path, and the WFC
     manifest sits on the OUTER shell while the sources it describes live inside
     the base64 child. Capture it before unwrapping — dropping it costs you
     every real filename. */
  const outerManifest = readManifest(html);

  const wrapper = unwrapForge(html);
  if (wrapper) {
    notes.push(`Input is a Forge iframe-srcdoc wrapper (${wrapper.kind}); decompiled the embedded child app.`);
    fs.mkdirSync(outDir, { recursive: true });
    fs.writeFileSync(path.join(outDir, "_forge-parent-shell.html"), html, "utf8");
    notes.push("Parent shell kept as _forge-parent-shell.html — it is where the sandboxing lives.");
    html = wrapper.child;
    fs.writeFileSync(path.join(outDir, "_child-app-as-shipped.html"), html, "utf8");
  }

  const innerManifest = readManifest(html);
  const manifest = innerManifest || outerManifest;
  if (wrapper && !innerManifest && outerManifest) {
    notes.push("Child carries no manifest of its own; used the outer shell's manifest to name its sources.");
  }
  if (!manifest) notes.push("No WFC-MANIFEST found — files fall back to positional names.");
  const manifestFiles = manifest ? manifest.files : [];

  const blocks = scanRawTextBlocks(html);
  const unterminated = blocks.filter((b) => b.unterminated).length;
  if (unterminated) {
    notes.push(`${unterminated} block(s) had no close tag and were read to end-of-file — the package may be truncated.`);
  }

  /* Blocks are identified in the override map by the sha256-8 of their exact
     content, so a map stays valid across reruns and can be reviewed in a diff. */
  const hashOf = new Map(blocks.map((b) => [b, hash8(blockText(b))]));
  const pinnedFor = (block) => map[hashOf.get(block)] || null;
  const usedOverrides = new Set();

  const cssAlign = bestAlignment(blocks.filter((b) => b.kind === "css"), manifestFiles, "css", pinnedFor);
  const jsAlign = bestAlignment(blocks.filter((b) => b.kind === "js"), manifestFiles, "js", pinnedFor);
  const pairOf = new Map();
  for (const pair of [...cssAlign.pairs, ...jsAlign.pairs]) pairOf.set(pair.block, pair);

  const written = [];
  const replacements = [];
  const untrackedCount = { css: 0, js: 0 };
  const inlineCount = { css: 0, js: 0 };
  const usedPaths = new Set();

  for (const block of blocks) {
    const isCss = block.kind === "css";
    const pair = pairOf.get(block) || { entry: null, evidence: EVIDENCE.NONE };
    const entry = pair.entry;
    const explicit = dataSource(block.attrs);
    const pinned = pinnedFor(block);
    let relPath = explicit || (entry && entry.path) || pinned;
    let tier = explicit ? "data-source" : entry ? pair.evidence.tier : pinned ? "map override" : pair.evidence.tier;
    if (pinned) usedOverrides.add(hashOf.get(block));

    if (!relPath) {
      const where = block.inHead ? "head" : "body";
      const suffix = isCss ? "css" : "js";
      relPath = `_app-inline/${where}-${String(++untrackedCount[block.kind]).padStart(2, "0")}.${suffix}`;
      tier = "untracked";
    } else if (/^https?:\/\//i.test(relPath)) {
      /* An inlined CDN dependency. Keep the remote origin visible in the tree
         rather than pretending it was ever a local file. */
      const url = new URL(relPath);
      relPath = path.posix.join("_vendor-cdn", url.hostname, url.pathname.replace(/^\//, ""));
    } else if (relPath.includes("#")) {
      /* index.html#inline / index.html#style — was inline in the source
         index.html, so it has no file of its own, and '#' is not a filename. */
      const suffix = isCss ? "css" : "js";
      const named = `assets/${suffix}/inline-${String(++inlineCount[block.kind]).padStart(2, "0")}.${suffix}`;
      notes.push(`${relPath} was inline in the original index.html → extracted to ${named}`);
      relPath = named;
    }

    /* Manifest paths can carry a `?v=` cache-buster (see notes/03-postmortems.md
       — those query strings are a trap on SharePoint, and at least one packager
       records them verbatim). Strip it, and keep everything inside outDir so a
       `../` path in the manifest cannot write outside the output tree. */
    relPath = relPath.split("?")[0].split("#")[0].replace(/^(\.\.\/)+/, "_external/");
    while (usedPaths.has(relPath)) relPath = relPath.replace(/(\.[^./]+)$/, "-dup$1");
    usedPaths.add(relPath);

    const text = blockText(block);
    const target = path.join(outDir, relPath);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, text.endsWith("\n") ? text : text + "\n", "utf8");

    const expected = entry && entry.hash;
    /* First line of real content — the thing a human needs in order to check a
       name at a glance (a license banner, a `// File:` header, a declaration). */
    const firstLine = (text.split(/\r?\n/).find((l) => l.trim().length > 2) || "").trim().slice(0, 110);
    written.push({
      relPath,
      kind: block.kind,
      bytes: Buffer.byteLength(text),
      hash: hash8(text),
      expected: expected || null,
      match: expected ? hashMatches(expected, hashCandidates(text, isCss, rawBlockBody(block))) : null,
      tier,
      firstLine,
    });

    replacements.push({
      raw: block.raw,
      tag: isCss ? `<link rel="stylesheet" href="${relPath}">` : `<script src="${relPath}"></script>`,
    });
  }

  /* Rebuild an index.html referencing the extracted files instead of carrying
     them inline. Replace each block exactly once, in document order. */
  let indexHtml = html.replace(/<!--\s*WFC-MANIFEST:[A-Za-z0-9+/=]+\s*-->\r?\n?/, "");
  let cursor = 0;
  for (const { raw, tag } of replacements) {
    const at = indexHtml.indexOf(raw, cursor);
    if (at === -1) continue;
    indexHtml = indexHtml.slice(0, at) + tag + indexHtml.slice(at + raw.length);
    cursor = at + tag.length;
  }
  /* Mark the reconstructed document. The header goes after any doctype so the
     document does not start with a comment. */
  if (/<!doctype html>/i.test(indexHtml)) {
    indexHtml = indexHtml.replace(/(<!doctype html>\r?\n?)/i, "$1<!-- UNCLASSIFIED -->\n");
  } else {
    indexHtml = `<!-- UNCLASSIFIED -->\n${indexHtml}`;
  }
  if (!/UNCLASSIFIED\s*-->\s*$/.test(indexHtml)) indexHtml += "\n<!-- UNCLASSIFIED -->\n";
  fs.writeFileSync(path.join(outDir, "index.html"), indexHtml, "utf8");

  const dataUris = (html.match(/data:[a-z0-9.+-]+\/[a-z0-9.+-]+;base64,/gi) || []).length;
  if (dataUris) {
    notes.push(
      `${dataUris} data: URI(s) left in place. The manifest does not record their original paths, so ` +
        `binary assets (images, fonts) come back as bytes inside the CSS/HTML, not as named files.`
    );
  }
  if (manifest) {
    fs.writeFileSync(path.join(outDir, "WFC-MANIFEST.json"), JSON.stringify(manifest, null, 2) + "\n", "utf8");
  }

  const stale = Object.keys(map).filter((h) => !usedOverrides.has(h));
  if (stale.length) {
    notes.push(`${stale.length} override(s) matched no block in this package (stale hash): ${stale.join(", ")}`);
  }

  /* An inventory a human can actually audit: every block, the name it got, the
     evidence tier behind that name, and its opening line. Anything named
     "positional" or "untracked" is worth eyeballing here; if a name is wrong,
     copy the hash into a map file and rerun with --map. */
  const inventory = [
    "<!-- UNCLASSIFIED -->",
    "# Block inventory",
    "",
    `Package: \`${path.basename(pkgPath)}\`${manifest ? ` · project \`${manifest.project}\` · packaged ${manifest.generated}` : ""}`,
    "",
    "Names carrying evidence (`hash`, `sourcemap`, `file-header`, `library banner`,",
    "`declaration`) are trustworthy. `positional` means order was the only clue, and",
    "`untracked` means the manifest never listed the block. To correct a name, put",
    "`{\"<hash>\": \"<path>\"}` in a JSON file and rerun with `--map <file>`.",
    "",
    "| # | hash | bytes | named by | path | opens with |",
    "|---:|---|---:|---|---|---|",
    ...written.map((f, i) => {
      const cell = (s) => String(s).replace(/\|/g, "\\|");
      return `| ${i + 1} | \`${f.hash}\` | ${f.bytes} | ${TIER_LABEL[f.tier] || f.tier} | \`${cell(f.relPath)}\` | \`${cell(f.firstLine)}\` |`;
    }),
    "",
    "<!-- UNCLASSIFIED -->",
    "",
  ].join("\n");
  fs.writeFileSync(path.join(outDir, "BLOCK-INVENTORY.md"), inventory, "utf8");

  return { manifest, written, notes, alignment: { css: cssAlign, js: jsAlign } };
}

/* ── Report ──────────────────────────────────────────────────────────────── */

const TIER_LABEL = {
  "map override": "pinned (operator map)",
  "data-source": "exact (data-source attr)",
  hash: "exact (manifest hash)",
  sourcemap: "strong (sourceMappingURL)",
  "file-header": "strong (// File: header)",
  declaration: "good (matching declaration)",
  mention: "weak (filename mentioned)",
  "title comment": "good (self-titling banner)",
  positional: "inferred (position only)",
  untracked: "untracked (not in manifest)",
};

function report(pkgPath, outDir, result, quiet) {
  const { manifest, written, notes, alignment } = result;
  const pad = (s, n) => String(s).padEnd(n);

  console.log("\nUNCLASSIFIED");
  console.log("═".repeat(96));
  console.log(`  ${path.basename(pkgPath)}  →  ${outDir}`);
  if (manifest) {
    console.log(`  project ${manifest.project} · manifest v${manifest.version} · packaged ${manifest.generated}`);
  }
  console.log("═".repeat(96));
  for (const kind of ["css", "js"]) {
    const a = alignment[kind];
    if (!a.entries.length) continue;
    console.log(`  ${kind.toUpperCase()} pairing: ${a.entries.length} manifest entries · ${a.label}${a.degraded ? " · DEGRADED" : ""}`);
  }
  console.log("─".repeat(96));

  if (!quiet) {
    console.log(`  ${pad("FILE", 50)}${pad("BYTES", 10)}${pad("HASH", 10)}NAMED BY`);
    console.log("─".repeat(96));
    for (const f of written) {
      const flag = f.match === false ? ` ✗ manifest ${f.expected}` : "";
      console.log(`  ${pad(f.relPath.slice(-48), 50)}${pad(f.bytes, 10)}${pad(f.hash, 10)}${TIER_LABEL[f.tier] || f.tier}${flag}`);
    }
    console.log("─".repeat(96));
  }

  const tally = {};
  for (const f of written) tally[f.tier] = (tally[f.tier] || 0) + 1;
  console.log(`  ${written.length} file(s) extracted`);
  for (const [tier, n] of Object.entries(tally).sort((a, b) => b[1] - a[1])) {
    console.log(`    ${String(n).padStart(3)}  ${TIER_LABEL[tier] || tier}`);
  }

  const mismatched = written.filter((f) => f.match === false);
  if (mismatched.length) {
    console.log(`\n  ${mismatched.length} file(s) named positionally with a hash that does NOT match:`);
    for (const f of mismatched) console.log(`    · ${f.relPath}`);
    console.log("  Expected when the packager rewrote a file on the way in (UMD→global wrap,");
    console.log("  local asset path swapped for a data: URI). On your own application code,");
    console.log("  treat it as a real finding and explain it before trusting the name.");
  }

  const unique = [...new Set(notes)];
  for (const n of unique.slice(0, 14)) console.log(`\n  NOTE: ${n}`);
  if (unique.length > 14) console.log(`\n  NOTE: … and ${unique.length - 14} more of the same kind.`);
  console.log("\nUNCLASSIFIED\n");
}

function main() {
  const argv = process.argv.slice(2);
  const quiet = argv.includes("--quiet");
  const mapFlag = argv.indexOf("--map");
  const mapPath = mapFlag !== -1 ? argv[mapFlag + 1] : null;
  const args = argv.filter((a, i) => !a.startsWith("--") && !(mapFlag !== -1 && i === mapFlag + 1));
  const pkgPath = args[0];

  if (!pkgPath) {
    console.log("UNCLASSIFIED\n");
    console.log("Usage: node wfc-decompile.js <package.html> [output-dir] [--map <overrides.json>] [--quiet]\n");
    console.log("  --map  JSON of {\"<block-hash>\": \"<output/path.js>\"} pinning names the");
    console.log("         evidence cannot settle. Hashes come from BLOCK-INVENTORY.md.\n");
    console.log("UNCLASSIFIED");
    process.exit(1);
  }
  if (!fs.existsSync(pkgPath)) fail(`no such file: ${pkgPath}`);

  let overrides = null;
  if (mapPath) {
    if (!fs.existsSync(mapPath)) fail(`no such map file: ${mapPath}`);
    try {
      overrides = JSON.parse(fs.readFileSync(mapPath, "utf8"));
    } catch (err) {
      fail(`could not parse map file ${mapPath}: ${err.message}`);
    }
    /* Allow comment keys so a map can explain itself. */
    for (const key of Object.keys(overrides)) if (!/^[0-9a-f]{6,}$/i.test(key)) delete overrides[key];
  }

  const peek = readManifest(fs.readFileSync(pkgPath, "utf8"));
  const outDir = path.resolve(args[1] || `${(peek && peek.project) || path.basename(pkgPath, ".html")}-decompiled`);

  report(pkgPath, outDir, decompile(pkgPath, outDir, overrides), quiet);
}

if (require.main === module) main();
module.exports = { decompile, readManifest, unwrapForge, scanRawTextBlocks };

/* UNCLASSIFIED */
