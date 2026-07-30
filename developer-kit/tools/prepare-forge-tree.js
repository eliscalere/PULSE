#!/usr/bin/env node
/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   prepare-forge-tree.js — regenerate developer-kit/forge-ide-decompiled/

   Runs wfc-decompile.js against forge/Forge.html and then trims the three kinds
   of byte that are not worth committing, because they are all derivable from an
   input that is already in the repo:

     1. _vendor-cdn/       2.2 MB of unmodified public library releases →
                           replaced by a README listing them with versions
     2. _child-app-as-shipped.html   5.8 MB = Forge.html minus one base64 decode
     3. the CHILD_HTML_B64 payload inside _forge-parent-shell.html →
                           elided, keeping the 38 KB of wrapper logic

   Hand-written files in the output directory (README.md) are preserved: the
   decompile runs into a scratch directory and is copied in, rather than the
   target being deleted. Deleting it is how a hand-written README gets lost.

   Usage:
     node developer-kit/tools/prepare-forge-tree.js

   Zero dependencies. Node 14+.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const { decompile, readManifest } = require("./wfc-decompile.js");

const KIT = path.resolve(__dirname, "..");
const PACKAGE = path.join(KIT, "forge", "Forge.html");
const TARGET = path.join(KIT, "forge-ide-decompiled");

/* Files in TARGET that are authored, not generated. Never overwritten, never
   removed. */
const AUTHORED = new Set(["README.md"]);

function rmrf(target) {
  fs.rmSync(target, { recursive: true, force: true });
}

function copyTree(from, to) {
  fs.mkdirSync(to, { recursive: true });
  for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
    const src = path.join(from, entry.name);
    const dst = path.join(to, entry.name);
    if (entry.isDirectory()) copyTree(src, dst);
    else fs.copyFileSync(src, dst);
  }
}

function main() {
  if (!fs.existsSync(PACKAGE)) {
    console.error(`\n  ERROR: ${PACKAGE} not found.\n`);
    process.exit(1);
  }

  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), "forge-decompile-"));
  console.log("\nUNCLASSIFIED\n");
  console.log(`  decompiling ${path.relative(KIT, PACKAGE)} → scratch`);

  const result = decompile(PACKAGE, scratch);
  const verified = result.written.filter((f) => f.match === true).length;
  const untracked = result.written.filter((f) => f.match === null).length;
  console.log(`  ${result.written.length} file(s): ${verified} hash-verified, ${untracked} untracked`);

  /* ── Trim 1: the shipped-copy of the child ──────────────────────────────── */
  rmrf(path.join(scratch, "_child-app-as-shipped.html"));

  /* ── Trim 2: elide the base64 payload in the parent shell ───────────────── */
  const shellPath = path.join(scratch, "_forge-parent-shell.html");
  if (fs.existsSync(shellPath)) {
    const shell = fs.readFileSync(shellPath, "utf8");
    const marker = "CHILD_HTML_B64";
    const declared = shell.indexOf(marker);
    const open = declared === -1 ? -1 : shell.indexOf('"', declared + marker.length);
    const close = open === -1 ? -1 : shell.indexOf('"', open + 1);
    if (close !== -1) {
      const size = close - open - 1;
      const elided =
        shell.slice(0, open + 1) +
        `[[[ CHILD_HTML_B64 payload elided by prepare-forge-tree.js: ${size} base64 chars ` +
        `= the entire Forge app. Regenerate from developer-kit/forge/Forge.html ]]]` +
        shell.slice(close);
      fs.writeFileSync(shellPath, elided, "utf8");
      console.log(`  elided ${size} base64 chars from _forge-parent-shell.html`);
    }
  }

  /* ── Trim 3: vendor bundles → an inventory ──────────────────────────────── */
  const manifest = result.manifest || readManifest(fs.readFileSync(PACKAGE, "utf8"));
  const external = (manifest && manifest.files ? manifest.files : []).filter((f) => f.external);
  rmrf(path.join(scratch, "_vendor-cdn"));
  fs.mkdirSync(path.join(scratch, "_vendor-cdn"), { recursive: true });
  fs.writeFileSync(
    path.join(scratch, "_vendor-cdn", "README.md"),
    [
      "<!-- UNCLASSIFIED -->",
      "# _vendor-cdn — extracted, then deliberately not committed",
      "",
      "Forge fetches these at compile time and inlines them into the shipped file,",
      "recording them in the manifest with `external: true`. `wfc-decompile.js` writes",
      "them here, but they are unmodified public library releases, so the bytes are not",
      `kept in git. What Forge shipped on ${(manifest && manifest.generated) || "this build"}:`,
      "",
      "| kind | URL |",
      "|---|---|",
      ...external.map((f) => `| ${f.kind} | ${f.path} |`),
      "",
      "Two things worth noticing:",
      "",
      "1. Forge inlines CDN dependencies rather than rejecting them. The PULSE packagers",
      "   throw on any `https://` reference instead. Both end up dependency-free at",
      "   runtime; Forge just does the fetching for you at compile time.",
      "2. It has to. The Firepit host injects `default-src 'none'` into the child frame,",
      "   so a runtime CDN fetch would be blocked outright. See",
      "   `../../notes/02-firepit-webpart-internals.md`.",
      "",
      "Regenerate the full tree, vendor bundles included:",
      "",
      "```bash",
      "node developer-kit/tools/wfc-decompile.js developer-kit/forge/Forge.html /tmp/forge-full",
      "```",
      "",
      "<!-- UNCLASSIFIED -->",
      "",
    ].join("\n"),
    "utf8"
  );
  console.log(`  replaced _vendor-cdn/ with an inventory of ${external.length} external dependency(ies)`);

  /* ── Publish, preserving authored files ─────────────────────────────────── */
  if (fs.existsSync(TARGET)) {
    for (const entry of fs.readdirSync(TARGET)) {
      if (!AUTHORED.has(entry)) rmrf(path.join(TARGET, entry));
    }
  }
  copyTree(scratch, TARGET);
  rmrf(scratch);

  const preserved = fs.existsSync(TARGET)
    ? fs.readdirSync(TARGET).filter((e) => AUTHORED.has(e))
    : [];
  console.log(`  wrote ${path.relative(KIT, TARGET)}${preserved.length ? ` (preserved ${preserved.join(", ")})` : ""}`);
  console.log("\nUNCLASSIFIED\n");
}

if (require.main === module) main();

/* UNCLASSIFIED */
