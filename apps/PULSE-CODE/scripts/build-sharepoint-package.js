/**
 * PULSE CODE — SharePoint package builder
 *
 * PULSE CODE is a self-contained single HTML file (all JS/CSS inline).
 * This script wraps it with a WFC-MANIFEST header that matches the format
 * used by PULSE, Tickets, and Calendar releases.
 *
 * Usage:
 *   node scripts/build-sharepoint-package.js [output-path]
 *
 * Output defaults to: releases/PULSE-CODE-v1.0.0.html
 */

"use strict";

const fs     = require("fs");
const path   = require("path");
const crypto = require("crypto");

const ROOT    = path.resolve(__dirname, "..");
const VERSION = "1.0.0";
const OUTPUT  = process.argv[2]
  || path.resolve(ROOT, "..", "..", "releases", `PULSE-CODE-v${VERSION}.html`);

const index  = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const hash   = (t) => crypto.createHash("sha256").update(t).digest("hex").slice(0, 8);
const escape = (s) => String(s).replace(/<\/script/gi, "<\\/script");

// ── Extract title ────────────────────────────────────────────────────────────
const title = (index.match(/<title>([^<]+)<\/title>/i) || [])[1] || "PULSE CODE";

// ── Extract head inline <script> blocks (theme init, etc.) ──────────────────
const headBlock = (index.match(/<head[^>]*>([\s\S]*?)<\/head>/i) || ["", ""])[1];

const headInlineScripts = [...headBlock.matchAll(
  /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi
)].map(m => `<script>\n${escape(m[1].trim())}\n</script>`).join("\n");

// ── Extract inline <style> blocks ───────────────────────────────────────────
const inlineStyles = [...index.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/gi)]
  .map(m => `<style>\n${m[1].trim()}\n</style>`)
  .join("\n");

// ── Collect CDN <script src="https://..."> tags from head (Monaco loader) ───
// These must stay external — Monaco loads its workers from the same CDN base URL.
const cdnScriptTags = [...headBlock.matchAll(
  /<script\s+src=["'](https?:\/\/[^"']+)["'][^>]*><\/script>/gi
)].map(m => m[0]).join("\n");

// ── Extract body HTML (strip all src-based script tags) ─────────────────────
const bodyInner = (index.match(/<body[^>]*>([\s\S]*?)<\/body>/i) || ["", ""])[1]
  .replace(/<script\s+src=["'][^"']+["'][^>]*><\/script>/gi, "")
  .trim();

// Separate body structure (non-script HTML) from the large inline script block
const bodyHtml = bodyInner.replace(/<script(?![^>]*\bsrc=)[^>]*>[\s\S]*?<\/script>/gi, "").trim();

const bodyScripts = [...bodyInner.matchAll(
  /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi
)].map(m => `<script>\n${escape(m[1].trim())}\n</script>`).join("\n");

// ── Build manifest ───────────────────────────────────────────────────────────
const manifest = [
  { kind: "js", path: "index.html#inline", external: false, hash: hash(bodyScripts) },
  { kind: "css", path: "index.html#style", external: false, hash: hash(inlineStyles) },
  {
    kind: "js-cdn", path: "https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/min/vs/loader.js",
    external: true, hash: "monaco-0.47.0"
  }
];

const manifestB64 = Buffer.from(JSON.stringify({
  version: 1, project: "PULSE-CODE",
  generated: new Date().toISOString(), index: "index.html", files: manifest
})).toString("base64");

// ── Assemble output ──────────────────────────────────────────────────────────
const html = [
  `<!--WFC-MANIFEST:${manifestB64}-->`,
  `<!DOCTYPE html><html lang="en"><head>`,
  `<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">`,
  `<title>${title}</title>`,
  headInlineScripts,
  inlineStyles,
  cdnScriptTags,
  `</head><body>`,
  bodyHtml,
  bodyScripts,
  `</body></html>`,
].filter(Boolean).join("\n");

fs.mkdirSync(path.dirname(OUTPUT), { recursive: true });
fs.writeFileSync(OUTPUT, html, "utf8");
console.log(OUTPUT);
