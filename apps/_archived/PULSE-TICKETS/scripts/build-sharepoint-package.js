const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

// Root of THIS app (PULSE-TICKETS/)
const root = path.resolve(__dirname, "..");

// The apps/ parent — needed to resolve ../PULSE/ cross-app imports
const appsDir = path.resolve(root, "..");

const VERSION = "1.0.0";
// ARCHIVED app — one directory deeper than a live app/, and the output
// filename is deliberately distinct from releases/PULSE-Tickets-v*.html
// so this can never again silently overwrite the real Tickets package
// (built from apps/PULSE/tickets.html via build-travel-packages.js).
const output = process.argv[2]
  || path.join(root, "..", "..", "..", "releases", "_legacy", `PULSE-Tickets-STANDALONE-v${VERSION}.html`);

const index = fs.readFileSync(path.join(root, "index.html"), "utf8");
const manifest = [];

const mime = {
  ".css": "text/css", ".js": "text/javascript",
  ".woff": "font/woff", ".woff2": "font/woff2", ".ttf": "font/ttf",
  ".svg": "image/svg+xml", ".png": "image/png", ".ico": "image/x-icon"
};

const hash = (text) =>
  crypto.createHash("sha256").update(text).digest("hex").slice(0, 8);

const dataUri = (file) =>
  `data:${mime[path.extname(file).toLowerCase()] || "application/octet-stream"};base64,${fs.readFileSync(file).toString("base64")}`;

// Resolve an href from index.html — supports ../PULSE/ cross-dir paths
function source(href) {
  const cleaned = href.split("?")[0].split("#")[0];
  // Try relative to root first, then relative to appsDir
  const fromRoot = path.resolve(root, cleaned);
  if (fs.existsSync(fromRoot)) return fromRoot;
  const fromApps = path.resolve(appsDir, cleaned);
  if (fs.existsSync(fromApps)) return fromApps;
  return fromRoot; // will fail with clear error
}

function inlineCss(href) {
  const file = source(href);
  let css = fs.readFileSync(file, "utf8");
  css = css.replace(/url\(([^)]+)\)/g, (whole, raw) => {
    const url = String(raw).trim().replace(/^['"]|['"]$/g, "");
    if (!url || /^(data:|https?:|#)/i.test(url)) return whole;
    const asset = path.resolve(path.dirname(file), url.split("#")[0].split("?")[0]);
    return fs.existsSync(asset) ? `url("${dataUri(asset)}")` : whole;
  });
  manifest.push({ kind: "css", path: href, external: false, hash: hash(css) });
  return `<style>\n${css}\n</style>`;
}

function inlineJs(src) {
  const file = source(src);
  const js = fs.readFileSync(file, "utf8").replace(/<\/script/gi, "<\\/script");
  manifest.push({ kind: "js", path: src, external: false, hash: hash(js) });
  return `<script>\n${js}\n<\/script>`;
}

const styles = [...index.matchAll(/<link\s+rel=["']stylesheet["']\s+href=["']([^"']+)["'][^>]*>/gi)]
  .map((m) => inlineCss(m[1])).join("\n");

const scripts = [...index.matchAll(/<script\s+src=["']([^"']+)["'][^>]*><\/script>/gi)]
  .map((m) => inlineJs(m[1])).join("\n");

const headMatch = index.match(/<head[^>]*>([\s\S]*?)<\/head>/i);
const headInlineScripts = headMatch
  ? [...headMatch[1].matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)]
      .map((m) => `<script>\n${m[1].trim().replace(/<\/script/gi, "<\\/script")}\n<\/script>`)
      .join("\n")
  : "";

const body = index.match(/<body[^>]*>([\s\S]*?)<\/body>/i)[1]
  .replace(/<script\s+src=["'][^"']+["'][^>]*><\/script>/gi, "").trim();

const titleMatch = index.match(/<title>([^<]+)<\/title>/i);
const title = titleMatch ? titleMatch[1] : "PULSE Tickets";

const metadata = Buffer.from(JSON.stringify({
  version: 1, project: path.basename(root),
  generated: new Date().toISOString(), index: "index.html", files: manifest
})).toString("base64");

fs.mkdirSync(path.dirname(output), { recursive: true });
fs.writeFileSync(output, [
  `<!--WFC-MANIFEST:${metadata}-->`,
  `<!DOCTYPE html><html lang="en"><head>`,
  `<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">`,
  `<title>${title}</title>`,
  headInlineScripts,
  styles,
  `</head><body>`,
  body,
  scripts,
  `</body></html>`
].join("\n"), "utf8");

console.log(output);
