const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const ROOT = path.resolve(__dirname, "..");
const INDEX_PATH = process.argv[3] ? path.resolve(ROOT, process.argv[3]) : path.join(ROOT, "index.html");
const VERSION = "1.0.0";
const OUTPUT_PATH = process.argv[2] || path.join(ROOT, "..", "..", "releases", `PULSE-v${VERSION}.html`);

const MIME_TYPES = {
  ".css": "text/css",
  ".js": "application/javascript",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".ttf": "font/ttf",
  ".eot": "application/vnd.ms-fontobject",
  ".ico": "image/x-icon",
  ".html": "text/html",
  ".map": "application/json"
};

function readText(filePath) {
  return fs.readFileSync(filePath, "utf8");
}

function readBuffer(filePath) {
  return fs.readFileSync(filePath);
}

function fileExists(filePath) {
  try {
    fs.accessSync(filePath, fs.constants.R_OK);
    return true;
  } catch {
    return false;
  }
}

function hashText(value) {
  return crypto.createHash("sha256").update(value).digest("hex").slice(0, 8);
}

function extToMime(filePath) {
  return MIME_TYPES[path.extname(filePath).toLowerCase()] || "application/octet-stream";
}

function toDataUri(filePath) {
  return `data:${extToMime(filePath)};base64,${readBuffer(filePath).toString("base64")}`;
}

function escapeScriptContent(value) {
  return String(value).replace(/<\/script/gi, "<\\/script");
}

function normalizeCssUrl(rawValue) {
  return String(rawValue || "").trim().replace(/^['"]|['"]$/g, "");
}

/* index.html uses ?devcache=N for local cache-busting; strip that before
   resolving filesystem paths so the SharePoint package build can read them. */
function stripAssetQuery(href) {
  return String(href || "").split("?")[0].split("#")[0];
}

function inlineCssUrls(cssText, cssPath) {
  return cssText.replace(/url\(([^)]+)\)/g, (match, rawUrl) => {
    const cleaned = normalizeCssUrl(rawUrl);
    if (!cleaned || cleaned.startsWith("data:") || cleaned.startsWith("#") || /^https?:/i.test(cleaned)) {
      return match;
    }
    const targetPath = path.resolve(path.dirname(cssPath), cleaned);
    if (!fileExists(targetPath)) {
      return match;
    }
    return `url("${toDataUri(targetPath)}")`;
  });
}

function collectFavicon(indexHtml) {
  const match = indexHtml.match(/<link\s+rel=["']icon["'][^>]*href=["']([^"']+)["'][^>]*>/i);
  if (!match) return "";
  const href = match[1];
  if (/^data:/i.test(href)) return match[0];
  const iconPath = path.resolve(ROOT, href);
  if (!fileExists(iconPath)) return "";
  return `<link rel="icon" type="image/svg+xml" href="${toDataUri(iconPath)}">`;
}

function collectStyles(indexHtml, manifestFiles) {
  const styleHrefRegex = /<link\s+rel=["']stylesheet["']\s+href=["']([^"']+)["'][^>]*>/gi;
  const blocks = [];
  let match;
  while ((match = styleHrefRegex.exec(indexHtml))) {
    const href = stripAssetQuery(match[1]);
    if (/^https?:/i.test(href)) {
      throw new Error(`External stylesheet not supported for SharePoint package: ${href}`);
    }
    const cssPath = path.resolve(ROOT, href);
    const cssText = inlineCssUrls(readText(cssPath), cssPath);
    manifestFiles.push({ kind: "css", path: href, external: false, hash: hashText(cssText) });
    blocks.push(`<style>\n${cssText}\n</style>`);
  }
  return blocks.join("\n");
}

function collectHeadInlineScripts(indexHtml) {
  const headMatch = indexHtml.match(/<head[^>]*>([\s\S]*?)<\/head>/i);
  if (!headMatch) return "";
  return [...headMatch[1].matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)]
    .map((match) => `<script>\n${escapeScriptContent(match[1].trim())}\n<\/script>`)
    .join("\n");
}

/* SharePoint (and other hosts) often define AMD `define`. PptxGenJS's UMD then
   registers JSZip via define() instead of window.JSZip, and the rest of the
   bundle throws "JSZip is not defined" — leaving window.PptxGenJS unset. */
function wrapUmdAsBrowserGlobal(jsText, globalNames) {
  const names = Array.isArray(globalNames) ? globalNames : [globalNames];
  const assign = names
    .map((name) => `if (typeof ${name} !== "undefined") root.${name} = ${name};`)
    .join("\n");
  const stripped = String(jsText).replace(/\/\/# sourceMappingURL=.*$/gm, "");
  return [
    "(function (root) {",
    "  var define = undefined;",
    "  var module = undefined;",
    "  var exports = undefined;",
    stripped,
    assign,
    "})(typeof window !== \"undefined\" ? window : this);"
  ].join("\n");
}

function collectScripts(indexHtml, manifestFiles) {
  const scriptSrcRegex = /<script\s+src=["']([^"']+)["'][^>]*><\/script>/gi;
  const blocks = [];
  let match;
  while ((match = scriptSrcRegex.exec(indexHtml))) {
    const src = stripAssetQuery(match[1]);
    if (/^https?:/i.test(src)) {
      throw new Error(`External script not supported for SharePoint package: ${src}`);
    }
    const scriptPath = path.resolve(ROOT, src);
    let jsText = readText(scriptPath);
    jsText = jsText.replace(/assets\/images\/aewttr-seal-template\.png/g, toDataUri(path.join(ROOT, "assets", "images", "aewttr-seal-template.png")));
    jsText = jsText.replace(/assets\/images\/ttsd-seal-template\.png/g, toDataUri(path.join(ROOT, "assets", "images", "ttsd-seal-template.png")));
    jsText = jsText.replace(/assets\/images\/navair-seal-template\.png/g, toDataUri(path.join(ROOT, "assets", "images", "navair-seal-template.png")));
    jsText = jsText.replace(/assets\/images\/navair-wordmark-template\.png/g, toDataUri(path.join(ROOT, "assets", "images", "navair-wordmark-template.png")));
    jsText = jsText.replace(/assets\/images\/aewttr-seal\.png/g, toDataUri(path.join(ROOT, "assets", "images", "aewttr-seal.png")));
    jsText = jsText.replace(/assets\/images\/pulse-wordmark-white\.png/g, toDataUri(path.join(ROOT, "assets", "images", "pulse-wordmark-white.png")));
    jsText = jsText.replace(/assets\/images\/pulse-wordmark-black\.png/g, toDataUri(path.join(ROOT, "assets", "images", "pulse-wordmark-black.png")));
    jsText = jsText.replace(/assets\/images\/overview-seal\.png/g, toDataUri(path.join(ROOT, "assets", "images", "overview-seal.png")));
    const ttsdSvgPath = path.join(ROOT, "assets", "images", "ttsd-seal.svg");
    if (fs.existsSync(ttsdSvgPath)) jsText = jsText.replace(/assets\/images\/ttsd-seal\.svg/g, toDataUri(ttsdSvgPath));
    const navairSvgPath = path.join(ROOT, "assets", "images", "navair-seal.svg");
    if (fs.existsSync(navairSvgPath)) jsText = jsText.replace(/assets\/images\/navair-seal\.svg/g, toDataUri(navairSvgPath));
    if (/pptxgen/i.test(src)) {
      jsText = wrapUmdAsBrowserGlobal(jsText, ["JSZip", "PptxGenJS"]);
    }
    if (/jspdf/i.test(src)) {
      jsText = wrapUmdAsBrowserGlobal(jsText, ["jspdf"]);
    }
    manifestFiles.push({ kind: "js", path: src, external: false, hash: hashText(jsText) });
    blocks.push(`<script>\n${escapeScriptContent(jsText)}\n<\/script>`);
  }
  return blocks.join("\n");
}

function buildPackage() {
  const indexHtml = readText(INDEX_PATH);
  const titleMatch = indexHtml.match(/<title>([^<]+)<\/title>/i);
  const bodyMatch = indexHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);

  if (!bodyMatch) {
    throw new Error("Could not find <body> in index.html.");
  }

  const manifestFiles = [];
  const styleBlocks = collectStyles(indexHtml, manifestFiles);
  const scriptBlocks = collectScripts(indexHtml, manifestFiles);
  const headInlineScripts = collectHeadInlineScripts(indexHtml);
  const faviconTag = collectFavicon(indexHtml);
  const title = titleMatch ? titleMatch[1] : "AEWTTR-PULSE";
  const bodyInner = bodyMatch[1].replace(/<script\s+src=["'][^"']+["'][^>]*><\/script>/gi, "").trim();

  const manifest = {
    version: 1,
    project: path.basename(ROOT),
    generated: new Date().toISOString(),
    index: path.basename(INDEX_PATH),
    files: manifestFiles
  };
  const manifestB64 = Buffer.from(JSON.stringify(manifest), "utf8").toString("base64");

  return [
    `<!--WFC-MANIFEST:${manifestB64}-->`,
    "<!DOCTYPE html>",
    '<html lang="en">',
    "<head>",
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
    `  <title>${title}</title>`,
    faviconTag ? `  ${faviconTag}` : "",
    "",
    headInlineScripts,
    "",
    styleBlocks,
    "</head>",
    "<body>",
    "",
    bodyInner,
    "",
    scriptBlocks,
    "</body>",
    "</html>"
  ].join("\n");
}

const finalHtml = buildPackage();
fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
fs.writeFileSync(OUTPUT_PATH, finalHtml, "utf8");
console.log(OUTPUT_PATH);
