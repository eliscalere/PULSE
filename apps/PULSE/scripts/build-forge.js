const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const INDEX_PATH = path.join(ROOT, "index.html");
const TEMPLATE_PATH = process.argv[2];
const OUTPUT_PATH = process.argv[3] || path.join(ROOT, "releases", "forge-builds", "Forge.html");

if (!TEMPLATE_PATH) {
  throw new Error("Usage: node scripts/build-forge.js <template-forge.html> [output-file]");
}

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
  ".html": "text/html"
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

function escapeScriptContent(value) {
  return String(value).replace(/<\/script/gi, "<\\/script");
}

function escapeTemplateLiteral(value) {
  return String(value)
    .replace(/\\/g, "\\\\")
    .replace(/`/g, "\\`")
    .replace(/\$\{/g, "\\${");
}

function toDataUri(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const mimeType = MIME_TYPES[ext] || "application/octet-stream";
  const data = readBuffer(filePath).toString("base64");
  return `data:${mimeType};base64,${data}`;
}

function inlineCssUrls(cssText, cssPath) {
  return cssText.replace(/url\(([^)]+)\)/g, (match, rawUrl) => {
    const cleaned = String(rawUrl).trim().replace(/^['"]|['"]$/g, "");
    if (!cleaned || cleaned.startsWith("data:") || /^https?:/i.test(cleaned) || cleaned.startsWith("#")) {
      return match;
    }

    const assetPath = path.resolve(path.dirname(cssPath), cleaned);
    if (!fileExists(assetPath)) {
      return match;
    }

    return `url("${toDataUri(assetPath)}")`;
  });
}

function bundleStyles(indexHtml) {
  const styleHrefRegex = /<link\s+rel=["']stylesheet["']\s+href=["']([^"']+)["'][^>]*>/gi;
  let match;
  const blocks = [];

  while ((match = styleHrefRegex.exec(indexHtml))) {
    const href = match[1];
    if (/^https?:/i.test(href)) {
      throw new Error(`External stylesheet not supported for Forge bundle: ${href}`);
    }
    const cssPath = path.resolve(ROOT, href.split("?")[0]);
    const cssText = inlineCssUrls(readText(cssPath), cssPath);
    blocks.push(`<style data-source="${href}">\n${cssText}\n</style>`);
  }

  return blocks.join("\n");
}

function bundleScripts(indexHtml) {
  const scriptSrcRegex = /<script\s+src=["']([^"']+)["'][^>]*><\/script>/gi;
  let match;
  const blocks = [];

  while ((match = scriptSrcRegex.exec(indexHtml))) {
    const src = match[1];
    if (/^https?:/i.test(src)) {
      throw new Error(`External script not supported for Forge bundle: ${src}`);
    }
    const scriptPath = path.resolve(ROOT, src.split("?")[0]);
    let jsText = readText(scriptPath);

    // Replace local asset references that would break once the child app is srcdoc-based.
    jsText = jsText.replace(/assets\/images\/aewttr-seal\.png/g, toDataUri(path.join(ROOT, "assets", "images", "aewttr-seal.png")));

    blocks.push(`<script data-source="${src}">\n${escapeScriptContent(jsText)}\n</script>`);
  }

  return blocks.join("\n");
}

function buildChildHtml() {
  const indexHtml = readText(INDEX_PATH);
  const titleMatch = indexHtml.match(/<title>([^<]+)<\/title>/i);
  const title = titleMatch ? titleMatch[1] : "AEWTTR-PULSE";
  const bodyMatch = indexHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);

  if (!bodyMatch) {
    throw new Error("Could not find <body> in index.html");
  }

  const bodyInner = bodyMatch[1]
    .replace(/<script\s+src=["'][^"']+["'][^>]*><\/script>/gi, "")
    .trim();

  const styles = bundleStyles(indexHtml);
  const scripts = bundleScripts(indexHtml);

  return [
    "<!DOCTYPE html>",
    '<html lang="en">',
    "<head>",
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
    `<title>${title}</title>`,
    styles,
    "</head>",
    "<body>",
    bodyInner,
    scripts,
    "</body>",
    "</html>"
  ].join("\n");
}

function buildForgeFile() {
  const template = readText(path.resolve(TEMPLATE_PATH));
  const childHtml = buildChildHtml();
  const childHtmlB64 = Buffer.from(childHtml, "utf8").toString("base64");

  if (!/const CHILD_HTML_B64 = ".*?";/s.test(template)) {
    throw new Error("Could not find CHILD_HTML_B64 in Forge template.");
  }

  const output = template.replace(/const CHILD_HTML_B64 = ".*?";/s, `const CHILD_HTML_B64 = "${childHtmlB64}";`);

  fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
  fs.writeFileSync(OUTPUT_PATH, output, "utf8");
  const readablePath = OUTPUT_PATH.replace(/\.html$/i, "-readable.html");
  const readableOutput = output
    .replace(/const CHILD_HTML_B64 = ".*?";/s, `const CHILD_HTML = String.raw\`${escapeTemplateLiteral(childHtml)}\`;`)
    .replace(/decodeB64Utf8\(CHILD_HTML_B64\)/g, "CHILD_HTML");
  fs.writeFileSync(readablePath, readableOutput, "utf8");
  return { packed: OUTPUT_PATH, readable: readablePath };
}

const result = buildForgeFile();
console.log(result.packed);
console.log(result.readable);
