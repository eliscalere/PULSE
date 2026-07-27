import { buildSync } from "esbuild";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { createHash } from "node:crypto";
import path from "node:path";
import { pathToFileURL } from "node:url";

const ROOT = path.resolve(import.meta.dirname, "..");
const CLIENT_ROOT = path.join(ROOT, "dist", "client");
const SERVER_ENTRY = path.join(ROOT, "dist", "server", "index.js");
const PUBLIC_ROOT = path.join(ROOT, "public");
const VERSION = "1.0.0";
const OUTPUT_PATH = path.resolve(
  process.argv[2] ?? path.join(ROOT, "..", "..", "releases", `PULSE-Documentation-v${VERSION}.html`),
);

const MIME_TYPES = {
  ".css": "text/css",
  ".gif": "image/gif",
  ".html": "text/html",
  ".ico": "image/x-icon",
  ".jpeg": "image/jpeg",
  ".jpg": "image/jpeg",
  ".js": "application/javascript",
  ".json": "application/json",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".zip": "application/zip",
};

function hash(value) {
  return createHash("sha256").update(value).digest("hex").slice(0, 12);
}

function mimeFor(filePath) {
  return MIME_TYPES[path.extname(filePath).toLowerCase()] ?? "application/octet-stream";
}

function dataUri(filePath) {
  return `data:${mimeFor(filePath)};base64,${readFileSync(filePath).toString("base64")}`;
}

function escapeScript(value) {
  return String(value).replace(/<\/script/gi, "<\\/script");
}

function publicFilesUnder(relativeDirectory) {
  const start = path.join(PUBLIC_ROOT, relativeDirectory);
  if (!existsSync(start)) return [];
  const files = [];
  const visit = (directory) => {
    for (const entry of readdirSync(directory)) {
      const fullPath = path.join(directory, entry);
      if (statSync(fullPath).isDirectory()) visit(fullPath);
      else files.push(fullPath);
    }
  };
  visit(start);
  return files;
}

function resolveClientAsset(assetPath, cssPath) {
  const clean = assetPath.split("?")[0].split("#")[0];
  if (clean.startsWith("/")) return path.join(CLIENT_ROOT, clean.slice(1));
  return path.resolve(path.dirname(cssPath), clean);
}

function inlineCssUrls(cssText, cssPath, manifestFiles) {
  return cssText.replace(/url\(([^)]+)\)/g, (match, rawValue) => {
    const value = String(rawValue).trim().replace(/^['"]|['"]$/g, "");
    if (!value || /^(data:|https?:|#)/i.test(value)) return match;
    const assetPath = resolveClientAsset(value, cssPath);
    if (!existsSync(assetPath)) return match;
    const bytes = readFileSync(assetPath);
    const relativePath = path.relative(CLIENT_ROOT, assetPath).replaceAll(path.sep, "/");
    if (!manifestFiles.some((file) => file.path === relativePath)) {
      manifestFiles.push({ kind: "asset", path: relativePath, external: false, hash: hash(bytes) });
    }
    return `url("${dataUri(assetPath)}")`;
  });
}

function collectPublicAssets(manifestFiles) {
  const sourceText = {};
  const assetUrls = {};
  const files = [
    ...publicFilesUnder("brand-assets"),
    ...publicFilesUnder("source-pdfs"),
    ...publicFilesUnder("source-text"),
    path.join(PUBLIC_ROOT, "favicon.svg"),
  ].filter((filePath) => existsSync(filePath));

  for (const filePath of files) {
    const relativePath = path.relative(PUBLIC_ROOT, filePath).replaceAll(path.sep, "/");
    const publicPath = `/${relativePath}`;
    const bytes = readFileSync(filePath);
    manifestFiles.push({ kind: "public", path: relativePath, external: false, hash: hash(bytes) });
    if (relativePath.startsWith("source-text/")) {
      sourceText[publicPath] = bytes.toString("utf8");
    } else {
      assetUrls[publicPath] = dataUri(filePath);
    }
  }

  return { sourceText, assetUrls };
}

function createAssetRuntime(sourceText, assetUrls) {
  const sourceJson = JSON.stringify(sourceText).replaceAll("<", "\\u003c");
  const assetJson = JSON.stringify(assetUrls).replaceAll("<", "\\u003c");
  return `
(function () {
  var sourceText = ${sourceJson};
  var assetUrls = ${assetJson};
  var nativeFetch = typeof window.fetch === "function" ? window.fetch.bind(window) : null;

  function pathnameFor(value) {
    try {
      return new URL(value, window.location.href).pathname;
    } catch (_) {
      return String(value || "").split("#")[0].split("?")[0];
    }
  }

  window.fetch = function (input, init) {
    var value = typeof input === "string" || input instanceof URL ? String(input) : input && input.url;
    var pathname = pathnameFor(value);
    if (Object.prototype.hasOwnProperty.call(sourceText, pathname)) {
      return Promise.resolve(new Response(sourceText[pathname], {
        status: 200,
        headers: { "Content-Type": "text/plain; charset=utf-8" }
      }));
    }
    if (!nativeFetch) return Promise.reject(new Error("No network fetch implementation is available."));
    return nativeFetch(input, init);
  };

  function rewriteAsset(element, attribute) {
    var value = element.getAttribute(attribute);
    if (!value || value.indexOf("data:") === 0) return;
    var url;
    try { url = new URL(value, window.location.href); } catch (_) { return; }
    var replacement = assetUrls[url.pathname];
    if (replacement) element.setAttribute(attribute, replacement + (url.hash || ""));
  }

  function rewriteTree(root) {
    if (root.nodeType === 1) {
      if (root.matches("[src]")) rewriteAsset(root, "src");
      if (root.matches("[href]")) rewriteAsset(root, "href");
    }
    if (root.querySelectorAll) {
      root.querySelectorAll("[src]").forEach(function (element) { rewriteAsset(element, "src"); });
      root.querySelectorAll("[href]").forEach(function (element) { rewriteAsset(element, "href"); });
    }
  }

  function startAssetRewriter() {
    rewriteTree(document);
    new MutationObserver(function (records) {
      records.forEach(function (record) {
        if (record.type === "attributes") {
          rewriteAsset(record.target, record.attributeName);
        } else {
          record.addedNodes.forEach(rewriteTree);
        }
      });
    }).observe(document.documentElement, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ["src", "href"]
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () {
      window.setTimeout(startAssetRewriter, 0);
    }, { once: true });
  } else {
    window.setTimeout(startAssetRewriter, 0);
  }
})();`;
}

function inlineHtmlAssetTags(html, assetUrls) {
  return html.replace(/<(?:a|img|link)\b[^>]*>/gi, (tag) =>
    tag.replace(/\b(src|href)=(["'])(.*?)\2/gi, (attribute, name, quote, value) => {
      if (!value || value.startsWith("data:")) return attribute;
      let url;
      try {
        url = new URL(value, "http://firepit.local/");
      } catch {
        return attribute;
      }
      const replacement = assetUrls[url.pathname];
      return replacement ? `${name}=${quote}${replacement}${url.hash || ""}${quote}` : attribute;
    }),
  );
}

async function renderHtml() {
  const workerUrl = pathToFileURL(SERVER_ENTRY);
  workerUrl.searchParams.set("firepit", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  const response = await worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
  if (!response.ok) throw new Error(`Unable to render documentation HTML (${response.status}).`);
  return response.text();
}

function bundleClient(html, manifestFiles) {
  const entryMatch = html.match(/<script\s+id=["']_R_["'][^>]*>\s*import\(["'`]\/assets\/([^"'`]+\.js)["'`]\)\s*<\/script>/i);
  if (!entryMatch) throw new Error("Unable to locate the vinext browser entry module.");
  const entryPath = path.join(CLIENT_ROOT, "assets", entryMatch[1]);
  const result = buildSync({
    entryPoints: [entryPath],
    bundle: true,
    format: "esm",
    platform: "browser",
    target: ["es2022"],
    minify: true,
    sourcemap: false,
    write: false,
  });
  const bundle = result.outputFiles[0].text;
  manifestFiles.push({ kind: "js", path: `assets/${entryMatch[1]}`, external: false, hash: hash(bundle) });
  const inlined = `<script id="_R_" type="module">\n${escapeScript(bundle)}\n</script>`;
  return {
    html: html.replace(entryMatch[0], () => inlined),
    bundle,
  };
}

function inlineStyles(html, manifestFiles) {
  let result = html.replace(
    /<link\s+rel=["']stylesheet["']\s+href=["']([^"']+)["'][^>]*>/gi,
    (tag, href) => {
      const cssPath = path.join(CLIENT_ROOT, href.replace(/^\//, ""));
      const css = inlineCssUrls(readFileSync(cssPath, "utf8"), cssPath, manifestFiles);
      manifestFiles.push({ kind: "css", path: href.replace(/^\//, ""), external: false, hash: hash(css) });
      return `<style data-firepit-bundle>\n${css}\n</style>`;
    },
  );

  result = result.replace(/<style([^>]*)>([\s\S]*?)<\/style>/gi, (tag, attributes, css) => {
    const syntheticPath = path.join(CLIENT_ROOT, "inline.css");
    return `<style${attributes}>${inlineCssUrls(css, syntheticPath, manifestFiles)}</style>`;
  });
  return result;
}

async function buildPackage() {
  if (!existsSync(SERVER_ENTRY)) {
    throw new Error("Missing dist output. Run npm run build before packaging.");
  }

  const manifestFiles = [];
  let html = await renderHtml();
  html = inlineStyles(html, manifestFiles);
  html = html
    .replace(/<link\b(?=[^>]*\brel=["'](?:modulepreload|preload)["'])[^>]*>/gi, "")
    .replace(/<meta\b(?=[^>]*\b(?:property|name)=["'](?:og:image|twitter:image)["'])[^>]*>/gi, "");

  const { sourceText, assetUrls } = collectPublicAssets(manifestFiles);
  const favicon = assetUrls["/favicon.svg"];
  if (favicon) {
    html = html.replace(/(<link\b[^>]*\bhref=["'])\/favicon\.svg(["'][^>]*>)/gi, `$1${favicon}$2`);
  }
  html = inlineHtmlAssetTags(html, assetUrls);

  const bundled = bundleClient(html, manifestFiles);
  const runtimeBlock = `<script>\n${escapeScript(createAssetRuntime(sourceText, assetUrls))}\n</script>\n<script id="_R_"`;
  html = bundled.html.replace(/<script\s+id=["']_R_["']/i, () => runtimeBlock);

  const manifest = {
    version: 1,
    project: "PULSE-Documentation",
    generated: new Date().toISOString(),
    index: "index.html",
    files: manifestFiles,
  };
  const manifestB64 = Buffer.from(JSON.stringify(manifest), "utf8").toString("base64");
  const finalHtml = `<!--WFC-MANIFEST:${manifestB64}-->\n${html}`;

  mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
  writeFileSync(OUTPUT_PATH, finalHtml, "utf8");
  console.log(OUTPUT_PATH);
}

await buildPackage();
