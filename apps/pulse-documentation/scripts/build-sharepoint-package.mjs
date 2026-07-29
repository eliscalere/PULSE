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
    ...publicFilesUnder("screenshots"),
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
  window.__PULSE_ASSETS__ = assetUrls;
  /* Exposed so the reader can populate itself synchronously instead of waiting
     on fetch() — the text is already in the document, so there is nothing to
     wait for and the boot overlay can clear on the first commit. */
  window.__PULSE_SOURCE_TEXT__ = sourceText;
  window.getAssetUrl = function (path) {
    return assetUrls[path] || path;
  };
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
    if (!root) return;
    if (root.nodeType === 1) {
      if (root.matches && root.matches("[src]")) rewriteAsset(root, "src");
      if (root.matches && root.matches("[href]")) rewriteAsset(root, "href");
    }
    if (root.querySelectorAll) {
      root.querySelectorAll("[src]").forEach(function (element) { rewriteAsset(element, "src"); });
      root.querySelectorAll("[href]").forEach(function (element) { rewriteAsset(element, "href"); });
    }
  }

  function startAssetRewriter() {
    rewriteTree(document);
    if (document.documentElement) {
      var observer = new MutationObserver(function (records) {
        records.forEach(function (record) {
          if (record.type === "attributes") {
            rewriteAsset(record.target, record.attributeName);
          } else {
            record.addedNodes.forEach(rewriteTree);
          }
        });
      });
      observer.observe(document.documentElement, {
        subtree: true,
        childList: true,
        attributes: true,
        attributeFilter: ["src", "href"]
      });
    }
  }

  startAssetRewriter();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startAssetRewriter, { once: true });
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
  /* Records every stylesheet we fold into the document so the RSC payload can
     be re-pointed at a self-contained copy (see neutralizeExternalAssetRefs). */
  const cssDataUris = {};
  let result = html.replace(
    /<link\s+rel=["']stylesheet["']\s+href=["']([^"']+)["'][^>]*>/gi,
    (tag, href) => {
      const cssPath = path.join(CLIENT_ROOT, href.replace(/^\//, ""));
      const css = inlineCssUrls(readFileSync(cssPath, "utf8"), cssPath, manifestFiles);
      manifestFiles.push({ kind: "css", path: href.replace(/^\//, ""), external: false, hash: hash(css) });
      cssDataUris[href] = `data:text/css;base64,${Buffer.from(css, "utf8").toString("base64")}`;
      return `<style data-firepit-bundle>\n${css}\n</style>`;
    },
  );

  result = result.replace(/<style([^>]*)>([\s\S]*?)<\/style>/gi, (tag, attributes, css) => {
    const syntheticPath = path.join(CLIENT_ROOT, "inline.css");
    return `<style${attributes}>${inlineCssUrls(css, syntheticPath, manifestFiles)}</style>`;
  });
  return { html: result, cssDataUris };
}

/* The rendered RSC payload still names the on-disk build chunks: a
   `precedence`-managed <link rel="stylesheet"> for the route CSS, a :HL preload
   hint for the same file, and the dynamic-import chunk table. None of those
   paths exist inside a single-file package.

   The stylesheet is the damaging one. React suspends the commit until a
   precedence-managed stylesheet resolves, so it will not reveal the app until
   that request settles. Served over plain HTTP the missing file 404s quickly
   and React continues, which is why the package looks fine locally. Inside a
   SharePoint/Firepit iframe the same path resolves against the host site and
   comes back as a slow auth/HTML response instead of a fast error, so React
   stays suspended and nothing paints until a user interaction forces a commit
   — the "I have to click the page before it opens" symptom.

   Re-point the stylesheet at an inlined data: URI (loads instantly, keeps
   React's resource bookkeeping intact) and make the JS chunk names inert. */
function neutralizeExternalAssetRefs(html, cssDataUris) {
  let result = html.replace(/:HL\[\\"\/assets\/[A-Za-z0-9_.\/-]+\.css\\",\\"style\\"\]\\n/g, "");

  const byBasename = {};
  for (const [href, dataUri] of Object.entries(cssDataUris)) {
    byBasename[path.posix.basename(href)] = dataUri;
  }
  return result.replace(
    /\/assets\/([A-Za-z0-9_.-]+\.css)/g,
    (match, basename) => byBasename[basename] ?? match,
  );
}

/* The bundled client keeps a chunk-filename table, consumed only to build
   preload hints — the modules themselves resolve inline through
   `Promise.resolve().then(...)` once esbuild has folded them in. Those names no
   longer exist next to the package, so every hint is a wasted request (a 404
   locally, a slow host-site response inside SharePoint). Collapse the lookup to
   an empty list so nothing is preloaded; module resolution is untouched.

   Runs after bundling, since that is when the table enters the document. */
function neutralizeChunkTable(html) {
  const chunkTable = /(\w+)=\((\w+),(\w+)=\1,(\w+)=\3\.f\|\|\(\3\.f=\[[^\]]*\]\)\)=>\2\.map\(\w+=>\4\[\w+\]\)/g;
  if (!chunkTable.test(html)) {
    console.warn("Warning: chunk preload table not found — packaged build may request missing chunks.");
    return html;
  }
  chunkTable.lastIndex = 0;
  return html.replace(chunkTable, (match, fn, arg, self, table) => `${fn}=(${arg},${self}=${fn},${table}=${self}.f||(${self}.f=[]))=>[]`);
}

async function buildPackage() {
  if (!existsSync(SERVER_ENTRY)) {
    throw new Error("Missing dist output. Run npm run build before packaging.");
  }

  const manifestFiles = [];
  let html = await renderHtml();
  const styled = inlineStyles(html, manifestFiles);
  html = styled.html;
  html = neutralizeExternalAssetRefs(html, styled.cssDataUris);
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
  html = neutralizeChunkTable(html);

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
