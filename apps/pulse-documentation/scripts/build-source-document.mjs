/* Builds a controlled source document from a single content file, emitting both
   artifacts the documentation library depends on:

     public/source-pdfs/<slug>.pdf    the controlled document
     public/source-text/<slug>.txt    the extraction the site searches

   Both come from the same block list, so the PDF and the searchable text cannot
   drift apart — which is the failure mode that left the library describing an
   older feature set than the application.

   Usage:
     node apps/pulse-documentation/scripts/build-source-document.mjs 09

   Requires playwright for PDF rendering. Set PLAYWRIGHT_NODE_PATH if playwright
   is installed outside this package.

   The plain-text output deliberately follows the conventions the reader's parser
   expects (see app/document-reader.tsx): a running header and a deeply indented
   uppercase footer so both are filtered out, an "NN / TITLE" line to carry the
   section title, blank lines between blocks, three or more spaces between table
   columns, and bullet or numbered markers for lists. */

import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const ROOT = path.resolve(import.meta.dirname, "..");
const DOCUMENTS_DIR = path.join(ROOT, "documents");
const PDF_OUT = path.join(ROOT, "public", "source-pdfs");
const TEXT_OUT = path.join(ROOT, "public", "source-text");

/* Classification marking carried on every page, top and bottom, as required for
   a controlled document. Change here and regenerate every document. */
const CLASSIFICATION = "UNCLASSIFIED";

const PAGE_WIDTH = 96; // characters of usable width in the text extraction
const FOOTER_INDENT = 44; // >= 40 so the reader treats the footer as furniture

const BRAND = {
  ink: "#070708",
  field: "#ffffff",
  graphite: "#51545a",
  mist: "#ecedef",
  line: "#d5d7dc",
  blue: "#2f66ff",
};

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* ---------- plain-text emitter ---------- */

function wrap(text, width) {
  const words = String(text).split(/\s+/).filter(Boolean);
  const lines = [];
  let line = "";
  for (const word of words) {
    if (!line.length) line = word;
    else if (`${line} ${word}`.length <= width) line += ` ${word}`;
    else { lines.push(line); line = word; }
  }
  if (line) lines.push(line);
  return lines.length ? lines : [""];
}

function textTable(rows) {
  const columns = Math.max(...rows.map((row) => row.length));
  const widths = Array.from({ length: columns }, (_, index) =>
    Math.max(...rows.map((row) => String(row[index] ?? "").length)),
  );
  /* Cap the widest column so a long cell cannot push the table past the page. */
  const budget = PAGE_WIDTH - 3 * (columns - 1);
  const total = widths.reduce((sum, width) => sum + width, 0);
  if (total > budget) {
    const widest = widths.indexOf(Math.max(...widths));
    widths[widest] = Math.max(12, widths[widest] - (total - budget));
  }

  const lines = [];
  for (const row of rows) {
    /* Wrap every cell, then emit the row across as many lines as the tallest
       cell needs, keeping column starts aligned so the parser reads columns. */
    const cells = row.map((cell, index) => wrap(cell ?? "", widths[index]));
    const height = Math.max(...cells.map((cell) => cell.length));
    for (let line = 0; line < height; line += 1) {
      const parts = cells.map((cell, index) =>
        (cell[line] ?? "").padEnd(widths[index], " "),
      );
      lines.push(parts.join("   ").replace(/\s+$/, ""));
    }
  }
  return lines;
}

function blockToText(block) {
  if (block.kind === "h4") return [block.text.toUpperCase()];
  /* A figure cannot go into a text extraction, so it contributes its caption and
     description instead — which is what makes it findable by search. */
  if (block.kind === "figure") {
    const lines = [`FIGURE  ${block.caption}`];
    if (block.meta) lines.push(block.meta);
    if (block.description) lines.push(...wrap(block.description, PAGE_WIDTH));
    if (block.text) lines.push(...block.text.filter((line) => line.trim()));
    return lines;
  }
  if (block.kind === "p") return wrap(block.text, PAGE_WIDTH);
  if (block.kind === "callout") {
    return [block.label.toUpperCase(), "", ...wrap(block.text, PAGE_WIDTH)];
  }
  if (block.kind === "ul") {
    return block.items.flatMap((item) => {
      const [first, ...rest] = wrap(item, PAGE_WIDTH - 2);
      return [`• ${first}`, ...rest.map((line) => `  ${line}`)];
    });
  }
  if (block.kind === "ol") {
    return block.items.flatMap((item, index) => {
      const marker = `${index + 1}. `;
      const [first, ...rest] = wrap(item, PAGE_WIDTH - marker.length);
      return [`${marker}${first}`, ...rest.map((line) => `${" ".repeat(marker.length)}${line}`)];
    });
  }
  if (block.kind === "table") return textTable(block.rows);
  throw new Error(`Unknown block kind: ${block.kind}`);
}

function coverToText(doc, pageNumber) {
  const { cover, meta } = doc;
  const lines = [
    cover.kicker.toUpperCase(),
    "",
    "",
    cover.title,
    ...wrap(cover.standfirst, PAGE_WIDTH),
    "",
    "",
  ];
  if (cover.callout) {
    lines.push(cover.callout.label.toUpperCase(), "", ...wrap(cover.callout.text, PAGE_WIDTH), "");
  }
  if (cover.spine?.length) {
    lines.push("", ...textTable(cover.spine));
  }
  return [centered(CLASSIFICATION), "", ...lines, "", footerLine(meta, pageNumber), "", centered(CLASSIFICATION)];
}

/* Centred and uppercase, so the reader's furniture filter drops it from the
   on-screen view the same way it drops running headers. */
function centered(text) {
  const pad = Math.max(44, Math.floor((PAGE_WIDTH - text.length) / 2));
  return `${" ".repeat(pad)}${text}`;
}

function footerLine(meta, pageNumber) {
  return `${" ".repeat(FOOTER_INDENT)}${meta.footer} | ${String(pageNumber).padStart(2, "0")}`.toUpperCase();
}

function runningHeader(meta) {
  /* "PULSE" plus 6+ spaces is filtered as a running header by the reader. */
  return `PULSE${" ".repeat(Math.max(6, PAGE_WIDTH - 5 - meta.runningHeader.length))}${meta.runningHeader}`;
}

function pageToText(doc, page, pageNumber) {
  const lines = [centered(CLASSIFICATION), "", runningHeader(doc.meta), "", "", page.title, ""];
  page.blocks.forEach((block, index) => {
    if (index) lines.push("");
    lines.push(...blockToText(block));
  });
  lines.push("", footerLine(doc.meta, pageNumber), "", centered(CLASSIFICATION));
  return lines;
}

function buildText(doc) {
  const pages = [coverToText(doc, 1).join("\n")];
  doc.pages.forEach((page, index) => {
    pages.push(pageToText(doc, page, index + 2).join("\n"));
  });
  return `${pages.join("\n\f")}\n`;
}

/* ---------- PDF renderer ---------- */

function blockToHtml(block) {
  if (block.kind === "h4") return `<h4>${escapeHtml(block.text)}</h4>`;
  if (block.kind === "figure") {
    const file = path.join(ROOT, "public", block.file.replace(/^\//, ""));
    if (!existsSync(file)) throw new Error(`Figure not found: ${block.file}`);
    const mime = block.file.endsWith(".svg") ? "image/svg+xml" : block.file.endsWith(".webp") ? "image/webp" : "image/png";
    const uri = `data:${mime};base64,${readFileSync(file).toString("base64")}`;
    return `<figure class="fig${block.file.endsWith(".svg") ? " fig--diagram" : ""}">
      <img src="${uri}" alt="${escapeHtml(block.caption)}" />
      <figcaption>${block.hideCaption ? "" : `<b>${escapeHtml(block.caption)}</b>`}${block.meta ? `<code>${escapeHtml(block.meta)}</code>` : ""}${block.description ? `<span>${escapeHtml(block.description)}</span>` : ""}</figcaption>
    </figure>`;
  }
  if (block.kind === "p") return `<p>${escapeHtml(block.text)}</p>`;
  if (block.kind === "callout") {
    return `<aside class="callout"><strong>${escapeHtml(block.label)}</strong><p>${escapeHtml(block.text)}</p></aside>`;
  }
  if (block.kind === "ul") {
    return `<ul>${block.items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
  }
  if (block.kind === "ol") {
    return `<ol>${block.items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ol>`;
  }
  if (block.kind === "table") {
    const [head, ...body] = block.rows;
    return `<table><thead><tr>${head.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr></thead><tbody>${body
      .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`)
      .join("")}</tbody></table>`;
  }
  throw new Error(`Unknown block kind: ${block.kind}`);
}

function buildHtml(doc) {
  const { meta, cover, pages } = doc;
  const coverHtml = `
    <section class="sheet cover">
      <div class="classification classification--top">${escapeHtml(CLASSIFICATION)}</div>
      <div class="cover-body">
        <div class="kicker">${escapeHtml(cover.kicker)}</div>
        <h1>${escapeHtml(cover.title)}</h1>
        <p class="standfirst">${escapeHtml(cover.standfirst)}</p>
        ${cover.callout ? `<aside class="callout callout--cover"><strong>${escapeHtml(cover.callout.label)}</strong><p>${escapeHtml(cover.callout.text)}</p></aside>` : ""}
      </div>
      ${cover.spine?.length ? `<dl class="spine">${cover.spine.map(([term, value]) => `<div><dt>${escapeHtml(term)}</dt><dd>${escapeHtml(value)}</dd></div>`).join("")}</dl>` : ""}
      <footer class="sheet-foot"><span>${escapeHtml(meta.footer)}</span><b>01</b></footer>
      <div class="classification classification--bottom">${escapeHtml(CLASSIFICATION)}</div>
    </section>`;

  const pageHtml = pages
    .map((page, index) => `
    <section class="sheet">
      <div class="classification classification--top">${escapeHtml(CLASSIFICATION)}</div>
      <header class="sheet-head"><span class="mark">PULSE</span><span>${escapeHtml(meta.runningHeader)}</span></header>
      <div class="sheet-body">
        <div class="kicker">${escapeHtml(page.kicker)}</div>
        <h2>${escapeHtml(page.title)}</h2>
        ${page.blocks.map(blockToHtml).join("\n")}
      </div>
      <footer class="sheet-foot"><span>${escapeHtml(meta.footer)}</span><b>${String(index + 2).padStart(2, "0")}</b></footer>
      <div class="classification classification--bottom">${escapeHtml(CLASSIFICATION)}</div>
    </section>`)
    .join("\n");

  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>${escapeHtml(meta.title)}</title><style>
  @page { size: Letter ${meta.orientation === "landscape" ? "landscape" : "portrait"}; margin: 0; }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body { font-family: Inter, -apple-system, "Helvetica Neue", Arial, sans-serif; color: ${BRAND.ink}; background: ${BRAND.field}; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .sheet { position: relative; width: 8.5in; height: 11in; padding: 0.56in 0.78in 0.5in; display: flex; flex-direction: column; break-after: page; page-break-after: always; overflow: hidden; }
  .sheet:last-child { break-after: auto; page-break-after: auto; }
  .sheet-head { display: flex; justify-content: space-between; align-items: baseline; padding-bottom: 9px; border-bottom: 1px solid ${BRAND.line}; font-size: 7.5pt; letter-spacing: 1.1px; text-transform: uppercase; color: ${BRAND.graphite}; }
  .sheet-head .mark { color: ${BRAND.ink}; font-weight: 800; letter-spacing: 3.4px; }
  .sheet-body { flex: 1; padding-top: 26px; }
  .sheet-foot { display: flex; justify-content: space-between; align-items: baseline; padding-top: 9px; border-top: 1px solid ${BRAND.line}; font-size: 7pt; letter-spacing: 1px; text-transform: uppercase; color: ${BRAND.graphite}; }
  .sheet-foot b { color: ${BRAND.ink}; font-size: 8pt; }
  .classification { position: absolute; left: 0; right: 0; text-align: center; font-size: 8.5pt; font-weight: 800; letter-spacing: 2.4px; color: ${BRAND.ink}; }
  .classification--top { top: 0.24in; }
  .classification--bottom { bottom: 0.2in; }
  .cover .classification { color: ${BRAND.field}; }
  .kicker { color: ${BRAND.blue}; font-size: 7.5pt; font-weight: 800; letter-spacing: 1.4px; text-transform: uppercase; }
  h1 { margin: 16px 0 14px; font-size: 34pt; line-height: 1.02; letter-spacing: -1.4px; }
  h2 { margin: 9px 0 18px; font-size: 17pt; line-height: 1.16; letter-spacing: -0.5px; }
  h4 { margin: 20px 0 8px; font-size: 8pt; font-weight: 800; letter-spacing: 1.2px; text-transform: uppercase; color: ${BRAND.ink}; }
  p { margin: 0 0 11px; font-size: 10pt; line-height: 1.62; color: ${BRAND.graphite}; }
  ul, ol { margin: 0 0 12px; padding-left: 17px; }
  li { margin-bottom: 6px; font-size: 10pt; line-height: 1.58; color: ${BRAND.graphite}; }
  table { width: 100%; border-collapse: collapse; margin: 4px 0 14px; }
  th, td { padding: 7px 9px; border-bottom: 1px solid ${BRAND.line}; text-align: left; font-size: 8.6pt; line-height: 1.45; vertical-align: top; }
  th { border-bottom: 1px solid ${BRAND.ink}; font-size: 7.2pt; font-weight: 800; letter-spacing: 1px; text-transform: uppercase; color: ${BRAND.ink}; }
  td { color: ${BRAND.graphite}; }
  td:first-child { color: ${BRAND.ink}; font-weight: 600; }
  .callout { margin: 14px 0 15px; padding: 15px 17px; background: ${BRAND.mist}; border-left: 2px solid ${BRAND.blue}; }
  .callout strong { display: block; margin-bottom: 6px; font-size: 7.2pt; font-weight: 800; letter-spacing: 1.2px; text-transform: uppercase; color: ${BRAND.blue}; }
  .callout p { margin: 0; font-size: 9.4pt; color: ${BRAND.ink}; }
  .fig { margin: 16px 0 18px; page-break-inside: avoid; break-inside: avoid; }
  .fig img { display: block; width: 100%; height: auto; border: 1px solid ${BRAND.line}; }
  .fig--diagram img { border: 0; }
  .fig figcaption { margin-top: 9px; display: grid; gap: 3px; }
  .fig figcaption b { font-size: 8.4pt; }
  .fig figcaption code { font: 7.6pt ui-monospace, Menlo, monospace; color: ${BRAND.graphite}; }
  .fig figcaption span { font-size: 8.4pt; line-height: 1.5; color: ${BRAND.graphite}; }
  .cover { background: ${BRAND.ink}; color: ${BRAND.field}; }
  .cover .cover-body { flex: 1; display: flex; flex-direction: column; justify-content: center; }
  .cover h1 { color: ${BRAND.field}; max-width: 6.2in; }
  .cover .kicker { color: #9bb6ff; }
  .cover .standfirst { max-width: 5.5in; font-size: 11pt; line-height: 1.6; color: #c8c8cc; }
  .callout--cover { margin-top: 30px; background: #18181a; border-left-color: #9bb6ff; }
  .callout--cover strong { color: #9bb6ff; }
  .callout--cover p { color: ${BRAND.field}; }
  .spine { display: grid; grid-template-columns: repeat(4, 1fr); gap: 0; margin: 0 0 26px; border-top: 1px solid #303035; }
  .spine > div { padding: 15px 14px 0 0; }
  .spine dt { font-size: 7pt; font-weight: 800; letter-spacing: 1.2px; text-transform: uppercase; color: #85868b; }
  .spine dd { margin: 6px 0 0; font-size: 9pt; color: ${BRAND.field}; }
  .cover .sheet-foot { border-top-color: #303035; color: #85868b; }
  .cover .sheet-foot b { color: ${BRAND.field}; }
</style></head><body>${coverHtml}${pageHtml}</body></html>`;
}

/* ---------- driver ---------- */

/* ESM imports ignore NODE_PATH, so when playwright lives outside this package
   accept an explicit directory via PLAYWRIGHT_NODE_PATH. */
async function loadPlaywright() {
  try {
    return await import("playwright");
  } catch (error) {
    const external = process.env.PLAYWRIGHT_NODE_PATH;
    if (!external) {
      throw new Error(
        "playwright is required to render the PDF. Install it in this package, or set PLAYWRIGHT_NODE_PATH to a directory containing node_modules/playwright.",
      );
    }
    const entry = path.join(external, "playwright", "index.js");
    const loaded = await import(pathToFileURL(entry).href);
    /* playwright is CommonJS; named exports are not always detected through an
       absolute ESM import, so fall back to the default wrapper. */
    return loaded.chromium ? loaded : loaded.default;
  }
}

function resolveDocument(selector) {
  const files = readdirSync(DOCUMENTS_DIR).filter((file) => file.endsWith(".mjs"));
  const match = files.find((file) => file.startsWith(`${selector}-`) || file === selector || file === `${selector}.mjs`);
  if (!match) {
    throw new Error(`No document source matching "${selector}" in ${DOCUMENTS_DIR}. Available: ${files.join(", ")}`);
  }
  return path.join(DOCUMENTS_DIR, match);
}

async function main() {
  const selector = process.argv[2];
  if (!selector) throw new Error("Pass a document selector, for example: 09");

  const sourcePath = resolveDocument(selector);
  const doc = await import(pathToFileURL(sourcePath).href);
  if (!doc.meta?.slug) throw new Error(`${sourcePath} does not export meta.slug`);

  mkdirSync(TEXT_OUT, { recursive: true });
  mkdirSync(PDF_OUT, { recursive: true });

  const textPath = path.join(TEXT_OUT, `${doc.meta.slug}.txt`);
  writeFileSync(textPath, buildText(doc), "utf8");

  const html = buildHtml(doc);
  /* Set DOCUMENT_HTML_OUT to keep the intermediate HTML for layout inspection. */
  if (process.env.DOCUMENT_HTML_OUT) {
    writeFileSync(process.env.DOCUMENT_HTML_OUT, html, "utf8");
    console.log(`  html  ${process.env.DOCUMENT_HTML_OUT}`);
  }
  const { chromium } = await loadPlaywright();
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: "load" });
  const pdfPath = path.join(PDF_OUT, `${doc.meta.slug}.pdf`);
  await page.pdf({ path: pdfPath, format: "Letter", landscape: doc.meta.orientation === "landscape", printBackground: true, preferCSSPageSize: true });
  await browser.close();

  /* The web reader renders from the text extraction, which cannot carry an
     image. Emit a manifest keyed by section so the reader can place the same
     figures the PDF shows, in the same order, without a second source of truth. */
  const figures = {};
  doc.pages.forEach((page, index) => {
    const pageFigures = page.blocks.filter((block) => block.kind === "figure");
    if (pageFigures.length) {
      figures[index + 2] = pageFigures.map(({ file, caption, meta, description, hideCaption }) => ({ file, caption, meta, description, hideCaption: Boolean(hideCaption) }));
    }
  });
  const manifestPath = path.join(ROOT, "app", "generated", "figures", `${doc.meta.slug}.json`);
  mkdirSync(path.dirname(manifestPath), { recursive: true });
  writeFileSync(manifestPath, `${JSON.stringify(figures, null, 2)}\n`, "utf8");

  const pageCount = doc.pages.length + 1;
  console.log(`${doc.meta.slug}`);
  console.log(`  text  ${path.relative(ROOT, textPath)}  (${pageCount} pages)`);
  console.log(`  pdf   ${path.relative(ROOT, pdfPath)}`);
  console.log(`  content pages (excluding cover): ${doc.pages.length}`);
}

await main();
