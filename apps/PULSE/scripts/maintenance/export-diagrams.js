#!/usr/bin/env node
/**
 * PULSE Diagram PNG Exporter
 * Renders SVG files to 4K PNG (3840×2160) using Puppeteer or system tools.
 * Usage: node export-diagrams.js
 */

const { execSync, spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const WORKSPACE = path.resolve(__dirname);
const DIAGRAMS = [
  {
    svg: path.join(WORKSPACE, 'pulse-executive-architecture-final.svg'),
    png: path.join(WORKSPACE, 'pulse-executive-architecture-final.png'),
    label: 'Executive Architecture Final'
  },
  {
    svg: path.join(WORKSPACE, 'pulse-data-relationship-graph-final.svg'),
    png: path.join(WORKSPACE, 'pulse-data-relationship-graph-final.png'),
    label: 'Data Relationship Graph Final'
  }
];

const W = 3840, H = 2160;

// ── Strategy 1: Puppeteer (preferred) ──────────────────────────────────────
async function exportWithPuppeteer() {
  let puppeteer;
  try {
    puppeteer = require('puppeteer');
  } catch (e) {
    return false;
  }

  console.log('Using Puppeteer...');
  const browser = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-web-security']
  });
  const page = await browser.newPage();
  await page.setViewport({ width: W, height: H, deviceScaleFactor: 1 });

  for (const { svg, png, label } of DIAGRAMS) {
    const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html, body { width: ${W}px; height: ${H}px; overflow: hidden; background: transparent; }
  svg { width: ${W}px; height: ${H}px; display: block; }
</style>
</head>
<body>${fs.readFileSync(svg, 'utf8')}</body>
</html>`;

    const tmpHtml = svg.replace('.svg', '_tmp_export.html');
    fs.writeFileSync(tmpHtml, html);
    await page.goto(`file://${tmpHtml}`, { waitUntil: 'networkidle0' });
    await page.screenshot({
      path: png,
      type: 'png',
      clip: { x: 0, y: 0, width: W, height: H },
      omitBackground: false
    });
    fs.unlinkSync(tmpHtml);
    console.log(`  ✓  ${label} → ${path.basename(png)}`);
  }

  await browser.close();
  return true;
}

// ── Strategy 2: rsvg-convert (librsvg) ─────────────────────────────────────
function exportWithRsvg() {
  try {
    execSync('which rsvg-convert', { stdio: 'pipe' });
  } catch {
    return false;
  }
  console.log('Using rsvg-convert...');
  for (const { svg, png, label } of DIAGRAMS) {
    const result = spawnSync('rsvg-convert', [
      '-w', String(W), '-h', String(H),
      '-o', png, svg
    ], { stdio: 'inherit' });
    if (result.status !== 0) {
      console.warn(`  ✗  ${label}: rsvg-convert failed`);
      return false;
    }
    console.log(`  ✓  ${label} → ${path.basename(png)}`);
  }
  return true;
}

// ── Strategy 3: inkscape ────────────────────────────────────────────────────
function exportWithInkscape() {
  try {
    execSync('which inkscape', { stdio: 'pipe' });
  } catch {
    return false;
  }
  console.log('Using Inkscape...');
  for (const { svg, png, label } of DIAGRAMS) {
    const result = spawnSync('inkscape', [
      `--export-filename=${png}`,
      `--export-width=${W}`,
      `--export-height=${H}`,
      svg
    ], { stdio: 'inherit' });
    if (result.status !== 0) {
      console.warn(`  ✗  ${label}: inkscape failed`);
      return false;
    }
    console.log(`  ✓  ${label} → ${path.basename(png)}`);
  }
  return true;
}

// ── Strategy 4: Chrome / Chromium headless ──────────────────────────────────
function exportWithChrome() {
  let chromeBin = null;
  const candidates = [
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    'google-chrome',
    'chromium-browser',
    'chromium'
  ];
  for (const c of candidates) {
    try {
      if (c.startsWith('/')) {
        if (fs.existsSync(c)) { chromeBin = c; break; }
      } else {
        execSync(`which ${c}`, { stdio: 'pipe' });
        chromeBin = c;
        break;
      }
    } catch {}
  }
  if (!chromeBin) return false;

  console.log(`Using Chrome: ${chromeBin}`);
  for (const { svg, png, label } of DIAGRAMS) {
    const html = `<!DOCTYPE html>
<html>
<head><style>
* { margin:0; padding:0; }
html,body { width:${W}px; height:${H}px; overflow:hidden; }
svg { width:${W}px; height:${H}px; }
</style></head>
<body>${fs.readFileSync(svg, 'utf8')}</body>
</html>`;
    const tmpHtml = svg.replace('.svg', '_tmp.html');
    fs.writeFileSync(tmpHtml, html);

    const result = spawnSync(chromeBin, [
      '--headless=new',
      '--no-sandbox',
      '--disable-gpu',
      '--disable-software-rasterizer',
      `--window-size=${W},${H}`,
      `--screenshot=${png}`,
      `--force-device-scale-factor=1`,
      `file://${tmpHtml}`
    ], { stdio: 'inherit' });

    try { fs.unlinkSync(tmpHtml); } catch {}
    if (result.status !== 0) {
      console.warn(`  ✗  ${label}: Chrome failed`);
      return false;
    }
    console.log(`  ✓  ${label} → ${path.basename(png)}`);
  }
  return true;
}

// ── Main ────────────────────────────────────────────────────────────────────
(async () => {
  console.log('\n🎨  PULSE Diagram PNG Exporter\n');
  console.log(`Output directory: ${WORKSPACE}\n`);

  // Check SVGs exist
  for (const { svg } of DIAGRAMS) {
    if (!fs.existsSync(svg)) {
      console.error(`ERROR: SVG not found: ${svg}`);
      process.exit(1);
    }
  }

  let success = false;

  // Try each strategy in order
  if (!success) success = await exportWithPuppeteer();
  if (!success) success = exportWithRsvg();
  if (!success) success = exportWithInkscape();
  if (!success) success = exportWithChrome();

  if (success) {
    console.log('\n✅  All PNGs exported successfully!\n');
    for (const { png } of DIAGRAMS) {
      if (fs.existsSync(png)) {
        const stat = fs.statSync(png);
        const mb = (stat.size / 1024 / 1024).toFixed(2);
        console.log(`   📦  ${path.basename(png)}  (${mb} MB)`);
      }
    }
    console.log('\nFiles ready for PowerPoint insertion.\n');
  } else {
    console.error('\n❌  No suitable renderer found.\n');
    console.error('Please install one of:');
    console.error('  • puppeteer: npm install puppeteer');
    console.error('  • rsvg-convert: brew install librsvg');
    console.error('  • inkscape: brew install inkscape');
    console.error('  • Google Chrome (already installed on most Macs)\n');
    process.exit(1);
  }
})();
