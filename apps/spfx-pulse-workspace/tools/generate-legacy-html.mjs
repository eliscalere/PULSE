import fs from 'node:fs';
import path from 'node:path';

const workspaceRoot = process.cwd();
const legacyRoot = path.resolve(workspaceRoot, '../PULSE');
const outputFile = path.resolve(workspaceRoot, 'src/legacy/LegacyWorkspaceHtml.ts');

const cssFiles = [
  'vendor/bootstrap/css/bootstrap.min.css',
  'vendor/boxicons/css/boxicons.min.css',
  'vendor/fullcalendar/css/main.min.css',
  'vendor/flatpickr/flatpickr.min.css',
  'assets/css/style.css'
];

const jsFiles = [
  'vendor/fullcalendar/js/main.min.js',
  'vendor/flatpickr/flatpickr.min.js',
  'vendor/pptxgenjs/pptxgen.browser.js',
  'assets/js/app-config.js',
  'assets/js/sharepoint-adapter.js',
  'assets/js/sharepoint-schema.js',
  'assets/js/sharepoint-repo.js',
  'assets/js/data.js',
  'assets/js/export.js',
  'assets/js/office-desktop.js',
  'assets/js/project-pptx-export.js',
  'assets/js/app.js',
  'assets/js/audit-log.js',
  'assets/js/notify.js',
  'assets/js/pages/dashboard.js',
  'assets/js/pages/overview.js',
  'assets/js/pages/people.js',
  'assets/js/pages/projects.js',
  'assets/js/pages/project-documents.js',
  'assets/js/pages/project-photos.js',
  'assets/js/pages/weekly.js',
  'assets/js/pages/travel.js',
  'assets/js/pages/docreview.js',
  'assets/js/pages/checklists.js',
  'assets/js/pages/tickets.js',
  'assets/js/pages/notification-settings.js',
  'assets/js/pages/users.js',
  'assets/js/pages/workload.js',
  'assets/js/pages/logs.js',
  'assets/js/sample-project-seed.js',
  'assets/js/pages/admin.js'
];

function mimeTypeFor(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  switch (extension) {
    case '.css':
      return 'text/css';
    case '.js':
      return 'text/javascript';
    case '.png':
      return 'image/png';
    case '.woff':
      return 'font/woff';
    case '.woff2':
      return 'font/woff2';
    case '.ttf':
      return 'font/ttf';
    case '.eot':
      return 'application/vnd.ms-fontobject';
    case '.svg':
      return 'image/svg+xml';
    default:
      return 'application/octet-stream';
  }
}

function toDataUri(filePath) {
  const data = fs.readFileSync(filePath);
  return `data:${mimeTypeFor(filePath)};base64,${data.toString('base64')}`;
}

function inlineCss(filePath) {
  const absolutePath = path.resolve(legacyRoot, filePath);
  let css = fs.readFileSync(absolutePath, 'utf8');

  css = css.replace(/url\((['"]?)([^'")]+)\1\)/g, (_, quote, assetPath) => {
    if (/^(data:|https?:|#)/.test(assetPath)) {
      return `url(${quote}${assetPath}${quote})`;
    }

    const sanitizedAssetPath = assetPath.split('#')[0].split('?')[0];
    const resolvedPath = path.resolve(path.dirname(absolutePath), sanitizedAssetPath);

    if (!fs.existsSync(resolvedPath)) {
      console.warn(`Skipping missing asset ${path.relative(legacyRoot, resolvedPath)}`);
      return `url(${quote}${assetPath}${quote})`;
    }

    return `url(${quote}${toDataUri(resolvedPath)}${quote})`;
  });

  return css;
}

function inlineJs(filePath) {
  const absolutePath = path.resolve(legacyRoot, filePath);
  let js = fs.readFileSync(absolutePath, 'utf8');
  const sealPath = path.resolve(legacyRoot, 'assets/images/aewttr-seal.png');
  js = js.replaceAll('assets/images/aewttr-seal.png', toDataUri(sealPath));
  return js;
}

function buildHtml() {
  const styles = cssFiles
    .map((filePath) => `<style data-source="${filePath}">\n${inlineCss(filePath)}\n</style>`)
    .join('\n');

  const scripts = jsFiles
    .map((filePath) => `<script data-source="${filePath}">\n${inlineJs(filePath)}\n</script>`)
    .join('\n');

  const resizeScript = `
    <script>
      (function () {
        function reportHeight() {
          var height = Math.max(
            document.documentElement.scrollHeight,
            document.body ? document.body.scrollHeight : 0,
            document.documentElement.offsetHeight
          );
          parent.postMessage({ source: 'aewttr-legacy-frame', height: height }, '*');
        }

        window.addEventListener('load', reportHeight);
        window.addEventListener('resize', reportHeight);
        window.addEventListener('hashchange', function () {
          setTimeout(reportHeight, 50);
        });

        var observer = new MutationObserver(function () {
          reportHeight();
        });

        observer.observe(document.documentElement, {
          childList: true,
          subtree: true,
          attributes: true
        });

        setInterval(reportHeight, 1000);
      })();
    </script>
  `;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AEWTTR-PULSE</title>
${styles}
</head>
<body>
<div id="aewttr-root"></div>
${scripts}
${resizeScript}
</body>
</html>`;
}

const html = buildHtml();
const ts = `export const legacyWorkspaceHtml: string = ${JSON.stringify(html)};\n`;

fs.mkdirSync(path.dirname(outputFile), { recursive: true });
fs.writeFileSync(outputFile, ts);

console.log(`Generated ${path.relative(workspaceRoot, outputFile)}`);
