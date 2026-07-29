const { execSync } = require("child_process");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const APPS = path.resolve(ROOT, "..");
const RELEASES = path.resolve(ROOT, "..", "..", "releases");
const FS_PACKAGES = path.join(ROOT, "FS packages");
const buildScript = path.join(ROOT, "scripts", "build-sharepoint-package.js");

// PULSE-native packages: built with the PULSE SharePoint package builder
const pulsePackages = [
  {
    name: "Travel Request Forms",
    entry: "travel-forms.html",
    fsOutput: "Travel-Request-Forms-v1.0.0.html",
    releaseOutput: "PULSE-Travel-Request-Forms-v1.0.0.html"
  },
  {
    name: "My Travel",
    entry: "my-travel.html",
    fsOutput: "Travel-MyTravel-v1.0.0.html",
    releaseOutput: "PULSE-Travel-MyTravel-v1.0.0.html"
  },
  {
    name: "Travel Calendar",
    entry: "travel-calendar.html",
    fsOutput: "Travel-Calendar-v1.0.0.html",
    releaseOutput: "PULSE-Travel-Calendar-v1.0.0.html"
  },
  {
    name: "Tickets",
    entry: "tickets.html",
    fsOutput: "Tickets-v1.0.0.html",
    releaseOutput: "PULSE-Tickets-v1.0.0.html"
  },
  {
    name: "Main PULSE Application",
    entry: "index.html",
    fsOutput: "AEWTTR-PULSE_v.20260728.html",
    releaseOutput: "PULSE-v1.0.0.html"
  }
];

// External app packages: each has its own build script, accepts output path as argv[2]
const externalPackages = [
  {
    name: "PULSE Calendar",
    script: path.join(APPS, "PULSE-TRAVEL-CALENDAR", "scripts", "build-sharepoint-package.js"),
    fsOutput: "PULSE-Calendar-v1.0.0.html",
    releaseOutput: "PULSE-Calendar-v1.0.0.html",
    useNode: "node"
  },
  {
    name: "PULSE CODE",
    script: path.join(APPS, "PULSE-CODE", "scripts", "build-sharepoint-package.js"),
    fsOutput: "PULSE-CODE-v1.0.0.html",
    releaseOutput: "PULSE-CODE-v1.0.0.html",
    useNode: "node"
  },
  {
    name: "PULSE Documentation",
    script: path.join(ROOT, "..", "..", "output", "pulse-documentation", "scripts", "build-sharepoint-package.mjs"),
    fsOutput: "PULSE-Documentation-v1.0.0.html",
    releaseOutput: "PULSE-Documentation-v1.0.0.html",
    useNode: "node"
  }
];

console.log("Building PULSE FS & Release Packages...\n");

for (const pkg of pulsePackages) {
  const fsPath = path.join(FS_PACKAGES, pkg.fsOutput);
  const relPath = path.join(RELEASES, pkg.releaseOutput);
  const entryPath = path.join(ROOT, pkg.entry);
  console.log(`Building [${pkg.name}]...`);
  console.log(`  Source: ${pkg.entry}`);
  try {
    execSync(`node "${buildScript}" "${fsPath}" "${entryPath}"`, { stdio: "inherit" });
    console.log(`  FS Output: FS packages/${pkg.fsOutput}`);
    execSync(`node "${buildScript}" "${relPath}" "${entryPath}"`, { stdio: "inherit" });
    console.log(`  Release Output: releases/${pkg.releaseOutput}`);
    console.log(`  ✓ Success\n`);
  } catch (err) {
    console.error(`  ✗ Build failed for ${pkg.name}:`, err.message);
    process.exit(1);
  }
}

for (const pkg of externalPackages) {
  const fsPath = path.join(FS_PACKAGES, pkg.fsOutput);
  const relPath = path.join(RELEASES, pkg.releaseOutput);
  console.log(`Building [${pkg.name}]...`);
  try {
    execSync(`${pkg.useNode} "${pkg.script}" "${fsPath}"`, { stdio: "inherit" });
    console.log(`  FS Output: FS packages/${pkg.fsOutput}`);
    execSync(`${pkg.useNode} "${pkg.script}" "${relPath}"`, { stdio: "inherit" });
    console.log(`  Release Output: releases/${pkg.releaseOutput}`);
    console.log(`  ✓ Success\n`);
  } catch (err) {
    console.error(`  ✗ Build failed for ${pkg.name}:`, err.message);
    process.exit(1);
  }
}

console.log("All PULSE FS & Release Packages built successfully!");

