const { execSync } = require("child_process");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const buildScript = path.join(ROOT, "scripts", "build-sharepoint-package.js");

const packages = [
  {
    name: "Travel Request Forms",
    entry: "travel-forms.html",
    output: "FS packages/Travel-Request-Forms-v1.0.0.html",
    releaseOutput: "../../releases/PULSE-Travel-Request-Forms-v1.0.0.html"
  },
  {
    name: "My Travel",
    entry: "my-travel.html",
    output: "FS packages/Travel-MyTravel-v1.0.0.html",
    releaseOutput: "../../releases/PULSE-Travel-MyTravel-v1.0.0.html"
  },
  {
    name: "Travel Calendar",
    entry: "travel-calendar.html",
    output: "FS packages/Travel-Calendar-v1.0.0.html",
    releaseOutput: "../../releases/PULSE-Travel-Calendar-v1.0.0.html"
  },
  {
    name: "Main PULSE Application",
    entry: "index.html",
    output: "FS packages/AEWTTR-PULSE_v.20260727.html",
    releaseOutput: "../../releases/PULSE-v1.0.0.html"
  }
];

console.log("Building PULSE FS & Release Packages...\n");

for (const pkg of packages) {
  const outPath = path.join(ROOT, pkg.output);
  const entryPath = path.join(ROOT, pkg.entry);
  console.log(`Building [${pkg.name}]...`);
  console.log(`  Source: ${pkg.entry}`);
  console.log(`  FS Output: ${pkg.output}`);
  try {
    execSync(`node "${buildScript}" "${outPath}" "${entryPath}"`, { stdio: "inherit" });
    if (pkg.releaseOutput) {
      const relPath = path.join(ROOT, pkg.releaseOutput);
      console.log(`  Release Output: ${pkg.releaseOutput}`);
      execSync(`node "${buildScript}" "${relPath}" "${entryPath}"`, { stdio: "inherit" });
    }
    console.log(`  ✓ Success\n`);
  } catch (err) {
    console.error(`  ✗ Build failed for ${pkg.name}:`, err.message);
    process.exit(1);
  }
}

console.log("All PULSE FS & Release Packages built successfully!");

