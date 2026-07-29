import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workspaceRoot = path.resolve(import.meta.dirname, "..", "..");
const outputDir = path.join(workspaceRoot, "artifacts/generated/spo-list-templates-20260630");
const workbookPath = path.join(outputDir, "PULSE_SharePoint_List_Templates.xlsx");
const renderDir = path.join(outputDir, "renders");

async function main() {
  await fs.mkdir(renderDir, { recursive: true });

  const fileBlob = await FileBlob.load(workbookPath);
  const workbook = await SpreadsheetFile.importXlsx(fileBlob);

  const summary = await workbook.inspect({
    kind: "sheet,table",
    maxChars: 8000,
    tableMaxRows: 6,
    tableMaxCols: 8
  });
  console.log(summary.ndjson);

  const sheetNames = workbook.worksheets.items.map((sheet) => sheet.name);
  for (const sheetName of sheetNames) {
    const png = await workbook.render({ sheetName, autoCrop: "all", scale: 1.2, format: "png" });
    const bytes = new Uint8Array(await png.arrayBuffer());
    const safeName = sheetName.replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
    await fs.writeFile(path.join(renderDir, `${safeName}.png`), bytes);
  }

  const csvFiles = (await fs.readdir(outputDir))
    .filter((name) => name.endsWith(".csv"))
    .sort();
  console.log(`CSV files: ${csvFiles.length}`);
  csvFiles.forEach((name) => console.log(name));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
