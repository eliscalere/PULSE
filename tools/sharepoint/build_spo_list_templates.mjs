import fs from "node:fs/promises";
import path from "node:path";
import vm from "node:vm";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const workspaceRoot = path.resolve(import.meta.dirname, "..", "..");
const schemaPath = path.join(workspaceRoot, "apps/PULSE/assets/js/sharepoint-schema.js");
const outputDir = path.join(workspaceRoot, "artifacts/generated/spo-list-templates-20260630");

function csvEscape(value) {
  const text = value == null ? "" : String(value);
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function parseSchemaObject(sourceText) {
  const marker = "const SHAREPOINT_SCHEMA =";
  const start = sourceText.indexOf(marker);
  if (start === -1) {
    throw new Error("Could not find SHAREPOINT_SCHEMA in sharepoint-schema.js");
  }

  const braceStart = sourceText.indexOf("{", start);
  let depth = 0;
  let end = -1;
  for (let i = braceStart; i < sourceText.length; i += 1) {
    const char = sourceText[i];
    if (char === "{") {
      depth += 1;
    } else if (char === "}") {
      depth -= 1;
      if (depth === 0) {
        end = i;
        break;
      }
    }
  }

  if (end === -1) {
    throw new Error("Could not parse SHAREPOINT_SCHEMA object block");
  }

  const objectLiteral = sourceText.slice(braceStart, end + 1);
  return vm.runInNewContext("(" + objectLiteral + ")");
}

function sampleValue(field, listTitle, rowIndex) {
  if (!field) {
    return `Sample ${listTitle} ${rowIndex}`;
  }

  switch (field.type) {
    case "Number":
      return rowIndex;
    case "Boolean":
      return rowIndex % 2 === 1 ? "TRUE" : "FALSE";
    case "DateTime":
      return rowIndex === 1 ? "2026-07-01" : "2026-07-15";
    case "Choice":
      return field.choices && field.choices.length ? field.choices[0] : "Choice";
    case "Note":
      return `${field.name} sample text`;
    case "Text":
    default:
      if (field.name === "ProjectCode") return rowIndex === 1 ? "P-1001" : "P-1002";
      if (field.name === "RequestCode") return rowIndex === 1 ? "REQ-1001" : "REQ-1002";
      if (field.name === "UserEmail") return rowIndex === 1 ? "member1@example.invalid" : "member2@example.invalid";
      if (field.name === "UserDisplayName") return rowIndex === 1 ? "Member One" : "Member Two";
      if (field.name === "LoginName") return rowIndex === 1 ? "member1" : "member2";
      if (field.name === "Title") return `${listTitle} Sample ${rowIndex}`;
      return `${field.name} sample ${rowIndex}`;
  }
}

function buildRows(listTitle, definition) {
  const headers = ["Title", ...(definition.fields || []).map((field) => field.name)];
  const firstRow = headers.map((header, index) => {
    if (index === 0) return `${listTitle} Sample 1`;
    return sampleValue(definition.fields[index - 1], listTitle, 1);
  });
  const secondRow = headers.map((header, index) => {
    if (index === 0) return `${listTitle} Sample 2`;
    return sampleValue(definition.fields[index - 1], listTitle, 2);
  });

  return { headers, rows: [firstRow, secondRow] };
}

function cleanSheetName(name) {
  return name.replace(/[:\\/?*\[\]]/g, " ").slice(0, 31);
}

function createInstructionsRows(schema) {
  return [
    ["SharePoint List Import Templates", "", "", ""],
    ["How to use", "1. Open the matching CSV or workbook tab.", "", ""],
    ["", "2. Use it as a template for Create list from Excel or manual list setup.", "", ""],
    ["", "3. Delete the sample rows after import if you do not want them.", "", ""],
    ["", "4. Keep the built-in SharePoint Title column.", "", ""],
    ["List", "Description", "Suggested Import File", "Notes"]
  ].concat(Object.entries(schema).map(([listTitle, definition]) => {
    return [
      listTitle,
      definition.description || "",
      `${listTitle.replace(/[^A-Za-z0-9]+/g, "_").toLowerCase()}.csv`,
      "Headers include Title plus custom fields."
    ];
  }));
}

async function buildWorkbook(schema) {
  const workbook = Workbook.create();
  const instructions = workbook.worksheets.add("Instructions");
  const instructionRows = createInstructionsRows(schema);
  instructions.getRangeByIndexes(0, 0, instructionRows.length, 4).values = instructionRows;
  instructions.getRange("A1:D1").merge();
  instructions.getRange("A1").format.font = { bold: true, size: 16, color: "#143757" };
  instructions.getRange("A6:D6").format.font = { bold: true, color: "#FFFFFF" };
  instructions.getRange("A6:D6").format.fill = { color: "#005A9C" };
  instructions.getRange(`A6:D${instructionRows.length}`).format.borders = { preset: "all", style: "thin", color: "#D9E1EA" };
  instructions.getRange(`A1:D${instructionRows.length}`).format.wrapText = true;
  instructions.freezePanes.freezeRows(6);
  instructions.showGridLines = false;
  instructions.getUsedRange().format.autofitColumns();

  for (const [listTitle, definition] of Object.entries(schema)) {
    const sheet = workbook.worksheets.add(cleanSheetName(listTitle));
    const { headers, rows } = buildRows(listTitle, definition);
    const fieldTypeRow = ["SharePoint Title", ...(definition.fields || []).map((field) => {
      const extras = field.type === "Choice" && field.choices ? ` (${field.choices.join(" | ")})` : "";
      const required = field.required ? " required" : "";
      return `${field.type}${required}${extras}`;
    })];

    sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
    sheet.getRangeByIndexes(1, 0, 1, headers.length).values = [fieldTypeRow];
    sheet.getRangeByIndexes(2, 0, rows.length, headers.length).values = rows;

    sheet.getRangeByIndexes(0, 0, 1, headers.length).format.font = { bold: true, color: "#FFFFFF" };
    sheet.getRangeByIndexes(0, 0, 1, headers.length).format.fill = { color: "#005A9C" };
    sheet.getRangeByIndexes(1, 0, 1, headers.length).format.font = { italic: true, color: "#425466" };
    sheet.getRangeByIndexes(1, 0, 1, headers.length).format.fill = { color: "#EAF2FB" };
    sheet.getRangeByIndexes(0, 0, rows.length + 2, headers.length).format.borders = { preset: "all", style: "thin", color: "#D9E1EA" };
    sheet.getRangeByIndexes(0, 0, rows.length + 2, headers.length).format.wrapText = true;
    sheet.freezePanes.freezeRows(2);
    sheet.showGridLines = false;
    sheet.getUsedRange().format.autofitColumns();
  }

  return workbook;
}

async function writeCsvFiles(schema) {
  await fs.mkdir(outputDir, { recursive: true });

  for (const [listTitle, definition] of Object.entries(schema)) {
    const { headers, rows } = buildRows(listTitle, definition);
    const lines = [headers, ...rows].map((row) => row.map(csvEscape).join(","));
    const fileName = `${listTitle.replace(/[^A-Za-z0-9]+/g, "_").toLowerCase()}.csv`;
    await fs.writeFile(path.join(outputDir, fileName), lines.join("\n") + "\n", "utf8");
  }
}

async function main() {
  await fs.mkdir(outputDir, { recursive: true });
  const sourceText = await fs.readFile(schemaPath, "utf8");
  const schema = parseSchemaObject(sourceText);

  await writeCsvFiles(schema);
  const workbook = await buildWorkbook(schema);

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(path.join(outputDir, "PULSE_SharePoint_List_Templates.xlsx"));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
