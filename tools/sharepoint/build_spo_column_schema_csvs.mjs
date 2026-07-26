import fs from "node:fs/promises";
import path from "node:path";
import vm from "node:vm";

const workspaceRoot = "/Users/eliscalere/Downloads/AEWTTR PAS";
const schemaPath = path.join(workspaceRoot, "apps/PULSE/assets/js/sharepoint-schema.js");
const outputDir = path.join(workspaceRoot, "artifacts/generated/spo-column-schemas-20260630");

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
    if (char === "{") depth += 1;
    if (char === "}") {
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

  return vm.runInNewContext("(" + sourceText.slice(braceStart, end + 1) + ")");
}

function toFileName(listTitle) {
  return listTitle.replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "").toLowerCase() + "_schema.csv";
}

function buildRows(listTitle, definition) {
  const rows = [
    ["ListName", listTitle],
    ["Description", definition.description || ""],
    [],
    ["ColumnName", "ColumnType", "Required", "Choices", "Notes"]
  ];

  rows.push(["Title", "Built-in SharePoint Title", "Yes", "", "Do not recreate this column."]);

  for (const field of definition.fields || []) {
    rows.push([
      field.name,
      field.type || "Text",
      field.required ? "Yes" : "No",
      field.choices ? field.choices.join(" | ") : "",
      field.numLines ? `Suggested lines: ${field.numLines}` : ""
    ]);
  }

  return rows;
}

async function main() {
  await fs.mkdir(outputDir, { recursive: true });
  const sourceText = await fs.readFile(schemaPath, "utf8");
  const schema = parseSchemaObject(sourceText);

  const indexRows = [["ListName", "SchemaFile", "Description"]];

  for (const [listTitle, definition] of Object.entries(schema)) {
    const fileName = toFileName(listTitle);
    const rows = buildRows(listTitle, definition);
    const csvText = rows.map((row) => row.map(csvEscape).join(",")).join("\n") + "\n";
    await fs.writeFile(path.join(outputDir, fileName), csvText, "utf8");
    indexRows.push([listTitle, fileName, definition.description || ""]);
  }

  const indexCsv = indexRows.map((row) => row.map(csvEscape).join(",")).join("\n") + "\n";
  await fs.writeFile(path.join(outputDir, "spo_schema_file_index.csv"), indexCsv, "utf8");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
