/* Writes each process flow to public/figures/flow-<id>.svg. Both the generated
   PDFs and the web reader display these files, so the diagrams cannot drift. */
import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { flows, renderFlowSvg } from "../documents/flows.mjs";
import { architectureFigures, scanSummary } from "../documents/architecture.mjs";

const OUT = path.join(import.meta.dirname, "..", "public", "figures");
mkdirSync(OUT, { recursive: true });
for (const flow of flows) {
  const file = path.join(OUT, `flow-${flow.id}.svg`);
  writeFileSync(file, renderFlowSvg(flow), "utf8");
  console.log(`  flow-${flow.id}.svg`);
}
for (const figure of architectureFigures) {
  const file = path.join(OUT, `${figure.id}.svg`);
  writeFileSync(file, figure.render(), "utf8");
  console.log(`  ${figure.id}.svg`);
}
console.log(`${flows.length + architectureFigures.length} figures written`);
console.log(`  scan: ${scanSummary.nodes} nodes, ${scanSummary.edges} edges, groups: ${scanSummary.groups.join(", ")}`);
