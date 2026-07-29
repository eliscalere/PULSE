/* Writes each process flow to public/figures/flow-<id>.svg. Both the generated
   PDFs and the web reader display these files, so the diagrams cannot drift. */
import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { flows, renderFlowSvg } from "../documents/flows.mjs";

const OUT = path.join(import.meta.dirname, "..", "public", "figures");
mkdirSync(OUT, { recursive: true });
for (const flow of flows) {
  const file = path.join(OUT, `flow-${flow.id}.svg`);
  writeFileSync(file, renderFlowSvg(flow), "utf8");
  console.log(`  flow-${flow.id}.svg`);
}
console.log(`${flows.length} figures written`);
