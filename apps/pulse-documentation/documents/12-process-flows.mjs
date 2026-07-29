/* Source content for document 12, "Process Flows".

   The diagrams come from documents/flows.mjs by way of
   scripts/build-figures.mjs, which writes public/figures/flow-<id>.svg. The PDF,
   the searchable text, and the web reader all display that same file, so no
   diagram is drawn twice and none can drift.

     node apps/pulse-documentation/scripts/build-figures.mjs
     node apps/pulse-documentation/scripts/build-source-document.mjs 12

   To change a flow, change flows.mjs and regenerate both. */

import { flowToText, flows } from "./flows.mjs";

const byId = Object.fromEntries(flows.map((flow) => [flow.id, flow]));

function flowPage(id, kicker, index, extra = []) {
  const flow = byId[id];
  return {
    kicker,
    title: `${String(index).padStart(2, "0")} / ${flow.title.toUpperCase()}`,
    blocks: [
      { kind: "p", text: flow.blurb },
      {
        kind: "figure",
        file: `/figures/flow-${flow.id}.svg`,
        caption: flow.title,
        hideCaption: true,
        meta: `Source: ${flow.source}`,
        text: flowToText(flow),
      },
      ...extra,
    ],
  };
}

export const meta = {
  number: "12",
  slug: "12_PULSE_Process_Flows",
  title: "Process Flows",
  runningHeader: "PROCESS FLOWS",
  footer: "PULSE PROCESS FLOWS | VERSION 1.0 | JULY 2026",
  orientation: "portrait",
};

export const cover = {
  kicker: "HOW WORK MOVES",
  title: "Process Flows",
  standfirst:
    "Four sequences that answer most of what people ask: what happens to a travel request, how a document reaches signature, where a ticket goes, and how a hosted tool finds its SharePoint site.",
  callout: {
    label: "NOTE",
    text: "A diagram shows the route, not the rules. Who may move a record between these states, and what evidence each move requires, is in the Standard Operating Procedures.",
  },
  spine: [
    ["Flows", "Four sequences"],
    ["States", "As the app displays them"],
    ["Exceptions", "Shown, not hidden"],
    ["Authority", "SOPs, not this document"],
  ],
};

export const pages = [
  {
    kicker: "READING THESE DIAGRAMS",
    title: "01 / HOW TO READ A FLOW",
    blocks: [
      { kind: "p", text: "Each diagram runs from the entry point to the closed state. States and labels are the values the application actually displays, so a diagram can be checked against a record rather than trusted." },
      {
        kind: "table",
        rows: [
          ["SHOWN AS", "MEANS"],
          ["Blue outline", "Entry point — where the record is created"],
          ["Plain outline", "In progress — the record is being worked"],
          ["Green outline", "Closed with an outcome recorded"],
          ["Dashed outline", "Exception — the record left the normal route"],
          ["Solid arrow", "The normal progression"],
          ["Dashed arrow", "An exception or a loop back for rework"],
        ],
      },
      {
        kind: "callout",
        label: "IMPORTANT",
        text: "A state on a diagram is not a permission. Reaching a state says nothing about who was entitled to move the record there; role boundaries are defined in the procedures and enforced by the application.",
      },
    ],
  },
  flowPage("travel", "TRAVEL", 2, [
    { kind: "h4", text: "WHAT CLOSES A TRIP" },
    { kind: "p", text: "Approval is not the end of the record. A trip stays flagged in My Travel until every traveller has filed a debrief, which is why the closed state here is the debrief rather than the approval." },
  ]),
  flowPage("docreview", "REVIEW", 3, [
    { kind: "h4", text: "WHY THE LOOP MATTERS" },
    { kind: "p", text: "Requested changes produce a new revision and the review runs again; the reviewed revision is never overwritten. A file stored under a project has not been through this loop and is not review evidence." },
  ]),
  flowPage("ticket", "SUPPORT", 4, [
    { kind: "h4", text: "ONE RECORD, TWO DOORS" },
    { kind: "p", text: "Tickets and issue reports are the same records in the PULSE Issues list. A ticket raised in the standalone Tickets package is the one the full application shows, so the same problem must not be opened twice." },
  ]),
  flowPage("site", "HOSTING", 5, [
    { kind: "h4", text: "WHAT LOCAL-ONLY MEANS" },
    { kind: "p", text: "Reaching the last step is not a failure to handle — it is the deliberate signal that the tool was loaded outside a SharePoint page. Local state is a development interface mode and is never an authoritative record." },
    {
      kind: "callout",
      label: "MAINTENANCE RULE",
      text: "If a hosted tool shows local-only behaviour, open the full PULSE page once on the same site and reload the tool; that lets the cached-address step succeed. Persistent failure is a ticket with the page address attached.",
    },
  ]),
];
