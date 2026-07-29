/* Process-flow definitions and a standalone SVG renderer.

   One source, two consumers: `build-figures.mjs` writes each flow to
   public/figures/flow-<id>.svg, and both the generated PDF and the web reader
   display that same file. Nothing redraws these independently, so a diagram in
   the PDF cannot disagree with the one on the page.

   Because the SVG is loaded through <img>, it cannot inherit page CSS — the
   style block below is self-contained and uses the brand hex values directly.

   States and labels are taken from the values the application actually
   displays. If a status vocabulary changes in the app, change it here and
   regenerate; do not redraw a prettier version. */

const BRAND = {
  ink: "#070708",
  field: "#ffffff",
  graphite: "#51545a",
  mist: "#ecedef",
  line: "#d5d7dc",
  blue: "#2f66ff",
  green: "#2b7a5e",
  muted: "#9aa0a8",
  rail: "#c3c7cd",
  index: "#a4a8b0",
};

export const flows = [
  {
    id: "travel",
    title: "Travel request",
    blurb: "A request moves through approval and funding before travel, and is not closed until every traveller has filed a debrief.",
    source: "Travel request status values and the travel module",
    layout: "row",
    steps: [
      { label: "New request", note: "Travel → Submit", tone: "start" },
      { label: "Pending", note: "Awaiting approval" },
      { label: "Pending Finance", note: "Funding and charge object" },
      { label: "Approved", note: "Appears on the calendar" },
      { label: "Debrief filed", note: "One per traveller", tone: "done" },
    ],
    exceptions: [
      { fromStep: 1, label: "Denied · Cancelled · Withdrawn" },
      { fromStep: 3, label: "Revoked" },
    ],
  },
  {
    id: "docreview",
    title: "Document review",
    blurb: "Formal review is a controlled loop: requested changes produce a new revision and the review runs again. Storing a file under a project is not review.",
    source: "Document Review workflow values",
    layout: "row",
    steps: [
      { label: "Not Started", tone: "start" },
      { label: "In Review", note: "Reviewers decide" },
      { label: "Review Complete" },
      { label: "Awaiting Final Pack", note: "Then Signing in Progress" },
      { label: "Signed", note: "Then Archived", tone: "done" },
    ],
    loop: { atStep: 1, label: "Changes Requested → new revision" },
  },
  {
    id: "ticket",
    title: "Support ticket",
    blurb: "Tickets and issue reports are the same records in the PULSE Issues list, so a ticket raised in the standalone tool is the one the full application shows.",
    source: "Ticket status values and the repository mapping",
    layout: "row",
    steps: [
      { label: "Open", note: "Raised by any user", tone: "start" },
      { label: "In Progress", note: "Being worked" },
      { label: "Resolved", note: "Outcome recorded", tone: "done" },
    ],
  },
  {
    id: "site",
    title: "Site resolution in a web part",
    blurb: "A focused tool in a SharePoint iframe cannot see the page context directly. It tries six sources in order, and local-only mode is the deliberate signal that none succeeded.",
    source: "getSiteUrl in the SharePoint adapter",
    layout: "column",
    steps: [
      { label: "Own page context", note: "Direct SharePoint page", tone: "start", via: "not found" },
      { label: "Configured site", note: "manualSharePointSiteUrl", via: "not set" },
      { label: "Parent window", note: "Normal web-part path", via: "not found" },
      { label: "Top window", note: "Nested frame", via: "not found" },
      { label: "Cached address", note: "From a previous boot", via: "not cached" },
      { label: "Local only", note: "Not a system of record", tone: "exception" },
    ],
  },
];

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function styleBlock() {
  return `<style>
    .n rect { fill:${BRAND.field}; stroke:${BRAND.line}; stroke-width:1; }
    .n.start rect { fill:#eaf0ff; stroke:${BRAND.blue}; }
    .n.done rect { fill:#e7f2ed; stroke:${BRAND.green}; }
    .n.exception rect { fill:${BRAND.mist}; stroke:${BRAND.muted}; stroke-dasharray:4 3; }
    .lbl { fill:${BRAND.ink}; font:650 12.5px Inter,Arial,Helvetica,sans-serif; }
    .note { fill:${BRAND.graphite}; font:10.5px Inter,Arial,Helvetica,sans-serif; }
    .edge { fill:${BRAND.graphite}; font:10px Inter,Arial,Helvetica,sans-serif; }
    .edge.muted { fill:#767b83; }
    .idx { fill:${BRAND.index}; font:10px ui-monospace,SFMono-Regular,Menlo,monospace; }
  </style>`;
}

function markers(id) {
  return `<defs>
    <marker id="a-${id}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,1 L9,5 L0,9 z" fill="${BRAND.graphite}"/></marker>
    <marker id="am-${id}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,1 L9,5 L0,9 z" fill="${BRAND.muted}"/></marker>
  </defs>`;
}

function renderRow(flow) {
  const BOX_W = 176, BOX_H = 74, GAP = 46, TOP = 12;
  const count = flow.steps.length;
  const width = count * BOX_W + (count - 1) * GAP;
  const hasExceptions = Boolean(flow.exceptions?.length);
  const hasLoop = Boolean(flow.loop);
  const exceptionY = TOP + BOX_H + 66;
  const height = TOP + BOX_H + (hasExceptions ? 92 : hasLoop ? 74 : 14);
  const xFor = (index) => index * (BOX_W + GAP);
  const midY = TOP + BOX_H / 2;
  const parts = [];

  for (let index = 0; index < count - 1; index += 1) {
    const x1 = xFor(index) + BOX_W;
    const x2 = xFor(index + 1);
    parts.push(`<line x1="${x1}" y1="${midY}" x2="${x2 - 3}" y2="${midY}" stroke="${BRAND.graphite}" stroke-width="1.5" marker-end="url(#a-${flow.id})"/>`);
  }

  flow.steps.forEach((step, index) => {
    const x = xFor(index);
    parts.push(`<g class="n ${step.tone ?? "step"}">
      <rect x="${x}" y="${TOP}" width="${BOX_W}" height="${BOX_H}" rx="2"/>
      <text class="lbl" x="${x + 14}" y="${TOP + (step.note ? 30 : 42)}">${escapeXml(step.label)}</text>
      ${step.note ? `<text class="note" x="${x + 14}" y="${TOP + 51}">${escapeXml(step.note)}</text>` : ""}
    </g>`);
  });

  if (hasLoop) {
    const x = xFor(flow.loop.atStep);
    const y = TOP + BOX_H;
    parts.push(`<path d="M ${x + 30} ${y} V ${y + 32} H ${x + BOX_W - 20} V ${y + 3}" fill="none" stroke="${BRAND.muted}" stroke-width="1.5" stroke-dasharray="4 3" marker-end="url(#am-${flow.id})"/>`);
    parts.push(`<text class="edge muted" x="${x + 30}" y="${y + 50}">${escapeXml(flow.loop.label)}</text>`);
  }

  if (hasExceptions) {
    const columns = flow.exceptions.map((exception) => exception.fromStep);
    const railStart = xFor(Math.min(...columns)) + BOX_W / 2;
    const railEnd = xFor(Math.max(...columns)) + BOX_W / 2;
    if (railEnd > railStart) {
      parts.push(`<line x1="${railStart}" y1="${exceptionY}" x2="${railEnd}" y2="${exceptionY}" stroke="${BRAND.rail}" stroke-width="1.5" stroke-dasharray="4 3"/>`);
    }
    flow.exceptions.forEach((exception) => {
      const x = xFor(exception.fromStep) + BOX_W / 2;
      parts.push(`<path d="M ${x} ${TOP + BOX_H} V ${exceptionY - 2}" fill="none" stroke="${BRAND.muted}" stroke-width="1.5" stroke-dasharray="4 3" marker-end="url(#am-${flow.id})"/>`);
      parts.push(`<circle cx="${x}" cy="${exceptionY}" r="2.5" fill="${BRAND.muted}"/>`);
      parts.push(`<text class="edge muted" x="${x}" y="${exceptionY + 17}" text-anchor="middle">${escapeXml(exception.label)}</text>`);
    });
    parts.push(`<text class="edge muted" x="${railEnd + 14}" y="${exceptionY + 4}">leaves the flow</text>`);
  }

  return { width: width + 4, height, viewBox: `-2 0 ${width + 4} ${height}`, body: parts.join("\n") };
}

function renderColumn(flow) {
  const BOX_W = 296, BOX_H = 64, GAP = 38;
  const width = BOX_W + 46;
  const height = flow.steps.length * (BOX_H + GAP) - GAP;
  const yFor = (index) => index * (BOX_H + GAP);
  const rail = 34;
  const parts = [];

  flow.steps.forEach((step, index) => {
    if (index < flow.steps.length - 1) {
      const y1 = yFor(index) + BOX_H;
      const y2 = yFor(index + 1);
      parts.push(`<line x1="${rail}" y1="${y1}" x2="${rail}" y2="${y2 - 3}" stroke="${BRAND.graphite}" stroke-width="1.5" marker-end="url(#a-${flow.id})"/>`);
      if (step.via) parts.push(`<text class="edge" x="${rail + 18}" y="${(y1 + y2) / 2 + 3}">${escapeXml(step.via)}</text>`);
    }
    const y = yFor(index);
    parts.push(`<g class="n ${step.tone ?? "step"}">
      <rect x="0" y="${y}" width="${BOX_W}" height="${BOX_H}" rx="2"/>
      <text class="lbl" x="15" y="${y + (step.note ? 26 : 38)}">${escapeXml(step.label)}</text>
      ${step.note ? `<text class="note" x="15" y="${y + 45}">${escapeXml(step.note)}</text>` : ""}
      <text class="idx" x="${BOX_W + 16}" y="${y + (step.note ? 33 : 38)}">${String(index + 1).padStart(2, "0")}</text>
    </g>`);
  });

  return { width, height: height + 4, viewBox: `-2 -2 ${width + 4} ${height + 4}`, body: parts.join("\n") };
}

export function renderFlowSvg(flow) {
  const geometry = flow.layout === "column" ? renderColumn(flow) : renderRow(flow);
  const description = `${flow.title}: ${flow.steps.map((step) => step.label).join(", then ")}`;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${geometry.viewBox}" width="${geometry.width}" height="${geometry.height}" role="img" aria-label="${escapeXml(description)}">
${styleBlock()}
${markers(flow.id)}
${geometry.body}
</svg>
`;
}

/* Plain-text rendering of a flow, for the searchable extraction. A reader
   searching "changes requested" should reach the diagram's document. */
export function flowToText(flow) {
  const lines = [];
  flow.steps.forEach((step, index) => {
    const bits = [`${String(index + 1).padStart(2, "0")}  ${step.label}`];
    if (step.note) bits.push(`— ${step.note}`);
    if (step.via) bits.push(`(next if ${step.via})`);
    lines.push(bits.join(" "));
  });
  if (flow.loop) lines.push(`Loop at ${flow.steps[flow.loop.atStep].label}: ${flow.loop.label}`);
  if (flow.exceptions?.length) {
    flow.exceptions.forEach((exception) => {
      lines.push(`Leaves the flow from ${flow.steps[exception.fromStep].label}: ${exception.label}`);
    });
  }
  return lines;
}
