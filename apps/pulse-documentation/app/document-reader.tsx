import type { ReactNode } from "react";

type TextBlock =
  | { kind: "heading"; text: string; level: 2 | 3 | 4 }
  | { kind: "paragraph"; text: string }
  | { kind: "callout"; label: string; text: string }
  | { kind: "list"; ordered: boolean; items: string[] }
  | { kind: "table"; rows: string[][]; header: boolean }
  /* The extraction carries a figure's caption and description as text so search
     can reach it. The reader replaces that text with the real figure, so the
     parser emits a slot here and the caller fills it in order. */
  | { kind: "figureSlot"; index: number };

const spacedHeadingFixes: Array<[RegExp, string]> = [
  [/I N T E R N A L A S S U R A N C E PAC K AG E/gi, "INTERNAL ASSURANCE PACKAGE"],
  [/P R O C E D U R E V E R I F I CAT I O N/gi, "PROCEDURE VERIFICATION"],
  [/P R O C E D U R E M AT R I X/gi, "PROCEDURE MATRIX"],
  [/A P P R OVA L R E C O R D/gi, "APPROVAL RECORD"],
  [/F E AT U R E V E R I F I CAT I O N/gi, "FEATURE VERIFICATION"],
  [/F E AT U R E M AT R I X/gi, "FEATURE MATRIX"],
  [/L I V E V E R I F I CAT I O N/gi, "LIVE VERIFICATION"],
  [/O P E R AT I O N A L E V I D E N C E/gi, "OPERATIONAL EVIDENCE"],
  [/G OV E R N A N C E D E C I S I O N S/gi, "GOVERNANCE DECISIONS"],
  [/G OV E R N A N C E/gi, "GOVERNANCE"],
  [/T R AC E A B I L I T Y/gi, "TRACEABILITY"],
  [/C L A S S I F I CAT I O N/gi, "CLASSIFICATION"],
];

const calloutLabels = [
  "BRAND PROMISE",
  "CORE INTERPRETATION",
  "CONSTRUCTION RULES",
  "THE RESULT",
  "A CRITICAL DISTINCTION",
  "BRAND IN ONE SENTENCE",
  "MEETING DISCIPLINE",
  "FINAL CHECK",
  "FORMAL DOCUMENT REVIEW ROUTE",
  "OPERATING STANDARD",
  "RELEASE READINESS",
  "MAINTENANCE RULE",
  "IMPORTANT",
  "NOTE",
  "CAUTION",
  "REMEMBER",
];

function tidy(value: string) {
  let text = value
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\s+([,.;:!])/g, "$1")
    .replace(/([([])\s+/g, "$1")
    .replace(/\s+([)\]])/g, "$1")
    .trim();
  for (const [pattern, replacement] of spacedHeadingFixes) {
    text = text.replace(pattern, replacement);
  }
  return text.replace(/^(\d)\s+(\d)\s*\/\s*/, "$1$2 / ");
}

function isFooterOrRunningHeader(line: string) {
  const compact = tidy(line);
  if (!compact) return false;
  return (
    (/^\s{40,}\S/.test(line) && isUpperLabel(compact)) ||
    /^PULSE\s{6,}\S/.test(line) ||
    /^PULSE BRAND IDENTITY\b.*\b\d{2}$/i.test(compact) ||
    /^PULSE PRODUCT BRIEF\b.*\b\d{2}$/i.test(compact) ||
    /^PULSE GENERAL OVERVIEW\b.*\b\d{2}$/i.test(compact) ||
    /^PULSE \| INTERNAL USE\b/i.test(compact) ||
    /^PULSE TECHNICAL REFERENCE\b.*\b\d{2}$/i.test(compact) ||
    /^PULSE - STANDARD OPERATING PROCEDURES\b/i.test(compact) ||
    /^PULSE DOCUMENTATION ASSURANCE\b.*\b\d{2}$/i.test(compact) ||
    /^PULSE \| INTERNAL TECHNICAL GUIDE\b/i.test(compact) ||
    /^PULSE\s{1,}.*(?:INTERNAL OVERVIEW|INTERNAL USER GUIDE|CONTROLLED INTERNAL PROCEDURE|CODEBASE SETUP & FIREPIT PACKAGING)$/i.test(compact)
  );
}

function cleanLines(raw: string) {
  return raw
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+$/, ""))
    .filter((line) => !isFooterOrRunningHeader(line));
}

function isUpperLabel(text: string) {
  const letters = text.replace(/[^A-Za-z]/g, "");
  return letters.length >= 4 && text === text.toUpperCase();
}

function isSectionHeading(text: string) {
  return /^\d{1,2}\s*\/\s*\S/.test(text) || /^SOP\s+\d{1,2}\b/.test(text);
}

function splitCells(line: string) {
  const cells: Array<{ text: string; start: number }> = [];
  const matcher = /\S(?:.*?\S)?(?=\s{3,}|$)/g;
  for (const match of line.matchAll(matcher)) {
    const text = tidy(match[0]);
    if (text) cells.push({ text, start: match.index ?? 0 });
  }
  return cells;
}

function tableFrom(lines: string[]): TextBlock | null {
  if (lines.length < 2) return null;
  const parsed = lines.filter((line) => line.trim()).map(splitCells);
  const multi = parsed.filter((row) => row.length >= 2);
  if (multi.length < 2 || multi.length / parsed.length < 0.55) return null;

  const anchor = multi.reduce((best, row) => (row.length > best.length ? row : best), multi[0]);
  const starts = anchor.map((cell) => cell.start);
  if (starts.length > 6) return null;

  const positioned = parsed.map((row) => {
    const result = Array(starts.length).fill("");
    for (const cell of row) {
      let column = 0;
      let distance = Number.POSITIVE_INFINITY;
      starts.forEach((start, index) => {
        const nextDistance = Math.abs(start - cell.start);
        if (nextDistance < distance) {
          distance = nextDistance;
          column = index;
        }
      });
      result[column] = result[column] ? `${result[column]} ${cell.text}` : cell.text;
    }
    return result;
  });

  const rows: string[][] = [];
  for (const row of positioned) {
    const firstFilled = row.findIndex(Boolean);
    const filled = row.filter(Boolean);
    const continuation =
      rows.length > 0 &&
      (firstFilled > 0 || (filled.length > 0 && filled.every((cell) => /^[a-z([]/.test(cell))));
    if (continuation) {
      row.forEach((cell, index) => {
        if (cell) rows[rows.length - 1][index] = tidy(`${rows[rows.length - 1][index]} ${cell}`);
      });
    } else {
      rows.push(row.map(tidy));
    }
  }

  const header =
    rows[0].filter(Boolean).length > 1 &&
    (
      rows[0].filter(Boolean).every((cell) => isUpperLabel(cell.replace(/^\d{1,2}\s*\/\s*/, ""))) ||
      (
        rows.length > 1 &&
        rows[0].filter(Boolean).every((cell) => cell.length < 48 && /^[A-Z]/.test(cell) && !/[.!?]$/.test(cell))
      )
    );
  return { kind: "table", rows, header };
}

function listFrom(lines: string[]): TextBlock | null {
  const marker = /^\s*(?:[•●▪✓□]\s*|\d{1,2}[.)]\s+|\d{1,2}\s*\/\s+)/;
  const markerCount = lines.filter((line) => marker.test(line)).length;
  if (markerCount < 2) return null;
  const ordered = lines.some((line) => /^\s*\d/.test(line));
  const items: string[] = [];
  for (const rawLine of lines) {
    const line = tidy(rawLine);
    if (!line) continue;
    if (marker.test(rawLine)) {
      items.push(tidy(rawLine.replace(marker, "")));
    } else if (items.length) {
      items[items.length - 1] = tidy(`${items[items.length - 1]} ${line}`);
    }
  }
  return items.length ? { kind: "list", ordered, items } : null;
}

function calloutFrom(text: string): TextBlock | null {
  const label = calloutLabels.find((candidate) => text.toUpperCase().startsWith(candidate));
  if (!label) return null;
  const body = tidy(text.slice(label.length).replace(/^[:.\s-]+/, ""));
  return { kind: "callout", label, text: body };
}

export function parsePage(raw: string) {
  const lines = cleanLines(raw);
  const chunks: string[][] = [];
  let current: string[] = [];
  for (const line of lines) {
    if (!line.trim()) {
      if (current.length) chunks.push(current);
      current = [];
    } else {
      current.push(line);
    }
  }
  if (current.length) chunks.push(current);

  const groupedChunks: string[][] = [];
  for (let index = 0; index < chunks.length;) {
    const hasColumns = (chunk: string[]) => chunk.some((line) => splitCells(line).length >= 2);
    const hasListMarker = (chunk: string[]) =>
      chunk.some((line) => /^\s*(?:[•●▪✓□]\s*|\d{1,2}[.)]\s+)/.test(line));
    const mode = hasColumns(chunks[index]) ? "table" : hasListMarker(chunks[index]) ? "list" : null;
    if (!mode) {
      groupedChunks.push(chunks[index]);
      index += 1;
      continue;
    }

    const merged = [...chunks[index]];
    let next = index + 1;
    while (
      next < chunks.length &&
      (mode === "table" ? hasColumns(chunks[next]) : hasListMarker(chunks[next]))
    ) {
      merged.push(...chunks[next]);
      next += 1;
    }
    groupedChunks.push(merged);
    index = next;
  }

  const blocks: TextBlock[] = [];
  let figureIndex = 0;
  groupedChunks.forEach((chunk, chunkIndex) => {
    const joined = tidy(chunk.join(" "));
    if (!joined) return;

    /* Emitted by the document generator as one contiguous block. */
    if (/^FIGURE\s/.test(chunk[0].trim())) {
      blocks.push({ kind: "figureSlot", index: figureIndex });
      figureIndex += 1;
      return;
    }

    const table = tableFrom(chunk);
    if (table) {
      blocks.push(table);
      return;
    }

    const list = listFrom(chunk);
    if (list) {
      blocks.push(list);
      return;
    }

    const callout = calloutFrom(joined);
    if (callout) {
      blocks.push(callout);
      return;
    }

    if (chunk.length === 1 && isSectionHeading(joined)) {
      blocks.push({ kind: "heading", text: joined, level: 3 });
      return;
    }
    if (chunk.length === 1 && isUpperLabel(joined) && joined.length <= 90) {
      blocks.push({ kind: "heading", text: joined, level: 4 });
      return;
    }
    if (chunk.length === 1 && joined.length <= 92 && !/[.!?]$/.test(joined)) {
      blocks.push({ kind: "heading", text: joined, level: chunkIndex < 2 ? 2 : 3 });
      return;
    }
    blocks.push({ kind: "paragraph", text: joined });
  });

  for (let index = 0; index < blocks.length - 1; index += 1) {
    const block = blocks[index];
    const next = blocks[index + 1];
    if (block.kind === "callout" && !block.text && next.kind === "paragraph") {
      block.text = next.text;
      blocks.splice(index + 1, 1);
    }
  }

  const sourceTitle = lines.map(tidy).find(isSectionHeading);
  const titleIndex = blocks.findIndex((block) => block.kind === "heading" && block.level === 3);
  const fallbackTitleIndex = blocks.findIndex((block) => block.kind === "heading" && block.level === 2);
  const anyTitleIndex = blocks.findIndex((block) => block.kind === "heading");
  const resolvedIndex = titleIndex >= 0 ? titleIndex : fallbackTitleIndex >= 0 ? fallbackTitleIndex : anyTitleIndex;
  const title =
    sourceTitle ??
    (resolvedIndex >= 0 && blocks[resolvedIndex].kind === "heading" ? blocks[resolvedIndex].text : "Document page");

  return {
    title,
    blocks: blocks.filter((block, index) => index !== resolvedIndex && !(block.kind === "heading" && block.text === title)),
  };
}

function renderBlock(block: TextBlock, key: string, figures?: ReactNode[]): ReactNode {
  if (block.kind === "figureSlot") return figures?.[block.index] ?? null;
  if (block.kind === "heading") {
    if (block.level === 2) return <h2 key={key}>{block.text}</h2>;
    if (block.level === 3) return <h3 key={key}>{block.text}</h3>;
    return <h4 key={key}>{block.text}</h4>;
  }
  if (block.kind === "paragraph") return <p key={key}>{block.text}</p>;
  if (block.kind === "callout") {
    return <aside className="doc-callout" key={key}><strong>{block.label}</strong>{block.text && <p>{block.text}</p>}</aside>;
  }
  if (block.kind === "list") {
    const List = block.ordered ? "ol" : "ul";
    return <List className="doc-list" key={key}>{block.items.map((item, index) => <li key={`${key}-${index}`}>{item}</li>)}</List>;
  }
  const bodyRows = block.header ? block.rows.slice(1) : block.rows;
  return <div className="doc-table-wrap" key={key}><table>{block.header && <thead><tr>{block.rows[0].map((cell, index) => <th key={`${key}-h-${index}`}>{cell}</th>)}</tr></thead>}<tbody>{bodyRows.map((row, rowIndex) => <tr key={`${key}-r-${rowIndex}`}>{row.map((cell, cellIndex) => <td key={`${key}-${rowIndex}-${cellIndex}`}>{cell}</td>)}</tr>)}</tbody></table></div>;
}

export function DocumentPage({ raw, pageNumber, documentTitle, figures }: { raw: string; pageNumber: number; documentTitle: string; figures?: ReactNode }) {
  const parsed = parsePage(raw);
  /* Figures arrive as a single node containing one child per figure, in document
     order; split it so each slot receives its own. */
  const figureNodes = Array.isArray(figures) ? figures : figures ? [figures] : [];
  return <article className="formatted-page" id={`page-${pageNumber}`}>
    <header className="formatted-page-header">
      <span>{documentTitle}</span>
      <b>Source page {String(pageNumber).padStart(2, "0")}</b>
    </header>
    <div className="formatted-page-body">{parsed.blocks.map((block, index) => renderBlock(block, `${pageNumber}-${index}`, figures))}</div>
    <footer>Source page {pageNumber}</footer>
  </article>;
}
