/* PULSE Overview PPTX export.
   The visual contract is normalized from the supplied My Overview and Team
   Overview OOXML references in overview-pptx-template.js. */
(function () {
  window.AEWTTR = window.AEWTTR || {};

  const T = window.PULSE_OVERVIEW_PPTX_TEMPLATE;
  if (!T) {
    console.error("PULSE: overview PPTX template specification is unavailable.");
    return;
  }

  const C = T.colors;
  const G = T.chrome;
  const CUI = T.cui;
  const COVER = T.cover;
  const FONT = T.fontFace;
  const W = T.slide.width;
  const H = T.slide.height;

  function safe(value) {
    return String(value == null ? "" : value)
      .replace(/[‐-―]/g, "-")
      .replace(/…/g, "...")
      .replace(/[""]/g, '"')
      .replace(/['']/g, "'")
      .normalize("NFKD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/[^\x20-\x7E]/g, "");
  }

  function pptxCtor() {
    const ctor = window.PptxGenJS || window.pptxgenjs || window.pptxgen;
    if (!ctor) throw new Error("PptxGenJS is not loaded");
    return typeof ctor === "function" ? ctor : ctor.default;
  }

  function isoDate(value) {
    if (!value) return "-";
    const match = String(value).match(/^\d{4}-\d{2}-\d{2}/);
    return match ? match[0] : safe(value);
  }

  function generatedTimestamp(date) {
    const d = date || new Date();
    const offset = -d.getTimezoneOffset();
    const sign = offset >= 0 ? "+" : "-";
    const hh = String(Math.floor(Math.abs(offset) / 60)).padStart(2, "0");
    const mm = String(Math.abs(offset) % 60).padStart(2, "0");
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}T${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:${String(d.getSeconds()).padStart(2, "0")}${sign}${hh}:${mm}`;
  }

  function monthYear(date) {
    return (date || new Date()).toLocaleDateString("en-US", {
      month: "long",
      year: "numeric"
    }).toUpperCase();
  }

  function metricValue(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return safe(value || "0");
    return n >= 0 && n < 10 ? String(n).padStart(2, "0") : String(n);
  }

  function percent(numerator, denominator) {
    const den = Number(denominator) || 0;
    if (!den) return 0;
    return Math.max(0, Math.min(100, Math.round((Number(numerator || 0) / den) * 100)));
  }

  function chunk(items, size) {
    const source = Array.isArray(items) ? items : [];
    if (!source.length) return [[]];
    const groups = [];
    for (let i = 0; i < source.length; i += size) groups.push(source.slice(i, i + size));
    return groups;
  }

  function taskIsDone(task) {
    const status = String(task && task.status || "").toLowerCase();
    return !!(task && task.done) || ["done", "completed", "complete", "closed"].includes(status);
  }

  function taskIsBlocked(task) {
    const status = String(task && task.status || "").toLowerCase();
    const health = String(task && task.health || "").toLowerCase();
    return status === "blocked" || status === "on hold" || health === "blocked";
  }

  function taskIsAtRisk(task) {
    const status = String(task && task.status || "").toLowerCase();
    const health = String(task && task.health || "").toLowerCase();
    return status.includes("risk") || health.includes("risk") || health.includes("off track");
  }

  function taskDisplayStatus(task) {
    const status = String(task && task.status || "").trim();
    if (taskIsDone(task)) return "Complete";
    if (taskIsBlocked(task)) return "Blocked";
    if (taskIsAtRisk(task)) return "At Risk";
    if (/review/i.test(status)) return "In Review";
    return status || "On Track";
  }

  function taskHealth(task) {
    const health = String(task && task.health || "").trim();
    if (taskIsBlocked(task)) return "Blocked";
    if (taskIsAtRisk(task)) return "At Risk";
    if (!health || /healthy|on track/i.test(health)) return "Healthy";
    return health;
  }

  function dateIsOverdue(value, now) {
    if (!value) return false;
    const d = new Date(String(value).slice(0, 10) + "T23:59:59");
    return !Number.isNaN(d.getTime()) && d < (now || new Date());
  }

  function riskIsOpen(risk) {
    const status = String(risk && risk.status || "").toLowerCase();
    return !["closed", "resolved", "accepted"].includes(status);
  }

  function statusStyle(value) {
    const status = String(value || "").toLowerCase();
    if (status.includes("block") || status.includes("off track")) {
      return { text: C.red, fill: C.redTint };
    }
    if (status.includes("risk") || status.includes("attention") || status.includes("medium") || status.includes("high")) {
      return { text: C.amber, fill: C.amberTint };
    }
    if (status.includes("review") || status.includes("complete") || status.includes("done")) {
      return { text: C.accent, fill: C.blueTint };
    }
    if (status.includes("track") || status.includes("active") || status.includes("healthy") || status.includes("low")) {
      return { text: C.green, fill: C.greenTint };
    }
    return { text: C.muted, fill: C.neutralTint };
  }

  function presentationProjectStatus(project, tasks) {
    const active = (tasks || []).filter(t => !taskIsDone(t));
    if (active.some(taskIsBlocked)) return "Blocked";
    if (active.some(taskIsAtRisk) || active.some(t => dateIsOverdue(t.end))) return "At Risk";
    if (!active.length) return (tasks || []).length ? "Complete" : "No Active Work";
    return "On Track";
  }

  function weightedWidths(total, weights) {
    const sum = weights.reduce((acc, value) => acc + value, 0) || 1;
    return weights.map(value => Number((total * value / sum).toFixed(3)));
  }

  function bytesToBase64(bytes) {
    let binary = "";
    const chunkSize = 0x8000;
    for (let index = 0; index < bytes.length; index += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
    }
    return btoa(binary);
  }

  async function asDataUri(source) {
    if (!source) return "";
    if (/^data:/i.test(source)) return source;
    const response = await fetch(source, { credentials: "same-origin" });
    if (!response.ok) throw new Error(`Could not load presentation asset: ${source}`);
    const bytes = new Uint8Array(await response.arrayBuffer());
    const mimeType = response.headers.get("content-type") || "application/octet-stream";
    return `data:${mimeType};base64,${bytesToBase64(bytes)}`;
  }

  async function loadTemplateAssets() {
    const [wordmarkWhite, wordmarkBlack, seal] = await Promise.all([
      asDataUri(T.assets.wordmarkWhite),
      asDataUri(T.assets.wordmarkBlack),
      asDataUri(T.assets.seal)
    ]);
    return { wordmarkWhite, wordmarkBlack, seal };
  }

  function addText(slide, text, box, options) {
    slide.addText(text, {
      ...box,
      fontFace: FONT,
      margin: 0,
      breakLine: false,
      shrinkText: true,
      ...options
    });
  }

  function addImage(slide, data, box) {
    if (data) slide.addImage({ data, ...box });
  }

  function addRule(slide, box, color, width) {
    slide.addShape("line", {
      ...box,
      line: { color: color || C.line, pt: width || 0.7 }
    });
  }

  function addCuiRails(slide, darkBackground) {
    const labelColor = darkBackground ? C.white : C.ink;
    const railColor = darkBackground ? C.coverPanelLine : C.tableHead;
    addText(slide, "CUI", CUI.topLabel, {
      fontSize: 6.2,
      color: labelColor,
      bold: true,
      align: "center",
      valign: "middle"
    });
    addRule(slide, CUI.topRule, railColor, 0.45);
    addRule(slide, CUI.bottomRule, railColor, 0.45);
    addText(slide, "CUI", CUI.bottomLabel, {
      fontSize: 6.2,
      color: labelColor,
      bold: true,
      align: "center",
      valign: "middle"
    });
  }

  function addContentChrome(slide, assets, options) {
    slide.background = { color: C.white };
    addCuiRails(slide, false);
    addImage(slide, assets.wordmarkBlack, G.wordmark);
    addText(slide, safe(options.sectionLabel || "OVERVIEW"), G.section, {
      fontSize: 6.6,
      color: C.accent,
      bold: true,
      valign: "middle"
    });
    addImage(slide, assets.seal, G.seal);
    addRule(slide, G.topRule, C.line, 0.7);
    addText(slide, safe(options.title || ""), G.title, {
      fontSize: 23.5,
      color: C.ink,
      bold: true,
      valign: "middle"
    });
    addRule(slide, G.footerRule, C.line, 0.7);
    addText(slide, safe(options.footerLeft || "PULSE OVERVIEW"), G.footerLeft, {
      fontSize: 6.1,
      color: C.quiet,
      bold: true
    });
    addText(slide, `SCHEMA ${T.schemaVersion}  |  ${monthYear()}`, G.footerRight, {
      fontSize: 6.1,
      color: C.quiet,
      align: "right"
    });
  }

  function addSectionLabel(slide, text, x, y, w) {
    addText(slide, safe(text), { x, y, w: w || 1.8, h: 0.16 }, {
      fontSize: 6.6,
      color: C.accent,
      bold: true
    });
  }

  function addMetric(slide, metric, x, y) {
    addText(slide, metricValue(metric.value), { x, y, w: 0.82, h: 0.36 }, {
      fontSize: 21.5,
      color: metric.color || C.ink
    });
    addText(slide, safe(metric.label).toUpperCase(), { x, y: y + 0.42, w: 0.96, h: 0.18 }, {
      fontSize: 5.9,
      color: C.muted,
      bold: true
    });
  }

  function addBar(slide, options) {
    const max = Math.max(1, Number(options.max) || 1);
    const value = Math.max(0, Number(options.value) || 0);
    const fraction = Math.max(0, Math.min(1, value / max));
    addText(slide, safe(options.label), { x: options.x, y: options.y, w: options.labelW, h: 0.16 }, {
      fontSize: options.fontSize || 7,
      color: C.muted
    });
    slide.addShape("rect", {
      x: options.barX,
      y: options.y + 0.11,
      w: options.barW,
      h: 0.10,
      line: { color: C.barTrack, transparency: 100 },
      fill: { color: C.barTrack }
    });
    if (fraction > 0) {
      slide.addShape("rect", {
        x: options.barX,
        y: options.y + 0.11,
        w: Math.max(0.03, options.barW * fraction),
        h: 0.10,
        line: { color: options.color || C.accent, transparency: 100 },
        fill: { color: options.color || C.accent }
      });
    }
    addText(slide, safe(options.valueLabel == null ? value : options.valueLabel), {
      x: options.valueX,
      y: options.y,
      w: options.valueW || 0.32,
      h: 0.16
    }, {
      fontSize: options.fontSize || 7,
      color: options.valueColor || C.ink,
      align: "right"
    });
  }

  function addCallout(slide, text, box) {
    slide.addShape("roundRect", {
      ...box,
      rectRadius: 0.05,
      line: { color: C.ink, transparency: 100 },
      fill: { color: C.ink }
    });
    addText(slide, safe(text), {
      x: box.x + 0.20,
      y: box.y + 0.19,
      w: box.w - 0.40,
      h: Math.max(0.14, box.h - 0.32)
    }, {
      fontSize: 7.3,
      color: C.white,
      bold: true,
      align: "center",
      valign: "middle"
    });
  }

  function addNarrative(slide, lines, box) {
    const text = (lines || []).filter(Boolean).map(line => safe(line)).join("\n");
    addText(slide, text || "No summary items require attention.", box, {
      fontSize: 9.4,
      color: C.ink,
      bullet: { indent: 13 },
      paraSpaceAfter: 4,
      breakLine: true,
      valign: "top"
    });
  }

  function tableCell(value, options) {
    const style = options || {};
    return {
      text: safe(value),
      options: {
        fill: { color: style.fill || C.white },
        color: style.color || C.muted,
        fontFace: FONT,
        fontSize: style.fontSize || 6.7,
        bold: !!style.bold,
        align: style.align || "left",
        valign: "middle",
        margin: style.margin || [2, 3, 2, 3]
      }
    };
  }

  function headerRow(labels, fontSize) {
    return labels.map((label, index) => tableCell(label, {
      fill: C.tableHead,
      color: C.muted,
      fontSize: fontSize || 6.8,
      bold: true,
      align: index === 0 ? "left" : "center"
    }));
  }

  function rowFill(index) {
    return index % 2 === 0 ? C.white : C.tableAlt;
  }

  function statusCell(value, fontSize) {
    const style = statusStyle(value);
    return tableCell(value, {
      fill: style.fill,
      color: style.text,
      fontSize: fontSize || 6.2,
      bold: true,
      align: "center"
    });
  }

  function addTable(slide, rows, options) {
    if (!rows || !rows.length) return;
    slide.addTable(rows, {
      x: options.x,
      y: options.y,
      w: options.w,
      colW: options.colW,
      rowH: options.rowH,
      border: { type: "solid", color: C.line, pt: 0.45 },
      autoPage: false,
      margin: 0
    });
  }

  function addCoverSlide(prs, model, assets) {
    const slide = prs.addSlide();
    slide.background = { color: C.ink };
    addCuiRails(slide, true);
    addImage(slide, assets.wordmarkWhite, COVER.wordmark);
    addRule(slide, COVER.accentRule, C.accent, 3.5);
    addText(slide, model.isMyOverview ? "My Overview" : "Team Overview", COVER.title, {
      fontSize: 34,
      color: C.white,
      bold: true,
      valign: "middle"
    });
    addText(slide, `${model.isMyOverview ? "PULSE personal workload snapshot" : "PULSE portfolio snapshot"} | Reporting as of ${model.reportingAsOf}`, COVER.subtitle, {
      fontSize: 11.5,
      color: C.coverMuted,
      valign: "middle"
    });
    slide.addShape("roundRect", {
      ...COVER.tag,
      rectRadius: 0.04,
      line: { color: C.coverTag, transparency: 100 },
      fill: { color: C.coverTag }
    });
    addText(slide, model.isMyOverview ? "PERSONAL OVERVIEW EXPORT" : "TEAM OVERVIEW EXPORT", {
      x: COVER.tag.x + 0.19,
      y: COVER.tag.y + 0.08,
      w: COVER.tag.w - 0.38,
      h: 0.11
    }, {
      fontSize: 7.3,
      color: C.accent,
      bold: true,
      align: "center"
    });
    addText(slide, [
      `Generated by ${safe(model.generatedBy || "PULSE user")}`,
      `Reporting as of ${model.reportingAsOf}`,
      `Generated ${model.generatedAt}`
    ].join("\n"), COVER.meta, {
      fontSize: 8.4,
      color: C.coverMeta,
      breakLine: true,
      valign: "top"
    });
    addImage(slide, assets.seal, COVER.seal);
    slide.addShape("roundRect", {
      ...COVER.classificationPanel,
      rectRadius: 0.04,
      line: { color: C.coverPanelLine, width: 0.65 },
      fill: { color: C.coverPanel }
    });
    slide.addShape("rect", {
      ...COVER.classificationStripe,
      line: { color: C.accent, transparency: 100 },
      fill: { color: C.accent }
    });
    addText(slide, "CUI DESIGNATION INDICATOR", COVER.classificationTitle, {
      fontSize: 5.8,
      color: C.accent,
      bold: true,
      valign: "middle"
    });
    addText(slide, [
      "Controlled information",
      "Authorized users and approved distribution only",
      "Handle and store according to applicable policy"
    ].join("\n"), COVER.classificationDetail, {
      fontSize: 5.8,
      color: C.white,
      breakLine: true,
      valign: "top"
    });
    addText(slide, "Current information turned into coordinated execution.", COVER.statementDetail, {
      fontSize: 8.3,
      color: C.coverQuiet,
      align: "center"
    });
  }

  function addTeamSummarySlide(prs, model, assets) {
    const slide = prs.addSlide();
    addContentChrome(slide, assets, {
      sectionLabel: "TEAM OVERVIEW",
      title: "Portfolio metrics and status distribution",
      footerLeft: "PULSE TEAM OVERVIEW"
    });
    const m = model.metrics;
    const metrics = [
      { label: "Projects", value: m.totalProjects },
      { label: "Active Proj.", value: m.activeProjects },
      { label: "Total Tasks", value: m.totalTasks },
      { label: "Open Tasks", value: m.activeTasks },
      { label: "Closed Tasks", value: m.closedTasks },
      { label: "Risks", value: m.openRisks, color: C.amber },
      { label: "People", value: m.teamSize },
      { label: "Review", value: m.documentsInReview }
    ];
    metrics.forEach((metric, index) => addMetric(slide, metric, 0.46 + index * 1.11, 1.78));

    addSectionLabel(slide, "PORTFOLIO READOUT", 0.50, 3.06, 1.7);
    addNarrative(slide, model.summaryNarrative, { x: 0.54, y: 3.42, w: 4.40, h: 1.62 });

    addSectionLabel(slide, "DERIVED STATUS", 5.70, 3.06, 1.7);
    const statuses = ["On Track", "At Risk", "Blocked"];
    const maximum = Math.max(1, ...statuses.map(status => Number(model.statusGroups[status] || 0)));
    statuses.forEach((status, index) => {
      const style = statusStyle(status);
      addBar(slide, {
        label: status,
        value: model.statusGroups[status] || 0,
        max: maximum,
        x: 5.72,
        y: 3.49 + index * 0.43,
        labelW: 0.85,
        barX: 6.76,
        barW: 2.05,
        valueX: 8.98,
        valueW: 0.24,
        color: style.text
      });
    });
    addCallout(slide, `${m.documentsInReview} in review / ${m.pendingTravel} pending travel`, {
      x: 5.66, y: 5.15, w: 3.55, h: 0.62
    });
  }

  function addMySummarySlide(prs, model, assets) {
    const slide = prs.addSlide();
    addContentChrome(slide, assets, {
      sectionLabel: "MY OVERVIEW",
      title: "Personal workload and activity",
      footerLeft: "PULSE MY OVERVIEW"
    });
    addSectionLabel(slide, "CURRENT WORKLOAD", 0.46, 1.70, 1.7);
    const m = model.metrics;
    const metrics = [
      { label: "Projects", value: m.totalProjects },
      { label: "Total Tasks", value: m.totalTasks },
      { label: "Open Tasks", value: m.activeTasks },
      { label: "Closed Tasks", value: m.completedTasks },
      { label: "Reviewed", value: m.documentsReviewed },
      { label: "Signed", value: m.documentsSigned },
      { label: "Travel", value: m.travelRequests }
    ];
    metrics.forEach((metric, index) => addMetric(slide, metric, 0.48 + index * 1.12, 2.02));

    addSectionLabel(slide, "EXECUTIVE READOUT", 0.46, 3.24, 1.7);
    addNarrative(slide, model.summaryNarrative, { x: 0.52, y: 3.60, w: 4.65, h: 1.56 });

    addSectionLabel(slide, "COMPLETION SIGNAL", 5.86, 3.24, 1.7);
    model.completionSignals.forEach((signal, index) => {
      addBar(slide, {
        label: signal.label,
        value: signal.percent,
        valueLabel: `${signal.percent}%`,
        max: 100,
        x: 5.88,
        y: 3.67 + index * 0.48,
        labelW: 1.55,
        barX: 7.58,
        barW: 1.30,
        valueX: 8.91,
        valueW: 0.34,
        color: signal.color
      });
    });
    addCallout(slide, `${m.activeTasks} open tasks / ${m.completedTasks} closed tasks`, {
      x: 5.84, y: 5.42, w: 3.45, h: 0.56
    });
  }

  function addEndItemSlide(prs, rows, assets, continuation) {
    const slide = prs.addSlide();
    addContentChrome(slide, assets, {
      sectionLabel: "TEAM OVERVIEW",
      title: continuation ? "Work grouped by end item — continued" : "Work grouped by end item",
      footerLeft: "PULSE TEAM OVERVIEW"
    });
    const tableRows = [
      headerRow(["End item", "Projects", "Active tasks", "Open risks", "Teams"], 7.0),
      ...rows.map((row, index) => {
        const fill = rowFill(index);
        return [
          tableCell(row.name, { fill, fontSize: 7.2, align: "center" }),
          tableCell(row.projectCount, { fill, color: C.ink, fontSize: 7.2, align: "center" }),
          tableCell(row.activeTaskCount, { fill, fontSize: 7.2, align: "center" }),
          tableCell(row.openRiskCount, { fill, fontSize: 7.2, align: "center" }),
          tableCell(row.teams.join(", ") || "-", { fill, fontSize: 7.2, align: "right" })
        ];
      })
    ];
    addTable(slide, tableRows, {
      x: 0.42,
      y: 1.63,
      w: 9.16,
      colW: weightedWidths(9.16, [2.4, 0.95, 1.10, 0.95, 3.76]),
      rowH: 0.30
    });
    addSectionLabel(slide, "TASK CONCENTRATION BY END ITEM", 0.50, 3.62, 2.5);
    const maximum = Math.max(1, ...rows.map(row => row.activeTaskCount));
    rows.slice(0, 5).forEach((row, index) => {
      addBar(slide, {
        label: row.name,
        value: row.activeTaskCount,
        max: maximum,
        x: 0.54,
        y: 4.07 + index * 0.46,
        labelW: 2.30,
        barX: 3.04,
        barW: 5.20,
        valueX: 8.48,
        valueW: 0.28,
        color: C.accent
      });
    });
  }

  function addTeamInventorySlide(prs, projects, assets, continuation) {
    const slide = prs.addSlide();
    addContentChrome(slide, assets, {
      sectionLabel: "TEAM OVERVIEW",
      title: continuation ? "Project inventory — continued" : "Project inventory",
      footerLeft: "PULSE TEAM OVERVIEW"
    });
    const rows = [
      headerRow(["Project", "End item", "PM", "Status", "Open", "Closed", "Risk", "Due"], 5.9),
      ...projects.map((project, index) => {
        const fill = rowFill(index);
        return [
          tableCell(project.name, { fill, color: C.ink, fontSize: 5.65 }),
          tableCell(project.endItem, { fill, fontSize: 5.65, align: "center" }),
          tableCell(project.pm || "-", { fill, fontSize: 5.65, align: "center" }),
          statusCell(project.derivedStatus, 5.65),
          tableCell(project.activeTasks, { fill, fontSize: 5.65, align: "center" }),
          tableCell(project.closedTasks, { fill, fontSize: 5.65, align: "center" }),
          tableCell(project.openRisks, { fill, fontSize: 5.65, align: "center" }),
          tableCell(isoDate(project.dueDate), { fill, fontSize: 5.65, align: "center" })
        ];
      })
    ];
    addTable(slide, rows, {
      x: 0.18,
      y: 1.50,
      w: 9.64,
      colW: weightedWidths(9.64, [2.70, 1.96, 0.92, 0.83, 0.62, 0.62, 0.52, 0.82]),
      rowH: 0.35
    });
    addText(slide, "Inventory continues onto additional pages before project-detail slides.", {
      x: 0.42, y: 6.03, w: 6.80, h: 0.22
    }, {
      fontSize: 8.2,
      color: C.muted
    });
  }

  function addMyProjectsSlide(prs, projects, assets, continuation) {
    const slide = prs.addSlide();
    addContentChrome(slide, assets, {
      sectionLabel: "MY OVERVIEW",
      title: continuation ? "Assigned projects — continued" : "Assigned projects",
      footerLeft: "PULSE MY OVERVIEW"
    });
    const rows = [
      headerRow(["Project", "Team", "Role", "Status", "My active", "Due"], 6.8),
      ...projects.map((project, index) => {
        const fill = rowFill(index);
        return [
          tableCell(project.name, { fill, color: C.ink, fontSize: 6.7 }),
          tableCell(project.team || "-", { fill, fontSize: 6.7, align: "center" }),
          tableCell(project.role || "Member", { fill, fontSize: 6.7, align: "center" }),
          statusCell(project.derivedStatus, 6.7),
          tableCell(project.activeTasks, { fill, fontSize: 6.7, align: "center" }),
          tableCell(isoDate(project.dueDate), { fill, fontSize: 6.7, align: "center" })
        ];
      })
    ];
    addTable(slide, rows, {
      x: 0.42,
      y: 1.70,
      w: 9.18,
      colW: weightedWidths(9.18, [2.87, 1.25, 1.45, 1.05, 0.75, 0.95]),
      rowH: 0.34
    });
    addSectionLabel(slide, "PROJECT SIGNAL", 0.48, 4.30, 1.7);
    const maximum = Math.max(1, ...projects.map(project => project.activeTasks));
    projects.slice(0, 5).forEach((project, index) => {
      const style = statusStyle(project.derivedStatus);
      addBar(slide, {
        label: project.name,
        value: project.activeTasks,
        max: maximum,
        x: 0.52,
        y: 4.72 + index * 0.36,
        labelW: 2.30,
        barX: 3.06,
        barW: 4.55,
        valueX: 7.85,
        valueW: 0.25,
        color: style.text
      });
    });
  }

  function addTaskDetailSlide(prs, tasks, assets, continuation) {
    const slide = prs.addSlide();
    addContentChrome(slide, assets, {
      sectionLabel: "MY OVERVIEW",
      title: continuation ? "Active task detail — continued" : "Active task detail",
      footerLeft: "PULSE MY OVERVIEW"
    });
    const rows = [
      headerRow(["ID", "Task", "Project", "Status", "Health", "Due", "%", "Next action"], 6.0),
      ...tasks.map((task, index) => {
        const fill = rowFill(index);
        const status = taskDisplayStatus(task);
        const health = taskHealth(task);
        const rawPercent = task.percentComplete ?? task.progress ?? (task.done ? 100 : 0);
        const pc = Number.isFinite(Number(rawPercent)) ? Math.max(0, Math.min(100, Number(rawPercent))) : 0;
        return [
          tableCell(task.id || "-", { fill, fontSize: 5.65, align: "center" }),
          tableCell(task.title || "Untitled task", { fill, color: C.ink, fontSize: 5.65 }),
          tableCell(task.projectName || task.projectId || "-", { fill, fontSize: 5.65 }),
          statusCell(status, 5.65),
          statusCell(health, 5.65),
          tableCell(isoDate(task.end), { fill, fontSize: 5.65, align: "center" }),
          tableCell(`${pc}%`, { fill, fontSize: 5.65, align: "center" }),
          tableCell(task.nextAction || task.notes || task.description || "-", { fill, fontSize: 5.65 })
        ];
      })
    ];
    addTable(slide, rows, {
      x: 0.20,
      y: 1.62,
      w: 9.60,
      colW: weightedWidths(9.60, [0.52, 1.82, 1.55, 0.82, 0.76, 0.72, 0.38, 1.93]),
      rowH: 0.39
    });
    addText(slide, "Task detail continues onto additional slides when the current page is full.", {
      x: 0.42, y: 6.14, w: 6.80, h: 0.22
    }, {
      fontSize: 8.3,
      color: C.muted
    });
  }

  function addProjectDetailSlide(prs, project, assets, index) {
    const slide = prs.addSlide();
    addContentChrome(slide, assets, {
      sectionLabel: "PROJECT DETAIL",
      title: project.name,
      footerLeft: `PULSE TEAM OVERVIEW | PROJECT DETAIL ${index}`
    });
    addText(slide, `${project.endItem}  /  PM: ${project.pm || "-"}`, {
      x: 0.42, y: 1.58, w: 6.80, h: 0.18
    }, {
      fontSize: 7.8,
      color: C.muted
    });

    const overviewRows = [
      headerRow(["Team", "Lifecycle", "Funding", "Task order", "Due", "Status"], 6.3),
      [
        tableCell(project.team || "-", { fontSize: 6.5, align: "center" }),
        tableCell(project.lifecycleStatus || "-", { color: C.ink, fontSize: 6.5 }),
        tableCell(project.fundingType || "-", { fontSize: 6.5, align: "center" }),
        tableCell(project.taskOrder || "-", { fontSize: 6.5, align: "center" }),
        tableCell(isoDate(project.dueDate), { fontSize: 6.5, align: "center" }),
        statusCell(project.derivedStatus, 6.5)
      ]
    ];
    addTable(slide, overviewRows, {
      x: 0.42,
      y: 2.02,
      w: 8.75,
      colW: weightedWidths(8.75, [1.25, 1.20, 1.40, 1.50, 1.05, 2.03]),
      rowH: 0.29
    });

    addSectionLabel(slide, "ACTIVE TASKS", 0.48, 2.95, 1.7);
    const taskRows = [
      headerRow(["Task", "Owner", "Status", "Due"], 6.2),
      ...(project.tasks.length ? project.tasks : [{ title: "No active tasks", assignee: "-", status: "-", end: "" }])
        .slice(0, T.pagination.projectTasks)
        .map((task, rowIndex) => {
          const fill = rowFill(rowIndex);
          return [
            tableCell(task.title || "Untitled task", { fill, fontSize: 6.2, align: "center" }),
            tableCell(task.assignee || task.owner || "-", { fill, color: C.ink, fontSize: 6.2 }),
            statusCell(taskDisplayStatus(task), 6.2),
            tableCell(isoDate(task.end), { fill, fontSize: 6.2, align: "center" })
          ];
        })
    ];
    addTable(slide, taskRows, {
      x: 0.42,
      y: 3.30,
      w: 4.72,
      colW: weightedWidths(4.72, [2.20, 0.76, 0.86, 0.90]),
      rowH: 0.32
    });

    addSectionLabel(slide, "OPEN RISKS", 5.48, 2.95, 1.7);
    const risks = project.risks.length ? project.risks : [{
      id: "-",
      title: "No open risks",
      likelihood: "-",
      impact: "-",
      dueDate: ""
    }];
    const riskRows = [
      headerRow(["ID", "Risk", "Likely", "Impact", "Due"], 6.0),
      ...risks.slice(0, T.pagination.projectRisks).map((risk, rowIndex) => {
        const fill = rowFill(rowIndex);
        return [
          tableCell(risk.id || "-", { fill, fontSize: 5.9, align: "center" }),
          tableCell(risk.title || risk.description || "Risk", { fill, color: C.ink, fontSize: 5.9 }),
          statusCell(risk.likelihood || "-", 5.9),
          statusCell(risk.impact || "-", 5.9),
          tableCell(isoDate(risk.dueDate), { fill, fontSize: 5.9, align: "center" })
        ];
      })
    ];
    addTable(slide, riskRows, {
      x: 5.44,
      y: 3.30,
      w: 4.05,
      colW: weightedWidths(4.05, [0.46, 1.62, 0.72, 0.72, 0.88]),
      rowH: 0.32
    });

    addSectionLabel(slide, "MITIGATION / NEXT ACTION", 0.48, 5.10, 2.5);
    slide.addShape("roundRect", {
      x: 0.42, y: 5.50, w: 8.92, h: 0.56,
      rectRadius: 0.05,
      line: { color: C.tableHead, transparency: 100 },
      fill: { color: C.tableHead }
    });
    const nextAction = project.risks.find(risk => risk.mitigation)?.mitigation
      || project.tasks.find(task => task.nextAction)?.nextAction
      || "Maintain cadence and close assigned tasks before the project due date.";
    addText(slide, safe(nextAction), { x: 0.74, y: 5.68, w: 8.30, h: 0.14 }, {
      fontSize: 7.5,
      color: C.ink
    });
  }

  function addTeamWorkloadSlide(prs, workload, assets, continuation) {
    const slide = prs.addSlide();
    addContentChrome(slide, assets, {
      sectionLabel: "TEAM OVERVIEW",
      title: continuation ? "Team workload — continued" : "Team workload",
      footerLeft: "PULSE TEAM OVERVIEW"
    });
    const rows = [
      headerRow(["Person", "Open", "Closed", "Projects", "Next due"], 6.8),
      ...workload.map((entry, index) => {
        const fill = rowFill(index);
        return [
          tableCell(entry.person, { fill, fontSize: 6.8 }),
          tableCell(entry.activeTasks, { fill, color: C.ink, fontSize: 6.8, align: "center" }),
          tableCell(entry.closedTasks, { fill, fontSize: 6.8, align: "center" }),
          tableCell(entry.projectCount, { fill, fontSize: 6.8, align: "center" }),
          tableCell(isoDate(entry.nextDue), { fill, fontSize: 6.8, align: "center" })
        ];
      })
    ];
    addTable(slide, rows, {
      x: 0.42,
      y: 1.62,
      w: 5.20,
      colW: weightedWidths(5.20, [1.75, 0.70, 0.75, 0.78, 1.22]),
      rowH: 0.34
    });
    addSectionLabel(slide, "LOAD INDICATORS", 6.18, 1.72, 1.7);
    const topWorkload = [...workload].sort((a, b) => b.activeTasks - a.activeTasks).slice(0, 7);
    slide.addChart(prs.ChartType.bar, [{
      name: "Open Tasks",
      labels: topWorkload.map(entry => entry.person),
      values: topWorkload.map(entry => entry.activeTasks)
    }], {
      x: 6.10, y: 1.95, w: 3.65, h: 2.55,
      barDir: "bar",
      barGapWidthPct: 40,
      chartColors: [C.accent],
      showLegend: false,
      showValue: true,
      dataLabelColor: C.ink,
      dataLabelFontSize: 7,
      catAxisLabelColor: C.muted,
      catAxisLabelFontSize: 7,
      valAxisHidden: true,
      showTitle: false
    });

    const tiers = { Heavy: 0, Moderate: 0, Light: 0 };
    workload.forEach(entry => {
      if (entry.activeTasks >= 5) tiers.Heavy++;
      else if (entry.activeTasks >= 2) tiers.Moderate++;
      else tiers.Light++;
    });
    addSectionLabel(slide, "WORKLOAD DISTRIBUTION", 0.42, 4.55, 2.6);
    slide.addChart(prs.ChartType.doughnut, [{
      name: "People",
      labels: ["Heavy (5+)", "Moderate (2-4)", "Light (0-1)"],
      values: [tiers.Heavy, tiers.Moderate, tiers.Light]
    }], {
      x: 0.42, y: 4.78, w: 2.9, h: 1.95,
      chartColors: [C.red, C.amber, C.green],
      showLegend: true,
      legendPos: "r",
      legendFontSize: 7,
      showTitle: false,
      dataLabelColor: C.white,
      dataLabelFontSize: 7
    });

    addCallout(slide, "Review open workload distribution before assigning new work.", {
      x: 3.55, y: 5.05, w: 2.55, h: 0.9
    });
  }

  function allReviewRecords(db) {
    if (typeof getAllDocReviewRecords === "function") {
      try { return getAllDocReviewRecords() || []; } catch (_) { /* fall through */ }
    }
    return db.docReviews || db.documentReviews || [];
  }

  function memberIdentity(db) {
    const user = db.user || {};
    return {
      name: String(user.name || "").trim().toLowerCase(),
      email: String(user.email || "").trim().toLowerCase()
    };
  }

  function assigneeIsCurrentUser(db, assignee) {
    if (!assignee) return false;
    const identity = memberIdentity(db);
    const value = String(assignee).trim().toLowerCase();
    if (value === identity.name || value === identity.email) return true;
    if (typeof currentUserMember === "function" && typeof memberMatchesAssignee === "function") {
      try {
        const member = currentUserMember();
        if (member && memberMatchesAssignee(member, assignee)) return true;
      } catch (_) { /* fall through */ }
    }
    return false;
  }

  function projectRoleForUser(db, projectId) {
    const roster = db.projectPeople && db.projectPeople[projectId] || [];
    const entry = roster.find(person => assigneeIsCurrentUser(
      db,
      person.email || person.label || person.name || person.displayName
    ));
    return entry ? safe(entry.role || entry.type || "Member") : "Member";
  }

  function projectManager(db, projectId) {
    const roster = db.projectPeople && db.projectPeople[projectId] || [];
    const entry = roster.find(person => {
      const role = String(person.role || person.type || "").toLowerCase();
      return role.includes("project manager") || role === "pm" || role.includes("program manager");
    });
    return entry ? safe(entry.label || entry.name || entry.displayName || entry.email || "") : "";
  }

  function plainTasks(db, projectId) {
    const raw = db.ganttTasks && db.ganttTasks[projectId] || [];
    if (typeof trackerPlainTasks === "function") {
      try { return trackerPlainTasks(raw) || []; } catch (_) { /* fall through */ }
    }
    return raw.filter(task => String(task.itemType || "").toLowerCase() !== "divider");
  }

  function openProjectRisks(db, projectId) {
    const extra = db.projectExtra && db.projectExtra[projectId] || {};
    return (extra.risks || []).filter(riskIsOpen);
  }

  function buildTeamModel(db, visibleProjects) {
    const projects = Array.isArray(visibleProjects) ? visibleProjects : (db.projects || []);
    let totalTasks = 0;
    let activeTasks = 0;
    let closedTasks = 0;
    let openRisks = 0;
    const statusGroups = { "On Track": 0, "At Risk": 0, "Blocked": 0 };
    const blockedProjectNames = [];

    const projectRows = projects.map(project => {
      const tasks = plainTasks(db, project.id);
      const active = tasks.filter(task => !taskIsDone(task));
      const closed = tasks.filter(taskIsDone);
      const risks = openProjectRisks(db, project.id);
      const derivedStatus = presentationProjectStatus(project, tasks);
      if (statusGroups[derivedStatus] != null) statusGroups[derivedStatus]++;
      if (derivedStatus === "Blocked") blockedProjectNames.push(project.name || project.id);
      totalTasks += tasks.length;
      activeTasks += active.length;
      closedTasks += closed.length;
      openRisks += risks.length;
      return {
        id: project.id || "-",
        name: project.name || project.id || "Untitled project",
        endItem: project.configEndItem || project.endItem || project.portfolios && project.portfolios[0] || "Unclassified",
        pm: projectManager(db, project.id),
        team: project.team || "",
        fundingType: project.fundingType || "",
        fiscalYear: project.fiscalYear || "",
        taskOrder: project.taskOrder || "",
        lifecycleStatus: project.lifecycleStatus || project.status || "",
        dueDate: project.dueDate || "",
        derivedStatus,
        totalTasks: tasks.length,
        activeTasks: active.length,
        closedTasks: closed.length,
        openRisks: risks.length,
        tasks: active.map(task => ({
          ...task,
          id: task.id || task.taskId || "-",
          title: task.title || task.text || "Untitled task",
          assignee: task.assignee || task.owner || "",
          nextAction: task.nextAction || task.notes || ""
        })),
        closedTaskList: closed.map(task => ({
          ...task,
          assignee: task.assignee || task.owner || ""
        })),
        risks: risks.map(risk => ({
          ...risk,
          id: risk.id || risk.riskId || "-",
          title: risk.title || risk.description || "Risk",
          dueDate: risk.dueDate || risk.end || "",
          mitigation: risk.mitigation || risk.mitigationPlan || risk.nextAction || ""
        }))
      };
    });

    const people = {};
    function ensurePerson(person) {
      if (!people[person]) {
        people[person] = { person, activeTasks: 0, closedTasks: 0, projects: new Set(), nextDue: "" };
      }
      return people[person];
    }
    projectRows.forEach(project => {
      project.tasks.forEach(task => {
        const person = safe(task.assignee || "");
        if (!person) return;
        const entry = ensurePerson(person);
        entry.activeTasks++;
        entry.projects.add(project.id);
        const due = isoDate(task.end);
        if (due !== "-" && (!entry.nextDue || due < entry.nextDue)) entry.nextDue = due;
      });
      project.closedTaskList.forEach(task => {
        const person = safe(task.assignee || "");
        if (!person) return;
        const entry = ensurePerson(person);
        entry.closedTasks++;
        entry.projects.add(project.id);
      });
    });
    const workload = Object.values(people).map(entry => ({
      person: entry.person,
      activeTasks: entry.activeTasks,
      closedTasks: entry.closedTasks,
      projectCount: entry.projects.size,
      nextDue: entry.nextDue
    })).sort((a, b) => b.activeTasks - a.activeTasks || a.person.localeCompare(b.person));

    const groups = {};
    projectRows.forEach(project => {
      const name = project.endItem || "Unclassified";
      if (!groups[name]) groups[name] = {
        name,
        projectCount: 0,
        activeTaskCount: 0,
        openRiskCount: 0,
        teams: new Set()
      };
      groups[name].projectCount++;
      groups[name].activeTaskCount += project.activeTasks;
      groups[name].openRiskCount += project.openRisks;
      if (project.team) groups[name].teams.add(project.team);
    });
    const endItems = Object.values(groups).map(group => ({
      ...group,
      teams: Array.from(group.teams)
    })).sort((a, b) => b.activeTaskCount - a.activeTaskCount || a.name.localeCompare(b.name));

    const docs = allReviewRecords(db);
    const documentsInReview = docs.filter(doc =>
      !doc.isArchived && !["Review Complete", "Signed", "Archived"].includes(doc._column || "")
    ).length;
    const pendingTravel = (db.travelRequests || []).filter(request =>
      request.status === "Pending" || request.status === "Submitted"
    ).length;
    const activeProjects = projectRows.filter(project =>
      ["On Track", "At Risk", "Blocked"].includes(project.derivedStatus)
    ).length;

    const teamNames = projectRows.map(project => project.name).filter(Boolean);
    const activeSummary = teamNames.length
      ? `Portfolio work is active across ${activeProjects} of ${projectRows.length} projects.`
      : "No projects are visible in the current scope.";
    const blockedSummary = blockedProjectNames.length
      ? `Blocked work is concentrated in ${blockedProjectNames.slice(0, 3).join(", ")}.`
      : "No projects are currently blocked.";
    return {
      isMyOverview: false,
      title: "PULSE Team Overview",
      generatedBy: safe(db.user && db.user.name || ""),
      reportingAsOf: isoDate(new Date().toISOString()),
      generatedAt: generatedTimestamp(),
      metrics: {
        totalProjects: projectRows.length,
        activeProjects,
        totalTasks,
        activeTasks,
        closedTasks,
        openRisks,
        teamSize: workload.length,
        documentsInReview,
        pendingTravel
      },
      summaryNarrative: [
        activeSummary,
        blockedSummary,
        `Document review has ${documentsInReview} active item${documentsInReview === 1 ? "" : "s"}; ${pendingTravel} travel request${pendingTravel === 1 ? " is" : "s are"} pending.`,
        `${closedTasks} of ${totalTasks} total task${totalTasks === 1 ? "" : "s"} are closed; ${openRisks} open risk${openRisks === 1 ? "" : "s"} require routine monitoring.`
      ],
      statusGroups,
      endItems,
      projects: projectRows,
      workload
    };
  }

  function buildMyModel(db) {
    const user = db.user || {};
    const now = new Date();
    const allMyTasks = [];
    const projects = [];

    (db.projects || []).forEach(project => {
      const tasks = plainTasks(db, project.id);
      const mine = tasks.filter(task => assigneeIsCurrentUser(db, task.assignee || task.owner));
      const roster = db.projectPeople && db.projectPeople[project.id] || [];
      const hasRosterRole = roster.some(person => assigneeIsCurrentUser(
        db,
        person.email || person.label || person.name || person.displayName
      ));
      if (!mine.length && !hasRosterRole) return;
      const active = mine.filter(task => !taskIsDone(task));
      mine.forEach(task => allMyTasks.push({
        ...task,
        id: task.id || task.taskId || "-",
        title: task.title || task.text || "Untitled task",
        projectId: project.id,
        projectName: project.name || project.id,
        nextAction: task.nextAction || task.notes || task.description || ""
      }));
      projects.push({
        id: project.id || "-",
        name: project.name || project.id || "Untitled project",
        team: project.team || "",
        role: projectRoleForUser(db, project.id),
        dueDate: project.dueDate || "",
        derivedStatus: presentationProjectStatus(project, tasks),
        activeTasks: active.length
      });
    });

    const active = allMyTasks.filter(task => !taskIsDone(task));
    const completed = allMyTasks.filter(taskIsDone);
    const docs = allReviewRecords(db);
    const identity = memberIdentity(db);
    const isMe = value => {
      const normalized = String(value || "").trim().toLowerCase();
      return !!normalized && (normalized === identity.name || normalized === identity.email);
    };
    const year = String(now.getFullYear());
    const reviewed = docs.filter(doc =>
      Array.isArray(doc.reviewers) && doc.reviewers.some(reviewer =>
        isMe(reviewer.email || reviewer.name || reviewer.label) &&
        String(reviewer.signedAt || reviewer.decidedAt || doc.updatedAt || doc.createdAt || "").slice(0, 4) === year
      )
    );
    const signed = docs.filter(doc =>
      Array.isArray(doc.reviewers) && doc.reviewers.some(reviewer =>
        reviewer.isSigner && reviewer.signedAt &&
        isMe(reviewer.email || reviewer.name || reviewer.label) &&
        String(reviewer.signedAt).slice(0, 4) === year
      )
    );
    const pendingReviews = docs.filter(doc =>
      Array.isArray(doc.reviewers) && doc.reviewers.some(reviewer =>
        isMe(reviewer.email || reviewer.name || reviewer.label) && !reviewer.signedAt && !reviewer.decidedAt
      )
    ).length;
    const travel = (db.travelRequests || []).filter(request => {
      const requester = String(request.requester || "").trim().toLowerCase();
      const email = String(request.requesterEmail || "").trim().toLowerCase();
      const mine = requester === identity.name || email === identity.email;
      return mine && String(request.createdAt || request.startDate || "").slice(0, 4) === year;
    });
    const nextDue = active
      .filter(task => task.end)
      .sort((a, b) => String(a.end).localeCompare(String(b.end)))[0];
    const projectNames = projects.slice(0, 3).map(project => project.name);
    return {
      isMyOverview: true,
      title: `PULSE My Overview${user.name ? ` - ${safe(user.name)}` : ""}`,
      generatedBy: safe(user.name || ""),
      reportingAsOf: isoDate(now.toISOString()),
      generatedAt: generatedTimestamp(now),
      metrics: {
        totalProjects: projects.length,
        totalTasks: allMyTasks.length,
        activeTasks: active.length,
        completedTasks: completed.length,
        documentsReviewed: reviewed.length,
        documentsSigned: signed.length,
        travelRequests: travel.length
      },
      summaryNarrative: [
        projectNames.length
          ? `Active work spans ${projectNames.join(", ")}${projects.length > 3 ? ", and additional assigned projects" : ""}.`
          : "No projects are currently assigned.",
        active.length
          ? `${active.length} assigned task${active.length === 1 ? " is" : "s are"} still open.`
          : "No assigned tasks are open.",
        nextDue ? `The next assigned task is due ${isoDate(nextDue.end)}.` : "No active task due dates are recorded.",
        `${completed.length} task${completed.length === 1 ? " is" : "s are"} closed and ${pendingReviews} document review${pendingReviews === 1 ? " is" : "s are"} awaiting action.`
      ],
      completionSignals: [
        { label: "Assigned tasks complete", percent: percent(completed.length, allMyTasks.length), color: C.green },
        { label: "Assigned tasks open", percent: percent(active.length, allMyTasks.length), color: C.accent },
        { label: "Documents reviewed", percent: percent(reviewed.length, reviewed.length + pendingReviews), color: C.amber }
      ],
      projects,
      activeTasks: active.sort((a, b) => String(a.end || "9999").localeCompare(String(b.end || "9999")))
    };
  }

  async function buildPresentation(isTeam, db, visibleProjects) {
    const PptxGen = pptxCtor();
    const prs = new PptxGen();
    prs.layout = T.slide.layout;
    prs.author = safe(db.user && db.user.name || "PULSE");
    prs.company = "PULSE";
    prs.subject = "PULSE Overview";
    prs.title = isTeam ? "PULSE Team Overview" : "PULSE My Overview";
    prs.lang = "en-US";
    const assets = await loadTemplateAssets();
    const model = isTeam ? buildTeamModel(db, visibleProjects) : buildMyModel(db);
    addCoverSlide(prs, model, assets);

    if (model.isMyOverview) {
      addMySummarySlide(prs, model, assets);
      chunk(model.projects, T.pagination.myProjectsPerSlide)
        .forEach((group, index) => addMyProjectsSlide(prs, group, assets, index > 0));
      chunk(model.activeTasks, T.pagination.tasksPerSlide)
        .forEach((group, index) => addTaskDetailSlide(prs, group, assets, index > 0));
    } else {
      addTeamSummarySlide(prs, model, assets);
      chunk(model.endItems, T.pagination.endItemsPerSlide)
        .forEach((group, index) => addEndItemSlide(prs, group, assets, index > 0));
      chunk(model.projects, T.pagination.teamProjectsPerSlide)
        .forEach((group, index) => addTeamInventorySlide(prs, group, assets, index > 0));
      model.projects.forEach((project, index) => addProjectDetailSlide(prs, project, assets, index + 1));
      chunk(model.workload, T.pagination.workloadPerSlide)
        .forEach((group, index) => addTeamWorkloadSlide(prs, group, assets, index > 0));
    }
    return { prs, model };
  }

  function closeReservedOpen(popup) {
    const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
    if (api && typeof api.closeReservedSharePointFileWindow === "function") {
      api.closeReservedSharePointFileWindow(popup);
    } else if (popup && !popup.closed) {
      popup.close();
    }
  }

  function openStoredOverviewFile(fileUrl, fileName, popup) {
    const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
    if (api && typeof api.openSharePointFileByPolicy === "function") {
      return api.openSharePointFileByPolicy(
        fileUrl,
        fileName,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        { popup }
      );
    }
    if (popup && !popup.closed) {
      popup.location.replace(fileUrl);
      return true;
    }
    window.open(fileUrl, "_blank", "noopener");
    return true;
  }

  async function generateOverviewPptx(isTeam, db, visibleProjects, options) {
    const opts = options || {};
    const { prs, model } = await buildPresentation(isTeam, db, visibleProjects);
    const date = new Date().toISOString().slice(0, 10);
    const fileName = model.isMyOverview
      ? `PULSE-My-Overview-${date}.pptx`
      : `PULSE-Team-Overview-${date}.pptx`;
    const blob = await prs.write({ outputType: "blob" });

    const canUpload = typeof isSharePointMode === "function"
      && isSharePointMode()
      && typeof sharePointAdapter !== "undefined"
      && typeof sharePointAdapter.resolvePulseDocumentsLibraryUrl === "function"
      && typeof sharePointAdapter.createPulseDocumentsSubfolder === "function"
      && typeof sharePointAdapter.uploadPulseDocumentsFile === "function";

    if (!canUpload) {
      closeReservedOpen(opts.popup);
      downloadOverviewPptxBlob(blob, fileName);
      return { mode: "download", fileName, blob };
    }

    try {
      const uploaded = await uploadOverviewBriefingToSharePoint(blob, fileName);
      const fileUrl = uploaded.fileUrl
        || (uploaded.serverRelativeUrl ? window.location.origin + uploaded.serverRelativeUrl : "");
      if (fileUrl) openStoredOverviewFile(fileUrl, uploaded.fileName || fileName, opts.popup);
      else closeReservedOpen(opts.popup);
      return {
        mode: "sharepoint",
        fileName: uploaded.fileName || fileName,
        blob,
        fileUrl
      };
    } catch (uploadError) {
      console.warn("PULSE: overview PPTX upload failed; downloading a local copy.", uploadError);
      closeReservedOpen(opts.popup);
      downloadOverviewPptxBlob(blob, fileName);
      return { mode: "download", fileName, blob, uploadError };
    }
  }

  function downloadOverviewPptxBlob(blob, fileName) {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName || "PULSE-Overview.pptx";
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 2000);
  }

  async function uploadOverviewBriefingToSharePoint(blob, fileName) {
    const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
    if (!siteUrl) throw new Error("No SharePoint site URL is available.");
    const libraryUrl = await sharePointAdapter.resolvePulseDocumentsLibraryUrl(siteUrl);
    const reportsRoot = await sharePointAdapter.createPulseDocumentsSubfolder(siteUrl, libraryUrl, "Reporting");
    const overviewFolder = await sharePointAdapter.createPulseDocumentsSubfolder(
      siteUrl,
      reportsRoot.serverRelativeUrl,
      "Overview"
    );
    const file = new File([blob], fileName, {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    });
    return sharePointAdapter.uploadPulseDocumentsFile(siteUrl, overviewFolder.serverRelativeUrl, file);
  }

  window.AEWTTR.OverviewPptxService = {
    generate: generateOverviewPptx,
    build: buildPresentation,
    buildModel: (isTeam, db, projects) => isTeam ? buildTeamModel(db, projects) : buildMyModel(db)
  };
}());
