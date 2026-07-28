/* Project status + risk assessment PPTX export (PptxGenJS).
   Layout fidelity targets the AEWTTR IPT June-2026 status deck (Slides.zip):
   4:3 (10×7.5), master4 seals/CUI, navy cross dividers, POC (6pt), left photo
   frame with aspect-preserving letterbox, Project Description, right Technical
   Status (#4E95D9) + Risk Category table + Milestones strip. Risk appendix follows.
   Geometry constants derived from slide/master EMU positions (÷ 914400). */
(function () {
  window.AEWTTR = window.AEWTTR || {};

  const CUI = "CONTROLLED UNCLASSIFIED INFORMATION";
  const NAVY = "0B3D91";
  const NAVY_DARK = "072A66";
  const TITLE_NAVY = "19264F";
  const DIVIDER = "002060";
  const GREEN = "00B050";
  const AMBER = "FFFF00";
  const RED = "FF0000";
  const GRAY = "6B7280";
  const CUI_GRAY = "808080";
  const LIGHT_BLUE = "4E95D9";
  const ROW_ALT = "EEF5FB";
  const TECH_FILL = "4E95D9";
  const PHOTO_FRAME_BG = "F5F6F8";
  /* Official status deck is 4:3 screen, not widescreen (presentation.xml sldSz). */
  const SLIDE_W = 10;
  const SLIDE_H = 7.5;

  /* Geometry measured from Slides.zip slideMaster4 + representative status slides
     (EMU ÷ 914400). Status slides use layout7 → master4. */
  const L = {
    ttsd: { x: 0.053, y: 0.053, w: 0.966, h: 0.967 },
    aewttr: { x: 8.914, y: 0.043, w: 1.041, h: 0.984 },
    navair: { x: 0.025, y: 7.232, w: 1.207, h: 0.215 },
    cuiTop: { x: 2.902, y: 0.02, w: 4.195, h: 0.2 },
    cuiBottom: { x: 2.902, y: 7.232, w: 4.195, h: 0.2 },
    pageNum: { x: 9.65, y: 7.19, w: 0.32, h: 0.24 },
    footerLineY: 7.173,
    title: { x: 0.95, y: 0.15, w: 6.05, h: 0.65 },
    poc: { x: 7.145, y: 0.01, w: 1.72, h: 0.5 },
    vDiv: { x: 4.731, y: 0.942, h: 6.05 },
    hDiv: { x: 0.083, y: 4.053, w: 9.65 },
    photo: { x: 0.28, y: 1.12, w: 4.2, h: 2.7 },
    techHead: { x: 4.91, y: 0.78, w: 5.0, h: 0.28 },
    techPanel: { x: 4.91, y: 1.08, w: 5.0, h: 1.15 },
    riskTable: { x: 4.91, y: 2.28, w: 5.0 },
    riskColW: [1.293, 0.839, 2.868],
    /* descHead / milesHead must sit BELOW the hDiv (y=4.053) to avoid the
       divider line cutting through the heading text. */
    descHead: { x: 0.15, y: 4.07, w: 4.4, h: 0.26 },
    milesHead: { x: 4.85, y: 4.07, w: 5.0, h: 0.26 },
    descBody: { x: 0.15, y: 4.38, w: 4.4, h: 2.65 },
    milesBody: { x: 4.85, y: 4.38, w: 5.0, h: 2.65 },
    asOf: { x: 6.45, y: 7.175, w: 1.45, h: 0.28 }
  };

  function pptxCtor() {
    const C = window.PptxGenJS || window.pptxgenjs || window.pptxgen;
    if (!C) throw new Error("PptxGenJS is not loaded");
    return typeof C === "function" ? C : C.default;
  }

  function todayLabel() {
    const d = new Date();
    return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
  }

  function shortDate(iso) {
    if (!iso) return "";
    const d = new Date(String(iso).slice(0, 10) + "T00:00:00");
    if (isNaN(d)) return String(iso).slice(0, 10);
    return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "2-digit" });
  }

  function bulletize(text, max) {
    const raw = String(text || "").trim();
    if (!raw) return [];
    const lines = raw.split(/\n+/).map((l) => l.replace(/^[\s•\-\*]+/, "").trim()).filter(Boolean);
    if (lines.length > 1) return lines.slice(0, max || 8);
    return raw
      .split(/(?<=[.!?])\s+/)
      .map((s) => s.trim())
      .filter(Boolean)
      .slice(0, max || 8);
  }

  /* Rough wrap estimate (Arial ≈ 0.5em average char width at 1pt ≈ 1/72").
     Returns { lines, fontSize, texts } that fit box height without mid-line clip. */
  function fitBulletsToBox(bullets, boxW, boxH, opts) {
    const options = opts || {};
    const maxFont = options.maxFont || 11;
    const minFont = options.minFont || 7;
    const paraGap = options.paraGap || 2;
    const pad = options.pad || 0.04;
    const usableH = Math.max(0.2, boxH - pad * 2);
    const usableW = Math.max(0.5, boxW - pad * 2);
    const source = (bullets || []).map((t) => String(t || "").trim()).filter(Boolean);
    if (!source.length) {
      return { texts: ["—"], fontSize: Math.min(maxFont, 9), paraGap };
    }

    function linesFor(fontPt, items) {
      const avgCharWidth = (fontPt * 0.48) / 72;
      const charsPerLine = Math.max(8, Math.floor(usableW / avgCharWidth));
      let total = 0;
      items.forEach((text) => {
        const len = Math.max(1, text.length);
        total += Math.max(1, Math.ceil(len / charsPerLine));
      });
      return total;
    }

    function heightFor(fontPt, items) {
      const lineH = (fontPt / 72) * 1.3;
      const gap = (paraGap / 72) * Math.max(0, items.length - 1);
      return linesFor(fontPt, items) * lineH + gap;
    }

    let items = source.slice();
    let fontSize = maxFont;
    for (let size = maxFont; size >= minFont; size -= 0.5) {
      if (heightFor(size, items) <= usableH) {
        fontSize = size;
        return { texts: items, fontSize, paraGap };
      }
      fontSize = size;
    }

    /* Still overflow at min font: drop bullets from the end, then ellipsis-trim. */
    while (items.length > 1 && heightFor(minFont, items) > usableH) {
      items = items.slice(0, -1);
    }
    fontSize = minFont;
    if (heightFor(fontSize, items) > usableH && items.length) {
      let text = items[items.length - 1];
      let guard = 0;
      while (
        text.length > 12
        && heightFor(fontSize, items.slice(0, -1).concat([`${text}…`])) > usableH
        && guard < 40
      ) {
        text = text.slice(0, Math.floor(text.length * 0.85)).replace(/\s+\S*$/, "").trim();
        guard += 1;
      }
      items = items.slice(0, -1).concat([text.endsWith("…") ? text : `${text}…`]);
    }
    return { texts: items, fontSize, paraGap };
  }

  function fitPlainTextToBox(text, boxW, boxH, opts) {
    const raw = String(text || "").trim() || "—";
    const fitted = fitBulletsToBox([raw], boxW, boxH, opts);
    return { text: fitted.texts[0] || "—", fontSize: fitted.fontSize };
  }

  function ragColor(value) {
    const v = String(value || "").toLowerCase();
    if (v === "red" || v === "high" || v === "off track") return RED;
    if (v === "amber" || v === "yellow" || v === "at risk" || v === "medium") return AMBER;
    return GREEN;
  }

  function ragFontColor(fill) {
    return fill === AMBER ? "111111" : "FFFFFF";
  }

  function ratingBand(likelihood, impact) {
    const score = Number(likelihood || 0) * Number(impact || 0);
    if (score >= 15) return "H";
    if (score >= 6) return "M";
    return "L";
  }

  function heatCellColor(likelihood, consequence) {
    const score = likelihood * consequence;
    if (score >= 15 || (likelihood >= 4 && consequence >= 4)) return RED;
    if (score >= 6) return AMBER;
    return GREEN;
  }

  async function fetchAsDataUri(url) {
    if (!url) return "";
    if (String(url).startsWith("data:")) return url;
    try {
      const res = await fetch(url);
      if (!res.ok) return "";
      const blob = await res.blob();
      return await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result || "");
        reader.onerror = () => resolve("");
        reader.readAsDataURL(blob);
      });
    } catch (_) {
      return "";
    }
  }

  async function svgToPngDataUri(svgText, size) {
    const sizePx = size || 160;
    const blob = new Blob([svgText], { type: "image/svg+xml" });
    const url = URL.createObjectURL(blob);
    try {
      const img = await new Promise((resolve, reject) => {
        const el = new Image();
        el.onload = () => resolve(el);
        el.onerror = reject;
        el.src = url;
      });
      const canvas = document.createElement("canvas");
      canvas.width = sizePx;
      canvas.height = sizePx;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, sizePx, sizePx);
      return canvas.toDataURL("image/png");
    } catch (_) {
      return "";
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  async function loadBrandLogos() {
    /* Prefer seals extracted from the official Slides.zip master. */
    const ttsd = await fetchAsDataUri("assets/images/ttsd-seal-template.png")
      || await fetchAsDataUri("assets/images/ttsd-seal.svg");
    const aewttr = await fetchAsDataUri("assets/images/aewttr-seal-template.png")
      || await fetchAsDataUri("assets/images/aewttr-seal.png");
    const navairMark = await fetchAsDataUri("assets/images/navair-wordmark-template.png");
    let navair = await fetchAsDataUri("assets/images/navair-seal-template.png");

    async function dataUriToSvgText(dataUri) {
      if (!dataUri) return "";
      if (dataUri.startsWith("data:image/svg+xml;base64,")) {
        try { return atob(dataUri.split(",")[1]); } catch (_) { return ""; }
      }
      if (dataUri.startsWith("data:image/svg+xml,")) {
        return decodeURIComponent(dataUri.slice("data:image/svg+xml,".length));
      }
      return "";
    }

    let ttsdOut = ttsd;
    let aewttrOut = aewttr;
    let navairOut = navair;
    if (ttsdOut && ttsdOut.indexOf("image/svg") !== -1) {
      const text = await dataUriToSvgText(ttsdOut);
      ttsdOut = text ? await svgToPngDataUri(text, 180) : "";
    }
    if (!ttsdOut) {
      ttsdOut = await svgToPngDataUri(
        `<svg xmlns="http://www.w3.org/2000/svg" width="160" height="160"><circle cx="80" cy="80" r="76" fill="#C41E3A"/><text x="80" y="88" text-anchor="middle" fill="#fff" font-family="Arial" font-size="28" font-weight="700">TTSD</text></svg>`,
        180
      );
    }
    if (!aewttrOut) aewttrOut = ttsdOut;
    if (navairOut && navairOut.indexOf("image/svg") !== -1) {
      const text = await dataUriToSvgText(navairOut);
      navairOut = text ? await svgToPngDataUri(text, 180) : "";
    }
    if (!navairOut) {
      navairOut = await fetchAsDataUri("assets/images/navair-seal.svg");
      if (navairOut && navairOut.indexOf("image/svg") !== -1) {
        const text = await dataUriToSvgText(navairOut);
        navairOut = text ? await svgToPngDataUri(text, 180) : "";
      }
    }
    return { ttsd: ttsdOut, aewttr: aewttrOut, navair: navairOut, navairMark };
  }

  function resolvePoc(proj) {
    const db = window.AEWTTR.db || {};
    const roster = (db.projectPeople && db.projectPeople[proj.id]) || [];
    const pmId = proj.pm;
    let entry = pmId ? roster.find((p) => p.id === pmId) : null;
    if (!entry) {
      entry = roster.find((p) => /pm|product manager|project manager/i.test(p.role || "")) || roster[0] || null;
    }
    if (!entry) {
      return {
        name: proj.pm || "",
        org: "AEWTTR IPT",
        phone: "",
        email: ""
      };
    }
    const member = entry.memberId ? (db.members || []).find((m) => m.id === entry.memberId) : null;
    return {
      name: entry.label || (member && member.name) || "",
      org: "AEWTTR IPT",
      phone: (member && (member.phone || member.workPhone)) || entry.phone || "",
      email: entry.email || (member && member.email) || ""
    };
  }

  function collectExportImages(proj) {
    const selected = (proj.exportSelections && proj.exportSelections.photos) || [];
    const cover = proj.coverImage ? [proj.coverImage] : [];
    const listed = Array.isArray(proj.images)
      ? proj.images.map((img) => (typeof img === "string" ? img : img.url || img.fileUrl || "")).filter(Boolean)
      : [];
    const urls = [];
    const seen = new Set();
    (selected.length ? selected : cover.concat(listed)).forEach((url) => {
      if (!url || seen.has(url)) return;
      seen.add(url);
      urls.push(url);
    });
    return urls.slice(0, 6);
  }

  function collectMilestones(proj) {
    const db = window.AEWTTR.db || {};
    const sharedMilestones = typeof getProjectReportingMilestones === "function"
      ? getProjectReportingMilestones(proj)
      : (((db.projectExtra || {})[proj.id] || {}).reportConfig || {}).taskMilestones || {};
    const tasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
    const plainTasks = typeof trackerPlainTasks === "function" ? trackerPlainTasks(tasks) : tasks;
    const typed = plainTasks
      .filter((task) => task && sharedMilestones[task.id] && sharedMilestones[task.id].inReport)
      .map((task) => {
        const values = Object.values(sharedMilestones[task.id] || {}).filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))).sort();
        if (!values.length) return null;
        return {
          title: task.title || task.name || "Milestone",
          start: values[0] || "",
          end: values[values.length - 1] || values[0] || "",
          done: /complete|closed|done/i.test(String(task.lifecycleStatus || task.status || "")),
          group: task.program || task.product || ""
        };
      }).filter(Boolean);
    if (typed.length) return typed;

    const dividers = typeof trackerDividers === "function"
      ? trackerDividers(tasks)
      : tasks.filter((t) => t && (t.itemType === "divider" || t.workItemLevel === "Divider"));
    return dividers
      .filter((d) => d && (d.isMilestone || (d.metadata && d.metadata.isMilestone)))
      .map((d) => ({
        title: d.title || "Milestone",
        start: d.startDate || d.start || "",
        end: d.dueDate || d.end || d.completionDate || d.startDate || d.start || "",
        done: /complete|closed|done/i.test(String(d.lifecycleStatus || d.status || "")),
        group: d.program || d.product || ""
      }));
  }

  function technicalStatusBullets(proj) {
    const db = window.AEWTTR.db || {};
    const extra = (db.projectExtra && db.projectExtra[proj.id]) || {};
    const configured = extra.reportConfig && extra.reportConfig.techBullets;
    /* The Reporting tab is the source of truth when a user has entered
       technical-status bullets.  Keep the exported slide in lockstep with
       the live preview instead of silently substituting generated content. */
    if (Array.isArray(configured) && configured.some((bullet) => String(bullet || "").trim())) {
      return configured.map((bullet) => String(bullet || "").trim()).filter(Boolean).slice(0, 8);
    }
    const bullets = [];
    bulletize(extra.handoff || "", 4).forEach((b) => bullets.push(b));
    bulletize(extra.meetingNotes || "", 3).forEach((b) => bullets.push(b));
    const tasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
    const plain = typeof trackerPlainTasks === "function" ? trackerPlainTasks(tasks) : tasks;
    const open = plain.filter((t) => t && t.itemType !== "divider" && t.status !== "Complete" && t.status !== "Done");
    const atRisk = open.filter((t) => t.health === "At Risk" || t.health === "Off Track" || t.status === "Blocked");
    if (atRisk.length) {
      bullets.push(`${atRisk.length} tracker item${atRisk.length === 1 ? "" : "s"} at risk / off track / blocked`);
    }
    const soon = open
      .filter((t) => t.end)
      .sort((a, b) => String(a.end).localeCompare(String(b.end)))
      .slice(0, 3);
    soon.forEach((t) => {
      bullets.push(`${t.title || "Task"} — due ${shortDate(t.end)}${t.assignee && t.assignee !== "Unassigned" ? ` (${t.assignee})` : ""}`);
    });
    if (!bullets.length) {
      bullets.push(`Project status: ${typeof computeProjectStatus === "function" ? computeProjectStatus(proj) : "In progress"}`);
      const deliverables = typeof projectDeliverables === "function" ? projectDeliverables(proj) : (proj.deliverables || []);
      deliverables.slice(0, 2).forEach((item) => bullets.push(item));
    }
    return bullets.slice(0, 8);
  }

  function riskSummaryRows(risks, proj) {
    /* Use manually-set overrides from reportConfig when available */
    const db = window.AEWTTR.db || {};
    const extra = proj ? (db.projectExtra && db.projectExtra[proj.id]) || {} : {};
    const overrides = (extra.reportConfig && extra.reportConfig.slideContent && extra.reportConfig.slideContent.riskRows) || null;

    const cats = [
      { key: "Schedule", match: /schedule|timing|delay/i },
      { key: "Budget", match: /budget|cost|funding|finance/i },
      { key: "Technical", match: /technical|tech|cyber|operational|staffing|contract|external/i }
    ];
    return cats.map((cat, i) => {
      /* If user set a manual status for this category, prefer it */
      const ov = overrides && overrides[i];
      if (ov && ov.rag) {
        return { category: cat.key, status: ov.rag === "Yellow" ? "Yellow" : ov.rag, comments: ov.notes || "" };
      }
      const matched = risks.filter((r) => cat.match.test(r.category || "") || (cat.key === "Technical" && !/schedule|budget|cost|funding/i.test(r.category || "")));
      const open = matched.filter((r) => r.status !== "Closed");
      const ranked = (open.length ? open : matched).slice().sort((a, b) => {
        const ra = (a.likelihood || 0) * (a.impact || 0);
        const rb = (b.likelihood || 0) * (b.impact || 0);
        return rb - ra;
      });
      const top = ranked[0];
      let status = "Green";
      if (top) {
        const score = (top.likelihood || 0) * (top.impact || 0);
        status = score >= 15 ? "Red" : score >= 6 ? "Yellow" : "Green";
      } else if (!matched.length) {
        status = "Green";
      }
      const comments = top
        ? (top.mitigationPlan || top.description || top.name || "").replace(/\s+/g, " ").trim().slice(0, 90)
        : "No open items in this category.";
      return { category: cat.key, status, comments };
    });
  }

  /* ----- Risk burn-down series (shared with Home UI) ----- */
  function riskBurnDownSeries(risks) {
    const pointsByDate = new Map();

    function bump(date, band) {
      if (!date) return;
      const key = String(date).slice(0, 10);
      if (!pointsByDate.has(key)) pointsByDate.set(key, { date: key, H: 0, M: 0, L: 0 });
      const row = pointsByDate.get(key);
      row[band] = (row[band] || 0) + 1;
    }

    const today = new Date().toISOString().slice(0, 10);
    risks.forEach((risk) => {
      const history = Array.isArray(risk.ratingHistory) && risk.ratingHistory.length ? risk.ratingHistory : null;
      if (history) {
        // Use real historical snapshots saved when likelihood/impact/status changed.
        history.forEach((snap) => {
          if (!snap.date) return;
          if (snap.status === "Closed") { bump(snap.date, "L"); return; }
          const b = snap.rating === "Red" ? "H" : snap.rating === "Amber" ? "M" : "L";
          bump(snap.date, b);
        });
        // Always anchor a current data point so the chart includes today.
        const bandNow = ratingBand(risk.likelihood, risk.impact);
        bump(today, risk.status === "Closed" ? "L" : bandNow);
      } else {
        // Legacy fallback: no history saved yet — use review notes with current rating.
        const reviews = Array.isArray(risk.reviewNotes) ? risk.reviewNotes.slice() : [];
        reviews.sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
        const opened = (reviews[0] && reviews[0].date) || risk.lastReviewedDate || risk.due || "";
        const bandNow = ratingBand(risk.likelihood, risk.impact);
        if (opened) bump(opened, risk.status === "Closed" ? "H" : bandNow);
        reviews.forEach((note) => {
          if (!note.date) return;
          bump(note.date, risk.status === "Closed" && note === reviews[reviews.length - 1] ? "L" : bandNow);
        });
        if (risk.lastReviewedDate) bump(risk.lastReviewedDate, risk.status === "Closed" ? "L" : bandNow);
        if (risk.status === "Closed" && risk.due) bump(risk.due, "L");
      }
    });

    // If history is thin, synthesize a best-effort path from open → mitigating → closed.
    if (pointsByDate.size < 2) {
      pointsByDate.clear();
      const today = new Date().toISOString().slice(0, 10);
      let h = 0;
      let m = 0;
      let l = 0;
      risks.forEach((risk) => {
        const band = ratingBand(risk.likelihood, risk.impact);
        if (risk.status === "Closed") l += 1;
        else if (band === "H") h += 1;
        else if (band === "M") m += 1;
        else l += 1;
      });
      const earliest = risks
        .map((r) => r.due || r.lastReviewedDate || "")
        .filter(Boolean)
        .sort()[0] || today;
      pointsByDate.set(earliest, { date: earliest, H: Math.max(h + m, 1), M: 0, L: 0 });
      const mid = new Date((new Date(earliest + "T00:00:00").getTime() + new Date(today + "T00:00:00").getTime()) / 2);
      const midIso = isNaN(mid) ? today : mid.toISOString().slice(0, 10);
      pointsByDate.set(midIso, { date: midIso, H: h, M: Math.max(m, 1), L: Math.max(l, 0) });
      pointsByDate.set(today, { date: today, H: h, M: m, L: l });
    }

    return Array.from(pointsByDate.values()).sort((a, b) => a.date.localeCompare(b.date));
  }

  function perRiskBurnPoints(risk) {
    const reviews = Array.isArray(risk.reviewNotes) ? risk.reviewNotes.slice() : [];
    reviews.sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
    const pts = [];
    const start = (reviews[0] && reviews[0].date) || risk.lastReviewedDate || risk.due || new Date().toISOString().slice(0, 10);
    pts.push({
      n: 1,
      date: start,
      pr: Math.min(5, Math.max(1, Number(risk.likelihood) || 3)),
      cr: Math.min(5, Math.max(1, Number(risk.impact) || 3)),
      desc: "Initial assessment",
      timeline: "Open"
    });
    reviews.slice(0, 3).forEach((note, i) => {
      pts.push({
        n: pts.length + 1,
        date: note.date || start,
        pr: Math.max(1, Math.min(5, (Number(risk.likelihood) || 3) - Math.floor((i + 1) / 2))),
        cr: Math.max(1, Math.min(5, (Number(risk.impact) || 3) - (i > 1 ? 1 : 0))),
        desc: (note.note || "Review update").slice(0, 80),
        timeline: "Mitigation"
      });
    });
    if (risk.mitigationPlan) {
      pts.push({
        n: pts.length + 1,
        date: risk.due || risk.lastReviewedDate || start,
        pr: risk.status === "Closed" ? 1 : Math.max(1, (Number(risk.likelihood) || 2) - 1),
        cr: risk.status === "Closed" ? 1 : Math.max(1, (Number(risk.impact) || 2) - 1),
        desc: String(risk.mitigationPlan).slice(0, 80),
        timeline: risk.status === "Closed" ? "Closed" : "In progress"
      });
    }
    // Deduplicate by keeping last 4.
    return pts.slice(-4).map((p, i) => Object.assign({}, p, { n: i + 1 }));
  }

  function addCuiBanner(slide, box) {
    slide.addText(CUI, {
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h,
      fontSize: 7.5,
      fontFace: "Arial Narrow",
      color: CUI_GRAY,
      align: "center",
      bold: false,
      valign: "middle"
    });
  }

  function addStatusChrome(slide, shapes, logos, pageNum) {
    addCuiBanner(slide, L.cuiTop);
    if (logos.ttsd) {
      slide.addImage({
        data: logos.ttsd,
        x: L.ttsd.x, y: L.ttsd.y, w: L.ttsd.w, h: L.ttsd.h
      });
    }
    if (logos.aewttr) {
      slide.addImage({
        data: logos.aewttr,
        x: L.aewttr.x, y: L.aewttr.y, w: L.aewttr.w, h: L.aewttr.h
      });
    }
    slide.addShape(shapes.line, {
      x: 0,
      y: L.footerLineY,
      w: SLIDE_W,
      h: 0,
      line: { color: "D9D9D9", width: 1 }
    });
    if (logos.navairMark) {
      slide.addImage({
        data: logos.navairMark,
        x: L.navair.x, y: L.navair.y, w: L.navair.w, h: L.navair.h,
        sizing: { type: "contain", w: L.navair.w, h: L.navair.h }
      });
    } else {
      slide.addText("NAVAIR", {
        x: L.navair.x, y: L.navair.y, w: L.navair.w, h: L.navair.h,
        fontSize: 10, bold: true, color: GRAY, fontFace: "Arial", valign: "middle"
      });
    }
    addCuiBanner(slide, L.cuiBottom);
    if (pageNum != null) {
      slide.addText(String(pageNum), {
        x: L.pageNum.x, y: L.pageNum.y, w: L.pageNum.w, h: L.pageNum.h,
        fontSize: 8, align: "center", valign: "middle",
        fontFace: "Arial Narrow", color: "808080"
      });
    }
  }

  /* Section headings mirror the centered title treatment in Slides.zip. */
  function addSectionHeading(slide, title, box) {
    slide.addText(title, {
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h,
      fontSize: 12,
      fontFace: "Arial",
      bold: true,
      color: "000000",
      align: "center",
      valign: "middle",
      margin: 0
    });
  }

  function probeImageSize(dataUri) {
    return new Promise((resolve) => {
      if (!dataUri) {
        resolve({ w: 0, h: 0 });
        return;
      }
      const img = new Image();
      img.onload = () => resolve({ w: img.naturalWidth || 0, h: img.naturalHeight || 0 });
      img.onerror = () => resolve({ w: 0, h: 0 });
      img.src = dataUri;
    });
  }

  /* Fit image inside frame preserving aspect ratio (letterbox). */
  function fitContainRect(frame, iw, ih) {
    if (!iw || !ih) {
      return { x: frame.x, y: frame.y, w: frame.w, h: frame.h };
    }
    const scale = Math.min(frame.w / iw, frame.h / ih);
    const w = iw * scale;
    const h = ih * scale;
    return {
      x: frame.x + (frame.w - w) / 2,
      y: frame.y + (frame.h - h) / 2,
      w,
      h
    };
  }

  function tileFrames(frame, count) {
    const n = Math.max(1, Math.min(count, 4));
    if (n === 1) return [Object.assign({}, frame)];
    if (n === 2) {
      const gap = 0.08;
      const w = (frame.w - gap) / 2;
      return [
        { x: frame.x, y: frame.y, w, h: frame.h },
        { x: frame.x + w + gap, y: frame.y, w, h: frame.h }
      ];
    }
    const gap = 0.08;
    const cols = 2;
    const rows = Math.ceil(n / cols);
    const cellW = (frame.w - gap) / cols;
    const cellH = (frame.h - gap * (rows - 1)) / rows;
    const out = [];
    for (let i = 0; i < n; i++) {
      const c = i % cols;
      const r = Math.floor(i / cols);
      out.push({
        x: frame.x + c * (cellW + gap),
        y: frame.y + r * (cellH + gap),
        w: cellW,
        h: cellH
      });
    }
    return out;
  }

  async function placePhotosInFrame(slide, shapes, imageDataUris, frame) {
    slide.addShape(shapes.rect, {
      x: frame.x,
      y: frame.y,
      w: frame.w,
      h: frame.h,
      fill: { color: PHOTO_FRAME_BG },
      line: { color: "D0D5DD", width: 0.75 }
    });
    if (!imageDataUris.length) {
      slide.addText("No project image selected", {
        x: frame.x,
        y: frame.y + frame.h / 2 - 0.15,
        w: frame.w,
        h: 0.3,
        fontSize: 10,
        color: GRAY,
        align: "center",
        fontFace: "Arial"
      });
      return;
    }
    const tiles = tileFrames(frame, imageDataUris.length);
    for (let i = 0; i < tiles.length; i++) {
      const data = imageDataUris[i];
      const size = await probeImageSize(data);
      const fit = fitContainRect(tiles[i], size.w, size.h);
      if (imageDataUris.length > 1) {
        slide.addShape(shapes.rect, {
          x: tiles[i].x,
          y: tiles[i].y,
          w: tiles[i].w,
          h: tiles[i].h,
          fill: { color: "ECEEF2" },
          line: { color: "D0D5DD", width: 0.5 }
        });
      }
      slide.addImage({
        data,
        x: fit.x,
        y: fit.y,
        w: fit.w,
        h: fit.h
      });
    }
  }

  /* MC milestone categories — mirrors the MC array in projects.js */
  const MC_PPTX = [
    { key: "contractAwarded", label: "Contract Award",     short: "Contract", color: "7F7F7F", sym: "star2" },
    { key: "fat",             label: "FAT",                short: "FAT",      color: "70D6C8", sym: "tri"   },
    { key: "sat",             label: "SAT",                short: "SAT",      color: "BDD7EE", sym: "tri"   },
    { key: "add",             label: "ADD / MFR",          short: "ADD",      color: "C55A11", sym: "tri"   },
    { key: "fielding",        label: "Fielding",           short: "Field",    color: "7030A0", sym: "tri"   },
    { key: "complete",        label: "Project Completion", short: "Complete", color: "FF0000", sym: "star2" }
  ];

  function mcSymText(sym) {
    if (sym === "star")  return "☆"; // ☆
    if (sym === "star2") return "★"; // ★
    return "△"; // △
  }

  /* Draw a Gantt-style milestone chart using typed MC dates from reportConfig.taskMilestones */
  function drawMcMilestoneChart(slide, shapes, proj, x, y, w, h) {
    const db = window.AEWTTR.db || {};
    const extra = (db.projectExtra && db.projectExtra[proj.id]) || {};
    const cfg = extra.reportConfig || {};
    const tm = typeof getProjectReportingMilestones === "function"
      ? getProjectReportingMilestones(proj)
      : cfg.taskMilestones || {};
    const tasks = (db.ganttTasks && db.ganttTasks[proj.id]) || [];
    const plain = typeof trackerPlainTasks === "function" ? trackerPlainTasks(tasks) : tasks;
    const reportTasks = plain.filter((t) => tm[t.id] && tm[t.id].inReport);

    if (!reportTasks.length) {
      slide.addText("No tasks selected in Reporting tab. Enable ‘In Report’ for tasks and set milestone dates.", {
        x, y: y + 0.1, w, h: 0.55, fontSize: 8, color: GRAY, fontFace: "Arial"
      });
      return;
    }

    const allMs = [];
    reportTasks.forEach((task) => {
      const t = tm[task.id] || {};
      MC_PPTX.forEach((c) => {
        const when = t[c.key] ? new Date(t[c.key]).getTime() : NaN;
        if (!isNaN(when)) allMs.push(when);
      });
    });
    if (!allMs.length) {
      slide.addText("Enter milestone dates in the Reporting tab to generate the timeline.", {
        x, y: y + 0.1, w, h: 0.55, fontSize: 8, color: GRAY, fontFace: "Arial"
      });
      return;
    }

    const firstDate = new Date(Math.min.apply(null, allMs));
    const lastDate = new Date(Math.max.apply(null, allMs));
    const minT = new Date(firstDate.getFullYear(), 0, 1).getTime();
    const maxT = new Date(lastDate.getFullYear() + 1, 0, 1).getTime();
    const span = maxT - minT;

    function toFrac(ds) { return (new Date(ds).getTime() - minT) / span; }

    const labelW = Math.min(1.4, w * 0.28);
    const chartX = x + labelW;
    const chartW = w - labelW;

    /* ── Year + Month axis ── */
    const yrH = 0.18;
    const moH = 0.14;
    const axisH = yrH + moH;
    slide.addShape(shapes.rect, {
      x, y, w: labelW, h: axisH,
      fill: { color: "FFFFFF" },
      line: { color: "000000", width: 0.5 }
    });
    slide.addText("Project", {
      x: x + 0.03, y: y + yrH, w: labelW - 0.06, h: moH,
      fontSize: 6.5, bold: true, color: "000000", fontFace: "Arial", valign: "middle"
    });
    // Year row background
    slide.addShape(shapes.rect, {
      x: chartX, y, w: chartW, h: yrH,
      fill: { color: "5B9BD5" },
      line: { color: "000000", width: 0.5 }
    });
    // Month row background
    slide.addShape(shapes.rect, {
      x: chartX, y: y + yrH, w: chartW, h: moH,
      fill: { color: "5B9BD5" },
      line: { color: "000000", width: 0.5 }
    });
    const MONTH_ABBR = ["J","F","M","A","M","J","J","A","S","O","N","D"];
    const startYear = new Date(minT).getFullYear();
    const endYear = new Date(maxT).getFullYear();
    for (let yr = startYear; yr <= endYear; yr++) {
      const yrStartT = new Date(yr, 0, 1).getTime();
      const yrEndT   = new Date(yr + 1, 0, 1).getTime();
      const fLeft  = Math.max(0, (yrStartT - minT) / span);
      const fRight = Math.min(1, (yrEndT - minT) / span);
      const yrW = (fRight - fLeft) * chartW;
      if (yrW < 0.08) continue;
      slide.addText(String(yr), {
        x: chartX + fLeft * chartW, y, w: (fRight - fLeft) * chartW, h: yrH,
        fontSize: 7, color: "000000", align: "center", valign: "middle", bold: true, fontFace: "Arial"
      });
      if (yrStartT > minT && fLeft > 0) {
        slide.addShape(shapes.line, {
          x: chartX + fLeft * chartW, y, w: 0, h: axisH,
          line: { color: "000000", width: 0.5 }
        });
      }
      // Month letters within this year
      for (let mo = 0; mo < 12; mo++) {
        const moStartT = new Date(yr, mo, 1).getTime();
        const moEndT   = new Date(yr, mo + 1, 1).getTime();
        const mfL = (moStartT - minT) / span;
        const mfR = (moEndT - minT) / span;
        if (mfR < 0 || mfL > 1) continue;
        const mfLc = Math.max(0, mfL);
        const mfRc = Math.min(1, mfR);
        const moW = (mfRc - mfLc) * chartW;
        if (moW < 0.02) continue;
        slide.addText(MONTH_ABBR[mo], {
          x: chartX + mfLc * chartW, y: y + yrH, w: (mfRc - mfLc) * chartW, h: moH,
          fontSize: 5.5, color: "000000", align: "center", valign: "middle", bold: true, fontFace: "Arial Narrow"
        });
        if (mo > 0 && mfL > 0 && mfL < 1) {
          slide.addShape(shapes.line, {
            x: chartX + mfLc * chartW, y: y + yrH, w: 0, h: moH,
            line: { color: "2F5597", width: 0.25 }
          });
        }
      }
    }

    /* ── Rows ── */
    const legendH = 0.18;
    const rows = reportTasks.slice(0, 8);
    const rowH = Math.min(0.32, (h - axisH - legendH - 0.14) / Math.max(rows.length, 1));
    const symSz = Math.max(7, Math.min(11, rowH * 26));

    rows.forEach((task, i) => {
      const rowY = y + axisH + i * rowH;
      const t = tm[task.id] || {};

      slide.addShape(shapes.rect, {
        x, y: rowY, w, h: rowH,
        fill: { color: "FFFFFF" }, line: { color: "000000", width: 0.5 }
      });

      slide.addText(task.title || task.name || "Task", {
        x: x + 0.03, y: rowY, w: labelW - 0.06, h: rowH,
        fontSize: 7, fontFace: "Arial", color: "000000", bold: true, valign: "middle", margin: 0
      });

      const markers = MC_PPTX.map((c) => ({ c, v: t[c.key] })).filter((m) => m.v);
      if (!markers.length) return;
      const fracs = markers.map((m) => toFrac(m.v));
      const minF = Math.min.apply(null, fracs);
      const maxF = Math.max.apply(null, fracs);

      /* Connector line between outermost markers */
      if (fracs.length >= 2) {
        slide.addShape(shapes.line, {
          x: chartX + minF * chartW,
          y: rowY + rowH / 2,
          w: (maxF - minF) * chartW,
          h: 0,
          line: { color: "5B9BD5", width: 1.25, beginArrowType: "none", endArrowType: "triangle" }
        });
      }

      /* Milestone symbols */
      markers.forEach((m, mi) => {
        const f = fracs[mi];
        const cx = chartX + f * chartW;
        const symY = rowY + rowH / 2 - (symSz / 72) * 0.55;
        slide.addText(mcSymText(m.c.sym), {
          x: cx - 0.14, y: symY, w: 0.28, h: (symSz / 72) * 1.2,
          fontSize: symSz, color: m.c.color, align: "center", valign: "middle", fontFace: "Arial"
        });
      });
    });

    for (let yr = startYear + 1; yr <= endYear; yr++) {
      const frac = (new Date(yr, 0, 1).getTime() - minT) / span;
      if (frac <= 0 || frac >= 1) continue;
      slide.addShape(shapes.line, {
        x: chartX + frac * chartW, y, w: 0, h: axisH + rows.length * rowH,
        line: { color: "000000", width: 0.5 }
      });
    }

    /* ── Today line ── */
    const todayT = Date.now();
    const todayF = (todayT - minT) / span;
    if (todayF >= 0 && todayF <= 1) {
      const tx = chartX + todayF * chartW;
      slide.addShape(shapes.line, {
        x: tx, y: y + axisH,
        w: 0, h: rows.length * rowH + 0.1,
        line: { color: "FFC000", width: 1.5 }
      });
      slide.addText("Today", {
        x: tx - 0.22, y: y + axisH + rows.length * rowH + 0.1, w: 0.44, h: 0.13,
        fontSize: 5.5, color: "9C6500", align: "center", fontFace: "Arial", bold: true
      });
    }

    /* ── Legend ── */
    const legY = y + h - legendH;
    const legendCellW = w / MC_PPTX.length;
    MC_PPTX.forEach((c, idx) => {
      const cellX = x + idx * legendCellW;
      slide.addShape(shapes.rect, {
        x: cellX, y: legY, w: legendCellW, h: legendH,
        fill: { color: "FFFFFF" }, line: { color: "000000", width: 0.5 }
      });
      slide.addText(mcSymText(c.sym), {
        x: cellX + 0.02, y: legY, w: 0.14, h: legendH,
        fontSize: 8, color: c.color, align: "center", valign: "middle", fontFace: "Arial"
      });
      slide.addText(c.label, {
        x: cellX + 0.16, y: legY, w: legendCellW - 0.18, h: legendH,
        fontSize: 5.5, color: "000000", bold: true, align: "center", valign: "middle", fontFace: "Arial"
      });
    });
  }

  /* Legacy bar-based gantt — kept for fallback when no MC data exists */
  function drawMilestoneGantt(slide, shapes, milestones, x, y, w, h) {
    if (!milestones.length) {
      slide.addText("No dividers flagged as milestones yet. Toggle the flag on a Project Divider in Tracker.", {
        x, y: y + 0.15, w, h: 0.55, fontSize: 9, color: GRAY, fontFace: "Arial"
      });
      return;
    }
    const dated = milestones.filter((m) => m.start || m.end);
    const starts = dated.map((m) => new Date((m.start || m.end) + "T00:00:00").getTime()).filter((n) => !isNaN(n));
    const ends = dated.map((m) => new Date((m.end || m.start) + "T00:00:00").getTime()).filter((n) => !isNaN(n));
    let minT = starts.length ? Math.min.apply(null, starts) : Date.now();
    let maxT = ends.length ? Math.max.apply(null, ends) : minT + 86400000 * 365;
    const pad = Math.max((maxT - minT) * 0.06, 86400000 * 45);
    minT -= pad; maxT += pad;
    const span = Math.max(maxT - minT, 86400000 * 90);
    const years = [];
    const startYear = new Date(minT).getFullYear();
    const endYear = new Date(maxT).getFullYear();
    for (let yr = startYear; yr <= endYear; yr++) years.push(yr);
    const axisH = 0.2;
    slide.addShape(shapes.rect, {
      x, y, w, h: axisH, fill: { color: "E8EEF5" }, line: { color: "C5D0DE", width: 0.5 }
    });
    years.forEach((yr, i) => {
      const left = x + (i / years.length) * w;
      slide.addText(String(yr), {
        x: left, y, w: w / years.length, h: axisH,
        fontSize: 8, color: "334155", align: "center", valign: "middle", bold: true, fontFace: "Arial"
      });
      if (i > 0) slide.addShape(shapes.line, { x: left, y, w: 0, h: axisH, line: { color: "C5D0DE", width: 0.75 } });
    });
    const labelW = Math.min(1.55, w * 0.3);
    const chartX = x + labelW;
    const chartW = w - labelW;
    const rows = milestones.slice(0, 6);
    const rowH = Math.min(0.34, (h - axisH - 0.12) / Math.max(rows.length, 1));
    rows.forEach((m, i) => {
      const rowY = y + axisH + 0.08 + i * rowH;
      slide.addText(m.title || "Milestone", {
        x, y: rowY, w: labelW - 0.06, h: rowH, fontSize: 8, fontFace: "Arial", color: "222222", valign: "middle"
      });
      const t0 = new Date((m.start || m.end) + "T00:00:00").getTime();
      const t1 = new Date((m.end || m.start) + "T00:00:00").getTime();
      if (isNaN(t0)) return;
      const left = chartX + ((Math.min(t0, t1) - minT) / span) * chartW;
      const width = Math.max(((Math.abs(t1 - t0) || 86400000 * 30) / span) * chartW, 0.12);
      slide.addShape(shapes.rect, {
        x: left, y: rowY + rowH * 0.28, w: width, h: Math.max(rowH * 0.42, 0.1),
        fill: { color: m.done ? GREEN : "2E75B6" }
      });
      slide.addShape(shapes.diamond || shapes.rect, {
        x: left + width - 0.06, y: rowY + rowH / 2 - 0.06, w: 0.12, h: 0.12,
        fill: { color: m.done ? "1F7A4D" : "C55A11" }
      });
    });
    const todayT = Date.now();
    if (todayT >= minT && todayT <= maxT) {
      const tx = chartX + ((todayT - minT) / span) * chartW;
      slide.addShape(shapes.line, {
        x: tx, y: y + axisH, w: 0, h: h - axisH - 0.05,
        line: { color: "C55A11", width: 1.25, dashType: "dash" }
      });
    }
  }

  function addHeatMap(slide, shapes, points, x, y, size, title) {
    slide.addText(title || "Risk Type", {
      x, y, w: size, h: 0.22, fontSize: 10, bold: true, fontFace: "Arial", color: "111111", align: "center"
    });
    const grid = size - 0.4;
    const cell = grid / 5;
    const ox = x + 0.28;
    const oy = y + 0.28;
    for (let pr = 5; pr >= 1; pr--) {
      for (let cr = 1; cr <= 5; cr++) {
        const cx = ox + (cr - 1) * cell;
        const cy = oy + (5 - pr) * cell;
        slide.addShape(shapes.rect, {
          x: cx,
          y: cy,
          w: cell - 0.02,
          h: cell - 0.02,
          fill: { color: heatCellColor(pr, cr) },
          line: { color: "FFFFFF", width: 0.5 }
        });
      }
      slide.addText(String(pr), {
        x: x,
        y: oy + (5 - pr) * cell,
        w: 0.24,
        h: cell,
        fontSize: 8,
        align: "center",
        valign: "middle",
        fontFace: "Arial"
      });
    }
    for (let cr = 1; cr <= 5; cr++) {
      slide.addText(String(cr), {
        x: ox + (cr - 1) * cell,
        y: oy + grid + 0.02,
        w: cell,
        h: 0.18,
        fontSize: 8,
        align: "center",
        fontFace: "Arial"
      });
    }
    slide.addText("Pr", {
      x: x,
      y: oy - 0.02,
      w: 0.24,
      h: 0.16,
      fontSize: 7,
      color: GRAY,
      align: "center",
      fontFace: "Arial"
    });
    slide.addText("Consequence (Cr)", {
      x: ox,
      y: oy + grid + 0.18,
      w: grid,
      h: 0.16,
      fontSize: 7,
      color: GRAY,
      align: "center",
      fontFace: "Arial"
    });
    points.forEach((p) => {
      const cx = ox + (p.cr - 1) * cell + cell / 2 - 0.1;
      const cy = oy + (5 - p.pr) * cell + cell / 2 - 0.1;
      slide.addShape(shapes.ellipse, {
        x: cx,
        y: cy,
        w: 0.2,
        h: 0.2,
        fill: { color: "111111" }
      });
      slide.addText(String(p.n), {
        x: cx,
        y: cy,
        w: 0.2,
        h: 0.2,
        fontSize: 8,
        color: "FFFFFF",
        align: "center",
        valign: "middle",
        bold: true,
        fontFace: "Arial"
      });
    });
  }

  function addBurnDownChart(slide, shapes, points, x, y, w, h) {
    slide.addText("Risk Burn Down", {
      x, y, w, h: 0.22, fontSize: 10, bold: true, fontFace: "Arial", align: "center"
    });
    const plotY = y + 0.28;
    const plotH = h - 0.55;
    const bandH = plotH / 3;
    [
      { label: "H", color: "F6D6D3" },
      { label: "M", color: "F7ECB8" },
      { label: "L", color: "D7EBDD" }
    ].forEach((band, i) => {
      slide.addShape(shapes.rect, {
        x: x + 0.28,
        y: plotY + i * bandH,
        w: w - 0.4,
        h: bandH,
        fill: { color: band.color },
        line: { color: "DDDDDD", width: 0.5 }
      });
      slide.addText(band.label, {
        x,
        y: plotY + i * bandH,
        w: 0.26,
        h: bandH,
        fontSize: 10,
        bold: true,
        align: "center",
        valign: "middle",
        fontFace: "Arial"
      });
    });

    const series = points.length ? points : [{ n: 1, pr: 4, cr: 4, date: "" }];
    const coords = series.map((p, i) => {
      const score = (p.pr || 1) * (p.cr || 1);
      const band = score >= 15 ? 0 : score >= 6 ? 1 : 2;
      const within = score >= 15 ? (score - 15) / 10 : score >= 6 ? (score - 6) / 9 : score / 6;
      const px = x + 0.4 + (i / Math.max(series.length - 1, 1)) * (w - 0.7);
      const py = plotY + band * bandH + (1 - Math.min(Math.max(within, 0.15), 0.85)) * bandH;
      return { x: px, y: py, n: p.n, date: p.date };
    });
    for (let i = 0; i < coords.length - 1; i++) {
      slide.addShape(shapes.line, {
        x: coords[i].x,
        y: coords[i].y,
        w: coords[i + 1].x - coords[i].x,
        h: coords[i + 1].y - coords[i].y,
        line: { color: NAVY_DARK, width: 2 }
      });
    }
    coords.forEach((c) => {
      slide.addShape(shapes.ellipse, {
        x: c.x - 0.08,
        y: c.y - 0.08,
        w: 0.16,
        h: 0.16,
        fill: { color: NAVY_DARK }
      });
      slide.addText(String(c.n), {
        x: c.x - 0.08,
        y: c.y - 0.08,
        w: 0.16,
        h: 0.16,
        fontSize: 8,
        color: "FFFFFF",
        align: "center",
        valign: "middle",
        fontFace: "Arial"
      });
    });
    series.forEach((p, i) => {
      slide.addText(shortDate(p.date) || `#${p.n}`, {
        x: x + 0.28 + (i / Math.max(series.length, 1)) * (w - 0.5),
        y: y + h - 0.22,
        w: (w - 0.5) / Math.max(series.length, 1),
        h: 0.2,
        fontSize: 7,
        color: GRAY,
        align: "center",
        fontFace: "Arial"
      });
    });
  }

  async function buildStatusSlide(pptx, shapes, proj, logos, imageDataUris, pageNum) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addStatusChrome(slide, shapes, logos, pageNum);

    /* Title stays left of the vertical divider (x≈4.73) so it never pushes over
       the Technical Status heading or POC column on the right side. */
    slide.addText(proj.name || "Project", {
      x: L.title.x,
      y: L.title.y,
      w: 3.6,
      h: L.title.h,
      fontSize: 20,
      bold: true,
      fontFace: "Arial Black",
      color: "000000",
      align: "left",
      valign: "middle",
      fit: "shrink"
    });

    const poc = resolvePoc(proj);
    const pocLines = [
      { text: poc.name ? `POC: ${poc.name}` : "POC:", options: { breakLine: true, bold: true } },
      { text: poc.org || "AEWTTR IPT", options: { breakLine: true } },
      { text: poc.phone ? `Phone: ${poc.phone}` : "", options: { breakLine: true } },
      { text: poc.email ? `E-mail: ${poc.email}` : "", options: { breakLine: false } }
    ].filter((line) => line.text);
    slide.addText(pocLines, {
      x: L.poc.x,
      y: L.poc.y,
      w: L.poc.w,
      h: L.poc.h,
      fontSize: 6,
      fontFace: "Arial",
      color: "222222",
      valign: "top",
      margin: 0
    });

    /* Navy cross dividers from the zip status layout. */
    slide.addShape(shapes.line, {
      x: L.vDiv.x,
      y: L.vDiv.y,
      w: 0,
      h: L.vDiv.h,
      line: { color: DIVIDER, width: 1.5 }
    });
    slide.addShape(shapes.line, {
      x: L.hDiv.x,
      y: L.hDiv.y,
      w: L.hDiv.w,
      h: 0,
      line: { color: DIVIDER, width: 1.5 }
    });

    await placePhotosInFrame(slide, shapes, imageDataUris.slice(0, 4), L.photo);

    addSectionHeading(slide, "Project Description", L.descHead);
    const descBullets = bulletize(
      proj.scope || proj.description || (typeof projectDeliverables === "function" ? projectDeliverables(proj).join("\n") : (proj.deliverables || []).join("\n")) || "",
      10
    );
    const descFit = fitBulletsToBox(
      descBullets.length ? descBullets : ["No project description captured yet."],
      L.descBody.w,
      L.descBody.h,
      { maxFont: 10, minFont: 7, paraGap: 2, pad: 0.02 }
    );
    slide.addText(
      descFit.texts.map((t, idx) => ({
        text: t,
        options: { bullet: true, breakLine: idx < descFit.texts.length - 1 }
      })),
      {
        x: L.descBody.x,
        y: L.descBody.y,
        w: L.descBody.w,
        h: L.descBody.h,
        fontSize: descFit.fontSize,
        fontFace: "Arial",
        color: "222222",
        valign: "top",
        margin: 0,
        paraSpacingAfter: descFit.paraGap,
        fit: "shrink"
      }
    );

    addSectionHeading(slide, "Technical Status", L.techHead);
    const tech = technicalStatusBullets(proj);
    /* Fixed-height tech panel: bounded by available space above risk table.
       risk table header + 3 rows at 0.32" each + gaps = 1.42".
       Available: hDiv.y (4.053) - techPanel.y (1.08) - 1.42 - 0.12 ≈ 1.41" for tech bullets.
       Cap at 1.4 so the risk table always starts at a predictable Y. */
    const TECH_BOX_H = 1.35;
    const RISK_TABLE_Y = L.techPanel.y + TECH_BOX_H + 0.12; // ≈ 2.55
    const techFit = fitBulletsToBox(tech, L.techPanel.w - 0.2, TECH_BOX_H - 0.1, {
      maxFont: 10, minFont: 7, paraGap: 2, pad: 0.02
    });
    slide.addText(
      techFit.texts.map((t, idx) => ({
        text: t,
        options: { bullet: true, breakLine: idx < techFit.texts.length - 1 }
      })),
      {
        x: L.techPanel.x + 0.1,
        y: L.techPanel.y + 0.06,
        w: L.techPanel.w - 0.2,
        h: TECH_BOX_H - 0.08,
        fontSize: techFit.fontSize,
        fontFace: "Arial",
        color: "111111",
        valign: "top",
        margin: 0,
        paraSpacingAfter: techFit.paraGap,
        fit: "shrink"
      }
    );

    const risks = typeof projectRisks === "function" ? projectRisks(proj.id) : [];
    const summary = riskSummaryRows(risks, proj);
    /* Place risk table at fixed Y below tech panel, clamped above hDiv. */
    let riskY = RISK_TABLE_Y;
    const riskMaxBottom = L.hDiv.y - 0.06;
    const riskEstH = 0.28 + summary.length * 0.34;
    if (riskY + riskEstH > riskMaxBottom) {
      riskY = Math.max(RISK_TABLE_Y, riskMaxBottom - riskEstH);
    }
    slide.addTable(
      [
        [
          {
            text: "Risk Category",
            options: { bold: true, fill: { color: LIGHT_BLUE }, color: "000000", align: "left" }
          },
          {
            text: "Status",
            options: { bold: true, fill: { color: LIGHT_BLUE }, color: "000000", align: "center" }
          },
          {
            text: "Comments",
            options: { bold: true, fill: { color: LIGHT_BLUE }, color: "000000", align: "left" }
          }
        ]
      ].concat(
        summary.map((row, idx) => {
          const fill = ragColor(row.status);
          const stripe = idx % 2 ? ROW_ALT : "FFFFFF";
          return [
            {
              text: row.category,
              options: { fontSize: 10, fill: { color: stripe }, color: "222222" }
            },
            {
              text: "",
              options: {
                fill: { color: fill },
                align: "center"
              }
            },
            {
              text: row.comments,
              options: { fontSize: 9, fill: { color: stripe }, color: "222222" }
            }
          ];
        })
      ),
      {
        x: L.riskTable.x,
        y: riskY,
        w: L.riskTable.w,
        colW: L.riskColW,
        border: [
          { pt: 0.75, color: "000000" },
          { pt: 0.75, color: "000000" },
          { pt: 0.75, color: "000000" },
          { pt: 0.75, color: "000000" }
        ],
        fontFace: "Arial",
        fontSize: 10,
        color: "222222",
        valign: "middle",
        rowH: 0.32
      }
    );

    addSectionHeading(slide, "Milestones", L.milesHead);
    drawMcMilestoneChart(slide, shapes, proj, L.milesBody.x, L.milesBody.y, L.milesBody.w, L.milesBody.h);

    slide.addText(`As of ${todayLabel()}`, {
      x: L.asOf.x, y: L.asOf.y, w: L.asOf.w, h: L.asOf.h,
      fontSize: 10, color: "333333", align: "center", fontFace: "Arial", valign: "middle"
    });
  }

  async function buildRiskSlide(pptx, shapes, proj, risk, logos, pageNum, totalPages) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    const points = perRiskBurnPoints(risk);
    addStatusChrome(slide, shapes, logos, pageNum);

    // ── Title block ──
    slide.addText("Risk Assessment", {
      x: L.title.x, y: L.title.y, w: 6.0, h: 0.35,
      fontSize: 20, bold: true, color: TITLE_NAVY, fontFace: "Arial Black"
    });
    slide.addText(proj.name || "Project", {
      x: L.title.x, y: 0.50, w: 6.0, h: 0.22, fontSize: 10, color: GRAY, fontFace: "Arial"
    });

    // ── Risk name bar + score badge (placed below seals at y≥1.04) ──
    const riskScore = (Number(risk.likelihood) || 0) * (Number(risk.impact) || 0);
    const scoreBand = riskScore >= 15 ? "HIGH" : riskScore >= 6 ? "MEDIUM" : "LOW";
    const scoreBadgeColor = riskScore >= 15 ? RED : riskScore >= 6 ? "FFC000" : GREEN;
    const scoreBadgeFontColor = (riskScore >= 6 && riskScore < 15) ? "333333" : "FFFFFF";
    slide.addShape(shapes.rect, {
      x: 0.28, y: 1.04, w: 9.44, h: 0.28,
      fill: { color: NAVY }, line: { color: NAVY, width: 0 }
    });
    slide.addText((risk.name || risk.title || "Untitled Risk").trim(), {
      x: 0.36, y: 1.04, w: 7.50, h: 0.28,
      fontSize: 11, bold: true, color: "FFFFFF", fontFace: "Arial",
      valign: "middle", margin: 0, fit: "shrink"
    });
    slide.addShape(shapes.rect, {
      x: 7.94, y: 1.04, w: 1.78, h: 0.28,
      fill: { color: scoreBadgeColor }, line: { color: scoreBadgeColor, width: 0 }
    });
    slide.addText(`Score: ${riskScore || "—"}  ${scoreBand}`, {
      x: 7.94, y: 1.04, w: 1.78, h: 0.28,
      fontSize: 8, bold: true, color: scoreBadgeFontColor,
      fontFace: "Arial", align: "center", valign: "middle"
    });

    // ── Meta strip: 6 attribute cells ──
    const metaY = 1.36;
    const metaH = 0.28;
    const metaItems = [
      { label: "Category",    value: risk.category || "—" },
      { label: "Status",      value: risk.status || "Open" },
      { label: "Likelihood",  value: risk.likelihood ? `${risk.likelihood}/5` : "—" },
      { label: "Impact",      value: risk.impact ? `${risk.impact}/5` : "—" },
      { label: "Due Date",    value: risk.due ? shortDate(risk.due) : "—" },
      { label: "Last Review", value: risk.lastReviewedDate ? shortDate(risk.lastReviewedDate) : "—" }
    ];
    const metaW = 9.44 / metaItems.length;
    metaItems.forEach((item, i) => {
      const cx = 0.28 + i * metaW;
      slide.addShape(shapes.rect, {
        x: cx, y: metaY, w: metaW, h: metaH,
        fill: { color: i % 2 ? "EEF5FB" : "F5F8FC" },
        line: { color: "D0D8E4", width: 0.5 }
      });
      slide.addText(item.label.toUpperCase(), {
        x: cx + 0.05, y: metaY + 0.02, w: metaW - 0.10, h: 0.10,
        fontSize: 5.5, color: GRAY, fontFace: "Arial", bold: true
      });
      slide.addText(item.value, {
        x: cx + 0.05, y: metaY + 0.12, w: metaW - 0.10, h: 0.16,
        fontSize: 8.5, color: "111111", fontFace: "Arial", valign: "middle", fit: "shrink"
      });
    });

    // ── Two-column body ──
    const LX = 0.28, LW = 4.72;
    const RX = 5.2,  RW = 4.5;
    const bodyY     = metaY + metaH + 0.08;  // ≈ 1.72
    const bodyBottom = 6.98;
    const labelH    = 0.18;
    const textH     = 1.40;
    const gap       = 0.06;

    // Consequence / Impact
    let ly = bodyY;
    slide.addText("Consequence / Impact:", {
      x: LX, y: ly, w: LW, h: labelH,
      fontSize: 9.5, bold: true, fontFace: "Arial", color: TITLE_NAVY, valign: "middle"
    });
    ly += labelH;
    const impBullets = bulletize(risk.description || "No consequence narrative captured.", 10);
    const impFit = fitBulletsToBox(impBullets, LW - 0.10, textH, { maxFont: 9, minFont: 7, paraGap: 1, pad: 0.02 });
    slide.addText(
      impFit.texts.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < impFit.texts.length - 1 } })),
      { x: LX + 0.05, y: ly, w: LW - 0.10, h: textH, fontSize: impFit.fontSize, fontFace: "Arial",
        color: "222222", valign: "top", margin: 0, paraSpacingAfter: impFit.paraGap, fit: "shrink" }
    );
    ly += textH + gap;

    // Mitigation Plan
    slide.addText("Mitigation Plan:", {
      x: LX, y: ly, w: LW, h: labelH,
      fontSize: 9.5, bold: true, fontFace: "Arial", color: TITLE_NAVY, valign: "middle"
    });
    ly += labelH;
    const mitBullets = bulletize(risk.mitigationPlan || "No mitigation plan captured.", 10);
    const mitFit = fitBulletsToBox(mitBullets, LW - 0.10, textH, { maxFont: 9, minFont: 7, paraGap: 1, pad: 0.02 });
    slide.addText(
      mitFit.texts.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < mitFit.texts.length - 1 } })),
      { x: LX + 0.05, y: ly, w: LW - 0.10, h: textH, fontSize: mitFit.fontSize, fontFace: "Arial",
        color: "222222", valign: "top", margin: 0, paraSpacingAfter: mitFit.paraGap, fit: "shrink" }
    );
    ly += textH + gap;

    // Action items table — explicit rowH prevents overflow
    const tableY = ly;
    const tableH = Math.max(0.4, bodyBottom - tableY);
    const dataPoints = points.slice(0, 4);
    if (dataPoints.length) {
      const rowH = Math.min(0.30, Math.max(0.22, (tableH - 0.28) / dataPoints.length));
      const actionRows = [
        [
          { text: "#",         options: { bold: true, fill: { color: LIGHT_BLUE }, color: "FFFFFF", align: "center" } },
          { text: "Description", options: { bold: true, fill: { color: LIGHT_BLUE }, color: "FFFFFF" } },
          { text: "Timeline",  options: { bold: true, fill: { color: LIGHT_BLUE }, color: "FFFFFF" } },
          { text: "Est. Comp", options: { bold: true, fill: { color: LIGHT_BLUE }, color: "FFFFFF" } },
          { text: "Pr",        options: { bold: true, fill: { color: LIGHT_BLUE }, color: "FFFFFF", align: "center" } },
          { text: "Cr",        options: { bold: true, fill: { color: LIGHT_BLUE }, color: "FFFFFF", align: "center" } }
        ]
      ].concat(
        dataPoints.map((p, i) => [
          { text: String(p.n), options: { align: "center", fill: { color: i % 2 ? ROW_ALT : "FFFFFF" } } },
          { text: String(p.desc || "").slice(0, 80), options: { fill: { color: i % 2 ? ROW_ALT : "FFFFFF" }, fontSize: 7.5 } },
          { text: p.timeline,        options: { fill: { color: i % 2 ? ROW_ALT : "FFFFFF" }, fontSize: 7.5 } },
          { text: shortDate(p.date), options: { fill: { color: i % 2 ? ROW_ALT : "FFFFFF" }, fontSize: 7.5 } },
          { text: String(p.pr), options: { align: "center", fill: { color: i % 2 ? ROW_ALT : "FFFFFF" } } },
          { text: String(p.cr), options: { align: "center", fill: { color: i % 2 ? ROW_ALT : "FFFFFF" } } }
        ])
      );
      slide.addTable(actionRows, {
        x: LX, y: tableY, w: LW,
        colW: [0.32, 2.05, 0.82, 0.72, 0.41, 0.40],
        border: [{ pt: 0.5, color: "BBBBBB" }, { pt: 0.5, color: "BBBBBB" }, { pt: 0.5, color: "BBBBBB" }, { pt: 0.5, color: "BBBBBB" }],
        fontFace: "Arial", fontSize: 8, color: "222222", valign: "middle", rowH
      });
    }

    // Right column — heat map then burn-down, no overlap
    const HEAT_SIZE = 2.4;
    addHeatMap(slide, shapes, points, RX, bodyY, HEAT_SIZE, `${String(risk.category || "Operational").toUpperCase()} RISK`);
    const burnY = bodyY + HEAT_SIZE + 0.34;  // clears "Consequence (Cr)" axis label
    const burnH = Math.max(1.0, bodyBottom - burnY);
    addBurnDownChart(slide, shapes, points, RX, burnY, RW, burnH);

    slide.addText(`As of ${todayLabel()}`, {
      x: L.asOf.x, y: L.asOf.y, w: L.asOf.w, h: L.asOf.h,
      fontSize: 10, color: "333333", align: "center", fontFace: "Arial", valign: "middle"
    });
    void totalPages;
  }

  function groupTypeLabel(groupType) {
    if (groupType === "portfolio") return "Portfolio";
    if (groupType === "program") return "Program Office";
    if (groupType === "eic") return "End Item Configuration";
    return "Project";
  }

  /* Cover slide matches Slides.zip slide 1: white background, corner seals, CUI chrome,
     centered black title, date line, distribution box. */
  function buildBriefingCoverSlide(pptx, shapes, logos, groupName, groupType, projectCount) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };

    /* Reuse the same chrome as status slides (seals, CUI banners, footer line, NAVAIR) */
    addStatusChrome(slide, shapes, logos, null);

    /* Main title — large centered, black, Arial Black matching ZIP */
    slide.addText("AEWTTR Project Updates", {
      x: 1.5, y: 1.6, w: 7.0, h: 0.9,
      fontSize: 36, bold: true, color: "000000", fontFace: "Arial Black",
      align: "center", valign: "middle", margin: 0, fit: "shrink"
    });

    /* Date line */
    slide.addText(todayLabel(), {
      x: 1.5, y: 2.6, w: 7.0, h: 0.42,
      fontSize: 20, bold: true, color: "111111", fontFace: "Arial",
      align: "center", margin: 0
    });

    /* Presented by / group info */
    const scope = groupTypeLabel(groupType);
    slide.addText(`${scope}: ${groupName}  ·  ${projectCount} project${projectCount === 1 ? "" : "s"}`, {
      x: 1.5, y: 3.1, w: 7.0, h: 0.3,
      fontSize: 12, color: GRAY, fontFace: "Arial", align: "center", margin: 0
    });

    /* Distribution notice box — mirrors ZIP layout */
    slide.addShape(shapes.rect, {
      x: 0.5, y: 3.6, w: 9.0, h: 2.8,
      fill: { color: "FFFFFF" },
      line: { color: "000000", width: 0.75 }
    });
    slide.addText(
      "Distribution Statement D: Distribution authorized to the Department of Defense and U.S. DoD contractors only.\n\n" +
      "DESTRUCTION NOTICE: The contents of this document are controlled unclassified information (CUI). Destroy in accordance with applicable regulations.\n\n" +
      "WARNING – EXPORT CONTROLLED: This document contains technical data whose export is restricted by the Arms Export Control Act (Section 2751 of Title 22, United States Code).",
      {
        x: 0.6, y: 3.7, w: 8.8, h: 2.6,
        fontSize: 8.5, fontFace: "Arial", color: "111111",
        valign: "top", margin: [4, 4, 4, 4], wrap: true
      }
    );

    /* Control info block — bottom right (mirrors ZIP) */
    slide.addText(
      "Controlled by: Department of the Navy\nControlled by: AEWTTR IPT\nCUI Category: CTI\nDistribution/Dissemination Control: D",
      {
        x: 5.5, y: 6.55, w: 4.35, h: 0.65,
        fontSize: 7.5, fontFace: "Arial", color: "111111",
        align: "right", valign: "top", margin: 0
      }
    );
  }

  async function exportProjectStatusPptx(proj) {
    const PptxGenJS = pptxCtor();
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: "AEWTTR_4x3", width: SLIDE_W, height: SLIDE_H });
    pptx.layout = "AEWTTR_4x3";
    pptx.author = "AEWTTR PULSE";
    pptx.title = `${proj.id || ""} ${proj.name || "Project"} Status`.trim();
    pptx.subject = "Project status and risk assessment";

    const shapes = pptx.ShapeType || PptxGenJS.shapes || {};
    const S = {
      rect: shapes.rect || shapes.RECTANGLE || "rect",
      line: shapes.line || shapes.LINE || "line",
      ellipse: shapes.ellipse || shapes.oval || shapes.OVAL || "ellipse",
      diamond: shapes.diamond || shapes.DIAMOND || shapes.rect || "rect"
    };

    const logos = await loadBrandLogos();
    const imageUrls = collectExportImages(proj);
    const imageDataUris = [];
    for (const url of imageUrls) {
      const data = await fetchAsDataUri(url);
      if (data) imageDataUris.push(data);
    }

    await buildStatusSlide(pptx, S, proj, logos, imageDataUris, 1);

    /* A project update is deliberately one slide.  Risk and work appendices
       belong only in an explicitly requested full deck. */
    const safeProjectName = String(proj.name || "project").replace(/[^\w.-]+/g, "-").replace(/^-+|-+$/g, "") || "project";
    const fileName = `${safeProjectName}-project-update-${new Date().toISOString().slice(0, 10)}.pptx`;
    const blob = await pptx.write({ outputType: "blob" });
    return { fileName, blob };
  }

  function downloadPptxBlob(blob, fileName) {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName || "status.pptx";
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 2000);
  }

  function closeReservedFileWindow(popup) {
    const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
    if (api && typeof api.closeReservedSharePointFileWindow === "function") {
      api.closeReservedSharePointFileWindow(popup);
    } else if (popup && !popup.closed) {
      popup.close();
    }
  }

  function openStoredFileByPolicy(fileUrl, fileName, popup) {
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

  function projectFilesLabel(proj) {
    if (typeof projectDocumentsLabel === "function") return projectDocumentsLabel(proj);
    return `${proj.id || "project"} - ${proj.name || "Untitled"}`;
  }

  async function uploadStatusPptxToSharePoint(proj, blob, fileName) {
    if (typeof isSharePointMode !== "function" || !isSharePointMode()) {
      throw new Error("SharePoint mode is required to save the status deck to project files.");
    }
    if (typeof sharePointAdapter === "undefined"
      || typeof sharePointAdapter.ensureProjectDocumentsRoot !== "function"
      || typeof sharePointAdapter.uploadPulseDocumentsFile !== "function") {
      throw new Error("SharePoint upload helpers are unavailable.");
    }
    const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
    if (!siteUrl) throw new Error("No SharePoint site URL is available.");
    const root = await sharePointAdapter.ensureProjectDocumentsRoot(siteUrl, projectFilesLabel(proj));
    const file = new File(
      [blob],
      fileName,
      { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
    );
    const uploaded = await sharePointAdapter.uploadPulseDocumentsFile(siteUrl, root.projectFolderUrl, file);
    return {
      fileName: uploaded.name || fileName,
      fileUrl: uploaded.fileUrl,
      serverRelativeUrl: uploaded.serverRelativeUrl,
      projectFolderUrl: root.projectFolderUrl
    };
  }

  async function uploadGroupStatusPptxToSharePoint(groupName, groupType, blob, fileName) {
    if (typeof isSharePointMode !== "function" || !isSharePointMode()) {
      throw new Error("SharePoint mode is required to save the briefing deck.");
    }
    if (typeof sharePointAdapter === "undefined"
      || typeof sharePointAdapter.resolvePulseDocumentsLibraryUrl !== "function"
      || typeof sharePointAdapter.createPulseDocumentsSubfolder !== "function"
      || typeof sharePointAdapter.uploadPulseDocumentsFile !== "function") {
      throw new Error("SharePoint reporting-folder helpers are unavailable.");
    }
    const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
    if (!siteUrl) throw new Error("No SharePoint site URL is available.");
    const libraryUrl = await sharePointAdapter.resolvePulseDocumentsLibraryUrl(siteUrl);
    const reportsRoot = await sharePointAdapter.createPulseDocumentsSubfolder(siteUrl, libraryUrl, "Reporting");
    const scopeType = await sharePointAdapter.createPulseDocumentsSubfolder(siteUrl, reportsRoot.serverRelativeUrl, groupTypeLabel(groupType));
    const scopeFolder = await sharePointAdapter.createPulseDocumentsSubfolder(siteUrl, scopeType.serverRelativeUrl, groupName || "General");
    const file = new File(
      [blob],
      fileName,
      { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
    );
    const uploaded = await sharePointAdapter.uploadPulseDocumentsFile(siteUrl, scopeFolder.serverRelativeUrl, file);
    return Object.assign({ reportFolderUrl: scopeFolder.serverRelativeUrl }, uploaded);
  }

  /* Build → upload to PULSE Documents/{project}/ → open in SharePoint.
     Falls back to a browser download if SPO upload is unavailable or fails. */
  async function exportProjectStatusPptxToSharePoint(proj, options) {
    const exportOptions = options || {};
    if (exportOptions.exportMode === "fullDeck") {
      return exportGroupStatusPptx([proj], Object.assign({}, exportOptions, {
        exportMode: "fullDeck",
        groupName: exportOptions.groupName || proj.name || "Project",
        groupType: "project"
      }));
    }
    const built = await exportProjectStatusPptx(proj);
    const { fileName, blob } = built;
    const canUpload = typeof isSharePointMode === "function"
      && isSharePointMode()
      && typeof sharePointAdapter !== "undefined"
      && typeof sharePointAdapter.ensureProjectDocumentsRoot === "function"
      && typeof sharePointAdapter.uploadPulseDocumentsFile === "function";

    if (!canUpload) {
      closeReservedFileWindow(exportOptions.popup);
      downloadPptxBlob(blob, fileName);
      return { mode: "download", fileName, blob };
    }

    try {
      const uploaded = await uploadStatusPptxToSharePoint(proj, blob, fileName);
      /* Some SharePoint adapters return a server-relative URL only.
         Build an absolute URL so the browser can open the stored file. */
      const fileUrl = uploaded.fileUrl
        || (uploaded.serverRelativeUrl
          ? window.location.origin + uploaded.serverRelativeUrl
          : "");
      if (fileUrl) openStoredFileByPolicy(fileUrl, uploaded.fileName || fileName, exportOptions.popup);
      else closeReservedFileWindow(exportOptions.popup);
      return {
        mode: "sharepoint",
        fileName: uploaded.fileName,
        blob,
        fileUrl,
        serverRelativeUrl: uploaded.serverRelativeUrl,
        projectFolderUrl: uploaded.projectFolderUrl
      };
    } catch (uploadError) {
      console.warn("PULSE: status PPTX SharePoint upload failed; downloading local copy", uploadError);
      closeReservedFileWindow(exportOptions.popup);
      downloadPptxBlob(blob, fileName);
      return {
        mode: "download",
        fileName,
        blob,
        uploadError
      };
    }
  }

  /* Build either one status slide or a complete briefing for a portfolio,
     program, end-item configuration, or project. */
  async function exportGroupStatusPptx(projects, config) {
    const cfg = config || {};
    const selectedProjects = Array.isArray(projects) ? projects.filter(Boolean) : [];
    const exportMode = cfg.exportMode || (cfg.reportLayout === "onePage" ? "projectSlide" : "fullDeck");
    const isProjectSlide = exportMode === "projectSlide";
    if (!selectedProjects.length) throw new Error("Select at least one project before exporting.");
    if (isProjectSlide && selectedProjects.length !== 1) {
      throw new Error("A project update export must contain exactly one project.");
    }
    const groupName = cfg.groupName || "Group";
    const PptxGenJS = pptxCtor();
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: "AEWTTR_4x3", width: SLIDE_W, height: SLIDE_H });
    pptx.layout = "AEWTTR_4x3";
    pptx.author = "AEWTTR PULSE";
    pptx.title = `${groupName} Status Overview`;
    pptx.subject = "AEWTTR project status briefing";

    const shapes = pptx.ShapeType || PptxGenJS.shapes || {};
    const S = {
      rect: shapes.rect || shapes.RECTANGLE || "rect",
      line: shapes.line || shapes.LINE || "line",
      ellipse: shapes.ellipse || shapes.oval || shapes.OVAL || "ellipse",
      diamond: shapes.diamond || shapes.DIAMOND || shapes.rect || "rect"
    };
    const logos = await loadBrandLogos();
    let page = 1;
    if (!isProjectSlide) {
      buildBriefingCoverSlide(pptx, S, logos, groupName, cfg.groupType, selectedProjects.length);
      page += 1;
    }

    /* Every output contains its project-status slide.  The concise output
       stops here; a full briefing can add focused appendices below. */
    for (const proj of selectedProjects) {
      const imageUrls = cfg.includePhotos === false ? [] : collectExportImages(proj);
      const imageDataUris = [];
      for (const url of imageUrls) {
        const data = await fetchAsDataUri(url);
        if (data) imageDataUris.push(data);
      }
      await buildStatusSlide(pptx, S, proj, logos, imageDataUris, page);
      page++;
    }

    /* Risk slides */
    if (!isProjectSlide && cfg.includeRisks) {
      const riskFilter = cfg.riskFilter || "open";
      const selectedRiskIds = new Set(Array.isArray(cfg.selectedRiskIds) ? cfg.selectedRiskIds : []);
      for (const proj of selectedProjects) {
        let risks = typeof projectRisks === "function" ? projectRisks(proj.id) : [];
        if (riskFilter === "open") risks = risks.filter(r => r.status !== "Closed");
        else if (riskFilter === "high") risks = risks.filter(r => r.status !== "Closed" && r.rating === "Red");
        if (selectedRiskIds.size) risks = risks.filter((risk) => selectedRiskIds.has(risk.id));
        risks = risks.sort((a, b) => (b.likelihood * b.impact) - (a.likelihood * a.impact)).slice(0, 3);
        for (const risk of risks) {
          await buildRiskSlide(pptx, S, proj, risk, logos, page, risks.length + page);
          page++;
        }
      }
    }

    /* Task table slides */
    if (!isProjectSlide && cfg.includeTasks) {
      const db = window.AEWTTR && window.AEWTTR.db;
      const taskFilter = cfg.taskFilter || "open";
      const selectedTaskIds = Array.isArray(cfg.selectedTaskIds) ? cfg.selectedTaskIds : [];
      const allTasks = [];
      selectedProjects.forEach(proj => {
        function flattenTasks(list) {
          (list || []).forEach(t => {
            const keep = taskFilter === "open" ? (t.status !== "Done" && t.status !== "Cancelled")
              : taskFilter === "high" ? ((t.priority === "Immediate" || t.priority === "High") && t.status !== "Done" && t.status !== "Cancelled")
              : true;
            if (keep && (!selectedTaskIds.length || selectedTaskIds.includes(t.id))) allTasks.push(Object.assign({}, t, { _projName: proj.name || "Untitled project" }));
            if (t.subtasks) flattenTasks(t.subtasks);
          });
        }
        flattenTasks((db && db.ganttTasks && db.ganttTasks[proj.id]) || []);
      });

      const ROWS = 12;
      const COL_X = [0.1, 3.2, 5.8, 7.2, 8.55];
      const COL_W = [3.0, 2.5, 1.35, 1.3, 1.4];
      const HEADERS = ["Task", "Project", "Assignee", "Priority", "Status"];
      const ROW_H = 0.35;

      for (let start = 0; start < Math.max(allTasks.length, 1); start += ROWS) {
        const chunk = allTasks.slice(start, start + ROWS);
        const slide = pptx.addSlide();
        slide.background = { color: "FFFFFF" };
        addStatusChrome(slide, S, logos, page);
        slide.addText("Open Work Summary", { x: 0.95, y: 0.15, w: 6.05, h: 0.55, color: "000000", fontSize: 22, bold: true, fontFace: "Arial Black", align: "center", margin: 0 });
        slide.addText(`${groupName} · ${todayLabel()}${allTasks.length > ROWS ? `  (${start + 1}–${Math.min(start + ROWS, allTasks.length)} of ${allTasks.length})` : ""}`, { x: 0.25, y: 0.83, w: 9.5, h: 0.24, color: "333333", fontSize: 10, fontFace: "Arial", align: "center", margin: 0 });
        slide.addShape(S.line, { x: 0.2, y: 1.12, w: 9.6, h: 0, line: { color: DIVIDER, width: 1.25 } });
        const TABLE_Y = 1.24;
        HEADERS.forEach((h, i) => {
          slide.addShape(S.rect, { x: COL_X[i], y: TABLE_Y, w: COL_W[i], h: ROW_H, fill: { color: LIGHT_BLUE } });
          slide.addText(h, { x: COL_X[i] + 0.04, y: TABLE_Y + 0.05, w: COL_W[i] - 0.06, h: ROW_H - 0.08, color: "000000", fontSize: 9, bold: true });
        });
        if (chunk.length) {
          chunk.forEach((t, ri) => {
            const rowY = TABLE_Y + ROW_H * (ri + 1);
            const fill = ri % 2 === 0 ? { color: "FFFFFF" } : { color: ROW_ALT };
            const vals = [
              String(t.name || t.title || "").slice(0, 55),
              String(t._projName || "").slice(0, 28),
              String(t.assignee || t.owner || "").slice(0, 22),
              String(t.priority || ""),
              String(t.status || "")
            ];
            vals.forEach((v, i) => {
              slide.addShape(S.rect, { x: COL_X[i], y: rowY, w: COL_W[i], h: ROW_H, fill });
              slide.addText(v, { x: COL_X[i] + 0.04, y: rowY + 0.05, w: COL_W[i] - 0.06, h: ROW_H - 0.08, color: "1F2937", fontSize: 8 });
            });
          });
        } else {
          slide.addText("No tasks match the selected filter.", { x: 0.1, y: TABLE_Y + ROW_H + 0.1, w: 9.8, h: 0.4, color: GRAY, fontSize: 10, italic: true });
        }
        page++;
        if (!chunk.length) break;
      }
    }

    const safeName = String(groupName).replace(/[^a-zA-Z0-9\-_]/g, "-").replace(/-{2,}/g, "-").slice(0, 40);
    const fileName = `${safeName}-${isProjectSlide ? "project-update" : "status-briefing"}-${new Date().toISOString().slice(0, 10)}.pptx`;
    const blob = await pptx.write({ outputType: "blob" });

    /* Try SharePoint upload; fall back to direct download */
    const canUpload = typeof isSharePointMode === "function" && isSharePointMode()
      && typeof sharePointAdapter !== "undefined"
      && typeof sharePointAdapter.resolvePulseDocumentsLibraryUrl === "function"
      && typeof sharePointAdapter.createPulseDocumentsSubfolder === "function"
      && typeof sharePointAdapter.uploadPulseDocumentsFile === "function";

    if (!canUpload) {
      closeReservedFileWindow(cfg.popup);
      downloadPptxBlob(blob, fileName);
      return { mode: "download", fileName, blob };
    }
    try {
      /* Group outputs are discoverable in PULSE Documents/Reporting/{scope}. */
      const uploaded = await uploadGroupStatusPptxToSharePoint(groupName, cfg.groupType, blob, fileName);
      const fileUrl = uploaded.fileUrl || (uploaded.serverRelativeUrl ? window.location.origin + uploaded.serverRelativeUrl : "");
      if (fileUrl) openStoredFileByPolicy(fileUrl, uploaded.name || fileName, cfg.popup);
      else closeReservedFileWindow(cfg.popup);
      return { mode: "sharepoint", fileName: uploaded.name || fileName, blob, fileUrl, serverRelativeUrl: uploaded.serverRelativeUrl, reportFolderUrl: uploaded.reportFolderUrl };
    } catch (uploadError) {
      console.warn("PULSE: group PPTX SharePoint upload failed; downloading local copy", uploadError);
      closeReservedFileWindow(cfg.popup);
      downloadPptxBlob(blob, fileName);
      return { mode: "download", fileName, blob, uploadError };
    }
  }

  window.AEWTTR.ProjectPptxExport = {
    exportProjectStatusPptx,
    exportProjectStatusPptxToSharePoint,
    exportGroupStatusPptx,
    downloadPptxBlob,
    uploadStatusPptxToSharePoint,
    uploadGroupStatusPptxToSharePoint,
    riskBurnDownSeries,
    ratingBand,
    collectMilestones,
    perRiskBurnPoints
  };
})();
