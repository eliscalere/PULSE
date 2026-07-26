/* PULSE Overview report generation and SharePoint delivery.
   The page supplies a normalized report model; this service owns the PDF
   layout, branded page chrome, file upload, and open-link behavior. */
(function () {
  "use strict";

  const REPORT_SEAL_SOURCE = "assets/images/aewttr-seal.png";
  const REPORT_LIBRARY_GROUP = "Overview Reports";
  const COLORS = {
    ink: [7, 7, 8],
    muted: [81, 84, 90],
    line: [217, 218, 221],
    soft: [247, 247, 248],
    blue: [47, 102, 255],
    red: [180, 35, 24],
    amber: [176, 101, 0],
    green: [43, 122, 94]
  };

  function safeText(value) {
    return String(value == null ? "" : value)
      .replace(/[\u2010-\u2015]/g, "-")
      .replace(/\u2026/g, "...")
      .replace(/\u00b7/g, " / ")
      .replace(/[“”]/g, "\"")
      .replace(/[‘’]/g, "'")
      .normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^\x20-\x7E]/g, "");
  }

  function safeFileSegment(value) {
    return safeText(value)
      .replace(/[^A-Za-z0-9._-]+/g, "-")
      .replace(/-+/g, "-")
      .replace(/^-|-$/g, "") || "overview";
  }

  function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(reader.error || new Error("Could not read the report seal."));
      reader.readAsDataURL(blob);
    });
  }

  async function loadSealDataUrl() {
    if (REPORT_SEAL_SOURCE.startsWith("data:")) return REPORT_SEAL_SOURCE;
    const response = await fetch(REPORT_SEAL_SOURCE, { credentials: "same-origin" });
    if (!response.ok) throw new Error("Could not load the report seal.");
    return blobToDataUrl(await response.blob());
  }

  function setTextColor(doc, color) {
    doc.setTextColor(color[0], color[1], color[2]);
  }

  function setFillColor(doc, color) {
    doc.setFillColor(color[0], color[1], color[2]);
  }

  function setDrawColor(doc, color) {
    doc.setDrawColor(color[0], color[1], color[2]);
  }

  function toneColor(tone) {
    if (tone === "danger") return COLORS.red;
    if (tone === "warning") return COLORS.amber;
    if (tone === "success") return COLORS.green;
    if (tone === "accent") return COLORS.blue;
    return COLORS.ink;
  }

  async function createPdfBlob(model) {
    const JsPdf = window.jspdf && window.jspdf.jsPDF;
    if (!JsPdf) throw new Error("The PDF engine is unavailable. Refresh PULSE and try again.");

    const sealDataUrl = await loadSealDataUrl();
    const doc = new JsPdf({ orientation: "landscape", unit: "pt", format: "letter", compress: true });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const marginX = 40;
    const contentWidth = pageWidth - (marginX * 2);
    const contentTop = 112;
    const contentBottom = pageHeight - 45;
    const generatedAt = safeText(model.generatedAt || new Date().toLocaleString());
    const generatedBy = safeText(model.generatedBy || "");
    const reportLabel = safeText(model.reportLabel || "Overview Report");
    let cursorY = contentTop;
    let currentSection = "";
    let currentSectionNote = "";
    let currentColumns = null;

    function addPageChrome() {
      setFillColor(doc, [255, 255, 255]);
      doc.rect(0, 0, pageWidth, pageHeight, "F");
      doc.addImage(sealDataUrl, "PNG", marginX, 22, 61, 60, undefined, "FAST");

      doc.setFont("helvetica", "bold");
      doc.setFontSize(19);
      setTextColor(doc, COLORS.ink);
      doc.text("PULSE", 116, 44, { charSpace: 4.2 });
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8.5);
      setTextColor(doc, COLORS.muted);
      doc.text(reportLabel.toUpperCase(), 116, 61, { charSpace: 1.1 });

      doc.setFont("helvetica", "bold");
      doc.setFontSize(9);
      setTextColor(doc, COLORS.ink);
      doc.text(safeText(model.scopeLabel || ""), pageWidth - marginX, 37, { align: "right" });
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8);
      setTextColor(doc, COLORS.muted);
      doc.text(`Generated ${generatedAt}${generatedBy ? ` by ${generatedBy}` : ""}`, pageWidth - marginX, 53, { align: "right" });

      setDrawColor(doc, COLORS.line);
      doc.setLineWidth(0.65);
      doc.line(marginX, 93, pageWidth - marginX, 93);
    }

    function addFooter(pageNumber) {
      setDrawColor(doc, COLORS.line);
      doc.setLineWidth(0.55);
      doc.line(marginX, pageHeight - 31, pageWidth - marginX, pageHeight - 31);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7.5);
      setTextColor(doc, COLORS.muted);
      doc.text(`Page ${pageNumber}`, marginX, pageHeight - 16);

      const brandX = pageWidth - marginX - 73;
      setFillColor(doc, COLORS.ink);
      doc.circle(brandX, pageHeight - 21, 2.7, "F");
      doc.circle(brandX + 8, pageHeight - 14, 2.7, "F");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(8.5);
      setTextColor(doc, COLORS.ink);
      doc.text("P U L S E", pageWidth - marginX, pageHeight - 15, { align: "right" });
    }

    function addNewPage() {
      doc.addPage("letter", "landscape");
      addPageChrome();
      cursorY = contentTop;
    }

    function ensureSpace(height, repeatTableHeader) {
      if (cursorY + height <= contentBottom) return;
      addNewPage();
      if (currentSection) addSectionHeading(currentSection, currentSectionNote, true);
      if (repeatTableHeader && currentColumns) drawTableHeader(currentColumns);
    }

    function addSectionHeading(title, note, continued) {
      const heading = safeText(title) + (continued ? " (continued)" : "");
      if (cursorY + (note ? 43 : 29) > contentBottom) addNewPage();
      doc.setFont("helvetica", "bold");
      doc.setFontSize(11.5);
      setTextColor(doc, COLORS.ink);
      doc.text(heading, marginX, cursorY + 11);
      setFillColor(doc, COLORS.blue);
      doc.rect(marginX, cursorY + 18, 34, 2.5, "F");
      cursorY += 27;
      if (note) {
        doc.setFont("helvetica", "normal");
        doc.setFontSize(8);
        setTextColor(doc, COLORS.muted);
        const noteLines = doc.splitTextToSize(safeText(note), contentWidth);
        doc.text(noteLines.slice(0, 2), marginX, cursorY);
        cursorY += Math.min(noteLines.length, 2) * 9 + 5;
      }
    }

    function addSummary(summary) {
      const items = Array.isArray(summary) ? summary : [];
      if (!items.length) return;
      const columns = Math.min(items.length, 6);
      const gap = 8;
      const cardWidth = (contentWidth - (gap * (columns - 1))) / columns;
      const cardHeight = 60;
      ensureSpace(cardHeight + 9, false);
      items.forEach((item, index) => {
        const x = marginX + (index % columns) * (cardWidth + gap);
        const row = Math.floor(index / columns);
        const y = cursorY + row * (cardHeight + gap);
        setFillColor(doc, COLORS.soft);
        doc.roundedRect(x, y, cardWidth, cardHeight, 5, 5, "F");
        setFillColor(doc, toneColor(item.tone));
        doc.rect(x, y, cardWidth, 3, "F");
        doc.setFont("helvetica", "bold");
        doc.setFontSize(8);
        setTextColor(doc, COLORS.muted);
        doc.text(safeText(item.label).toUpperCase(), x + 10, y + 18);
        doc.setFontSize(21);
        setTextColor(doc, toneColor(item.tone));
        doc.text(safeText(item.value), x + 10, y + 41);
        if (item.detail) {
          doc.setFont("helvetica", "normal");
          doc.setFontSize(7);
          setTextColor(doc, COLORS.muted);
          doc.text(safeText(item.detail), x + cardWidth - 10, y + 41, { align: "right" });
        }
      });
      cursorY += Math.ceil(items.length / columns) * (cardHeight + gap) + 9;
    }

    function normalizeColumns(columns) {
      const normalized = (columns || []).map((column) => {
        if (typeof column === "string") return { label: column, weight: 1 };
        return { label: column.label || "", weight: Number(column.weight) || 1, key: column.key };
      });
      const totalWeight = normalized.reduce((sum, column) => sum + column.weight, 0) || 1;
      let x = marginX;
      return normalized.map((column, index) => {
        const width = index === normalized.length - 1
          ? (pageWidth - marginX) - x
          : contentWidth * (column.weight / totalWeight);
        const result = { ...column, x, width };
        x += width;
        return result;
      });
    }

    function drawTableHeader(columns) {
      const height = 21;
      setFillColor(doc, COLORS.ink);
      doc.rect(marginX, cursorY, contentWidth, height, "F");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(7.2);
      setTextColor(doc, [255, 255, 255]);
      columns.forEach((column) => {
        doc.text(safeText(column.label).toUpperCase(), column.x + 6, cursorY + 14);
      });
      cursorY += height;
    }

    function addTable(section) {
      currentSection = safeText(section.title || "Details");
      currentSectionNote = safeText(section.note || "");
      currentColumns = normalizeColumns(section.columns);
      addSectionHeading(currentSection, currentSectionNote, false);
      drawTableHeader(currentColumns);

      const rows = Array.isArray(section.rows) ? section.rows : [];
      if (!rows.length) {
        doc.setFont("helvetica", "italic");
        doc.setFontSize(8.5);
        setTextColor(doc, COLORS.muted);
        doc.text(safeText(section.emptyMessage || "No records in this view."), marginX + 6, cursorY + 18);
        cursorY += 29;
        return;
      }

      rows.forEach((row, rowIndex) => {
        const values = Array.isArray(row) ? row : currentColumns.map((column) => row[column.key]);
        const wrapped = currentColumns.map((column, columnIndex) => {
          const value = safeText(values[columnIndex]);
          return doc.splitTextToSize(value || "-", Math.max(20, column.width - 12)).slice(0, 5);
        });
        const lineCount = Math.max(1, ...wrapped.map((lines) => lines.length));
        const rowHeight = Math.max(20, (lineCount * 8.2) + 8);
        ensureSpace(rowHeight, true);
        if (rowIndex % 2 === 1) {
          setFillColor(doc, COLORS.soft);
          doc.rect(marginX, cursorY, contentWidth, rowHeight, "F");
        }
        setDrawColor(doc, COLORS.line);
        doc.setLineWidth(0.35);
        doc.line(marginX, cursorY + rowHeight, pageWidth - marginX, cursorY + rowHeight);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7.5);
        setTextColor(doc, COLORS.ink);
        wrapped.forEach((lines, columnIndex) => {
          doc.text(lines, currentColumns[columnIndex].x + 6, cursorY + 12, { lineHeightFactor: 1.1 });
        });
        cursorY += rowHeight;
      });
      cursorY += 13;
    }

    addPageChrome();
    doc.setFont("helvetica", "bold");
    doc.setFontSize(22);
    setTextColor(doc, COLORS.ink);
    doc.text(safeText(model.title || "Overview Report"), marginX, cursorY + 18);
    cursorY += 28;
    if (model.subtitle) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      setTextColor(doc, COLORS.muted);
      const subtitleLines = doc.splitTextToSize(safeText(model.subtitle), contentWidth);
      doc.text(subtitleLines.slice(0, 3), marginX, cursorY);
      cursorY += Math.min(subtitleLines.length, 3) * 11 + 12;
    }
    addSummary(model.summary);
    (model.sections || []).forEach(addTable);

    const totalPages = doc.getNumberOfPages();
    for (let pageNumber = 1; pageNumber <= totalPages; pageNumber++) {
      doc.setPage(pageNumber);
      addFooter(pageNumber);
    }
    return doc.output("blob");
  }

  function downloadBlob(blob, fileName) {
    const href = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = href;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(href), 2000);
  }

  function openUploadedReport(popup, href, fileName, mimeType) {
    if (!href) return;
    const api = window.AEWTTR && window.AEWTTR.OfficeDesktop;
    if (api && typeof api.openSharePointFileByPolicy === "function") {
      api.openSharePointFileByPolicy(href, fileName, mimeType, { popup });
      return;
    }
    if (popup && !popup.closed) {
      popup.location.replace(href);
      return;
    }
    window.open(href, "_blank", "noopener");
  }

  async function generateStoreAndOpen(model, options) {
    const opts = options || {};
    const blob = await createPdfBlob(model);
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const fileName = `${safeFileSegment(model.fileBaseName || "PULSE-Overview-Report")}-${stamp}.pdf`;
    const sharePointReady = typeof isSharePointMode === "function" &&
      isSharePointMode() &&
      typeof sharePointAdapter !== "undefined" &&
      typeof sharePointAdapter.uploadProjectDocument === "function";

    if (!sharePointReady) {
      if (opts.popup && !opts.popup.closed) opts.popup.close();
      downloadBlob(blob, fileName);
      return { mode: "local", blob, fileName, fileUrl: "" };
    }

    const file = new File([blob], fileName, { type: "application/pdf", lastModified: Date.now() });
    try {
      const result = await sharePointAdapter.uploadProjectDocument(
        currentSiteUrl(),
        REPORT_LIBRARY_GROUP,
        safeFileSegment(model.storageFolder || model.scopeLabel || "Overview"),
        file
      );
      openUploadedReport(opts.popup, result.fileUrl, result.fileName || fileName, file.type);
      return { mode: "sharepoint", blob, fileName: result.fileName || fileName, fileUrl: result.fileUrl || "" };
    } catch (error) {
      if (opts.popup && !opts.popup.closed) opts.popup.close();
      downloadBlob(blob, fileName);
      error.localCopyDownloaded = true;
      throw error;
    }
  }

  window.AEWTTR = window.AEWTTR || {};
  window.AEWTTR.OverviewReportService = {
    createPdfBlob,
    generateStoreAndOpen
  };
})();
