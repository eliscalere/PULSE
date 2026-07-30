window.AEWTTR = window.AEWTTR || {};

function pulseDownloadBlob(blob, fileName) {
  const href = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = href;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(href), 2000);
}

function pulseExportOfficeApi() {
  return window.AEWTTR && window.AEWTTR.OfficeDesktop;
}

function pulseReserveExportWindow(fileName, mimeType) {
  const api = pulseExportOfficeApi();
  if (api && typeof api.reserveSharePointFileWindow === "function"
    && typeof isSharePointMode === "function" && isSharePointMode()) {
    return api.reserveSharePointFileWindow(fileName, mimeType);
  }
  return null;
}

function pulseCloseExportWindow(popup) {
  const api = pulseExportOfficeApi();
  if (api && typeof api.closeReservedSharePointFileWindow === "function") {
    api.closeReservedSharePointFileWindow(popup);
  } else if (popup && !popup.closed) {
    popup.close();
  }
}

async function pulseUploadGeneratedFile(file, folderName) {
  if (typeof isSharePointMode !== "function" || !isSharePointMode()
    || typeof sharePointAdapter === "undefined"
    || typeof sharePointAdapter.resolvePulseDocumentsLibraryUrl !== "function"
    || typeof sharePointAdapter.createPulseDocumentsSubfolder !== "function"
    || typeof sharePointAdapter.uploadPulseDocumentsFile !== "function") {
    return null;
  }
  const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
  if (!siteUrl) return null;
  const libraryUrl = await sharePointAdapter.resolvePulseDocumentsLibraryUrl(siteUrl);
  const reportsRoot = await sharePointAdapter.createPulseDocumentsSubfolder(siteUrl, libraryUrl, "Reporting");
  const exportFolder = await sharePointAdapter.createPulseDocumentsSubfolder(
    siteUrl,
    reportsRoot.serverRelativeUrl,
    folderName || "Exports"
  );
  return sharePointAdapter.uploadPulseDocumentsFile(siteUrl, exportFolder.serverRelativeUrl, file);
}

function pulseXlsxEscape(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function pulseXlsxColumnName(index) {
  let value = Number(index) + 1;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function pulseXlsxCell(ref, value, styleId) {
  const style = Number.isFinite(Number(styleId)) ? ` s="${Number(styleId)}"` : "";
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${ref}"${style} t="n"><v>${value}</v></c>`;
  }
  if (typeof value === "boolean") {
    return `<c r="${ref}"${style} t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  const text = pulseXlsxEscape(value);
  return `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`;
}

function pulseXlsxRow(rowNumber, cells, options = {}) {
  const height = options.height ? ` ht="${options.height}" customHeight="1"` : "";
  return `<row r="${rowNumber}"${height}>${cells.join("")}</row>`;
}

function pulseXlsxStatusStyle(header, value, fallbackStyle) {
  const field = String(header || "").trim().toLowerCase();
  if (!/^(status|derived status|health|likelihood|impact)$/.test(field)) return fallbackStyle;
  const text = String(value || "").trim().toLowerCase();
  if (/block|off track|critical|high|overdue|failed/.test(text)) return 11;
  if (/risk|attention|medium|pending|delayed/.test(text)) return 10;
  if (/track|healthy|complete|completed|done|active|low|approved/.test(text)) return 9;
  if (/review/.test(text)) return 12;
  return fallbackStyle;
}

function pulseXlsxColumnWidths(columns, rows) {
  return columns.map((column, index) => {
    let longest = Math.max(10, String(column || "").length + 2);
    rows.slice(0, 100).forEach((row) => {
      longest = Math.max(longest, String(row && row[index] != null ? row[index] : "").length + 2);
    });
    return Math.max(11, Math.min(42, Math.round(longest * 0.88)));
  });
}

function pulseXlsxStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="11">
    <font><sz val="10"/><color rgb="FF070708"/><name val="Inter"/><family val="2"/></font>
    <font><b/><sz val="9"/><color rgb="FFFFFFFF"/><name val="Inter"/><family val="2"/></font>
    <font><b/><sz val="22"/><color rgb="FFFFFFFF"/><name val="Inter"/><family val="2"/></font>
    <font><sz val="9"/><color rgb="FFC9CCD2"/><name val="Inter"/><family val="2"/></font>
    <font><b/><sz val="9"/><color rgb="FF51545A"/><name val="Inter"/><family val="2"/></font>
    <font><sz val="9"/><color rgb="FF070708"/><name val="Inter"/><family val="2"/></font>
    <font><b/><sz val="9"/><color rgb="FF2B7A5E"/><name val="Inter"/><family val="2"/></font>
    <font><b/><sz val="9"/><color rgb="FFC77800"/><name val="Inter"/><family val="2"/></font>
    <font><b/><sz val="9"/><color rgb="FFB42318"/><name val="Inter"/><family val="2"/></font>
    <font><b/><sz val="9"/><color rgb="FF2F66FF"/><name val="Inter"/><family val="2"/></font>
    <font><b/><sz val="9"/><color rgb="FF2F66FF"/><name val="Inter"/><family val="2"/></font>
  </fonts>
  <fills count="10">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF070708"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFECEDEF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8F9FA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEAF5F1"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF4E2"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFDECEA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEAF0FF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFDADDE2"/></left>
      <right style="thin"><color rgb="FFDADDE2"/></right>
      <top style="thin"><color rgb="FFDADDE2"/></top>
      <bottom style="thin"><color rgb="FFDADDE2"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="16">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="10" fillId="3" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="5" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="6" fillId="6" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="7" fillId="7" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="8" fillId="8" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="9" fillId="9" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
}

function pulseXlsxThemeXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="PULSE">
  <a:themeElements>
    <a:clrScheme name="PULSE">
      <a:dk1><a:srgbClr val="070708"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="51545A"/></a:dk2><a:lt2><a:srgbClr val="ECEDEF"/></a:lt2>
      <a:accent1><a:srgbClr val="2F66FF"/></a:accent1><a:accent2><a:srgbClr val="2B7A5E"/></a:accent2>
      <a:accent3><a:srgbClr val="C77800"/></a:accent3><a:accent4><a:srgbClr val="B42318"/></a:accent4>
      <a:accent5><a:srgbClr val="73767C"/></a:accent5><a:accent6><a:srgbClr val="DADDE2"/></a:accent6>
      <a:hlink><a:srgbClr val="2F66FF"/></a:hlink><a:folHlink><a:srgbClr val="1F4FD0"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="PULSE"><a:majorFont><a:latin typeface="Inter"/></a:majorFont><a:minorFont><a:latin typeface="Inter"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="PULSE">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

async function pulseBuildStyledWorkbook(columns, dataRows, options = {}) {
  const ZipCtor = window.JSZip || (typeof JSZip !== "undefined" ? JSZip : null);
  if (!ZipCtor) throw new Error("The Excel export engine is unavailable. Refresh PULSE and try again.");

  const safeColumns = (Array.isArray(columns) ? columns : []).map((value) => String(value == null ? "" : value));
  if (!safeColumns.length) throw new Error("The workbook export has no columns.");
  const safeRows = Array.isArray(dataRows) ? dataRows : [];
  const lastColumn = pulseXlsxColumnName(safeColumns.length - 1);
  const widths = pulseXlsxColumnWidths(safeColumns, safeRows);
  const dataStartRow = 9;
  const footerRow = Math.max(dataStartRow + safeRows.length + 1, 11);
  const finalRow = footerRow;
  const now = new Date();
  const monthYear = now.toLocaleDateString("en-US", { month: "long", year: "numeric" }).toUpperCase();
  const title = options.title || "PULSE Data Export";
  const subtitle = options.subtitle || `Generated ${now.toLocaleString()}`;
  const sectionLabel = String(options.sectionLabel || "DATA EXPORT").toUpperCase();
  const footerLabel = String(options.footerLabel || title).toUpperCase();
  const sheetName = String(options.sheetName || "PULSE Export").replace(/[\\/*?:[\]]/g, " ").slice(0, 31) || "PULSE Export";
  const sheetRows = [];

  function bandRow(rowNumber, text, styleId, height) {
    const cells = [];
    for (let index = 0; index < safeColumns.length; index += 1) {
      cells.push(pulseXlsxCell(`${pulseXlsxColumnName(index)}${rowNumber}`, index === 0 ? text : "", index === 0 ? styleId : 4));
    }
    sheetRows.push(pulseXlsxRow(rowNumber, cells, { height }));
  }

  bandRow(1, "CUI", 1, 16);
  bandRow(2, title, 2, 32);
  bandRow(3, "", 4, 8);
  bandRow(4, subtitle, 3, 20);
  bandRow(5, "", 4, 16);
  sheetRows.push(pulseXlsxRow(6, [], { height: 8 }));
  const sectionCells = [pulseXlsxCell("A7", sectionLabel, 5)];
  sheetRows.push(pulseXlsxRow(7, sectionCells, { height: 18 }));
  sheetRows.push(pulseXlsxRow(8, safeColumns.map((column, index) => pulseXlsxCell(`${pulseXlsxColumnName(index)}8`, column, 6)), { height: 28 }));

  safeRows.forEach((row, rowIndex) => {
    const rowNumber = dataStartRow + rowIndex;
    const baseStyle = rowIndex % 2 === 0 ? 7 : 8;
    const numericStyle = rowIndex % 2 === 0 ? 14 : 15;
    const cells = safeColumns.map((column, columnIndex) => {
      const value = row && row[columnIndex] != null ? row[columnIndex] : "";
      let style = typeof value === "number" ? numericStyle : baseStyle;
      style = pulseXlsxStatusStyle(column, value, style);
      return pulseXlsxCell(`${pulseXlsxColumnName(columnIndex)}${rowNumber}`, value, style);
    });
    sheetRows.push(pulseXlsxRow(rowNumber, cells, { height: 23 }));
  });

  if (footerRow > dataStartRow + safeRows.length) {
    sheetRows.push(pulseXlsxRow(footerRow - 1, [], { height: 8 }));
  }
  const footerText = `CUI  |  ${footerLabel}  |  SCHEMA 1.0  |  ${monthYear}`;
  bandRow(footerRow, footerText, 13, 18);

  const columnsXml = widths.map((width, index) =>
    `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`
  ).join("");
  const dataEndRow = Math.max(8, dataStartRow + safeRows.length - 1);
  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:${lastColumn}${finalRow}"/>
  <sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0"><pane ySplit="8" topLeftCell="A9" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A9" sqref="A9"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>${columnsXml}</cols>
  <sheetData>${sheetRows.join("")}</sheetData>
  <autoFilter ref="A8:${lastColumn}${dataEndRow}"/>
  <mergeCells count="5">
    <mergeCell ref="A1:${lastColumn}1"/><mergeCell ref="A2:${lastColumn}2"/><mergeCell ref="A4:${lastColumn}4"/>
    <mergeCell ref="A7:${lastColumn}7"/><mergeCell ref="A${footerRow}:${lastColumn}${footerRow}"/>
  </mergeCells>
  <pageMargins left="0.25" right="0.25" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>
  <pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0" paperSize="9"/>
</worksheet>`;

  const created = now.toISOString();
  const zip = new ZipCtor();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${pulseXlsxEscape(title)}</dc:title><dc:creator>PULSE</dc:creator><cp:lastModifiedBy>PULSE</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>
</cp:coreProperties>`);
  zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>PULSE</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>${pulseXlsxEscape(sheetName)}</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`);
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl"/><workbookPr defaultThemeVersion="164011"/>
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14000"/></bookViews>
  <sheets><sheet name="${pulseXlsxEscape(sheetName)}" sheetId="1" state="visible" r:id="rId1"/></sheets>
  <calcPr calcId="191029"/>
</workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);
  zip.file("xl/worksheets/sheet1.xml", worksheetXml);
  zip.file("xl/styles.xml", pulseXlsxStylesXml());
  zip.file("xl/theme/theme1.xml", pulseXlsxThemeXml());
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
    compressionOptions: { level: 6 }
  });
}

async function pulseFetchImageBytes(url) {
  const response = await fetch(url, { credentials: "same-origin" });
  if (!response.ok) throw new Error(`Could not load image asset: ${url}`);
  return new Uint8Array(await response.arrayBuffer());
}

const PULSE_XLSX_CHART_PALETTE = ["2F66FF", "2B7A5E", "C77800", "B42318", "73767C", "9B59B6", "16A085", "8E44AD"];

function pulseXlsxChartTitleXml(title) {
  if (!title) return "<c:autoTitleDeleted val=\"1\"/>";
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="070708"/></a:solidFill><a:latin typeface="Inter"/></a:defRPr></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="070708"/></a:solidFill><a:latin typeface="Inter"/></a:rPr><a:t>${pulseXlsxEscape(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
}

function pulseXlsxChartStrLit(values) {
  return `<c:strLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${pulseXlsxEscape(v)}</c:v></c:pt>`).join("")}</c:strLit>`;
}

function pulseXlsxChartNumLit(values) {
  return `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${Number(v) || 0}</c:v></c:pt>`).join("")}</c:numLit>`;
}

function pulseXlsxAxTextXml() {
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="51545A"/></a:solidFill><a:latin typeface="Inter"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

function pulseXlsxChartTx(name) {
  return `<c:tx><c:v>${pulseXlsxEscape(name)}</c:v></c:tx>`;
}

/* Builds a real, native Excel chart part (c:chartSpace XML) — a bar/column,
   pie, or doughnut chart with its data cached inline via c:numLit/c:strLit.
   This is what Excel itself writes when you select cells and insert a
   chart, so unlike a rendered image it stays crisp at any zoom, is themed
   like the rest of the workbook, and can be edited/recolored by the user.
   No embedded workbook / cell reference is needed for numLit/strLit data. */
function pulseXlsxChartXml(def) {
  const type = def.type === "pie" || def.type === "doughnut" ? def.type : "bar";
  const categories = (def.categories || []).map((c) => String(c == null ? "" : c));
  const series = Array.isArray(def.series) && def.series.length ? def.series : [{ name: "Series 1", values: [] }];
  const catAxId = 111111111;
  const valAxId = 222222222;

  if (type === "pie" || type === "doughnut") {
    const s = series[0] || { values: [] };
    const colors = def.colors && def.colors.length ? def.colors : PULSE_XLSX_CHART_PALETTE;
    const dPts = categories.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${colors[i % colors.length]}"/></a:solidFill></c:spPr></c:dPt>`).join("");
    const chartEl = type === "doughnut"
      ? `<c:doughnutChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${pulseXlsxChartTx(s.name || "Series 1")}${dPts}<c:cat>${pulseXlsxChartStrLit(categories)}</c:cat><c:val>${pulseXlsxChartNumLit(s.values)}</c:val></c:ser><c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`
      : `<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${pulseXlsxChartTx(s.name || "Series 1")}${dPts}<c:cat>${pulseXlsxChartStrLit(categories)}</c:cat><c:val>${pulseXlsxChartNumLit(s.values)}</c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    ${pulseXlsxChartTitleXml(def.title)}
    <c:plotArea><c:layout/>${chartEl}</c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/>${pulseXlsxAxTextXml()}</c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
</c:chartSpace>`;
  }

  const colors = def.colors && def.colors.length ? def.colors : PULSE_XLSX_CHART_PALETTE;
  const pointColors = def.pointColors && def.pointColors.length ? def.pointColors : null;
  const serXml = series.map((s, idx) => {
    const dPts = (pointColors && series.length === 1)
      ? categories.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${pointColors[i % pointColors.length]}"/></a:solidFill></c:spPr></c:dPt>`).join("")
      : "";
    return `<c:ser>
    <c:idx val="${idx}"/><c:order val="${idx}"/>
    ${pulseXlsxChartTx(s.name || `Series ${idx + 1}`)}
    ${dPts ? "" : `<c:spPr><a:solidFill><a:srgbClr val="${(colors[idx % colors.length])}"/></a:solidFill></c:spPr>`}
    ${dPts}
    <c:cat>${pulseXlsxChartStrLit(categories)}</c:cat>
    <c:val>${pulseXlsxChartNumLit(s.values)}</c:val>
  </c:ser>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    ${pulseXlsxChartTitleXml(def.title)}
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="${def.horizontal ? "bar" : "col"}"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="${pointColors && series.length === 1 ? "1" : "0"}"/>
        ${serXml}
        <c:gapWidth val="60"/>
        <c:axId val="${catAxId}"/><c:axId val="${valAxId}"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="${catAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>
        <c:axPos val="${def.horizontal ? "l" : "b"}"/>${pulseXlsxAxTextXml()}<c:crossAx val="${valAxId}"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="${valAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>
        <c:axPos val="${def.horizontal ? "b" : "l"}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="ECEDEF"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
        <c:numFmt formatCode="General" sourceLinked="0"/>${pulseXlsxAxTextXml()}<c:crossAx val="${catAxId}"/>
      </c:valAx>
    </c:plotArea>
    ${series.length > 1 ? `<c:legend><c:legendPos val="b"/><c:overlay val="0"/>${pulseXlsxAxTextXml()}</c:legend>` : ""}
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
</c:chartSpace>`;
}

/* Multi-sheet workbook builder — supports several tabs, each optionally with
   a data table and/or embedded pictures (brand imagery, rendered chart PNGs).
   Reuses the single-sheet builder's cell/row/style primitives; the sheet,
   drawing, and relationship plumbing is generalized to N sheets here. */
async function pulseBuildMultiSheetWorkbook(sheetDefs, options = {}) {
  const ZipCtor = window.JSZip || (typeof JSZip !== "undefined" ? JSZip : null);
  if (!ZipCtor) throw new Error("The Excel export engine is unavailable. Refresh PULSE and try again.");
  const defs = (Array.isArray(sheetDefs) ? sheetDefs : []).filter(Boolean);
  if (!defs.length) throw new Error("The workbook export has no sheets.");

  const now = new Date();
  const monthYear = now.toLocaleDateString("en-US", { month: "long", year: "numeric" }).toUpperCase();
  const created = now.toISOString();
  const workbookTitle = options.title || "PULSE Data Export";

  const zip = new ZipCtor();
  let mediaIndex = 0;
  let drawingIndex = 0;
  let chartIndex = 0;
  const chartOverrides = [];
  const sheetMeta = [];

  defs.forEach((def, sheetIdx) => {
    const sheetNumber = sheetIdx + 1;
    const sheetName = String(def.name || `Sheet${sheetNumber}`).replace(/[\\/*?:[\]]/g, " ").slice(0, 31) || `Sheet${sheetNumber}`;
    const columns = Array.isArray(def.columns) ? def.columns.map((v) => String(v == null ? "" : v)) : [];
    const rows = Array.isArray(def.rows) ? def.rows : [];
    const hasTable = columns.length > 0;
    const charts = Array.isArray(def.charts) ? def.charts.filter(Boolean) : [];
    const images = Array.isArray(def.images) ? def.images.filter(Boolean) : [];

    const colCount = Math.max(hasTable ? columns.length : 0, def.minCols || 7);
    const lastColumn = pulseXlsxColumnName(colCount - 1);
    const widths = hasTable ? pulseXlsxColumnWidths(columns, rows) : Array.from({ length: colCount }, () => 15);

    const title = def.title || workbookTitle;
    const subtitle = def.subtitle || "";
    const sectionLabel = String(def.sectionLabel || sheetName).toUpperCase();
    const footerLabel = String(def.footerLabel || workbookTitle).toUpperCase();

    const sheetRows = [];
    function bandRow(rowNumber, text, styleId, height) {
      const cells = [];
      for (let index = 0; index < colCount; index += 1) {
        cells.push(pulseXlsxCell(`${pulseXlsxColumnName(index)}${rowNumber}`, index === 0 ? text : "", index === 0 ? styleId : 4));
      }
      sheetRows.push(pulseXlsxRow(rowNumber, cells, { height }));
    }

    bandRow(1, "CUI", 1, 16);
    bandRow(2, title, 2, 32);
    bandRow(3, "", 4, 8);
    bandRow(4, subtitle, 3, 20);
    bandRow(5, "", 4, 16);
    sheetRows.push(pulseXlsxRow(6, [], { height: 8 }));

    const mergeCellRefs = [`A1:${lastColumn}1`, `A2:${lastColumn}2`, `A4:${lastColumn}4`];
    let dataEndRow = 6;
    let nextRow = 7;

    if (hasTable) {
      sheetRows.push(pulseXlsxRow(7, [pulseXlsxCell("A7", sectionLabel, 5)], { height: 18 }));
      mergeCellRefs.push(`A7:${lastColumn}7`);
      sheetRows.push(pulseXlsxRow(8, columns.map((column, index) => pulseXlsxCell(`${pulseXlsxColumnName(index)}8`, column, 6)), { height: 28 }));
      rows.forEach((row, rowIndex) => {
        const rowNumber = 9 + rowIndex;
        const baseStyle = rowIndex % 2 === 0 ? 7 : 8;
        const numericStyle = rowIndex % 2 === 0 ? 14 : 15;
        const cells = columns.map((column, columnIndex) => {
          const value = row && row[columnIndex] != null ? row[columnIndex] : "";
          let style = typeof value === "number" ? numericStyle : baseStyle;
          style = pulseXlsxStatusStyle(column, value, style);
          return pulseXlsxCell(`${pulseXlsxColumnName(columnIndex)}${rowNumber}`, value, style);
        });
        sheetRows.push(pulseXlsxRow(rowNumber, cells, { height: 23 }));
      });
      if (!rows.length && def.emptyMessage) {
        sheetRows.push(pulseXlsxRow(9, [pulseXlsxCell("A9", def.emptyMessage, 7)], { height: 23 }));
        mergeCellRefs.push(`A9:${lastColumn}9`);
        dataEndRow = 9;
      } else {
        dataEndRow = Math.max(8, 9 + rows.length - 1);
      }
      nextRow = dataEndRow + 2;
    }

    const footerRow = Math.max(nextRow, dataEndRow + 2);
    bandRow(footerRow, `CUI  |  ${footerLabel}  |  SCHEMA 1.0  |  ${monthYear}`, 13, 18);
    mergeCellRefs.push(`A${footerRow}:${lastColumn}${footerRow}`);

    let sheetHasDrawing = false;
    let finalRow = footerRow;
    if (charts.length || images.length) {
      sheetHasDrawing = true;
      drawingIndex += 1;
      const thisDrawingIndex = drawingIndex;
      let anchorRow = footerRow + 2;
      let nextRelId = 1;
      const anchorXml = [];
      const drawingRels = [];

      charts.forEach((chartDef, chartIdx) => {
        chartIndex += 1;
        const thisChartIndex = chartIndex;
        zip.file(`xl/charts/chart${thisChartIndex}.xml`, pulseXlsxChartXml(chartDef));
        chartOverrides.push(`<Override PartName="/xl/charts/chart${thisChartIndex}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
        const relId = `rId${nextRelId}`;
        nextRelId += 1;
        drawingRels.push(`<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${thisChartIndex}.xml"/>`);
        const pxW = chartDef.pxW || 560;
        const pxH = chartDef.pxH || 320;
        const emuW = Math.round(pxW * 9525);
        const emuH = Math.round(pxH * 9525);
        const fromRow = anchorRow - 1;
        anchorXml.push(`<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${emuW}" cy="${emuH}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${chartIdx + 1000}" name="Chart${chartIdx + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emuW}" cy="${emuH}"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`);
        anchorRow += Math.ceil(pxH / 20) + 2;
        finalRow = Math.max(finalRow, fromRow + Math.ceil(pxH / 20) + 1);
      });

      images.forEach((img, imgIdx) => {
        mediaIndex += 1;
        const mediaName = `image${mediaIndex}.png`;
        zip.file(`xl/media/${mediaName}`, img.bytes);
        const relId = `rId${nextRelId}`;
        nextRelId += 1;
        drawingRels.push(`<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`);
        const pxW = img.pxW || 480;
        const pxH = img.pxH || 260;
        const emuW = Math.round(pxW * 9525);
        const emuH = Math.round(pxH * 9525);
        const fromRow = anchorRow - 1;
        anchorXml.push(`<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${emuW}" cy="${emuH}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${imgIdx + 1}" name="Image${imgIdx + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emuW}" cy="${emuH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`);
        anchorRow += Math.ceil(pxH / 20) + 2;
        finalRow = Math.max(finalRow, fromRow + Math.ceil(pxH / 20) + 1);
      });

      zip.file(`xl/drawings/drawing${thisDrawingIndex}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchorXml.join("")}</xdr:wsDr>`);
      zip.file(`xl/drawings/_rels/drawing${thisDrawingIndex}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRels.join("")}</Relationships>`);
      zip.file(`xl/worksheets/_rels/sheet${sheetNumber}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${thisDrawingIndex}.xml"/>
</Relationships>`);
    }

    const columnsXml = widths.map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`).join("");
    const mergeCellsXml = `<mergeCells count="${mergeCellRefs.length}">${mergeCellRefs.map((r) => `<mergeCell ref="${r}"/>`).join("")}</mergeCells>`;
    const autoFilterXml = hasTable ? `<autoFilter ref="A8:${lastColumn}${dataEndRow}"/>` : "";
    const paneXml = hasTable ? `<pane ySplit="8" topLeftCell="A9" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A9" sqref="A9"/>` : "";
    const drawingTagXml = sheetHasDrawing ? `<drawing r:id="rIdDrawing"/>` : "";
    const tabColorXml = def.tabColor ? `<sheetPr><tabColor rgb="FF${String(def.tabColor).replace(/^#/, "").toUpperCase()}"/></sheetPr>` : "";

    const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${tabColorXml}
  <dimension ref="A1:${lastColumn}${finalRow}"/>
  <sheetViews><sheetView showGridLines="0"${sheetIdx === 0 ? " tabSelected=\"1\"" : ""} workbookViewId="0">${paneXml}</sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>${columnsXml}</cols>
  <sheetData>${sheetRows.join("")}</sheetData>
  ${autoFilterXml}
  ${mergeCellsXml}
  <pageMargins left="0.25" right="0.25" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>
  <pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0" paperSize="9"/>
  ${drawingTagXml}
</worksheet>`;

    zip.file(`xl/worksheets/sheet${sheetNumber}.xml`, worksheetXml);
    sheetMeta.push({ name: sheetName, sheetNumber, hasDrawing: sheetHasDrawing, drawingIndex: sheetHasDrawing ? drawingIndex : null });
  });

  const sheetOverrides = sheetMeta.map((s) => `<Override PartName="/xl/worksheets/sheet${s.sheetNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  const drawingOverrides = sheetMeta.filter((s) => s.hasDrawing).map((s) => `<Override PartName="/xl/drawings/drawing${s.drawingIndex}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join("");
  const hasAnyImages = mediaIndex > 0;

  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${hasAnyImages ? `<Default Extension="png" ContentType="image/png"/>` : ""}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetOverrides}
  ${drawingOverrides}
  ${chartOverrides.join("")}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);

  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);

  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${pulseXlsxEscape(workbookTitle)}</dc:title><dc:creator>PULSE</dc:creator><cp:lastModifiedBy>PULSE</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>
</cp:coreProperties>`);

  zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>PULSE</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheetMeta.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="${sheetMeta.length}" baseType="lpstr">${sheetMeta.map((s) => `<vt:lpstr>${pulseXlsxEscape(s.name)}</vt:lpstr>`).join("")}</vt:vector></TitlesOfParts>
</Properties>`);

  const sheetsXml = sheetMeta.map((s, idx) => `<sheet name="${pulseXlsxEscape(s.name)}" sheetId="${idx + 1}" state="visible" r:id="rId${idx + 1}"/>`).join("");
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl"/><workbookPr defaultThemeVersion="164011"/>
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14000"/></bookViews>
  <sheets>${sheetsXml}</sheets>
  <calcPr calcId="191029"/>
</workbook>`);

  const stylesRid = `rId${sheetMeta.length + 1}`;
  const themeRid = `rId${sheetMeta.length + 2}`;
  const sheetRelEntries = sheetMeta.map((s, idx) => `<Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${s.sheetNumber}.xml"/>`).join("");
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetRelEntries}
  <Relationship Id="${stylesRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="${themeRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);

  zip.file("xl/styles.xml", pulseXlsxStylesXml());
  zip.file("xl/theme/theme1.xml", pulseXlsxThemeXml());

  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
    compressionOptions: { level: 6 }
  });
}

window.AEWTTR.ExportService = {
  reserveOpen(fileName, mimeType) {
    return pulseReserveExportWindow(fileName, mimeType);
  },

  buildXlsx(columns, dataRows, options = {}) {
    return pulseBuildStyledWorkbook(columns, dataRows, options);
  },

  fetchImageBytes(url) {
    return pulseFetchImageBytes(url);
  },

  async exportMultiSheetXlsx(filename, sheetDefs, options = {}) {
    const blob = await pulseBuildMultiSheetWorkbook(sheetDefs, options);
    const fileName = filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`;
    const popup = options.popup || null;
    const mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    const file = new File([blob], fileName, { type: mimeType, lastModified: Date.now() });
    try {
      const uploaded = await pulseUploadGeneratedFile(file, options.folderName || "Excel Exports");
      if (uploaded && uploaded.fileUrl) {
        const api = pulseExportOfficeApi();
        if (api && typeof api.openSharePointFileByPolicy === "function") {
          api.openSharePointFileByPolicy(uploaded.fileUrl, uploaded.fileName || fileName, mimeType, { popup });
        } else {
          window.open(uploaded.fileUrl, "_blank", "noopener");
        }
        return {
          mode: "sharepoint",
          blob,
          fileName: uploaded.fileName || fileName,
          fileUrl: uploaded.fileUrl,
          serverRelativeUrl: uploaded.serverRelativeUrl || ""
        };
      }
      pulseCloseExportWindow(popup);
      pulseDownloadBlob(blob, fileName);
      return { mode: "download", blob, fileName, fileUrl: "" };
    } catch (uploadError) {
      pulseCloseExportWindow(popup);
      pulseDownloadBlob(blob, fileName);
      return { mode: "download", blob, fileName, fileUrl: "", uploadError };
    }
  },

  async exportCsv(filename, columns, dataRows, options = {}) {
    const rows = [
      columns,
      ...dataRows
    ];
    const csv = rows.map((row) => row.map((cell) => `"${String(cell || "").replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const fileName = filename.endsWith(".csv") ? filename : `${filename}.csv`;
    const popup = options.popup || null;
    const file = new File([blob], fileName, { type: "text/csv;charset=utf-8;", lastModified: Date.now() });
    try {
      const uploaded = await pulseUploadGeneratedFile(file, options.folderName || "CSV Exports");
      if (uploaded && uploaded.fileUrl) {
        const api = pulseExportOfficeApi();
        if (api && typeof api.openSharePointFileByPolicy === "function") {
          api.openSharePointFileByPolicy(uploaded.fileUrl, uploaded.fileName || fileName, file.type, { popup });
        } else {
          window.open(uploaded.fileUrl, "_blank", "noopener");
        }
        return {
          mode: "sharepoint",
          blob,
          fileName: uploaded.fileName || fileName,
          fileUrl: uploaded.fileUrl,
          serverRelativeUrl: uploaded.serverRelativeUrl || ""
        };
      }
      pulseCloseExportWindow(popup);
      pulseDownloadBlob(blob, fileName);
      return { mode: "download", blob, fileName, fileUrl: "" };
    } catch (uploadError) {
      pulseCloseExportWindow(popup);
      pulseDownloadBlob(blob, fileName);
      return { mode: "download", blob, fileName, fileUrl: "", uploadError };
    }
  },

  async exportXlsx(filename, columns, dataRows, options = {}) {
    const blob = await pulseBuildStyledWorkbook(columns, dataRows, options);
    const fileName = filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`;
    const popup = options.popup || null;
    const mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    const file = new File([blob], fileName, { type: mimeType, lastModified: Date.now() });
    try {
      const uploaded = await pulseUploadGeneratedFile(file, options.folderName || "Excel Exports");
      if (uploaded && uploaded.fileUrl) {
        const api = pulseExportOfficeApi();
        if (api && typeof api.openSharePointFileByPolicy === "function") {
          api.openSharePointFileByPolicy(uploaded.fileUrl, uploaded.fileName || fileName, mimeType, { popup });
        } else {
          window.open(uploaded.fileUrl, "_blank", "noopener");
        }
        return {
          mode: "sharepoint",
          blob,
          fileName: uploaded.fileName || fileName,
          fileUrl: uploaded.fileUrl,
          serverRelativeUrl: uploaded.serverRelativeUrl || ""
        };
      }
      pulseCloseExportWindow(popup);
      pulseDownloadBlob(blob, fileName);
      return { mode: "download", blob, fileName, fileUrl: "" };
    } catch (uploadError) {
      pulseCloseExportWindow(popup);
      pulseDownloadBlob(blob, fileName);
      return { mode: "download", blob, fileName, fileUrl: "", uploadError };
    }
  },

  exportPdf(filename, htmlContent, opts = {}) {
    if (typeof html2pdf !== "undefined") {
      html2pdf().set({
        margin: opts.margin || 0.5,
        filename: filename.endsWith(".pdf") ? filename : `${filename}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'in', format: 'letter', orientation: opts.orientation || 'portrait' }
      }).from(htmlContent).save();
    } else {
      console.warn("html2pdf is not loaded, falling back to print dialog.");
      window.print();
    }
  }
};
