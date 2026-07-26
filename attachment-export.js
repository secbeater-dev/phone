(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.PhoneAttachmentExport = factory();
})(typeof globalThis !== "undefined" ? globalThis : window, function () {
  "use strict";

  const SHEET_NAMES = ["時間分布圖", "熱點時間摘要", "通聯列表", "用戶資料", "電話統計-次數版", "電話統計-秒數版"];
  const PDF_SECTIONS = [
    { key: "hours", label: "時間分布圖" },
    { key: "hotspots", label: "熱點時間摘要" },
    { key: "calls", label: "通聯列表" },
    { key: "profile", label: "用戶資料" },
    { key: "stats_count", label: "電話統計-次數版" },
    { key: "stats_seconds", label: "電話統計-秒數版" },
  ];
  const COLORS = {
    ink: "FF171717",
    muted: "FF60646C",
    line: "FFD9DCE1",
    red: "FFB91C1C",
    redSoft: "FFFDECEC",
    header: "FF222222",
    white: "FFFFFFFF",
    surface: "FFF6F7F9",
  };

  async function createAttachmentXlsx(report, ExcelJS, chartDataUrl) {
    if (!ExcelJS?.Workbook) throw new Error("ExcelJS 尚未載入");
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Phone Workbench";
    workbook.lastModifiedBy = "Phone Workbench";
    workbook.created = new Date(report.meta.exported_at);
    workbook.modified = new Date(report.meta.exported_at);
    workbook.subject = "附卷檔案";
    workbook.title = "通聯資料附卷檔案";
    workbook.company = "SecBeater";

    addHoursSheet(workbook, report, chartDataUrl);
    addHotspotsSheet(workbook, report);
    addCallsSheet(workbook, report);
    addProfileSheet(workbook, report);
    addStatsSheet(workbook, report, "count", "電話統計-次數版");
    addStatsSheet(workbook, report, "seconds", "電話統計-秒數版");

    const bytes = await workbook.xlsx.writeBuffer();
    return new Uint8Array(bytes);
  }

  function setupSheet(workbook, report, name, maxColumn, orientation = "landscape") {
    const sheet = workbook.addWorksheet(name, {
      views: [{ state: "frozen", ySplit: 5, showGridLines: false }],
      pageSetup: {
        paperSize: 9,
        orientation,
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        margins: { left: 0.25, right: 0.25, top: 0.45, bottom: 0.45, header: 0.2, footer: 0.2 },
      },
      headerFooter: {
        oddFooter: "&LPhone Workbench&C第 &P / &N 頁&R附卷檔案",
      },
    });
    sheet.mergeCells(1, 1, 1, maxColumn);
    const title = sheet.getCell(1, 1);
    title.value = name;
    title.font = { name: "Microsoft JhengHei", size: 18, bold: true, color: COLORS.white };
    title.fill = solidFill(COLORS.header);
    title.alignment = { vertical: "middle", horizontal: "left" };
    sheet.getRow(1).height = 30;
    sheet.mergeCells(2, 1, 2, maxColumn);
    const scope = sheet.getCell(2, 1);
    scope.value = `匯出範圍：完整匯入資料｜匯出時間：${formatExportTime(report.meta.exported_at)}`;
    scope.font = { name: "Microsoft JhengHei", size: 10, color: COLORS.muted };
    scope.fill = solidFill(COLORS.surface);
    scope.alignment = { vertical: "middle", wrapText: true };
    sheet.getRow(2).height = 22;
    sheet.mergeCells(3, 1, 3, maxColumn);
    const sources = sheet.getCell(3, 1);
    sources.value = `來源檔案：${(report.meta.source_files || []).join("、") || "-"}`;
    sources.font = { name: "Microsoft JhengHei", size: 9, color: COLORS.muted };
    sources.alignment = { vertical: "middle", wrapText: true };
    sheet.getRow(3).height = 22;
    return sheet;
  }

  function addHoursSheet(workbook, report, chartDataUrl) {
    const sheet = setupSheet(workbook, report, "時間分布圖", 10);
    sheet.columns = [
      { width: 14 }, { width: 12 }, { width: 12 }, { width: 34 },
      { width: 3 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 },
    ];
    const headers = ["時段", "通聯次數", "占比", "視覺比例"];
    writeTable(sheet, 5, headers, report.hours.map((item) => [
      item.label,
      item.count,
      item.percent / 100,
      item.count ? "█".repeat(Math.max(1, Math.round(item.percent / 2))) : "",
    ]), { percentColumns: [3], integerColumns: [2], filter: true });
    if (chartDataUrl) {
      const imageId = workbook.addImage({ base64: chartDataUrl, extension: "png" });
      sheet.addImage(imageId, { tl: { col: 5, row: 4 }, br: { col: 10, row: 20 }, editAs: "oneCell" });
    }
    sheet.pageSetup.printArea = `A1:J${Math.max(29, sheet.rowCount)}`;
  }

  function addHotspotsSheet(workbook, report) {
    const sheet = setupSheet(workbook, report, "熱點時間摘要", 7);
    sheet.columns = [
      { width: 8 }, { width: 46 }, { width: 12 }, { width: 12 }, { width: 22 }, { width: 22 }, { width: 22 },
    ];
    const summaryRows = report.hotspots.map((item, index) => [
      index + 1, item.address, item.count, item.percent / 100, item.first_seen, item.last_seen, item.times.length,
    ]);
    let row = writeTable(sheet, 5, ["排名", "基地台地址", "次數", "占比", "首次時間", "末次時間", "時間明細數"], summaryRows, {
      percentColumns: [4], integerColumns: [1, 3, 7], filter: true,
    });
    row += 2;
    sheet.mergeCells(row, 1, row, 7);
    styleSectionTitle(sheet.getCell(row, 1), "全部發生時間明細");
    const detailRows = [];
    report.hotspots.forEach((item, hotspotIndex) => {
      item.times.forEach((time, timeIndex) => detailRows.push([hotspotIndex + 1, item.address, timeIndex + 1, time]));
    });
    writeTable(sheet, row + 1, ["熱點排名", "基地台地址", "序號", "發生時間"], detailRows, { integerColumns: [1, 3] });
    sheet.pageSetup.printArea = `A1:G${Math.max(sheet.rowCount, row + 2)}`;
  }

  function addCallsSheet(workbook, report) {
    const sheet = setupSheet(workbook, report, "通聯列表", 16);
    const widths = [18, 22, 11, 10, 16, 22, 16, 22, 10, 18, 18, 22, 22, 34, 34, 46];
    sheet.columns = widths.map((width) => ({ width }));
    const headers = [
      "來源檔案", "時間", "類型", "方向", "調閱/目標", "目標電話備註", "對象", "對象電話備註",
      "秒數", "IMEI", "IMSI", "外部IP", "內部IP", "基地台", "備註", "來源工作表/列號",
    ];
    const rows = report.calls.map((record) => [
      record.source_file, record.occurred_at, record.call_type, record.direction, record.target_phone, record.target_note,
      record.counterparty_phone, record.counterparty_note, record.duration_seconds, record.imei, record.imsi,
      record.external_ip, record.internal_ip, record.base_stations, record.note,
      [record.source_sheet, record.row_number ? `第 ${record.row_number} 列` : ""].filter(Boolean).join(" / "),
    ]);
    writeTable(sheet, 5, headers, rows, { integerColumns: [9], textColumns: [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16], filter: true });
    sheet.pageSetup.printArea = `A1:P${Math.max(sheet.rowCount, 6)}`;
    sheet.pageSetup.printTitlesRow = "1:5";
  }

  function addProfileSheet(workbook, report) {
    const sheet = setupSheet(workbook, report, "用戶資料", 6, "portrait");
    sheet.columns = [{ width: 24 }, { width: 42 }, { width: 24 }, { width: 42 }, { width: 20 }, { width: 20 }];
    let row = 5;
    sheet.mergeCells(row, 1, row, 6);
    styleSectionTitle(sheet.getCell(row, 1), "案件摘要");
    row += 1;
    const summaryEntries = Object.entries(report.profile.summary || {});
    summaryEntries.forEach(([label, value], index) => {
      const col = index % 2 === 0 ? 1 : 3;
      const targetRow = row + Math.floor(index / 2);
      setTextCell(sheet.getCell(targetRow, col), label, true);
      setTextCell(sheet.getCell(targetRow, col + 1), value);
    });
    row += Math.ceil(summaryEntries.length / 2) + 1;
    sheet.mergeCells(row, 1, row, 6);
    styleSectionTitle(sheet.getCell(row, 1), "用戶欄位");
    row = writeTable(sheet, row + 1, ["欄位", "內容"], report.profile.subject.map((item) => [item.key, item.value]));
    row += 1;
    sheet.mergeCells(row, 1, row, 6);
    styleSectionTitle(sheet.getCell(row, 1), "來源檔案");
    row = writeTable(sheet, row + 1, ["序號", "檔名"], report.meta.source_files.map((name, index) => [index + 1, name]), { integerColumns: [1] });
    row += 1;
    sheet.mergeCells(row, 1, row, 6);
    styleSectionTitle(sheet.getCell(row, 1), "所有 IMEI");
    writeTable(sheet, row + 1, ["序號", "IMEI"], report.profile.imeis.map((imei, index) => [index + 1, imei]), { integerColumns: [1], textColumns: [2] });
    sheet.pageSetup.printArea = `A1:F${Math.max(sheet.rowCount, row + 2)}`;
  }

  function addStatsSheet(workbook, report, mode, name) {
    const sheet = setupSheet(workbook, report, name, 5, "portrait");
    sheet.columns = [{ width: 9 }, { width: 20 }, { width: 42 }, { width: 14 }, { width: 16 }];
    const stats = report.stats[mode];
    let row = 5;
    [
      ["來電排行", stats.inboundRows],
      ["去電排行", stats.outboundRows],
      ["完整排行", stats.totalRows],
    ].forEach(([title, rows], sectionIndex) => {
      if (sectionIndex) row += 2;
      sheet.mergeCells(row, 1, row, 5);
      styleSectionTitle(sheet.getCell(row, 1), title);
      row = writeTable(sheet, row + 1, ["排名", "電話", "備註", "次數", "秒數"], rows.map((item, index) => [
        index + 1, item.phone, item.note, item.count, item.seconds,
      ]), { integerColumns: [1, 4, 5], textColumns: [2, 3] });
    });
    sheet.pageSetup.printArea = `A1:E${Math.max(sheet.rowCount, 6)}`;
  }

  function writeTable(sheet, startRow, headers, rows, options = {}) {
    const headerRow = sheet.getRow(startRow);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      cell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: COLORS.white };
      cell.fill = solidFill(COLORS.red);
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = { bottom: { style: "medium", color: { argb: COLORS.red } } };
    });
    headerRow.height = 24;
    rows.forEach((values, rowIndex) => {
      const row = sheet.getRow(startRow + rowIndex + 1);
      values.forEach((value, columnIndex) => {
        const cell = row.getCell(columnIndex + 1);
        const column = columnIndex + 1;
        if (typeof value === "number" && !(options.textColumns || []).includes(column)) cell.value = value;
        else cell.value = safeText(value);
        cell.font = { name: "Microsoft JhengHei", size: 9, color: COLORS.ink };
        cell.alignment = {
          vertical: "top",
          horizontal: typeof cell.value === "number" ? "right" : "left",
          wrapText: true,
        };
        cell.border = { bottom: { style: "hair", color: { argb: COLORS.line } } };
        if ((options.textColumns || []).includes(column)) cell.numFmt = "@";
        if ((options.integerColumns || []).includes(column)) cell.numFmt = "#,##0";
        if ((options.percentColumns || []).includes(column)) cell.numFmt = "0.0%";
      });
      if (rowIndex % 2 === 1) row.eachCell((cell) => { cell.fill = solidFill(COLORS.surface); });
      row.height = Math.min(72, Math.max(20, estimateRowHeight(values)));
    });
    if (options.filter && headers.length) sheet.autoFilter = {
      from: { row: startRow, column: 1 },
      to: { row: startRow + Math.max(rows.length, 1), column: headers.length },
    };
    return startRow + rows.length + 1;
  }

  function styleSectionTitle(cell, text) {
    cell.value = text;
    cell.font = { name: "Microsoft JhengHei", size: 12, bold: true, color: COLORS.ink };
    cell.fill = solidFill(COLORS.redSoft);
    cell.alignment = { vertical: "middle" };
  }

  function setTextCell(cell, value, bold = false) {
    cell.value = safeText(value);
    cell.font = { name: "Microsoft JhengHei", size: 9, bold, color: bold ? COLORS.muted : COLORS.ink };
    cell.alignment = { vertical: "top", wrapText: true };
    cell.numFmt = "@";
  }

  function solidFill(argb) {
    return { type: "pattern", pattern: "solid", fgColor: { argb } };
  }

  function safeText(value) {
    return String(value ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  }

  function estimateRowHeight(values) {
    const longest = Math.max(0, ...values.map((value) => safeText(value).length));
    return 18 + Math.floor(longest / 45) * 12;
  }

  async function createAttachmentPdf(report, sectionKey, PDFLib, fontkit, fontBytes) {
    if (!PDFLib?.PDFDocument || !fontkit || !fontBytes) throw new Error("PDF 產生資產尚未載入");
    // pdf-lib/fontkit subsetting corrupts some Traditional Chinese glyph maps;
    // embed this compact OFL font intact so every viewer renders searchable text.
    return buildPdf(report, sectionKey, PDFLib, fontkit, fontBytes, false);
  }

  async function buildPdf(report, sectionKey, PDFLib, fontkit, fontBytes, subset) {
    const section = PDF_SECTIONS.find((item) => item.key === sectionKey);
    if (!section) throw new Error("未知的 PDF 區段");
    const { PDFDocument, rgb } = PDFLib;
    const doc = await PDFDocument.create();
    doc.registerFontkit(fontkit);
    const font = await doc.embedFont(fontBytes, { subset });
    doc.setTitle(section.label);
    doc.setSubject("通聯資料附卷檔案");
    doc.setCreator("Phone Workbench");
    doc.setProducer("Phone Workbench");
    doc.setCreationDate(new Date(report.meta.exported_at));
    const landscape = sectionKey !== "profile";
    const size = landscape ? [841.89, 595.28] : [595.28, 841.89];
    const ctx = createPdfContext(doc, font, rgb, size, report.meta, section.label);

    if (sectionKey === "hours") drawHoursPdf(ctx, report);
    else if (sectionKey === "hotspots") drawHotspotsPdf(ctx, report);
    else if (sectionKey === "calls") drawCallsPdf(ctx, report);
    else if (sectionKey === "profile") drawProfilePdf(ctx, report);
    else if (sectionKey === "stats_count") drawStatsPdf(ctx, report.stats.count, "電話統計-次數版");
    else if (sectionKey === "stats_seconds") drawStatsPdf(ctx, report.stats.seconds, "電話統計-秒數版");

    addPdfFooters(ctx);
    return new Uint8Array(await doc.save());
  }

  function createPdfContext(doc, font, rgb, size, meta, title) {
    const ctx = {
      doc, font, rgb, size, meta, title,
      pages: [], page: null, y: 0,
      margin: 26, footerHeight: 22,
      ink: rgb(0.09, 0.09, 0.09),
      muted: rgb(0.38, 0.4, 0.44),
      red: rgb(0.73, 0.11, 0.11),
      redSoft: rgb(0.99, 0.92, 0.92),
      surface: rgb(0.965, 0.97, 0.98),
      line: rgb(0.82, 0.84, 0.87),
    };
    newPdfPage(ctx, true);
    return ctx;
  }

  function newPdfPage(ctx, first = false) {
    const page = ctx.doc.addPage(ctx.size);
    ctx.pages.push(page);
    ctx.page = page;
    const height = ctx.size[1];
    page.drawText(first ? ctx.title : `${ctx.title}（續）`, {
      x: ctx.margin, y: height - ctx.margin - 18, size: 15, font: ctx.font, color: ctx.ink,
    });
    page.drawText(`完整匯入資料｜${formatExportTime(ctx.meta.exported_at)}`, {
      x: ctx.margin, y: height - ctx.margin - 34, size: 7.5, font: ctx.font, color: ctx.muted,
    });
    ctx.y = height - ctx.margin - 48;
  }

  function ensurePdfSpace(ctx, height, onNewPage) {
    if (ctx.y - height >= ctx.margin + ctx.footerHeight) return;
    newPdfPage(ctx, false);
    if (onNewPage) onNewPage();
  }

  function drawHoursPdf(ctx, report) {
    const chartHeight = 205;
    const chartWidth = ctx.size[0] - ctx.margin * 2;
    const chartBottom = ctx.y - chartHeight;
    const max = Math.max(1, ...report.hours.map((item) => item.count));
    const gap = 4;
    const barWidth = (chartWidth - gap * 23) / 24;
    report.hours.forEach((item, index) => {
      const height = item.count ? (item.count / max) * 145 : 0;
      const x = ctx.margin + index * (barWidth + gap);
      ctx.page.drawRectangle({ x, y: chartBottom + 28, width: barWidth, height, color: ctx.ink });
      drawCentered(ctx.page, ctx.font, safeText(item.count), 6, x, chartBottom + 31 + height, barWidth, ctx.ink);
      drawCentered(ctx.page, ctx.font, safeText(index).padStart(2, "0"), 6, x, chartBottom + 13, barWidth, ctx.muted);
    });
    ctx.page.drawLine({ start: { x: ctx.margin, y: chartBottom + 27 }, end: { x: ctx.margin + chartWidth, y: chartBottom + 27 }, thickness: 0.6, color: ctx.line });
    ctx.y = chartBottom - 12;
    drawPdfTable(ctx, ["時段", "通聯次數", "占比"], report.hours.map((item) => [item.label, item.count, `${item.percent.toFixed(1)}%`]), [0.4, 0.3, 0.3]);
  }

  function drawHotspotsPdf(ctx, report) {
    drawPdfTable(ctx, ["排名", "基地台地址", "次數", "占比", "首次時間", "末次時間"], report.hotspots.map((item, index) => [
      index + 1, item.address, item.count, `${item.percent.toFixed(1)}%`, item.first_seen, item.last_seen,
    ]), [0.07, 0.35, 0.08, 0.08, 0.21, 0.21]);
    drawPdfSectionTitle(ctx, "全部發生時間明細");
    const details = [];
    report.hotspots.forEach((item, hotspotIndex) => item.times.forEach((time, timeIndex) => {
      details.push([hotspotIndex + 1, item.address, timeIndex + 1, time]);
    }));
    drawPdfTable(ctx, ["熱點排名", "基地台地址", "序號", "發生時間"], details, [0.1, 0.52, 0.1, 0.28]);
  }

  function drawCallsPdf(ctx, report) {
    const headers = ["來源", "時間", "類型", "目標", "目標備註", "對象", "對象備註", "秒", "IMEI", "基地台", "備註/IP"];
    const rows = report.calls.map((record) => [
      record.source_file, record.occurred_at, record.call_type, record.target_phone, record.target_note,
      record.counterparty_phone, record.counterparty_note, record.duration_seconds, record.imei,
      record.base_stations, [record.note, record.external_ip].filter(Boolean).join("；"),
    ]);
    drawPdfTable(ctx, headers, rows, [0.09, 0.12, 0.065, 0.09, 0.09, 0.09, 0.09, 0.045, 0.105, 0.13, 0.085], { fontSize: 5.4, padding: 2.2 });
  }

  function drawProfilePdf(ctx, report) {
    drawPdfSectionTitle(ctx, "案件摘要");
    drawPdfTable(ctx, ["欄位", "內容"], Object.entries(report.profile.summary || {}), [0.33, 0.67]);
    drawPdfSectionTitle(ctx, "用戶欄位");
    drawPdfTable(ctx, ["欄位", "內容"], report.profile.subject.map((item) => [item.key, item.value]), [0.33, 0.67]);
    drawPdfSectionTitle(ctx, "來源檔案");
    drawPdfTable(ctx, ["序號", "檔名"], report.meta.source_files.map((name, index) => [index + 1, name]), [0.18, 0.82]);
    drawPdfSectionTitle(ctx, "所有 IMEI");
    drawPdfTable(ctx, ["序號", "IMEI"], report.profile.imeis.map((imei, index) => [index + 1, imei]), [0.18, 0.82]);
  }

  function drawStatsPdf(ctx, stats, title) {
    [
      ["來電排行", stats.inboundRows],
      ["去電排行", stats.outboundRows],
      ["完整排行", stats.totalRows],
    ].forEach(([sectionTitle, rows], index) => {
      if (index) {
        newPdfPage(ctx, false);
        drawPdfSectionTitle(ctx, sectionTitle);
      } else drawPdfSectionTitle(ctx, sectionTitle);
      drawPdfTable(ctx, ["排名", "電話", "備註", "次數", "秒數"], rows.map((item, rank) => [
        rank + 1, item.phone, item.note, item.count, item.seconds,
      ]), [0.1, 0.22, 0.42, 0.12, 0.14]);
    });
  }

  function drawPdfSectionTitle(ctx, title) {
    ensurePdfSpace(ctx, 28);
    ctx.page.drawRectangle({ x: ctx.margin, y: ctx.y - 20, width: ctx.size[0] - ctx.margin * 2, height: 20, color: ctx.redSoft });
    ctx.page.drawText(title, { x: ctx.margin + 6, y: ctx.y - 14, size: 10, font: ctx.font, color: ctx.ink });
    ctx.y -= 26;
  }

  function drawPdfTable(ctx, headers, rows, fractions, options = {}) {
    const fontSize = options.fontSize || 7.2;
    const padding = options.padding || 3.5;
    const width = ctx.size[0] - ctx.margin * 2;
    const columnWidths = normalizeFractions(fractions, headers.length).map((fraction) => width * fraction);
    const drawHeader = () => {
      const height = 19;
      ctx.page.drawRectangle({ x: ctx.margin, y: ctx.y - height, width, height, color: ctx.red });
      let x = ctx.margin;
      headers.forEach((header, index) => {
        const lines = wrapPdfText(ctx.font, header, fontSize, columnWidths[index] - padding * 2);
        drawPdfLines(ctx.page, ctx.font, lines, fontSize, x + padding, ctx.y - padding - fontSize, columnWidths[index] - padding * 2, ctx.rgb(1, 1, 1));
        x += columnWidths[index];
      });
      ctx.y -= height;
    };
    ensurePdfSpace(ctx, 24, drawHeader);
    drawHeader();
    if (!rows.length) rows = [["尚無資料", ...Array(Math.max(0, headers.length - 1)).fill("")]];
    rows.forEach((values, rowIndex) => {
      const lineSets = headers.map((_, index) => wrapPdfText(ctx.font, values[index], fontSize, columnWidths[index] - padding * 2));
      const rowHeight = Math.max(16, ...lineSets.map((lines) => lines.length * (fontSize + 1.8) + padding * 2));
      ensurePdfSpace(ctx, rowHeight, drawHeader);
      if (rowIndex % 2 === 1) ctx.page.drawRectangle({ x: ctx.margin, y: ctx.y - rowHeight, width, height: rowHeight, color: ctx.surface });
      let x = ctx.margin;
      lineSets.forEach((lines, index) => {
        drawPdfLines(ctx.page, ctx.font, lines, fontSize, x + padding, ctx.y - padding - fontSize, columnWidths[index] - padding * 2, ctx.ink);
        x += columnWidths[index];
        if (index < lineSets.length - 1) ctx.page.drawLine({ start: { x, y: ctx.y }, end: { x, y: ctx.y - rowHeight }, thickness: 0.25, color: ctx.line });
      });
      ctx.page.drawLine({ start: { x: ctx.margin, y: ctx.y - rowHeight }, end: { x: ctx.margin + width, y: ctx.y - rowHeight }, thickness: 0.35, color: ctx.line });
      ctx.y -= rowHeight;
    });
    ctx.y -= 8;
  }

  function wrapPdfText(font, value, size, maxWidth) {
    const text = safeText(value).replace(/\r/g, "");
    if (!text) return [""];
    const lines = [];
    text.split("\n").forEach((paragraph) => {
      let line = "";
      Array.from(paragraph).forEach((character) => {
        const candidate = line + character;
        if (line && font.widthOfTextAtSize(candidate, size) > maxWidth) {
          lines.push(line);
          line = character;
        } else line = candidate;
      });
      lines.push(line);
    });
    return lines.length ? lines : [""];
  }

  function drawPdfLines(page, font, lines, size, x, y, _width, color) {
    lines.forEach((line, index) => page.drawText(safeText(line), { x, y: y - index * (size + 1.8), size, font, color }));
  }

  function drawCentered(page, font, text, size, x, y, width, color) {
    const textWidth = font.widthOfTextAtSize(text, size);
    page.drawText(text, { x: x + Math.max(0, (width - textWidth) / 2), y, size, font, color });
  }

  function normalizeFractions(fractions, count) {
    const values = Array.from({ length: count }, (_, index) => Number(fractions[index] || 1 / count));
    const total = values.reduce((sum, value) => sum + value, 0) || 1;
    return values.map((value) => value / total);
  }

  function addPdfFooters(ctx) {
    const total = ctx.pages.length;
    ctx.pages.forEach((page, index) => {
      const text = `第 ${index + 1} / ${total} 頁`;
      drawCentered(page, ctx.font, text, 7, ctx.margin, 13, ctx.size[0] - ctx.margin * 2, ctx.muted);
    });
  }

  function formatExportTime(value) {
    return safeText(value).replace("T", " ").replace(/\.\d{3}Z$/, " UTC");
  }

  return {
    SHEET_NAMES,
    PDF_SECTIONS,
    createAttachmentXlsx,
    createAttachmentPdf,
    safeText,
  };
});
