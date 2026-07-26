const test = require("node:test");
const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");

const XLSX = require("../vendor/xlsx.full.min.js");
const ExcelJS = require("../vendor/exceljs.min.js");
const PDFLib = require("../vendor/pdf-lib.min.js");
const fontkit = require("../vendor/fontkit.umd.min.js");
const fontBase64 = require("../vendor/open-huninn-data.js");
const AttachmentExport = require("../attachment-export.js");
const PhoneWorkbench = require("../app.js");

function workbookBytes(rows, sheetName = "工作表1") {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
}

function multiSheetWorkbookBytes(sheets) {
  const workbook = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name));
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
}

function chtRows() {
  const rows = [
    ["查詢結果：成功"],
    ["發文案號：SYNTHETIC-CASE"],
    ["備註：synthetic fixture"],
    [],
    ["查詢條件：設備號碼", "", "設備號碼：0900000001"],
    [],
    ["電話號碼：0900000001", "", "", "用戶名稱：測試用戶"],
    ["申裝地址：測試地址"],
    ["帳寄地址：測試地址"],
    ["申請日期：2026-01-01"],
    ["CDR類別", "主 叫 號 碼", "查詢狀態", "", "受 叫 號 碼", "", "", "始 話 日 期 時 間", "", "通話秒數", "IMEI", "指定轉接", "起始基地台-地址/終止基地台-地址"],
    ["發話", "0900000001", "完成", "", "0900000002", "", "", "2026/01/02 03:04:05", "", "30", "123456789012345", "無設定", "1001-臺北市測試路1號/1002-新北市測試路2號"],
    ["受話", "0900000003", "", "", "0900000001", "", "", "2026-01-02 04:05:06", "", "45", "", "", ""],
    ["進來CDR", "0900000004", "", "", "0900000001", "", "", "2026-01-02 05:06:07", "", "15", "", "", ""],
    ["發話", "0900000005", "", "", "0900000006", "", "", "2026-01-02 06:07:08", "", "12", "", "", ""],
    ["受話", "0900000007", "", "", "0900000001", "", "", "not-a-date", "", "5", "", "", ""],
    ["CDR類別", "主 叫 號 碼", "查詢狀態", "", "受 叫 號 碼", "", "", "始 話 日 期 時 間", "", "通話秒數", "IMEI", "指定轉接", "起始基地台-地址/終止基地台-地址"],
  ];
  return rows;
}

function fetProsecutorHeaders(withSpacer = false) {
  const headers = ["始話時間", "通話秒數", "調閱號碼", "IMEI", "通話類別", "通話對象", "轉接電話", "基地台/交換機"];
  if (withSpacer) headers.push("");
  headers.push("備註");
  return headers;
}

function fetProsecutorRows({ withSpacer = false, repeatedSection = false } = {}) {
  const row = (values) => withSpacer ? [...values.slice(0, 8), "", values[8] || ""] : values;
  const rows = [
    ["文號：SYNTHETIC-FET-CASE"],
    ["查詢日期：2026-07-26"],
    ["電信業者：", "遠傳電信"],
    ["通聯類別：語音", "查詢狀態：完成"],
    ["區段時間：2026-07-01T00:00:00 至 2026-07-02T00:00:00", "備註：合成資料"],
    ["電話號碼：0900000101"],
    fetProsecutorHeaders(withSpacer),
    row(["2026-07-01T01:02:03", "10", "0900000101", "111111111111111", "發話", "0900000102", "0900000199", "SYNTH-1001", "合成備註"]),
    row(["", "", "", "", "", "", "", "SYNTH-1002", ""]),
    row(["", "", "", "", "", "", "", "SYNTH-1002", ""]),
    row(["2026-07-01T02:03:04", "20", "0900000101", "", "受話", "0900000103", "", "", ""]),
    row(["invalid-synthetic-date", "5", "0900000101", "", "受話", "", "", "", ""]),
  ];
  if (repeatedSection) {
    rows.push(
      [],
      ["通聯類別：語音", "查詢狀態：完成"],
      ["區段時間：2026-07-03T00:00:00 至 2026-07-04T00:00:00", "備註：第二段合成資料"],
      ["電話號碼：", "0900000201"],
      fetProsecutorHeaders(!withSpacer),
      ["2026-07-03T03:04:05", "30", "0900000201", "222222222222222", "發話", "0900000202", "", "SYNTH-2001", "第二段合成備註"],
    );
  }
  return rows;
}

test("parses the Chunghwa prosecutor-office XLSX layout", () => {
  const workspace = PhoneWorkbench.parseImportFile("synthetic-chunghwa.xlsx", workbookBytes(chtRows()));

  assert.equal(workspace.case.carrier, "中華電信");
  assert.equal(workspace.case.source_format, "chunghwa_prosecutor_cdr_xlsx");
  assert.equal(workspace.case.header_row, 9);
  assert.equal(workspace.case.subject["電話號碼"], "0900000001");
  assert.equal(workspace.records.length, 5);

  assert.deepEqual(
    workspace.records.slice(0, 4).map((record) => [record.call_type, record.direction, record.target_phone, record.counterparty_phone]),
    [
      ["發話", "outbound", "0900000001", "0900000002"],
      ["受話", "inbound", "0900000001", "0900000003"],
      ["進來CDR", "inbound", "0900000001", "0900000004"],
      ["發話", "outbound", "0900000005", "0900000006"],
    ],
  );
  assert.equal(workspace.records[0].occurred_at, "2026-01-02T03:04:05");
  assert.equal(workspace.records[0].duration_seconds, 30);
  assert.equal(workspace.records[0].imei, "123456789012345");
  assert.match(workspace.records[0].note, /查詢狀態：完成/);
  assert.match(workspace.records[0].note, /指定轉接：無設定/);
  assert.equal(workspace.records[0].base_refs.length, 2);
  assert.deepEqual(workspace.records[0].base_refs.map((ref) => ref.role), ["start", "end"]);
  assert.equal(workspace.base_stations.length, 2);
  assert.equal(workspace.records[4].occurred_at, "not-a-date");
  assert.equal(workspace.parse_warnings.length, 1);
});

test("parses the Far EasTone prosecutor-office nine-column XLSX layout", () => {
  const workspace = PhoneWorkbench.parseImportFile("synthetic-fet-nine.xlsx", workbookBytes(fetProsecutorRows()));

  assert.equal(workspace.case.carrier, "遠傳電信");
  assert.equal(workspace.case.source_format, "fet_prosecutor_call_xlsx");
  assert.equal(workspace.case.subject["電話號碼"], "0900000101");
  assert.equal(workspace.records.length, 3);
  assert.deepEqual(workspace.records.map((record) => record.direction), ["outbound", "inbound", "inbound"]);
  assert.equal(workspace.records[0].occurred_at, "2026-07-01T01:02:03");
  assert.equal(workspace.records[0].duration_seconds, 10);
  assert.equal(workspace.records[0].imei, "111111111111111");
  assert.match(workspace.records[0].note, /轉接電話：0900000199/);
  assert.match(workspace.records[0].note, /合成備註/);
  assert.equal(workspace.records[0].base_refs.length, 2);
  assert.equal(workspace.base_stations.length, 2);
  assert.equal(workspace.records[2].occurred_at, "invalid-synthetic-date");
  assert.equal(workspace.parse_warnings.length, 1);
});

test("parses spacer columns and every repeated Far EasTone query section", () => {
  const workspace = PhoneWorkbench.parseImportFile(
    "synthetic-fet-repeated.xlsx",
    workbookBytes(fetProsecutorRows({ withSpacer: true, repeatedSection: true })),
  );

  assert.equal(workspace.case.source_format, "fet_prosecutor_call_xlsx");
  assert.equal(workspace.records.length, 4);
  assert.equal(workspace.case.subject["電話號碼"], "0900000101、0900000201");
  assert.match(workspace.case.subject["區段時間"], /、/);
  assert.equal(workspace.records[0].note, "轉接電話：0900000199；合成備註");
  assert.equal(workspace.records[3].target_phone, "0900000201");
  assert.equal(workspace.records[3].counterparty_phone, "0900000202");
  assert.equal(workspace.records[3].imei, "222222222222222");
  assert.equal(workspace.records[3].base_refs[0].role, "primary");
});

test("tracks the actual source sheet for Far EasTone records", () => {
  const firstRows = fetProsecutorRows().slice(0, 9);
  const secondRows = [
    fetProsecutorHeaders(),
    ["2026-07-05T04:05:06", "12", "0900000301", "", "發話", "0900000302", "", "SYNTH-3001", ""],
  ];
  const workspace = PhoneWorkbench.parseImportFile("synthetic-fet-sheets.xlsx", multiSheetWorkbookBytes([
    { name: "第一區段", rows: firstRows },
    { name: "第二區段", rows: secondRows },
  ]));

  assert.ok(workspace.records.some((record) => record.source_sheet === "第一區段"));
  assert.ok(workspace.records.some((record) => record.source_sheet === "第二區段"));
});

test("merges every selected workspace without overwriting earlier call records", () => {
  const first = PhoneWorkbench.parseImportFile("synthetic-fet-a.xlsx", workbookBytes(fetProsecutorRows()));
  const second = PhoneWorkbench.parseImportFile("synthetic-fet-b.xlsx", workbookBytes(fetProsecutorRows({ withSpacer: true, repeatedSection: true })));
  const merged = PhoneWorkbench.mergeWorkspaces([first, second]);

  assert.deepEqual(merged.case.source_files, ["synthetic-fet-a.xlsx", "synthetic-fet-b.xlsx"]);
  assert.equal(merged.records.length, first.records.length + second.records.length);
  assert.equal(merged.case.total_records, merged.records.length);
  assert.equal(merged.case.summary.records, merged.records.length);
  assert.ok(merged.records.some((record) => record.source_file === "synthetic-fet-a.xlsx"));
  assert.ok(merged.records.some((record) => record.source_file === "synthetic-fet-b.xlsx"));
  assert.match(merged.case.subject["電話號碼"], /、/);
  assert.ok(merged.records.every((record) => record.occurred_at || record.call_type || record.target_phone || record.counterparty_phone || record.imei || record.note || record.duration_seconds !== null));
});

test("builds a complete attachment report with both ranking modes and browser-only notes", () => {
  const first = PhoneWorkbench.parseImportFile("synthetic-fet-a.xlsx", workbookBytes(fetProsecutorRows()));
  const second = PhoneWorkbench.parseImportFile("synthetic-fet-b.xlsx", workbookBytes(fetProsecutorRows({ repeatedSection: true })));
  const merged = PhoneWorkbench.mergeWorkspaces([first, second]);
  const report = PhoneWorkbench.buildAttachmentReport(merged, { "0900000102": "合成電話備註" }, "2026-07-26T00:00:00.000Z");

  assert.equal(report.meta.scope, "complete_import");
  assert.equal(report.calls.length, merged.records.length);
  assert.equal(report.hours.length, 24);
  assert.equal(report.stats.count.totalRows.length, report.stats.seconds.totalRows.length);
  assert.ok(report.calls.some((record) => record.counterparty_note === "合成電話備註"));
  assert.equal(report.profile.imeis.length > 0, true);
});

test("creates a six-sheet attachment XLSX with identifiers stored as text", async () => {
  const workspace = PhoneWorkbench.parseImportFile("synthetic-fet.xlsx", workbookBytes(fetProsecutorRows({ repeatedSection: true })));
  workspace.records[0].note = "=SYNTHETIC_FORMULA";
  const report = PhoneWorkbench.buildAttachmentReport(workspace, {}, "2026-07-26T00:00:00.000Z");
  const chartDataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
  const bytes = await AttachmentExport.createAttachmentXlsx(report, ExcelJS, chartDataUrl);
  const workbook = XLSX.read(bytes, { type: "array", raw: true });
  const styledWorkbook = new ExcelJS.Workbook();
  await styledWorkbook.xlsx.load(bytes);

  assert.deepEqual(workbook.SheetNames, AttachmentExport.SHEET_NAMES);
  assert.ok(styledWorkbook.worksheets.every((sheet) => sheet.views[0].ySplit === 5));
  assert.equal(styledWorkbook.media.length, 1);
  assert.equal(styledWorkbook.getWorksheet("時間分布圖").getImages().length, 1);
  const callSheet = workbook.Sheets["通聯列表"];
  assert.equal(callSheet.A6.t, "s");
  assert.equal(callSheet.E6.t, "s");
  assert.equal(callSheet.O6.t, "s");
  assert.equal(callSheet.O6.f, undefined);
  assert.match(callSheet.O6.v, /^=/);
});

test("creates six searchable-font PDF document structures from synthetic data", async () => {
  const workspace = PhoneWorkbench.parseImportFile("synthetic-fet.xlsx", workbookBytes(fetProsecutorRows({ repeatedSection: true })));
  const report = PhoneWorkbench.buildAttachmentReport(workspace, { "0900000102": "合成電話備註" }, "2026-07-26T00:00:00.000Z");
  const fontBytes = Buffer.from(fontBase64, "base64");

  for (const section of AttachmentExport.PDF_SECTIONS) {
    const bytes = await AttachmentExport.createAttachmentPdf(report, section.key, PDFLib, fontkit, fontBytes);
    assert.equal(Buffer.from(bytes.subarray(0, 4)).toString("ascii"), "%PDF");
    const document = await PDFLib.PDFDocument.load(bytes);
    assert.ok(document.getPageCount() >= 1, `${section.key} PDF should contain at least one page`);
  }
});

test("explicit direction overrides raw call type in phone statistics", () => {
  const stats = PhoneWorkbench.computePhoneStats([
    { call_type: "發話", direction: "inbound", target_phone: "0900000001", counterparty_phone: "0900000002", duration_seconds: 10 },
    { call_type: "受話", direction: "outbound", target_phone: "0900000001", counterparty_phone: "0900000003", duration_seconds: 20 },
  ]);
  assert.deepEqual(stats.inboundRows.map((row) => row.phone), ["0900000002"]);
  assert.deepEqual(stats.outboundRows.map((row) => row.phone), ["0900000003"]);
});

test("phone statistics preserve every ranked phone beyond the former top 20 limit", () => {
  const records = Array.from({ length: 25 }, (_, index) => ({
    call_type: "發話",
    direction: "outbound",
    target_phone: "0900000000",
    counterparty_phone: `0911${String(index).padStart(6, "0")}`,
    duration_seconds: index + 1,
  }));
  const stats = PhoneWorkbench.computePhoneStats(records);
  assert.equal(stats.outboundRows.length, 25);
  assert.equal(stats.totalRows.filter((row) => row.phone !== "0900000000").length, 25);
  const source = fs.readFileSync(path.resolve(__dirname, "..", "app.js"), "utf8");
  assert.doesNotMatch(source, /rows\.slice\(0,\s*20\)\.map/);
});

test("keeps the existing Taiwan Mobile parser behavior", () => {
  const rows = [
    ["通話類別", "目標電話", "對象電話", "始話日期時間", "通話時間(秒)", "基地台編號1/位置1"],
    ["發話", "0900000011", "0900000012", "2026-01-02 03:04:05", "9", "2001 測試地址"],
  ];
  const workspace = PhoneWorkbench.parseImportFile("synthetic-twm.xlsx", workbookBytes(rows));
  assert.equal(workspace.case.source_format, "taiwan_mobile_call");
  assert.equal(workspace.records.length, 1);
  assert.equal(workspace.records[0].direction, "outbound");
});

test("HTML uses pinned local scripts and contains no analytics tag", () => {
  const root = path.resolve(__dirname, "..");
  const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
  assert.doesNotMatch(html, /googletagmanager|gtag\s*\(/i);
  assert.doesNotMatch(html, /tellows\.tw/i);
  assert.match(html, /今日重點（2026-07-26）/);
  assert.match(html, /附卷檔案匯出/);
  assert.match(html, /多檔匯入/);
  assert.match(html, /NT\$1,500/);
  assert.match(html, /家庭分享教學/);
  assert.match(html, /https:\/\/families\.google\/intl\/zh-TW_ALL\/families\//);
  assert.match(html, /https:\/\/t\.me\/tg_secbeater/);
  assert.match(html, /https:\/\/secbeater\.notion\.site\//);
  assert.doesNotMatch(html, /supportPasswordInput|SUPPORT_PASSWORD|目前支援檔案類型/);

  for (const relativePath of ["vendor/xlsx.full.min.js", "attachment-export.js", "app.js"]) {
    const bytes = fs.readFileSync(path.join(root, relativePath));
    const sri = `sha384-${crypto.createHash("sha384").update(bytes).digest("base64")}`;
    assert.match(html, new RegExp(`src="\\./${relativePath.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\?v=20260726-attachment-v1"[^>]+integrity="${sri.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"`));
    assert.ok(html.includes(`'${sri}'`));
  }
  const appSource = fs.readFileSync(path.join(root, "app.js"), "utf8");
  for (const relativePath of ["vendor/exceljs.min.js", "vendor/pdf-lib.min.js", "vendor/fontkit.umd.min.js", "vendor/open-huninn-data.js"]) {
    const bytes = fs.readFileSync(path.join(root, relativePath));
    const sri = `sha384-${crypto.createHash("sha384").update(bytes).digest("base64")}`;
    assert.ok(html.includes(`'${sri}'`));
    assert.ok(appSource.includes(`./${relativePath}?v=20260726-attachment-v1`));
    assert.ok(appSource.includes(sri));
  }
  assert.match(html, /connect-src 'none'/);
  assert.match(appSource, /CALL_PAGE_SIZE = 500/);
  assert.doesNotMatch(appSource, /rows\.slice\(0,\s*5000\)/);
});

test("Pages deploys attachment assets from an explicit file allowlist", () => {
  const root = path.resolve(__dirname, "..");
  const workflow = fs.readFileSync(path.join(root, ".github", "workflows", "pages.yml"), "utf8");
  const ignore = fs.readFileSync(path.join(root, ".gitignore"), "utf8");
  assert.match(workflow, /attachment-export\.js/);
  for (const asset of ["exceljs.min.js", "pdf-lib.min.js", "fontkit.umd.min.js", "open-huninn-data.js", "fontkit-NOTICE.txt", "open-huninn-LICENSE.txt"]) {
    assert.ok(workflow.includes(asset), `${asset} must be explicitly deployed`);
  }
  assert.doesNotMatch(workflow, /cp\s+-R\s+assets\s+vendor/);
  assert.match(ignore, /附卷\*\.pdf/);
});

test("private workbook integration is local-only and opt-in", { skip: !process.env.PRIVATE_CDR_XLSX }, () => {
  const bytes = fs.readFileSync(process.env.PRIVATE_CDR_XLSX);
  const workspace = PhoneWorkbench.parseImportFile(path.basename(process.env.PRIVATE_CDR_XLSX), bytes);
  const sourceWorkbook = XLSX.read(bytes, { type: "buffer", cellDates: false });
  const sourceRows = XLSX.utils.sheet_to_json(sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]], { header: 1, defval: "", raw: false, blankrows: false });
  const canonical = (value) => String(value ?? "").trim().replace(/\s+/g, "");
  const required = ["CDR類別", "主叫號碼", "受叫號碼", "始話日期時間", "通話秒數"];
  const headerIndex = sourceRows.findIndex((row) => required.every((header) => row.map(canonical).includes(header)));
  assert.ok(headerIndex >= 0);
  const sourceDataRows = sourceRows.slice(headerIndex + 1).filter((row) => {
    const cells = row.map((value) => String(value ?? "").trim());
    return [cells[0], cells[1], cells[4], cells[7]].some(Boolean) && canonical(cells[0]) !== "CDR類別";
  });
  assert.equal(workspace.case.source_format, "chunghwa_prosecutor_cdr_xlsx");
  assert.equal(workspace.records.length, sourceDataRows.length);
  assert.ok(workspace.records.every((record) => record.call_type || record.occurred_at || record.target_phone || record.counterparty_phone));
  assert.ok(workspace.records.every((record) => ["inbound", "outbound", "data", "other"].includes(record.direction)));
});

test("private Far EasTone workbook integration is local-only and opt-in", { skip: !process.env.PRIVATE_FET_XLSX }, () => {
  const bytes = fs.readFileSync(process.env.PRIVATE_FET_XLSX);
  const workspace = PhoneWorkbench.parseImportFile(path.basename(process.env.PRIVATE_FET_XLSX), bytes);
  const sourceWorkbook = XLSX.read(bytes, { type: "buffer", cellDates: false });
  const canonical = (value) => String(value ?? "").trim().replace(/\s+/g, "");
  const required = ["始話時間", "通話秒數", "調閱號碼", "IMEI", "通話類別", "通話對象", "轉接電話", "基地台/交換機", "備註"];
  const metadataKeys = new Set(["文號", "查詢日期", "電信業者", "通聯類別", "查詢狀態", "區段時間", "備註", "電話號碼"]);
  let expectedRecords = 0;

  sourceWorkbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(sourceWorkbook.Sheets[sheetName], { header: 1, defval: "", raw: false, blankrows: false });
    let headers = null;
    rows.forEach((row) => {
      const canonicalRow = row.map(canonical);
      if (required.every((header) => canonicalRow.includes(header))) {
        headers = canonicalRow;
        return;
      }
      const metadataRow = row.slice(0, 3).some((value) => {
        const match = String(value ?? "").trim().match(/^([^:：]+)\s*[:：]/);
        return Boolean(match && metadataKeys.has(canonical(match[1])));
      });
      if (!headers || metadataRow) return;
      const data = Object.fromEntries(headers.map((header, index) => [header, String(row[index] ?? "").trim()]).filter(([header]) => header));
      if (["始話時間", "通話秒數", "調閱號碼", "IMEI", "通話類別", "通話對象", "轉接電話", "備註"].some((key) => data[key])) expectedRecords += 1;
    });
  });

  assert.ok(workspace.case.source_format === "fet_prosecutor_call_xlsx", "private format check failed");
  assert.ok(workspace.records.length === expectedRecords, "private record completeness check failed");
  assert.ok(workspace.records.every((record) => ["inbound", "outbound", "data", "other"].includes(record.direction)), "private direction check failed");
  assert.ok(workspace.records.every((record) => Array.isArray(record.base_refs)), "private station shape check failed");
  assert.ok(workspace.records.every((record) => record.occurred_at || record.call_type || record.target_phone || record.counterparty_phone || record.imei || record.note || record.duration_seconds !== null), "private visible record check failed");
});
