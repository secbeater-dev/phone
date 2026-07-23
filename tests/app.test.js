const test = require("node:test");
const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");

const XLSX = require("../vendor/xlsx.full.min.js");
const PhoneWorkbench = require("../app.js");

function workbookBytes(rows, sheetName = "工作表1") {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
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

test("explicit direction overrides raw call type in phone statistics", () => {
  const stats = PhoneWorkbench.computePhoneStats([
    { call_type: "發話", direction: "inbound", target_phone: "0900000001", counterparty_phone: "0900000002", duration_seconds: 10 },
    { call_type: "受話", direction: "outbound", target_phone: "0900000001", counterparty_phone: "0900000003", duration_seconds: 20 },
  ]);
  assert.deepEqual(stats.inboundRows.map((row) => row.phone), ["0900000002"]);
  assert.deepEqual(stats.outboundRows.map((row) => row.phone), ["0900000003"]);
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

  for (const relativePath of ["vendor/xlsx.full.min.js", "app.js"]) {
    const bytes = fs.readFileSync(path.join(root, relativePath));
    const sri = `sha384-${crypto.createHash("sha384").update(bytes).digest("base64")}`;
    assert.match(html, new RegExp(`src="\\./${relativePath.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]+integrity="${sri.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"`));
    assert.ok(html.includes(`'${sri}'`));
  }
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
