"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const { deflateRawSync } = require("node:zlib");
const XLSX = require("../vendor/xlsx.full.min.js");
let parser, model;
try { parser = require("../streaming-xlsx.js"); model = require("../cdr-model.js"); } catch (error) { if (error.code !== "MODULE_NOT_FOUND") throw error; }
const callHeaders = ["查詢項目", "通話類別", "始話時間", "調閱門號", "對象門號", "通話期間", "開始基地台編號", "開始基地台", "結束基地台編號", "結束基地台", "imei", "主叫號碼"];
const dataHeaders = ["查詢項目", "開始時間", "連線期間", "通聯回應_用戶編號_帳號", "開始基地台編號", "開始基地台", "結束基地台編號", "結束基地台", "上網ipv4", "上網ipv6", "使用者內網ip"];
function workbook(sheets, options = {}) {
  const book = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
  if (options.date1904) book.Workbook = { WBProps: { date1904: true } };
  return new File([XLSX.write(book, { type: "buffer", bookType: "xlsx", bookSST: options.shared !== false, compression: true })], "synthetic.xlsx");
}
function store() {
  const strings = new Map();
  return { clears: 0, async putMany(items) { assert.ok(items.length <= 1000); for (const item of items) strings.set(item.id, item.value); }, async getMany(ids) { return new Map(ids.filter(id => strings.has(id)).map(id => [id, strings.get(id)])); }, async clear() { this.clears++; strings.clear(); } };
}
async function importAll(file, options = {}) {
  assert.equal(typeof parser?.importFile, "function", "streaming import API must exist");
  const records = [], users = [], stringStore = store();
  const summary = await parser.importFile(file, { stringStore, onRecords: async batch => records.push(...batch), onUsers: async batch => users.push(...batch), ...options });
  return { records, users, summary, stringStore };
}
test("streaming import keeps per-row targets, both record kinds and users; ignores integrated history", async () => {
  const file = workbook({
    "使用者資料": [["用戶名稱", "查詢項目", "用戶回應_用戶編號_帳號"], ["合成用戶一", "0911000001", ""], ["合成未匹配用戶", "0911000099", ""]],
    "通聯紀錄": [callHeaders, ["+886911000001", "發話", "2026/09/01 10:00:00", "0911000098", "0911000002", 5, 1001, "臺北市中正區合成路", "", "", 123456789012345, "0911000002"], ["", "受話", "2026/09/01 10:01:00", "0911000003", "0911000004", 8], ["", "發話", "invalid synthetic time", "", "0911000004", 8]],
    "網路歷程": [dataHeaders, ["", "2026/09/01 11:00:00", 60, "0911000005", 1002, "臺北市中正區合成路", 1003, "新北市板橋區合成路", "192.0.2.1", "2001:db8::1", "10.0.0.1"]],
    "整合歷程": [callHeaders, ["0911000001", "發話", "2026/09/01 10:00:00"]]
  });
  assert.equal(typeof parser?.inspect, "function", "streaming inspect API must exist");
  assert.equal((await parser.inspect(file)).supported, true);
  const { records, users, summary, stringStore } = await importAll(file);
  assert.equal(records.length, 4); assert.equal(users.length, 2);
  assert.deepEqual(records.map(r => r.target_phone), ["0911000001", "0911000003", "", "0911000005"]);
  assert.equal(records[0].direction, "inbound"); assert.equal(records[0].imei, "123456789012345");
  assert.equal(records[0].source_query, "+886911000001"); assert.equal(records[0].source_target, "0911000098");
  assert.equal(records[2].occurred_at, "invalid synthetic time");
  assert.equal(records[3].record_kind, "data"); assert.equal(records[3].external_ipv6, "2001:db8::1");
  assert.equal(records[0].base_refs[0].station_key, records[0].stations[0].station_key);
  assert.equal(summary.warning_count, 1); assert.equal(summary.call_count, 3); assert.equal(summary.data_count, 1);
  assert.equal(stringStore.clears, 1);
});
test("model aliases, direction, 1904 dates and blanks preserve raw identifier precision", () => {
  assert.equal(typeof model?.normalizeRow, "function", "normalization API must exist");
  const headers = ["查詢項目", "通話類別", "始話時間", "調閱門號", "對象門號", "通話期間", "imei"];
  const record = model.normalizeRow("call", headers, ["", "受話", { value: 1.5, type: "n", style: 1 }, 911000001, "0911000002", 8, { value: "1.2345678901234567E+16", type: "n" }], { fileName: "synthetic.xlsx", sheetName: "通聯紀錄", rowNumber: 19, date1904: true, styles: [0, 22] });
  assert.equal(record.occurred_at, "1904-01-02T12:00:00"); assert.equal(record.row_number, 19);
  assert.equal(record.imei, "12345678901234567"); assert.equal(record.target_phone, "0911000001");
  assert.equal(model.normalizeRow("call", headers, headers, {}), null);
  assert.equal(model.normalizeRow("call", headers, [], {}), null);
  assert.equal(model.canonicalHeader(" 通聯回應_用戶編號_帳號 "), model.canonicalHeader("通聯回應-用戶編號/帳號"));
  assert.equal(model.kindOf(record), "call");
});
test("awaited sinks receive at most 1000 rows and 4 MiB and honor cancellation", async () => {
  const rows = [callHeaders];
  for (let i = 0; i < 2200; i++) rows.push(["0911000001", "發話", "2026-09-01 10:00:00", "", "0911000002", i]);
  const file = workbook({ "通聯紀錄": rows });
  let active = false, count = 0, batches = 0;
  await importAll(file, { onRecords: async records => { assert.equal(active, false); active = true; assert.ok(records.length <= 1000); assert.ok(Buffer.byteLength(JSON.stringify(records)) <= 4 * 1024 * 1024); await new Promise(resolve => setTimeout(resolve, 2)); count += records.length; batches++; active = false; } });
  assert.equal(count, 2200); assert.ok(batches >= 3);
  const controller = new AbortController(); let writes = 0; const strings = store();
  await assert.rejects(importAll(file, { stringStore: strings, signal: controller.signal, onRecords: async () => { writes++; controller.abort(); } }), { name: "AbortError" });
  assert.equal(writes, 1); assert.equal(strings.clears, 1);
});
test("unsupported names return false; corrupt or invalid recognized structures fail without content", async () => {
  assert.equal(typeof parser?.inspect, "function", "streaming inspect API must exist");
  assert.equal((await parser.inspect(workbook({ "其他": [["synthetic"]] }))).supported, false);
  await assert.rejects(parser.inspect(new File(["not a zip"], "synthetic.xlsx")), /XLSX/);
  await assert.rejects(importAll(workbook({ "通聯紀錄": [["private-like synthetic marker"]] })), error => !error.message.includes("marker"));
  assert.equal((await importAll(workbook({ "網路歷程": [] }))).records.length, 0);
});

// A tiny fixture ZIP writer keeps XML edge cases explicit; all bytes are synthetic.
function zipFile(parts) {
  const locals = [], central = []; let offset = 0;
  for (const [name, text] of Object.entries(parts)) {
    const body = Buffer.from(text), packed = deflateRawSync(body), filename = Buffer.from(name);
    let crc = 0xffffffff;
    for (const byte of body) { crc ^= byte; for (let b = 0; b < 8; b++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0); }
    crc = (crc ^ 0xffffffff) >>> 0;
    const local = Buffer.alloc(30); local.writeUInt32LE(0x04034b50); local.writeUInt16LE(20, 4); local.writeUInt16LE(8, 8); local.writeUInt32LE(crc, 14); local.writeUInt32LE(packed.length, 18); local.writeUInt32LE(body.length, 22); local.writeUInt16LE(filename.length, 26);
    const cd = Buffer.alloc(46); cd.writeUInt32LE(0x02014b50); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6); cd.writeUInt16LE(8, 10); cd.writeUInt32LE(crc, 16); cd.writeUInt32LE(packed.length, 20); cd.writeUInt32LE(body.length, 24); cd.writeUInt16LE(filename.length, 28); cd.writeUInt32LE(offset, 42);
    locals.push(local, filename, packed); central.push(cd, filename); offset += local.length + filename.length + packed.length;
  }
  const c = Buffer.concat(central), end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50); end.writeUInt16LE(central.length / 2, 8); end.writeUInt16LE(central.length / 2, 10); end.writeUInt32LE(c.length, 12); end.writeUInt32LE(offset, 16);
  return new File([...locals, c, end], "synthetic.xlsx");
}
function xmlFile(sheet, extra = {}) {
  return zipFile({
    "xl/workbook.xml": '<workbook xmlns:r="urn:r"><workbookPr date1904="1"/><sheets><sheet name="通聯紀錄" sheetId="1" r:id="rId1"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels": '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="t" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": `<worksheet><sheetData>${sheet}</sheetData></worksheet>`,
    "xl/styles.xml": '<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="22"/></cellXfs></styleSheet>', ...extra
  });
}
function inline(ref, text) { return `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`; }
test("sparse physical rows, inline/rich shared text, and styled numeric dates stream correctly", async () => {
  const head = `<row r="3">${callHeaders.slice(0, 6).map((h, i) => inline(String.fromCharCode(65 + i) + "3", h)).join("")}</row>`;
  const row = '<row r="19"><c r="A19" t="s"><v>0</v></c><c r="B19" t="inlineStr"><is><r><t>發</t></r><r><t>話</t></r></is></c><c r="C19" s="1"><v>1.5</v></c><c r="F19"><v>5</v></c></row>';
  const file = xmlFile(head + row, { "xl/sharedStrings.xml": '<sst><si><r><t>0911</t></r><r><t>000001</t></r></si></sst>' });
  const { records } = await importAll(file);
  assert.equal(records[0].row_number, 19); assert.equal(records[0].target_phone, "0911000001"); assert.equal(records[0].occurred_at, "1904-01-02T12:00:00"); assert.equal(records[0].call_type, "發話"); assert.equal(records[0].counterparty_phone, "");
});
test("expanded shared strings cannot inflate a row beyond the row memory limit", async () => {
  const head = `<row r="1">${[...callHeaders.slice(0, 6), ...Array.from({ length: 14 }, (_, i) => `合成欄${i}`)].map((h, i) => inline(String.fromCharCode(65 + i) + "1", h)).join("")}</row>`;
  const row = `<row r="2">${Array.from({ length: 20 }, (_, i) => `<c r="${String.fromCharCode(65 + i)}2" t="s"><v>0</v></c>`).join("")}</row>`;
  const file = xmlFile(head + row, { "xl/sharedStrings.xml": `<sst><si><t>${"s".repeat(100000)}</t></si></sst>` });
  await assert.rejects(importAll(file), /XLSX/);
});
test("large record batches stop at 4 MiB and cancellation during shared-string writes stops parsing", async () => {
  const rows = [[...callHeaders, "備註"]];
  for (let i = 0; i < 90; i++) rows.push(["0911000001", "發話", "2026/09/01 10:00:00", "", "0911000002", 1, "", "", "", "", "", "", "合成內容".repeat(5000)]);
  let count = 0, batches = 0;
  await importAll(workbook({ "通聯紀錄": rows }), { onRecords: async records => { count += records.length; batches++; assert.ok(Buffer.byteLength(JSON.stringify(records)) <= 4 * 1024 * 1024); } });
  assert.equal(count, 90); assert.ok(batches >= 2);
  const strings = store(), controller = new AbortController(); let writes = 0;
  strings.putMany = async () => { writes++; controller.abort(); };
  await assert.rejects(importAll(workbook({ "通聯紀錄": rows }), { stringStore: strings, signal: controller.signal }), { name: "AbortError" });
  assert.equal(writes, 1); assert.equal(strings.clears, 1);
});
test("corrupt XML, missing shared-string references and oversized text fail with safe diagnostics", async () => {
  for (const body of ['<row r="1"><c r="A1" t="s"><v>888</v></c></row>', '<row r="1"><c r="A1"></row>', `<row r="1">${inline("A1", "synthetic-secret-marker".repeat(14000))}</row>`]) {
    await assert.rejects(importAll(xmlFile(body)), error => error.name === "XlsxFormatError" && !error.message.includes("marker"));
  }
  await assert.rejects(importAll(xmlFile('<row r="2"/><row r="1"/>')), /XLSX/);
});
test("inspection and import use Blob slices and never full-file arrayBuffer; pre-cancel reads nothing", async () => {
  const original = workbook({ "通聯紀錄": [callHeaders, ["0911000001", "發話", "2026/09/01 10:00:00", "", "0911000002", 1]] });
  const file = new Blob([original]);
  file.arrayBuffer = () => { throw new Error("full-file read forbidden"); };
  assert.equal((await parser.inspect(file)).supported, true);
  assert.equal((await importAll(file)).records.length, 1);
  const controller = new AbortController(); controller.abort();
  file.slice = () => { throw new Error("cancelled read forbidden"); };
  await assert.rejects(parser.inspect(file, { signal: controller.signal }), { name: "AbortError" });
});
test("numeric identifiers discard numeric decimal padding and invalid end time does not hide valid date bounds", async () => {
  assert.equal(model.normalizePhone({ value: "911000001.000", type: "n" }), "0911000001");
  const file = workbook({ "網路歷程": [[...dataHeaders, "結束時間"], ["0911000001", "2026/09/01 10:00:00", 60, "", 1, "臺北市中正區合成路", "", "", "", "", "", "invalid end"]] });
  const { summary } = await importAll(file);
  assert.equal(summary.warning_count, 1); assert.deepEqual(summary.date_bounds, { start: "2026-09-01", end: "2026-09-01" });
});
test("browser UMD parser inflates using only local JavaScript without Worker, fetch or WASM", async () => {
  const vm = require("node:vm"), fs = require("node:fs"), path = require("node:path");
  let prohibited = 0;
  const blocked = () => { prohibited++; throw new Error("unavailable browser network/worker"); };
  const sandbox = vm.createContext({ Blob, TextDecoder, TextEncoder, DOMException, AbortController, ReadableStream, WritableStream, TransformStream, Response, URL, setTimeout, clearTimeout, Worker: blocked, fetch: blocked, WebAssembly: undefined });
  for (const file of ["vendor/zip-no-worker-inflate-2.7.57.min.js", "vendor/sax-1.4.1.js", "cdr-model.js", "streaming-xlsx.js"]) vm.runInContext(fs.readFileSync(path.join(__dirname, "..", file), "utf8"), sandbox);
  const file = workbook({ "通聯紀錄": [callHeaders, ["0911000001", "發話", "2026/09/01 10:00:00", "", "0911000002", 1]] });
  const records = [];
  await sandbox.PhoneStreamingXlsx.importFile(file, { stringStore: store(), onRecords: async batch => records.push(...batch) });
  assert.equal(records[0].target_phone, "0911000001"); assert.equal(prohibited, 0);
});
test("encrypted ZIP members and checksum corruption are rejected", async () => {
  const good = xmlFile(""), encrypted = Buffer.from(await good.arrayBuffer());
  for (let offset = 0; offset < encrypted.length - 46; offset++) {
    const magic = encrypted.readUInt32LE(offset);
    if (magic === 0x04034b50) encrypted.writeUInt16LE(1, offset + 6);
    if (magic === 0x02014b50) encrypted.writeUInt16LE(1, offset + 8);
  }
  await assert.rejects(parser.inspect(new File([encrypted], "synthetic.xlsx")), /XLSX/);
  const corrupt = Buffer.from(await good.arrayBuffer());
  // Flip the workbook central-directory CRC, leaving the XML syntactically valid.
  for (let offset = 0; offset < corrupt.length - 46; offset++) if (corrupt.readUInt32LE(offset) === 0x02014b50) { corrupt.writeUInt32LE(42, offset + 16); break; }
  await assert.rejects(parser.inspect(new File([corrupt], "synthetic.xlsx")), /XLSX/);
});

test("inspection keeps legacy layouts on their parser even when sheet names overlap", async () => {
  const fet = [
    ["電話號碼：0900000001"],
    ["始話時間", "通話秒數", "調閱號碼", "IMEI", "通話類別", "通話對象", "轉接電話", "基地台/交換機", "備註"],
    ["2026-09-01T01:02:03", 10, "0900000001", "", "發話", "0900000002", "0900000009", "SYNTH-1", ""],
    ["", "", "", "", "", "", "", "SYNTH-2", ""]
  ];
  const cht = [["CDR類別", "主叫號碼", "查詢狀態", "受叫號碼", "始話日期時間", "通話秒數", "IMEI", "指定轉接", "起始基地台-地址/終止基地台-地址"]];
  for (const rows of [fet, cht]) assert.equal((await parser.inspect(workbook({ "通聯紀錄": rows }))).supported, false);
  const legacy = require("../app.js").parseImportFile("synthetic.xlsx", await workbook({ "通聯紀錄": fet }).arrayBuffer());
  assert.equal(legacy.records.length, 1); assert.equal(legacy.base_stations.length, 2);
  assert.match(legacy.records[0].note, /0900000009/);
});

test("inspection resolves late shared headers and stops before the worksheet body", async () => {
  const base = 2000;
  const headers = `<row r="1">${callHeaders.slice(0, 6).map((_, i) => `<c r="${String.fromCharCode(65 + i)}1" t="s"><v>${base + i}</v></c>`).join("")}</row>`;
  const harmlessRows = Array.from({ length: 79 }, (_, i) => `<row r="${i + 2}"/>`).join("");
  const file = xmlFile(headers + harmlessRows + '<row r="81"><c r="A81"><v>malformed body</row>', {
    "xl/sharedStrings.xml": `<sst>${Array.from({ length: base }, (_, i) => `<si><t>synthetic unused ${i}</t></si>`).join("")}${callHeaders.slice(0, 6).map(h => `<si><r><t>${h}</t></r></si>`).join("")}</sst>`
  });
  assert.equal((await parser.inspect(file)).supported, true);
  await assert.rejects(importAll(file), { name: "XlsxFormatError" });
});
