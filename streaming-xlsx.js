(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./vendor/zip-no-worker-inflate-2.7.57.min.js"), require("./vendor/sax-1.4.1.js"), require("./cdr-model.js"));
  else root.PhoneStreamingXlsx = factory(root.zip, root.sax, root.PhoneCdrModel);
})(typeof globalThis !== "undefined" ? globalThis : this, function (zip, sax, model) {
  "use strict";
  const MAX_BATCH_ROWS = 1000, MAX_BATCH_BYTES = 4 * 1024 * 1024;
  const MAX_PART_METADATA = 4 * 1024 * 1024, MAX_NODE = 256 * 1024, MAX_ROW = 2 * 1024 * 1024;
  const NAMES = new Set(["使用者資料", "通聯紀錄", "網路歷程"]);
  const encoder = new TextEncoder();
  zip.configure({ useWebWorkers: false, useCompressionStream: false, chunkSize: 4096, maxWorkers: 1 });
  sax.MAX_BUFFER_LENGTH = MAX_NODE;
  function abort(signal) { if (signal?.aborted) throw new DOMException("Import cancelled", "AbortError"); }
  function invalid(part = "結構", row) { const error = new Error(`XLSX ${part}${row ? `（列 ${row}）` : ""} 無法解析或超過安全限制`); error.name = "XlsxFormatError"; return error; }
  function rethrow(error, signal, part) { abort(signal); if (["AbortError", "QuotaExceededError", "XlsxFormatError", "XlsxInspectionComplete"].includes(error?.name)) throw error; throw invalid(part); }
  function name(value) { return value.split(":").pop(); }
  function attr(node, key) { return node.attributes[key] ?? Object.entries(node.attributes).find(([keyName]) => name(keyName) === key)?.[1] ?? ""; }
  function size(value) { return encoder.encode(JSON.stringify(value)).byteLength; }
  function safePath(base, target) {
    if (!target || /[\\\u0000]|^[a-z]+:/i.test(target)) throw invalid("關聯");
    const parts = (target.startsWith("/") ? target.slice(1) : base.slice(0, base.lastIndexOf("/") + 1) + target).split("/");
    const result = [];
    for (const part of parts) { if (part === "..") { if (!result.length) throw invalid("關聯"); result.pop(); } else if (part && part !== ".") result.push(part); }
    return result.join("/");
  }
  async function guardDirectory(file, signal) {
    abort(signal);
    if (!file || typeof file.slice !== "function" || file.size < 22) throw invalid("ZIP");
    const tail = new Uint8Array(await file.slice(Math.max(0, file.size - 65557)).arrayBuffer());
    const view = new DataView(tail.buffer, tail.byteOffset, tail.byteLength);
    for (let at = tail.length - 22; at >= 0; at--) {
      if (view.getUint32(at, true) !== 0x06054b50 || at + 22 + view.getUint16(at + 20, true) !== tail.length) continue;
      // Metadata is deliberately capped; the directory is the only ZIP-wide allocation.
      if (view.getUint16(at + 4, true) || view.getUint16(at + 6, true) || view.getUint16(at + 10, true) > 4096 || view.getUint32(at + 12, true) > MAX_PART_METADATA) throw invalid("ZIP 目錄");
      return;
    }
    throw invalid("ZIP");
  }
  async function xml(entry, handlers, { signal, maxBytes = Infinity, onBytes } = {}) {
    abort(signal);
    if (!entry || entry.encrypted || entry.uncompressedSize > maxBytes) throw invalid("XML 部件");
    let depth = 0, bytes = 0, decoder, first = true;
    const parser = sax.parser(true, { trim: false, normalize: false, strictEntities: true });
    parser.onerror = () => { throw invalid("XML 部件"); };
    parser.ondoctype = () => { throw invalid("XML 宣告"); };
    parser.onopentag = node => {
      if (++depth > 64) throw invalid("XML 深度");
      handlers.open?.(name(node.name), node, depth);
    };
    parser.onclosetag = tag => { handlers.close?.(name(tag), depth); depth--; };
    parser.ontext = parser.oncdata = value => { if (value.length > MAX_NODE) throw invalid("XML 文字"); handlers.text?.(value); };
    const write = async value => {
      // SAX is synchronous; draining after each small input slice bounds its output queue.
      for (let at = 0; at < value.length; at += 4096) { abort(signal); parser.write(value.slice(at, at + 4096)); await handlers.flush?.(); abort(signal); }
    };
    try {
      await entry.getData(new WritableStream({
        async write(chunk) {
          abort(signal); bytes += chunk.byteLength;
          if (bytes > maxBytes || bytes > entry.uncompressedSize) throw invalid("XML 大小");
          if (first) {
            first = false;
            const utf16le = (chunk[0] === 255 && chunk[1] === 254) || (chunk[0] === 60 && chunk[1] === 0);
            const utf16be = (chunk[0] === 254 && chunk[1] === 255) || (chunk[0] === 0 && chunk[1] === 60);
            decoder = new TextDecoder(utf16le ? "utf-16le" : utf16be ? "utf-16be" : "utf-8", { fatal: true });
          }
          await write(decoder.decode(chunk, { stream: true })); onBytes?.(bytes);
        },
        async close() { if (decoder) await write(decoder.decode()); parser.close(); await handlers.flush?.(); }
      }), { signal, useWebWorkers: false, useCompressionStream: false, checkSignature: true });
      if (bytes !== entry.uncompressedSize || depth !== 0) throw invalid("XML 大小");
      await handlers.end?.();
    } catch (error) { rethrow(error, signal, "XML 部件"); }
  }
  async function archive(file, signal) {
    await guardDirectory(file, signal);
    const reader = new zip.ZipReader(new zip.BlobReader(file), { useWebWorkers: false, useCompressionStream: false });
    try {
      const entries = new Map();
      for await (const entry of reader.getEntriesGenerator()) {
        abort(signal);
        if (entries.size >= 4096 || entries.has(entry.filename) || entry.encrypted) throw invalid("ZIP 目錄");
        entries.set(entry.filename, entry);
      }
      let workbookPath = "xl/workbook.xml";
      if (entries.has("_rels/.rels")) {
        await xml(entries.get("_rels/.rels"), { open(tag, node) {
          if (tag === "Relationship" && /\/officeDocument$/.test(attr(node, "Type"))) {
            if (attr(node, "TargetMode") === "External") throw invalid("關聯");
            workbookPath = safePath("", attr(node, "Target"));
          }
        } }, { signal, maxBytes: MAX_PART_METADATA });
      }
      const relations = new Map();
      const relationPath = workbookPath.slice(0, workbookPath.lastIndexOf("/") + 1) + "_rels/" + workbookPath.split("/").pop() + ".rels";
      await xml(entries.get(relationPath), { open(tag, node) {
        if (tag !== "Relationship") return;
        if (relations.size >= 4096) throw invalid("關聯");
        relations.set(attr(node, "Id"), { type: attr(node, "Type"), external: attr(node, "TargetMode") === "External", target: attr(node, "Target") });
      } }, { signal, maxBytes: MAX_PART_METADATA });
      const sheets = []; let date1904 = false, workbookRoot = false;
      await xml(entries.get(workbookPath), { open(tag, node, depth) {
        if (depth === 1) { if (tag !== "workbook") throw invalid("活頁簿"); workbookRoot = true; }
        if (tag === "workbookPr") date1904 = ["1", "true"].includes(attr(node, "date1904"));
        if (tag === "sheet") {
          const relation = relations.get(attr(node, "id"));
          if (!relation || relation.external || !/\/worksheet$/.test(relation.type)) throw invalid("工作表關聯");
          const path = safePath(workbookPath, relation.target), entry = entries.get(path);
          if (!entry || sheets.length >= 1024) throw invalid("工作表");
          sheets.push({ name: attr(node, "name"), path, uncompressedSize: entry.uncompressedSize });
        }
      } }, { signal, maxBytes: MAX_PART_METADATA });
      if (!workbookRoot) throw invalid("活頁簿");
      const relationEntry = type => {
        const relation = [...relations.values()].find(r => r.type.endsWith("/" + type));
        if (!relation) return undefined;
        if (relation.external) throw invalid("關聯");
        const entry = entries.get(safePath(workbookPath, relation.target));
        if (!entry) throw invalid("關聯");
        return entry;
      };
      const shared = relationEntry("sharedStrings"), styles = relationEntry("styles");
      const estimatedBytes = sheets.filter(s => NAMES.has(s.name)).reduce((sum, s) => sum + s.uncompressedSize, 0) + (shared?.uncompressedSize || 0) + (styles?.uncompressedSize || 0);
      const expandedBytes = [...entries.values()].reduce((total, entry) => total + entry.uncompressedSize, 0);
      return { reader, entries, sheets, date1904, shared, styles, estimatedBytes, expandedBytes, supported: sheets.some(s => s.name === "通聯紀錄" || s.name === "網路歷程") };
    } catch (error) { await reader.close(); rethrow(error, signal, "ZIP 部件"); }
  }
  async function inspect(file, { signal } = {}) {
    let book;
    try {
      book = await archive(file, signal);
      let supported = false;
      for (const sheet of book.sheets.filter(sheet => ["通聯紀錄", "網路歷程"].includes(sheet.name))) {
        const rows = await inspectionRows(book.entries.get(sheet.path), signal);
        await inspectionStrings(book.shared, rows, signal);
        const kind = sheet.name === "通聯紀錄" ? "call" : "data";
        if (rows.some(values => headerMatches(kind, values) && values.some(value => model.canonicalHeader(value) === "查詢項目"))) supported = true;
      }
      return { supported, sheets: book.sheets, estimatedBytes: book.estimatedBytes, expandedBytes: book.expandedBytes };
    }
    catch (error) { rethrow(error, signal, "結構"); }
    finally { if (book) await book.reader.close(); }
  }
  function inspectionComplete() { const error = new Error('Header prefix complete'); error.name = 'XlsxInspectionComplete'; throw error; }
  async function inspectionRows(entry, signal) {
    const rows = []; let row, cell, inside = false, bytes = 0, cells = 0, complete = false;
    const stop = () => { complete = true; inspectionComplete(); };
    try {
      await xml(entry, {
        open(tag, node) {
          if (tag === 'row') row = [];
          if (tag === 'c' && row) {
            cell = { type: attr(node, 't'), value: '', index: attr(node, 'r') ? column(attr(node, 'r')).index : row.length };
            if (++cells > 8192) stop();
          }
          if (cell && (tag === 'v' || tag === 't')) inside = true;
        },
        text(value) { if (inside && cell) { bytes += value.length * 2; if (bytes > 2 * 1024 * 1024) stop(); cell.value += value; } },
        close(tag) {
          if (tag === 'v' || tag === 't') inside = false;
          if (tag === 'c' && cell) { row[cell.index] = cell.type === 's' ? { shared: Number(cell.value) } : cell.value; cell = null; }
          if (tag === 'row' && row) { rows.push(row); row = null; if (rows.length >= 80) stop(); }
        }
      }, { signal });
    } catch (error) { abort(signal); if (!complete) throw error; }
    return rows;
  }
  async function inspectionStrings(entry, rows, signal) {
    const wanted = new Map(); let bytes = 0;
    for (const row of rows) row.forEach((cell, index) => {
      if (typeof cell === 'object') {
        if (!Number.isSafeInteger(cell.shared) || cell.shared < 0) throw invalid('共用字串索引');
        if (!wanted.has(cell.shared)) wanted.set(cell.shared, []);
        wanted.get(cell.shared).push([row, index]);
      }
    });
    if (!wanted.size) return;
    if (!entry) throw invalid('共用字串');
    let id = -1, value = '', inside = false, phonetic = false, complete = false;
    try {
      await xml(entry, {
        open(tag) { if (tag === 'si') { id++; value = ''; } if (tag === 'rPh') phonetic = true; if (tag === 't' && !phonetic) inside = true; },
        text(part) { if (inside && wanted.has(id)) { value += part; if (value.length > MAX_NODE) throw invalid('共用字串大小'); } },
        close(tag) {
          if (tag === 't') inside = false; if (tag === 'rPh') phonetic = false;
          if (tag === 'si' && wanted.has(id)) {
            for (const [row, index] of wanted.get(id)) { bytes += value.length * 2; if (bytes > 2 * 1024 * 1024) throw invalid('標題大小'); row[index] = value; }
            wanted.delete(id); if (!wanted.size) { complete = true; inspectionComplete(); }
          }
        }
      }, { signal });
    } catch (error) { abort(signal); if (!complete) throw error; }
    if (wanted.size) throw invalid('共用字串索引');
  }
  function batchSink(sink, signal) {
    let batch = [], bytes = 2;
    async function flush() {
      abort(signal); if (!batch.length) return;
      const current = batch; batch = []; bytes = 2;
      await sink(current); abort(signal);
    }
    return { flush, async push(item) {
      abort(signal); const itemBytes = size(item) + 1;
      if (itemBytes + 2 > MAX_BATCH_BYTES) throw invalid("資料列大小");
      if (batch.length >= MAX_BATCH_ROWS || bytes + itemBytes > MAX_BATCH_BYTES) await flush();
      batch.push(item); bytes += itemBytes;
    } };
  }
  async function sharedStrings(entry, store, signal, onBytes) {
    if (!entry) return;
    const batch = batchSink(items => store.putMany(items), signal);
    let value = "", inside = false, inText = false, phonetic = false, id = 0, queue = [], root = false;
    await xml(entry, {
      open(tag, _node, depth) {
        if (depth === 1) { if (tag !== "sst") throw invalid("共用字串"); root = true; }
        if (tag === "si") { if (inside) throw invalid("共用字串"); inside = true; value = ""; }
        if (tag === "rPh") phonetic = true;
        if (tag === "t" && inside && !phonetic) inText = true;
      },
      text(part) { if (inText) { value += part; if (value.length > MAX_NODE) throw invalid("共用字串大小"); } },
      close(tag) { if (tag === "t") inText = false; if (tag === "rPh") phonetic = false; if (tag === "si") { queue.push({ id: id++, value }); value = ""; inside = false; } },
      async flush() { for (const item of queue) await batch.push(item); queue = []; },
      async end() { if (!root || inside) throw invalid("共用字串"); await batch.flush(); }
    }, { signal, onBytes });
  }
  async function readStyles(entry, signal, onBytes) {
    if (!entry) return [];
    const formats = new Map(), styles = []; let inXfs = false;
    await xml(entry, { open(tag, node) {
      if (tag === "numFmt") formats.set(Number(attr(node, "numFmtId")), attr(node, "formatCode"));
      if (tag === "cellXfs") inXfs = true;
      if (tag === "xf" && inXfs) {
        if (styles.length >= 65536) throw invalid("樣式");
        const numFmtId = Number(attr(node, "numFmtId")), code = formats.get(numFmtId) || "";
        styles.push({ numFmtId, isDate: (numFmtId >= 14 && numFmtId <= 22) || (numFmtId >= 27 && numFmtId <= 36) || (numFmtId >= 45 && numFmtId <= 47) || /[ymdhis]/i.test(code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, "")) });
      }
    }, close(tag) { if (tag === "cellXfs") inXfs = false; } }, { signal, maxBytes: MAX_PART_METADATA, onBytes });
    return styles;
  }
  function headerMatches(kind, values) {
    const keys = new Set(values.map(model.canonicalHeader));
    const groups = kind === "call" ? [["通話類別", "CDR類別"], ["始話時間", "始話日期時間"], ["調閱門號", "目標電話", "調閱號碼"], ["對象門號", "對象電話", "通話對象"], ["通話期間", "通話秒數", "通話時間(秒)"]] : kind === "data" ? [["查詢項目"], ["開始時間", "啟始時間", "通聯起始時間"], ["連線期間", "通聯時間(秒)"], ["開始基地台編號", "起始基地台編號", "基地台 ID"], ["開始基地台", "起始基地台地址", "基地台位址"], ["結束基地台編號", "離開基地台編號", "最終基地台 ID"], ["結束基地台", "離開基地台地址", "最終基地台位址"]] : [["用戶名稱"], ["查詢項目", "用戶回應-用戶編號/帳號"]];
    return groups.every(aliases => aliases.some(alias => keys.has(model.canonicalHeader(alias))));
  }
  function column(reference) {
    const match = reference.match(/^([A-Z]{1,3})([1-9]\d*)$/);
    if (!match) throw invalid("儲存格位置");
    let index = 0; for (const char of match[1]) index = index * 26 + char.charCodeAt(0) - 64;
    if (index > 16384 || Number(match[2]) > 1048576) throw invalid("儲存格位置");
    return { index: index - 1, row: Number(match[2]) };
  }
  function stringLookup(store, signal) {
    const cache = new Map(); let bytes = 0;
    function remember(id, value) {
      if (cache.has(id)) return;
      const valueBytes = value.length * 2;
      while (cache.size && (cache.size >= 4096 || bytes + valueBytes > 2 * 1024 * 1024)) { const key = cache.keys().next().value; bytes -= cache.get(key).length * 2; cache.delete(key); }
      cache.set(id, value); bytes += valueBytes;
    }
    return async row => {
      let rowBytes = 0;
      const references = new Map();
      for (let index = 0; index < row.values.length; index++) {
        const cell = row.values[index];
        if (cell?.type === "s") {
          if (!references.has(cell.value)) references.set(cell.value, []);
          references.get(cell.value).push(index);
        } else rowBytes += String(cell?.value ?? cell ?? "").length * 2;
      }
      function replace(id, value) {
        if (typeof value !== "string" || value.length > MAX_NODE) throw invalid("共用字串大小");
        for (const index of references.get(id)) {
          rowBytes += value.length * 2;
          if (rowBytes > MAX_ROW) throw invalid("資料列大小", row.rowNumber);
          row.values[index] = value;
        }
      }
      const missing = [];
      for (const id of references.keys()) { if (cache.has(id)) replace(id, cache.get(id)); else missing.push(id); }
      // Even maximally sized values cannot make one storage response unbounded.
      for (let i = 0; i < missing.length; i += 16) {
        const ids = missing.slice(i, i + 16);
        abort(signal); const found = await store.getMany(ids); abort(signal);
        for (const id of ids) {
          if (!found.has(id)) throw invalid("共用字串索引");
          const value = found.get(id); replace(id, value); remember(id, value);
        }
      }
    };
  }
  async function worksheet(entry, kind, context, store, sink, progress, signal) {
    let row = null, cell = null, inValue = false, inText = false, phonetic = false, rowBytes = 0, previousRow = 0, queue = [], headers = null, hadValues = false, root = false;
    const resolve = stringLookup(store, signal);
    await xml(entry, {
      open(tag, node, depth) {
        if (depth === 1) { if (tag !== "worksheet") throw invalid("工作表"); root = true; }
        if (tag === "row") {
          if (row) throw invalid("資料列");
          const number = Number(attr(node, "r") || previousRow + 1);
          if (!Number.isInteger(number) || number <= previousRow || number > 1048576) throw invalid("資料列位置");
          row = { rowNumber: number, values: [] }; rowBytes = 0;
        }
        if (tag === "c") {
          if (!row || cell) throw invalid("儲存格");
          const position = attr(node, "r") ? column(attr(node, "r")) : { index: row.values.length, row: row.rowNumber };
          if (position.row !== row.rowNumber || position.index < row.values.length) throw invalid("儲存格位置", row.rowNumber);
          cell = { index: position.index, value: "", type: attr(node, "t") || "n", style: Number(attr(node, "s") || 0) };
        }
        if (tag === "v" && cell) inValue = true;
        if (tag === "rPh") phonetic = true;
        if (tag === "t" && cell && !phonetic) inText = true;
      },
      text(value) { if (cell && (inValue || inText)) { cell.value += value; rowBytes += value.length * 2; if (cell.value.length > MAX_NODE || rowBytes > MAX_ROW) throw invalid("資料列大小", row.rowNumber); } },
      close(tag) {
        if (tag === "v") inValue = false; if (tag === "t") inText = false; if (tag === "rPh") phonetic = false;
        if (tag === "c") {
          if (!cell || !row) throw invalid("儲存格");
          if (cell.type === "s") { if (!/^\d+$/.test(cell.value) || !Number.isSafeInteger(Number(cell.value))) throw invalid("共用字串索引", row.rowNumber); cell.value = Number(cell.value); }
          row.values[cell.index] = ["s", "n"].includes(cell.type) ? { value: cell.value, type: cell.type, style: cell.style } : cell.value;
          cell = null;
        }
        if (tag === "row") { if (!row || cell) throw invalid("資料列"); previousRow = row.rowNumber; queue.push(row); row = null; }
      },
      async flush() {
        if (!queue.length) return;

        for (const row of queue) {
          await resolve(row);
          abort(signal); progress.row();
          const values = Array.from(row.values, value => value ?? "");
          // Release resolved strings before expanding the next queued source row.
          row.values = [];
          if (!values.some(value => String(value?.value ?? value).trim())) continue;
          hadValues = true;
          if (headerMatches(kind, values)) { headers = values; continue; }
          if (!headers) continue;
          const rowContext = { ...context, rowNumber: row.rowNumber };
          const record = kind === "user" ? model.normalizeUser(headers, values, rowContext) : model.normalizeRow(kind, headers, values, rowContext);
          if (record) await sink(record);
        }
        queue = [];
      },
      end() { if (!root || row || cell || (hadValues && !headers)) throw invalid("工作表標題"); }
    }, { signal, onBytes: progress.bytes });
  }
  async function importFile(file, { signal, stringStore, onUsers = async () => {}, onRecords = async () => {}, onProgress = () => {} } = {}) {
    if (!stringStore?.putMany || !stringStore?.getMany || !stringStore?.clear) throw invalid("字串儲存");
    let book, processedRows = 0, completedBytes = 0;
    const summary = { source_format: "multi_phone_streaming_xlsx", source_file: file.name || "", source_files: [file.name || ""], sheet_name: "多工作表", total_records: 0, call_count: 0, data_count: 0, warning_count: 0, date_bounds: null };
    try {
      abort(signal); book = await archive(file, signal);
      if (!book.supported) throw invalid("不支援的工作表");
      const progressFor = stage => ({ row() { processedRows++; }, bytes(amount) { onProgress({ stage, processedRows, completedBytes: completedBytes + amount, totalBytes: book.estimatedBytes }); } });
      const runPart = async (entry, stage, action) => { if (!entry) return; const progress = progressFor(stage); await action(progress); completedBytes += entry.uncompressedSize; };
      await runPart(book.shared, "sharedStrings", progress => sharedStrings(book.shared, stringStore, signal, progress.bytes));
      let styles = [];
      await runPart(book.styles, "styles", async progress => { styles = await readStyles(book.styles, signal, progress.bytes); });
      const users = batchSink(onUsers, signal), records = batchSink(onRecords, signal);
      // User metadata is available before either record sheet, regardless of workbook order.
      const sheets = book.sheets.filter(sheet => NAMES.has(sheet.name)).sort((a, b) => (a.name === "使用者資料" ? -1 : 0) - (b.name === "使用者資料" ? -1 : 0));
      for (const sheet of sheets) {
        abort(signal); const kind = sheet.name === "使用者資料" ? "user" : sheet.name === "通聯紀錄" ? "call" : "data";
        const context = { fileName: file.name || "", sheetName: sheet.name, date1904: book.date1904, styles };
        const sink = kind === "user" ? item => users.push(item) : async record => {
          summary.total_records++; summary[record.record_kind === "data" ? "data_count" : "call_count"]++; summary.warning_count += record.warning_count;
          const timestamp = new Date(record.occurred_at + "Z");
          if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(record.occurred_at) && Number.isFinite(timestamp.getTime()) && timestamp.toISOString().slice(0, 19) === record.occurred_at) {
            const date = record.occurred_at.slice(0, 10);
            if (!summary.date_bounds) summary.date_bounds = { start: date, end: date };
            else { if (date < summary.date_bounds.start) summary.date_bounds.start = date; if (date > summary.date_bounds.end) summary.date_bounds.end = date; }
          }
          await records.push(record);
        };
        const entry = book.entries.get(sheet.path);
        await runPart(entry, kind === "user" ? "users" : "records", progress => worksheet(entry, kind, context, stringStore, sink, progress, signal));
        if (kind === "user") await users.flush();
      }
      await users.flush(); await records.flush(); abort(signal);
      onProgress({ stage: "complete", processedRows, completedBytes, totalBytes: book.estimatedBytes });
      return summary;
    } catch (error) { rethrow(error, signal, "匯入"); }
    finally { try { if (book) await book.reader.close(); } finally { await stringStore.clear(); } }
  }
  return { inspect, importFile };
});
