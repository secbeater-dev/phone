/* All data processing remains inside this same-origin dedicated worker. */
importScripts(
  './vendor/zip-no-worker-inflate-2.7.57.min.js?v=20260916-import-hours-v1', './vendor/sax-1.4.1.js?v=20260916-import-hours-v1',
  './cdr-model.js?v=20260916-import-hours-v1', './streaming-xlsx.js?v=20260916-import-hours-v1',
  './dataset-store.js?v=20260916-import-hours-v1', './dataset-report.js?v=20260916-import-hours-v1'
);
let store, queue = Promise.resolve(), legacyLoaded = false, xlsxLoaded = false, pdfLoaded = false;
const jobs = new Map();
function legacy() {
  if (!legacyLoaded) {
    importScripts('./vendor/xlsx.full.min.js?v=20260916-import-hours-v1', './attachment-export.js?v=20260916-import-hours-v1', './app.js?v=20260916-import-hours-v1');
    legacyLoaded = true;
  }
}
function safeError(error) {
  if (error.name === 'AbortError') return '已取消處理。';
  if (error.name === 'QuotaExceededError') return '本機暫存空間不足。請清除不需要的網站資料或釋放磁碟空間後重試。';
  return error.publicMessage || '檔案格式不符、資料損壞或本機儲存失敗；原有資料未變更。';
}
function publicError(message) { const error = new Error(message); error.publicMessage = message; return error; }
function sourceMetadata(source = {}) {
  const declared = Array.isArray(source.source_formats) ? source.source_formats : [];
  const fallback = source.source_format && !['workspace_volume', 'merged_datasets'].includes(source.source_format) ? [source.source_format] : [];
  const source_formats = [...new Set((declared.length ? declared : fallback).map(value => String(value || '').trim()).filter(Boolean))];
  if (!source_formats.length) source_formats.push('legacy');
  return { source_formats, ui_mode: source_formats.includes('multi_phone_streaming_xlsx') ? 'multi' : 'legacy' };
}
function mergedSourceMetadata(datasets = []) {
  const source_formats = [...new Set(datasets.flatMap(dataset => sourceMetadata(dataset).source_formats))];
  return { source_formats, ui_mode: source_formats.includes('multi_phone_streaming_xlsx') ? 'multi' : 'legacy' };
}
function workspaceVolumeCase(metadata, volume, totalRecords) {
  return { ...metadata, source_format: 'workspace_volume', volume, total_records: totalRecords, subject: {} };
}
async function importOne(file, id, signal, progress) {
  const info = /\.xlsx$/i.test(file.name) ? await PhoneStreamingXlsx.inspect(file, { signal }) : { supported: false };
  if (info.supported) {
    const estimate = await navigator.storage?.estimate?.();
    if (estimate?.quota && estimate.quota - (estimate.usage || 0) < Math.max(128 * 1024 * 1024, (info.estimatedBytes || 0) * 2.5)) throw publicError('本機暫存可用空間不足，請先釋放磁碟或網站儲存空間。');
  } else {
    if (file.size > 16 * 1024 * 1024) throw publicError('此舊格式超過 16 MiB，請先拆成較小檔案；多電話原始工作簿可直接使用大檔匯入。');
    if (info.expandedBytes > 64 * 1024 * 1024) throw publicError('此舊格式解壓後超過 64 MiB，請先拆成較小檔案；多電話原始工作簿可直接使用大檔匯入。');
  }
  await store.begin(id, { source_files: [file.name], source_file: file.name });
  try {
    let summary;
    if (info.supported) {
      summary = await PhoneStreamingXlsx.importFile(file, { signal, stringStore: store.stringStore(id), onUsers: rows => store.appendUsers(id, rows), onRecords: rows => store.appendRecords(id, rows), onProgress: progress });
      Object.assign(summary, sourceMetadata(summary));
    } else {
      legacy(); PhoneDatasetStore.check(signal);
      progress({ stage: '解析既有格式', processedRows: 0 });
      let workspace, subjects, sourceCase;
      if (/\.json$/i.test(file.name)) {
        const parsed = JSON.parse(await file.text());
        if (!parsed.case || !Array.isArray(parsed.records)) throw publicError('不是支援的 workspace JSON。');
        sourceCase = parsed.case;
        workspace = PhoneWorkbench.normalizeWorkspace(parsed);
        subjects = Array.isArray(parsed.subjects) ? parsed.subjects : null;
      } else {
        workspace = PhoneWorkbench.parseImportFile(file.name, await file.arrayBuffer());
        sourceCase = workspace.case;
      }
      PhoneDatasetStore.check(signal);
      const stations = new Map((workspace.base_stations || []).map(s => [s.station_key, s]));
      for (let i = 0; i < workspace.records.length; i += 1000) {
        PhoneDatasetStore.check(signal);
        await store.appendRecords(id, workspace.records.slice(i, i + 1000).map(row => ({ ...row, target_phone: PhoneCdrModel.normalizePhone(row.target_phone), record_kind: PhoneCdrModel.kindOf(row), stations: (row.base_refs || []).map(ref => stations.get(ref.station_key)).filter(Boolean) })));
        progress({ stage: '寫入本機暫存', processedRows: Math.min(i + 1000, workspace.records.length) });
      }
      const phones = [...new Set(workspace.records.map(r => PhoneCdrModel.normalizePhone(r.target_phone)))];
      if (subjects) {
        for (let i = 0; i < subjects.length; i += 500) {
          PhoneDatasetStore.check(signal);
          await store.appendUsers(id, subjects.slice(i, i + 500).map(row => ({ ...row, phone: PhoneCdrModel.normalizePhone(row.phone) })));
        }
      } else {
        const subject = workspace.case.subject || {};
        const explicit = [...new Set(['電話號碼', '申請號碼', '設備號碼', '調閱門號', '調閱號碼', '查詢項目'].flatMap(key => String(subject[key] || '').split(/[、,;；\s]+/)).map(PhoneCdrModel.normalizePhone).filter(phone => phone && phones.includes(phone)))];
        const phone = explicit.length === 1 ? explicit[0] : phones.length === 1 ? phones[0] : '';
        if (Object.keys(subject).length) await store.appendUsers(id, [{ phone, subject, source_file: file.name, source_sheet: workspace.case.sheet_name || '', row_number: 0 }]);
      }
      summary = { source_file: file.name, source_files: [file.name], source_format: workspace.case.source_format, sheet_name: workspace.case.sheet_name, warning_count: workspace.parse_warnings?.length || 0, ...sourceMetadata(sourceCase) };
    }
    return await store.finish(id, summary);
  } catch (error) { await store.remove(id).catch(() => {}); throw error; }
  finally { await store.stringStore(id).clear().catch(() => {}); }
}
async function importBatch(files, signal, progress) {
  const datasets = [], results = [];
  try {
    for (let index = 0; index < files.length; index++) {
      PhoneDatasetStore.check(signal);
      const file = files[index], id = crypto.randomUUID();
      const report = value => progress({ ...value, fileName: file.name, fileIndex: index + 1, fileCount: files.length });
      try { const dataset = await importOne(file, id, signal, report); datasets.push(dataset); results.push({ ok: true, fileName: file.name, format: dataset.source_format }); }
      catch (error) { if (error.name === 'AbortError') throw error; results.push({ ok: false, fileName: file.name, message: safeError(error) }); }
    }
    if (!datasets.length) throw publicError(results[0]?.message || '所有檔案匯入失敗，原有資料未變更。');
    let dataset = datasets[0];
    if (datasets.length > 1) {
      const id = crypto.randomUUID();
      await store.begin(id, { source_file: '多檔匯入', source_files: datasets.flatMap(d => d.source_files || []), source_format: 'merged_datasets', ...mergedSourceMetadata(datasets) });
      datasets.push({ id });
      for (const source of datasets.slice(0, -1)) {
        await store.scan({ datasetId: source.id }, rows => store.appendRecords(id, rows), { signal });
        let page = 1;
        while (true) {
          PhoneDatasetStore.check(signal);
          const users = await store.users({ datasetId: source.id }, { page });
          if (users.rows.length) await store.appendUsers(id, users.rows);
          if (page * users.pageSize >= users.total) break; page++;
        }
      }
      dataset = await store.finish(id);
      for (const source of datasets.slice(0, -1)) await store.remove(source.id);
    }
    return { dataset, results };
  } catch (error) { for (const d of datasets) await store.remove(d.id).catch(() => {}); throw error; }
}
async function handle(op, payload, signal, progress) {
  PhoneDatasetStore.check(signal);
  if (op === 'init') { store?.close(); store = await PhoneDatasetStore.open(payload.sessionId); await store.recover(payload.retainId); return true; }
  if (!store) throw publicError('本機暫存尚未初始化，請重新整理。');
  if (op === 'importBatch') return importBatch(payload.files, signal, progress);
  if (op === 'targets') return store.targets(payload.datasetId, payload.options);
  if (op === 'users') return store.users(payload.scope, payload.options);
  if (op === 'page') return store.page(payload.scope, payload.options, signal);
  if (op === 'summary') { legacy(); return store.summarize(payload.scope, { signal, onProgress: progress, classifyCounty: PhoneWorkbench.classifyTaiwanCounty }); }
  if (op === 'hotspots') { legacy(); return store.hotspotPage(payload.scope, payload.options, { signal, onProgress: progress, classifyCounty: PhoneWorkbench.classifyTaiwanCounty }); }
  if (op === 'ticketPhones') {
    legacy();
    const summary = await store.summarize({ datasetId: payload.datasetId }, { signal, classifyCounty: PhoneWorkbench.classifyTaiwanCounty });
    const result = await store.aggregatePage(summary.queryKey, 'phone:submission', { page: payload.page });
    return { phones: result.rows.map(row => row.key), total: result.total, pageSize: result.pageSize };
  }
  if (op === 'aggregate') return store.aggregatePage(payload.queryKey, payload.group, payload.options);
  if (op === 'volume') return store.volume(payload.scope, payload.options, signal);
  if (op === 'remove') { await store.remove(payload.datasetId); return true; }
  if (op === 'exportSummary') {
    legacy();
    const summary = await store.summarize(payload.scope, { signal, classifyCounty: PhoneWorkbench.classifyTaiwanCounty });
    const metadata = await store.metadata(payload.scope.datasetId);
    const users = await store.users(payload.scope, { page: payload.volume });
    let more = users.total > payload.volume * users.pageSize;
    const aggregate = async (group, mode = 'count') => {
      PhoneDatasetStore.check(signal);
      const result = await store.aggregatePage(summary.queryKey, group, { page: payload.volume, mode });
      more ||= result.total > payload.volume * result.pageSize;
      return result.rows.map((row, index) => ({ ...row, rank: (payload.volume - 1) * result.pageSize + index + 1 }));
    };
    const imeis = (await aggregate('imei')).map(row => row.key), hotspots = await aggregate('hotspot'), stats = { count: {}, seconds: {} };
    for (const mode of ['count', 'seconds']) for (const [group, key] of [['phone:inbound', 'inboundRows'], ['phone:outbound', 'outboundRows'], ['phone:total', 'totalRows']]) stats[mode][key] = (await aggregate(group, mode)).map(row => ({ ...row, phone: row.key, note: payload.notes?.[row.key] || '' }));
    const report = PhoneDatasetReport.buildSummary({ summary, metadata, users: users.rows, imeis, hotspots, stats, label: payload.label, volume: payload.volume });
    return { ...await handle('encodeAttachment', { ...payload, report }, signal, progress), more };
  }
  if (op === 'exportVolume' || op === 'jsonVolume') {
    legacy();
    const cursor = payload.after;
    const part = cursor?.recordsDone ? { rows: [], after: null } : await store.volume(payload.scope, { after: cursor?.records, maxBytes: op === 'jsonVolume' ? 8 * 1024 * 1024 : 16 * 1024 * 1024 }, signal);
    const metadata = await store.metadata(payload.scope.datasetId);
    const users = await store.users(payload.scope, { page: payload.volume });
    if (!part.rows.length && !users.rows.length) return { after: null };
    const after = part.after || users.total > payload.volume * users.pageSize ? { records: part.after, recordsDone: !part.after } : null;
    if (op === 'jsonVolume') {
      const stations = [...new Map(part.rows.flatMap(r => r.stations || []).map(s => [s.station_key, s])).values()];
      const records = part.rows.map(row => { const result = { ...row }; delete result.dataset_id; delete result.seq; delete result.time_key; delete result.stations; return result; });
      return { after, bytes: new TextEncoder().encode(JSON.stringify({ case: workspaceVolumeCase(metadata, payload.volume, records.length), records, base_stations: stations, subjects: users.rows, parse_warnings: [] })) };
    }
    const summary = await store.summarize(payload.scope, { signal, classifyCounty: PhoneWorkbench.classifyTaiwanCounty });
    const report = PhoneDatasetReport.buildVolume({ records: part.rows, users: users.rows, metadata, summary, volume: payload.volume, label: payload.label, notes: payload.notes }, PhoneWorkbench);
    const encoded = await handle('encodeAttachment', { format: payload.format, section: payload.section, report }, signal, progress);
    return { ...encoded, after };
  }
  if (op === 'encodeAttachment') {
    legacy();
    PhoneDatasetStore.check(signal);
    if (payload.format === 'xlsx') {
      if (!xlsxLoaded) { importScripts('./vendor/exceljs.min.js?v=20260916-import-hours-v1'); xlsxLoaded = true; }
      return { bytes: await PhoneAttachmentExport.createAttachmentXlsx(payload.report, ExcelJS), mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
    }
    if (!pdfLoaded) { importScripts('./vendor/pdf-lib.min.js?v=20260916-import-hours-v1', './vendor/fontkit.umd.min.js?v=20260916-import-hours-v1', './vendor/open-huninn-data.js?v=20260916-import-hours-v1'); pdfLoaded = true; }
    const binary = atob(PhoneExportFontBase64), font = Uint8Array.from(binary, c => c.charCodeAt(0));
    return { bytes: await PhoneAttachmentExport.createAttachmentPdf(payload.report, payload.section, PDFLib, fontkit, font), mime: 'application/pdf' };
  }
  throw publicError('未知的本機操作。');
}
self.onmessage = ({ data }) => {
  if (data.op === 'cancel') { for (const id of data.ids || []) jobs.get(id)?.abort(); return; }
  const controller = new AbortController(); jobs.set(data.id, controller);
  queue = queue.catch(() => {}).then(async () => {
    try {
      let lastProgress = 0, lastStage = '';
      const result = await handle(data.op, data.payload || {}, controller.signal, progress => {
        const now = performance.now();
        if (now - lastProgress >= 150 || progress.stage !== lastStage) {
          lastProgress = now; lastStage = progress.stage;
          self.postMessage({ id: data.id, progress });
        }
      });
      PhoneDatasetStore.check(controller.signal);
      const transfers = result?.bytes?.buffer instanceof ArrayBuffer ? [result.bytes.buffer] : result?.bytes instanceof ArrayBuffer ? [result.bytes] : [];
      self.postMessage({ id: data.id, ok: true, result }, transfers);
    } catch (error) { self.postMessage({ id: data.id, ok: false, name: error.name, message: safeError(error) }); }
    finally { jobs.delete(data.id); }
  });
};
