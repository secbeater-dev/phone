(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PhoneDatasetStore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const PREFIX = 'phone-workbench-session-';
  const PAGE = 500;
  const BATCH = 1000;
  const READ_BYTES = 4 * 1024 * 1024;
  const encodedSize = value => new TextEncoder().encode(JSON.stringify(value)).byteLength + 16;
  const MAX_KEY = '\uffff';
  const request = req => new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  function finished(tx) {
    return new Promise((resolve, reject) => {
      tx.oncomplete = resolve;
      tx.onabort = tx.onerror = () => reject(tx.error || new Error('Local storage transaction failed'));
    });
  }
  function check(signal) { if (signal?.aborted) throw new DOMException('Operation cancelled', 'AbortError'); }
  function prefixRange(prefix) { return IDBKeyRange.bound(prefix, [...prefix, []]); }
  function dateKey(value) {
    const s = String(value || '');
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})/);
    if (!m) return '';
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]));
    return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3] && +m[4] < 24 && +m[5] < 60 && +m[6] < 60 ? s.slice(0, 19).replace(' ', 'T') : '';
  }
  function groupKey(scope) {
    return JSON.stringify([scope.datasetId, scope.target ?? null, scope.kind || '', scope.date?.active ? [scope.date.start, scope.date.end] : null, scope.hours || null]);
  }
  function rangeFor(scope, after) {
    let index, prefix;
    if (scope.target !== undefined) {
      prefix = [scope.datasetId, scope.target];
      index = scope.kind ? 'targetKindTime' : 'targetTime';
      if (scope.kind) prefix.push(scope.kind);
    } else { index = 'datasetTime'; prefix = [scope.datasetId]; }
    const low = scope.date?.active ? scope.date.start + 'T00:00:00' : '';
    const high = scope.date?.active ? scope.date.end + 'T23:59:59' : MAX_KEY;
    return { index, range: IDBKeyRange.bound(after || [...prefix, low, 0], [...prefix, high, Number.MAX_SAFE_INTEGER], Boolean(after), false), prefix };
  }
  function stationAddresses(row) {
    const result = new Map();
    for (const s of row.stations || []) {
      if (s.is_virtual || !s.address) continue;
      const key = s.normalized_address || s.address.replace(/\s/g, '').replace(/台/g, '臺');
      if (key) result.set(key, s.address);
    }
    return result;
  }
  function matches(row, scope, options = {}) {
    if (scope.kind && row.record_kind !== scope.kind) return false;
    if (scope.hours && (!row.time_key || !scope.hours.includes(Number(row.time_key.slice(11, 13))))) return false;
    if (options.address && !stationAddresses(row).has(options.address)) return false;
    const text = String(options.search || '').trim().toLowerCase();
    if (!text) return true;
    return [row.source_file, row.source_sheet, row.occurred_at, row.ended_at, row.call_type, row.target_phone, row.counterparty_phone, row.imei, row.imsi, row.note, row.external_ip, row.external_ipv4, row.external_ipv6, row.internal_ip, options.notes?.[row.target_phone], options.notes?.[row.counterparty_phone], ...stationAddresses(row).values()].some(v => String(v || '').toLowerCase().includes(text));
  }
  function sortable(row, column, notes) {
    if (column === 'target_note') return String(notes?.[row.target_phone] || '').toLowerCase();
    if (column === 'counterparty_note') return String(notes?.[row.counterparty_phone] || '').toLowerCase();
    if (['duration_seconds', 'row_number', 'upload_bytes', 'download_bytes'].includes(column)) return Number(row[column] || 0);
    return String(row[column] || '').toLowerCase();
  }
  async function cursorPage(source, range, offset, limit, direction = 'next', predicate, project = row => row) {
    let total = 0;
    const rows = [];
    await new Promise((resolve, reject) => {
      const req = source.openCursor(range, direction);
      let advanced = false;
      req.onerror = () => reject(req.error);
      req.onsuccess = () => {
        const cursor = req.result;
        if (!cursor) return resolve();
        if (!predicate) {
          if (offset && !advanced) { advanced = true; cursor.advance(offset); return; }
          rows.push(project(cursor.value));
          if (rows.length >= limit) return resolve();
        } else if (predicate(cursor.value)) {
          if (total >= offset && rows.length < limit) rows.push(project(cursor.value));
          total++;
        }
        cursor.continue();
      };
    });
    return { rows, total };
  }
  async function open(sessionId) {
    const name = PREFIX + sessionId;
    const req = indexedDB.open(name, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      db.createObjectStore('datasets', { keyPath: 'id' });
      const records = db.createObjectStore('records', { keyPath: ['dataset_id', 'seq'] });
      records.createIndex('datasetTime', ['dataset_id', 'time_key', 'seq']);
      records.createIndex('targetTime', ['dataset_id', 'target_phone', 'time_key', 'seq']);
      records.createIndex('targetKindTime', ['dataset_id', 'target_phone', 'record_kind', 'time_key', 'seq']);
      db.createObjectStore('targets', { keyPath: ['dataset_id', 'phone'] });
      const users = db.createObjectStore('users', { keyPath: 'id', autoIncrement: true });
      users.createIndex('datasetTarget', ['dataset_id', 'phone']);
      db.createObjectStore('strings', { keyPath: ['dataset_id', 'id'] });
      db.createObjectStore('sorts', { keyPath: ['query', 'sort_value', 'seq'] });
      const agg = db.createObjectStore('aggregates', { keyPath: ['query', 'group', 'key'] });
      agg.createIndex('rank_count', ['query', 'group', 'neg_count', 'neg_seconds', 'key']);
      agg.createIndex('rank_seconds', ['query', 'group', 'neg_seconds', 'neg_count', 'key']);
      db.createObjectStore('summaries', { keyPath: 'queryKey' });
    };
    const db = await request(req);
    db.onversionchange = () => db.close();
    return new DatasetStore(db);
  }
  class DatasetStore {
    constructor(db) { this.db = db; this.sortQuery = ''; this.sortTotal = 0; this.summaryJobs = new Map(); }
    async metadata(id) { return request(this.db.transaction('datasets').objectStore('datasets').get(id)); }
    async begin(id, metadata) {
      const tx = this.db.transaction('datasets', 'readwrite'), done = finished(tx);
      tx.objectStore('datasets').put({ ...metadata, id, status: 'importing', total_records: 0, call_count: 0, data_count: 0, date_bounds: { start: '', end: '' }, warning_count: 0 });
      await done;
    }
    async finish(id, extra = {}) {
      const meta = await this.metadata(id);
      if (!meta) throw new Error('Dataset not found');
      const tx = this.db.transaction('datasets', 'readwrite'), done = finished(tx);
      const result = { ...meta, ...extra, id, total_records: meta.total_records, call_count: meta.call_count, data_count: meta.data_count, date_bounds: meta.date_bounds, status: 'ready' };
      tx.objectStore('datasets').put(result); await done; return result;
    }
    async appendUsers(id, rows) {
      if (!rows.length) return;
      const tx = this.db.transaction(['users', 'targets', 'datasets'], 'readwrite'), done = finished(tx);
      done.catch(() => {});
      const metaRequest = request(tx.objectStore('datasets').get(id));
      const store = tx.objectStore('targets');
      const keys = [...new Set(rows.map(r => r.phone || ''))];
      const existing = await Promise.all(keys.map(phone => request(store.get([id, phone]))));
      const meta = await metaRequest;
      const targets = new Map(keys.map((phone, i) => [phone, existing[i] || { dataset_id: id, phone, call_count: 0, data_count: 0, names: [] }]));
      for (const row of rows) {
        const target = targets.get(row.phone || '');
        const fullName = String(row.subject?.['用戶名稱'] || '');
        const name = fullName.length > 80 ? fullName.slice(0, 80) + '…' : fullName;
        if (name && !target.names.includes(name) && target.names.length < 20) target.names.push(name);
        const user = { ...row, phone: row.phone || '', dataset_id: id };
        delete user.id;
        meta.max_user_bytes = Math.max(meta.max_user_bytes || 0, encodedSize(user));
        if (meta.max_user_bytes > READ_BYTES) { tx.abort(); await done; return; }
        tx.objectStore('users').add(user);
      }
      for (const value of targets.values()) store.put(value);
      tx.objectStore('datasets').put(meta);
      await done;
    }
    async appendRecords(id, rows) {
      if (!rows.length) return;
      const tx = this.db.transaction(['datasets', 'records', 'targets'], 'readwrite'), done = finished(tx);
      done.catch(() => {});
      const metaReq = request(tx.objectStore('datasets').get(id));
      const keys = [...new Set(rows.map(r => r.target_phone || ''))];
      const targetsStore = tx.objectStore('targets');
      const values = await Promise.all(keys.map(phone => request(targetsStore.get([id, phone]))));
      const meta = await metaReq;
      if (!meta || meta.status !== 'importing') { tx.abort(); await done; return; }
      const targets = new Map(keys.map((phone, i) => [phone, values[i] || { dataset_id: id, phone, call_count: 0, data_count: 0, names: [] }]));
      for (const row of rows) {
        const kind = row.record_kind === 'data' || row.direction === 'data' || /數據|上網/.test(row.call_type || '') ? 'data' : 'call';
        const normalized = { ...row, dataset_id: id, seq: ++meta.total_records, target_phone: row.target_phone || '', record_kind: kind, time_key: dateKey(row.occurred_at) };
        meta.max_record_bytes = Math.max(meta.max_record_bytes || 0, encodedSize(normalized));
        if (meta.max_record_bytes > READ_BYTES) { tx.abort(); await done; return; }
        tx.objectStore('records').put(normalized);
        targets.get(normalized.target_phone)[kind + '_count']++;
        meta[kind + '_count']++;
        if (normalized.time_key) {
          const day = normalized.time_key.slice(0, 10);
          if (!meta.date_bounds.start || day < meta.date_bounds.start) meta.date_bounds.start = day;
          if (!meta.date_bounds.end || day > meta.date_bounds.end) meta.date_bounds.end = day;
        } else meta.warning_count++;
      }
      targets.forEach(value => targetsStore.put(value));
      tx.objectStore('datasets').put(meta);
      await done;
    }
    stringStore(datasetId) {
      return {
        putMany: async rows => {
          const tx = this.db.transaction('strings', 'readwrite'), done = finished(tx);
          for (const row of rows) tx.objectStore('strings').put({ ...row, dataset_id: datasetId });
          await done;
        },
        getMany: async ids => {
          const store = this.db.transaction('strings').objectStore('strings');
          const result = await Promise.all(ids.map(id => request(store.get([datasetId, id]))));
          return new Map(ids.map((id, index) => [id, result[index]?.value]));
        },
        clear: () => this.deleteRange('strings', prefixRange([datasetId]))
      };
    }
    async targets(datasetId, { search = '', page = 1 } = {}) {
      const source = this.db.transaction('targets').objectStore('targets');
      const range = prefixRange([datasetId]);
      const query = search.trim().toLowerCase();
      const total = query ? 0 : await request(source.count(range));
      const source2 = this.db.transaction('targets').objectStore('targets');
      const result = await cursorPage(source2, range, (Math.max(1, page) - 1) * PAGE, PAGE, 'next', query ? r => [r.phone, ...r.names].join(' ').toLowerCase().includes(query) : null);
      return { rows: result.rows, total: query ? result.total : total, pageSize: PAGE };
    }
    async users(scope, { page = 1 } = {}) {
      const pageSize = await this.readSize(scope.datasetId, PAGE, 'max_user_bytes');
      const range = prefixRange(scope.target === undefined ? [scope.datasetId] : [scope.datasetId, scope.target]);
      const store = this.db.transaction('users').objectStore('users').index('datasetTarget');
      const total = await request(store.count(range));
      const result = await cursorPage(this.db.transaction('users').objectStore('users').index('datasetTarget'), range, (page - 1) * pageSize, pageSize);
      return { rows: result.rows, total, pageSize };
    }
    async scan(scope, consume, { signal, batchSize = BATCH } = {}) {
      let after;
      const safeSize = await this.readSize(scope.datasetId, batchSize);
      while (true) {
        check(signal);
        const { index, range, prefix } = rangeFor(scope, after);
        const rows = await request(this.db.transaction('records').objectStore('records').index(index).getAll(range, safeSize));
        check(signal);
        if (!rows.length) return;
        const last = rows[rows.length - 1];
        after = [...prefix, last.time_key, last.seq];
        const selected = rows.filter(row => matches(row, scope));
        if (selected.length) await consume(selected);
        check(signal);
        if (rows.length < safeSize) return;
      }
    }
    async page(scope, options = {}, signal) {
      check(signal);
      const pageSize = await this.readSize(scope.datasetId, Math.min(PAGE, Math.max(1, options.pageSize || PAGE)));
      const page = Math.max(1, Number(options.page) || 1), offset = (page - 1) * pageSize;
      const sort = options.sort || { column: 'occurred_at', direction: 'asc' };
      const direct = !options.search?.trim() && !options.address && !scope.hours && (!scope.kind || scope.target !== undefined) && sort.column === 'occurred_at';
      if (direct) {
        const { index, range } = rangeFor(scope);
        const total = await request(this.db.transaction('records').objectStore('records').index(index).count(range));
        check(signal);
        const result = await cursorPage(this.db.transaction('records').objectStore('records').index(index), range, offset, pageSize, sort.direction === 'desc' ? 'prev' : 'next');
        check(signal); return { rows: result.rows, total, page, pageSize };
      }
      const query = JSON.stringify([groupKey(scope), options.search || '', options.address || '', sort.column, options.notes || {}]);
      if (this.sortQuery !== query) {
        this.sortQuery = ''; this.sortTotal = 0;
        const tx = this.db.transaction('sorts', 'readwrite'), done = finished(tx); tx.objectStore('sorts').clear(); await done;
        let total = 0;
        await this.scan(scope, async rows => {
          const selected = rows.filter(row => matches(row, scope, options));
          if (!selected.length) return;
          let batch = [], bytes = 0;
          const flush = async () => {
            if (!batch.length) return;
            check(signal);
            const tx = this.db.transaction('sorts', 'readwrite'), done = finished(tx);
            for (const entry of batch) tx.objectStore('sorts').put(entry);
            await done; batch = []; bytes = 0;
          };
          for (const row of selected) {
            const entry = { query: 'active', sort_value: sortable(row, sort.column, options.notes), seq: row.seq, dataset_id: scope.datasetId }, size = encodedSize(entry);
            if (size > READ_BYTES) throw new Error('備註欄位超過排序容量，請縮短備註。');
            if (bytes + size > READ_BYTES) await flush();
            batch.push(entry); bytes += size;
          }
          await flush(); total += selected.length;
        }, { signal });
        check(signal); this.sortQuery = query; this.sortTotal = total;
      }
      const result = await cursorPage(this.db.transaction('sorts').objectStore('sorts'), prefixRange(['active']), offset, pageSize, sort.direction === 'desc' ? 'prev' : 'next', null, row => ({ dataset_id: row.dataset_id, seq: row.seq }));
      const recordStore = this.db.transaction('records').objectStore('records');
      const rows = await Promise.all(result.rows.map(r => request(recordStore.get([r.dataset_id, r.seq]))));
      check(signal); return { rows, total: this.sortTotal, page, pageSize };
    }
    async volume(scope, { after = null, limit = 10000, maxBytes = 16 * 1024 * 1024 } = {}, signal) {
      const result = []; let bytes = 0, cursor = after;
      const safeSize = await this.readSize(scope.datasetId, BATCH);
      while (result.length < limit) {
        check(signal);
        const { index, range, prefix } = rangeFor(scope, cursor);
        const rows = await request(this.db.transaction('records').objectStore('records').index(index).getAll(range, Math.min(safeSize, limit - result.length)));
        check(signal);
        if (!rows.length) return { rows: result, after: null };
        for (const row of rows) {
          if (matches(row, scope)) {
            const size = new TextEncoder().encode(JSON.stringify(row)).byteLength;
            if (size > maxBytes) throw new Error('A single record exceeds the export volume limit');
            if (bytes + size > maxBytes && result.length) return { rows: result, after: cursor };
            result.push(row); bytes += size;
          }
          cursor = [...prefix, row.time_key, row.seq];
        }
      }
      return { rows: result, after: cursor };
    }
    async summarize(scope, { signal, onProgress, classifyCounty = () => '未辨識' } = {}) {
      const queryKey = groupKey(scope);
      const cached = await request(this.db.transaction('summaries').objectStore('summaries').get(queryKey));
      if (cached) return cached;
      await this.deleteRange('aggregates', prefixRange([queryKey]));
      const summary = { queryKey, call_count: 0, data_count: 0, call_seconds: 0, data_seconds: 0, invalid_dates: 0, first_seen: '', last_seen: '', hours: Array.from({ length: 24 }, (_, hour) => ({ hour, label: String(hour).padStart(2, '0') + '-' + String(hour + 1).padStart(2, '0'), count: 0, call_count: 0, data_count: 0 })), counties: {} };
      await this.scan(scope, async rows => {
        const groups = new Map();
        function add(group, key, seconds = 0, extra = {}) {
          const id = group + '\0' + key;
          const value = groups.get(id) || { query: queryKey, group, key, count: 0, seconds: 0, ...extra };
          value.count++; value.seconds += seconds;
          if (extra.first_seen && (!value.first_seen || extra.first_seen < value.first_seen)) value.first_seen = extra.first_seen;
          if (extra.last_seen && (!value.last_seen || extra.last_seen > value.last_seen)) value.last_seen = extra.last_seen;
          groups.set(id, value);
        }
        for (const row of rows) {
          const kind = row.record_kind, seconds = Number(row.duration_seconds || 0);
          summary[kind + '_count']++; summary[kind + '_seconds'] += seconds;
          if (!row.time_key) summary.invalid_dates++;
          else {
            const h = summary.hours[Number(row.time_key.slice(11, 13))]; h.count++; h[kind + '_count']++;
            if (!summary.first_seen || row.time_key < summary.first_seen) summary.first_seen = row.time_key;
            if (!summary.last_seen || row.time_key > summary.last_seen) summary.last_seen = row.time_key;
          }
          if (row.imei) add('imei', row.imei);
          for (const phone of new Set([row.target_phone, row.counterparty_phone].filter(Boolean))) add('phone:submission', phone);
          if (kind === 'call') {
            if (row.target_phone) add('phone:total', row.target_phone, seconds, { role: '目標' });
            if (row.counterparty_phone) {
              add('phone:total', row.counterparty_phone, seconds, { role: '對象' });
              if (['inbound', 'outbound'].includes(row.direction)) add('phone:' + row.direction, row.counterparty_phone, seconds);
            }
          }
          for (const [key, address] of stationAddresses(row)) {
            const county = classifyCounty(address) || '未辨識';
            add('hotspot', key, 0, { address, county, first_seen: row.time_key, last_seen: row.time_key });
            summary.counties[county] = (summary.counties[county] || 0) + 1;
          }
        }
        const tx = this.db.transaction('aggregates', 'readwrite'), done = finished(tx); done.catch(() => {});
        const store = tx.objectStore('aggregates');
        const updates = [...groups.values()];
        const previous = await Promise.all(updates.map(v => request(store.get([v.query, v.group, v.key]))));
        updates.forEach((value, i) => {
          const old = previous[i];
          if (old) {
            value.count += old.count; value.seconds += old.seconds;
            if (old.first_seen && (!value.first_seen || old.first_seen < value.first_seen)) value.first_seen = old.first_seen;
            if (old.last_seen && (!value.last_seen || old.last_seen > value.last_seen)) value.last_seen = old.last_seen;
          }
          value.neg_count = -value.count; value.neg_seconds = -value.seconds;
          summary.max_aggregate_bytes = Math.max(summary.max_aggregate_bytes || 0, encodedSize(value)); store.put(value);
        });
        await done; onProgress?.({ stage: '統計中', processedRows: summary.call_count + summary.data_count });
      }, { signal });
      check(signal);
      const tx = this.db.transaction('summaries', 'readwrite'), done = finished(tx); tx.objectStore('summaries').put(summary); await done;
      return summary;
    }
    async aggregatePage(queryKey, group, { page = 1, pageSize = PAGE, mode = 'count', search = '', counties } = {}) {
      const summary = await request(this.db.transaction('summaries').objectStore('summaries').get(queryKey));
      pageSize = Math.max(1, Math.min(PAGE, pageSize, Math.floor(READ_BYTES / (summary?.max_aggregate_bytes || READ_BYTES))));
      const index = mode === 'seconds' ? 'rank_seconds' : 'rank_count', range = prefixRange([queryKey, group]);
      const predicate = search || counties ? row => (!search || String(row.address || row.key).toLowerCase().includes(search.toLowerCase())) && (!counties || counties.includes(row.county)) : null;
      const total = predicate ? 0 : await request(this.db.transaction('aggregates').objectStore('aggregates').index(index).count(range));
      const result = await cursorPage(this.db.transaction('aggregates').objectStore('aggregates').index(index), range, (page - 1) * Math.min(PAGE, pageSize), Math.min(PAGE, pageSize), 'next', predicate);
      return { rows: result.rows, total: predicate ? result.total : total, page, pageSize: Math.min(PAGE, pageSize) };
    }
    async deleteRange(name, range, index) {
      const tx = this.db.transaction(name, 'readwrite'), done = finished(tx);
      const store = tx.objectStore(name);
      if (!index) store.delete(range);
      else {
        const req = store.index(index).openKeyCursor(range);
        req.onsuccess = () => { const c = req.result; if (c) { store.delete(c.primaryKey); c.continue(); } };
      }
      await done;
    }
    async remove(id) {
      for (const name of ['records', 'targets', 'strings']) await this.deleteRange(name, prefixRange([id]));
      await this.deleteRange('users', prefixRange([id]), 'datasetTarget');
      const tx = this.db.transaction(['datasets', 'sorts', 'aggregates', 'summaries'], 'readwrite'), done = finished(tx);
      tx.objectStore('datasets').delete(id);
      for (const name of ['sorts', 'aggregates', 'summaries']) tx.objectStore(name).clear();
      this.sortQuery = ''; await done;
    }
    async discardIncomplete() {
      const rows = await request(this.db.transaction('datasets').objectStore('datasets').getAll());
      for (const row of rows) if (row.status !== 'ready') await this.remove(row.id);
    }
    async readSize(id, limit, field = 'max_record_bytes') {
      const meta = await this.metadata(id);
      return Math.max(1, Math.min(BATCH, limit, Math.floor(READ_BYTES / (meta?.[field] || READ_BYTES))));
    }
    async recover(retainId) {
      const rows = await request(this.db.transaction('datasets').objectStore('datasets').getAll());
      for (const row of rows) if (row.id !== retainId || row.status !== 'ready') await this.remove(row.id);
    }
    close() { this.db.close(); }
    async destroy() { const name = this.db.name; this.db.close(); await request(indexedDB.deleteDatabase(name)); }
  }
  return { open, PREFIX, check, dateKey, stationAddresses };
});
