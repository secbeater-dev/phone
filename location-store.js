(function (root, factory) {
  const api = factory(
    typeof module === "object" && module.exports
      ? require("./dataset-store")
      : root.PhoneDatasetStore,
  );
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.PhoneLocationStore = api;
})(globalThis, function (D) {
  "use strict";
  const req = (r) =>
    new Promise((ok, no) => {
      r.onsuccess = () => ok(r.result);
      r.onerror = () => no(r.error);
    });
  const range = (p) =>
    IDBKeyRange.bound(
      p,
      Array.isArray(p[1])
        ? [p[0], [...p[1], []]]
        : p.length === 1
          ? [p[0], [[]]]
          : [...p, []],
    );
  const size = (v) => new TextEncoder().encode(JSON.stringify(v)).length + 32;
  async function digest(value) {
    const bytes = new TextEncoder().encode(value);
    const result = await crypto.subtle.digest("SHA-256", bytes);
    return Array.from(new Uint8Array(result), (byte) =>
      byte.toString(16).padStart(2, "0"),
    ).join("");
  }
  const counties = [
    "臺北市",
    "新北市",
    "桃園市",
    "臺中市",
    "臺南市",
    "高雄市",
    "基隆市",
    "新竹市",
    "嘉義市",
    "新竹縣",
    "苗栗縣",
    "彰化縣",
    "南投縣",
    "雲林縣",
    "嘉義縣",
    "屏東縣",
    "宜蘭縣",
    "花蓮縣",
    "臺東縣",
    "澎湖縣",
    "金門縣",
    "連江縣",
  ];
  class LocationStore {
    constructor(store) {
      this.store = store;
      this.db = store.db;
    }
    source(name, index) {
      const s = this.db.transaction(name).objectStore(name);
      return index ? s.index(index) : s;
    }
    get(name, id, key) {
      return req(this.source(name).get([id, key]));
    }
    async put(name, row) {
      const tx = this.db.transaction(name, "readwrite");
      const done = new Promise((ok, no) => {
        tx.oncomplete = ok;
        tx.onabort = tx.onerror = () =>
          no(tx.error || new Error("Storage failed"));
      });
      tx.objectStore(name).put(row);
      await done;
    }
    async putMany(name, rows) {
      if (!rows.length) return;
      const tx = this.db.transaction(name, "readwrite");
      const done = new Promise((ok, no) => {
        tx.oncomplete = ok;
        tx.onabort = tx.onerror = () =>
          no(tx.error || new Error("Storage failed"));
      });
      for (const row of rows) tx.objectStore(name).put(row);
      await done;
    }
    async next(name, r, index) {
      return new Promise((ok, no) => {
        const q = this.source(name, index).openCursor(r);
        q.onsuccess = () =>
          ok(q.result ? { row: q.result.value, key: q.result.key } : null);
        q.onerror = () => no(q.error);
      });
    }
    async *walk(name, p, index, bounds) {
      let after;
      const meta = await this.store.metadata(p[0]);
      const max =
        name === "locationMatches"
          ? 2048
          : name === "locationEvents"
            ? (meta?.max_record_bytes || 4194304) + 1024
            : name === "locationCounts"
              ? Math.max(1024, meta?.max_record_bytes || 1024)
              : (meta?.max_user_bytes || 4194304) * 3;
      const limit = Math.max(1, Math.min(1000, Math.floor(4194304 / max)));
      while (true) {
        if (after && indexedDB.cmp(after, (bounds || range(p)).upper) >= 0)
          return;
        const rows = await req(
          this.source(name, index).getAll(
            after
              ? IDBKeyRange.bound(after, (bounds || range(p)).upper, true)
              : bounds || range(p),
            limit,
          ),
        );
        if (!rows.length) return;
        const last = rows[rows.length - 1];
        after =
          index === "ordered"
            ? name === "locationEvents"
              ? [last.dataset_id, last.area, last.time, last.phone, last.seq]
              : [
                  last.dataset_id,
                  last.start_at,
                  last.county_rank,
                  last.district_order,
                  last.key,
                ]
            : [last.dataset_id, name === "targets" ? last.phone : last.key];
        for (const row of rows) yield row;
      }
    }
    async page(
      name,
      p,
      { page = 1, pageSize = 500, index, predicate, project, bounds } = {},
    ) {
      page = Math.max(1, Number(page) || 1);
      let total = 0;
      const rows = [];
      let max = 1;
      // First pass fixes a byte-safe page capacity for every page in this query.
      for await (const row of this.walk(name, p, index, bounds)) {
        if (predicate && !(await predicate(row))) continue;
        total++;
        max = Math.max(max, size(project ? await project(row) : row));
      }
      pageSize = Math.max(1, Math.min(pageSize, Math.floor(4194304 / max)));
      let i = 0;
      for await (const row of this.walk(name, p, index, bounds)) {
        if (predicate && !(await predicate(row))) continue;
        if (i++ < (page - 1) * pageSize) continue;
        if (rows.length === pageSize) break;
        rows.push(project ? await project(row) : row);
      }
      return { rows, total, page, pageSize };
    }
    async analyze(
      id,
      {
        signal,
        classify,
        normalizePhone = (x) => String(x || ""),
        onProgress,
      } = {},
    ) {
      D.check(signal);
      await this.store.finish(id, { cards_ready: false });
      const metadata = await this.store.metadata(id);
      const previewLimit = Math.max(
        1,
        Math.min(
          8,
          Math.floor(
            4190208 / Math.max(1024, metadata.max_record_bytes || 1024),
          ),
        ),
      );
      let maxMatchBytes = 1;
      for (const name of [
        "locationEvents",
        "locationCounts",
        "locationMatches",
        "cardFields",
        "cardSources",
      ])
        await this.store.deleteRange(name, range([id]));
      const areaNames = new Map();
      const excluded = {
        missing_phone: 0,
        invalid_time: 0,
        invalid_address: 0,
      };
      let eventSeq = 0,
        total = 0,
        eventBytes = 0;
      let eventBatch = [];
      const saveEvent = async (row) => {
        const n = size(row);
        if (
          eventBatch.length &&
          (eventBatch.length >= 1000 || eventBytes + n > 4194304)
        ) {
          await this.putMany("locationEvents", eventBatch);
          eventBatch = [];
          eventBytes = 0;
        }
        eventBatch.push(row);
        eventBytes += n;
      };
      await this.store.scan(
        { datasetId: id },
        async (rows) => {
          for (const row of rows) {
            D.check(signal);
            const phone = normalizePhone(row.target_phone);
            if (!phone) {
              excluded.missing_phone++;
              continue;
            }
            const time = D.dateKey(row.occurred_at);
            if (!time) {
              excluded.invalid_time++;
              continue;
            }
            const areas = new Map(),
              stations = new Map(
                (row.stations || []).map((s) => [s.station_key, s]),
              );
            const refs = row.base_refs || [];
            if (!refs.length) excluded.invalid_address++;
            for (const ref of refs) {
              const s = stations.get(ref.station_key),
                a = s && !s.is_virtual && classify(s.address);
              if (!a) {
                excluded.invalid_address++;
                continue;
              }
              const area = a.county + "\0" + a.district;
              areaNames.set(area, a.district);
              let event = areas.get(area);
              if (!event) {
                event = { ...a, area, matched_stations: [], seen: new Set() };
                areas.set(area, event);
              }
              const key = ref.role + "\0" + ref.station_key;
              if (!event.seen.has(key)) {
                event.seen.add(key);
                event.matched_stations.push({
                  role: ref.role || "",
                  cell_id: s.cell_id || "",
                  address: s.address || "",
                });
              }
            }
            for (const a of areas.values()) {
              const key = ++eventSeq;
              await saveEvent({
                dataset_id: id,
                key,
                seq: "location-" + key,
                phone,
                time,
                ...{ area: a.area, county: a.county, district: a.district },
                source_record: {
                  source_file: row.source_file || "",
                  source_sheet: row.source_sheet || "",
                  row_number: row.row_number ?? "",
                  occurred_at: time,
                  target_phone: phone,
                  counterparty_phone: normalizePhone(row.counterparty_phone),
                  time_source: row.location_time_source || "record",
                  matched_stations: a.matched_stations,
                },
              });
            }
          }
          onProgress?.({ stage: "建立位置索引" });
        },
        { signal },
      );
      await this.putMany("locationEvents", eventBatch);
      eventBatch = [];
      // The classifier admits only the fixed 368 administrative districts.
      const districtOrder = new Map(
        [...areaNames.entries()]
          .sort((a, b) =>
            a[1].localeCompare(b[1], "zh-Hant", { numeric: true }),
          )
          .map(([key], i) => [key, i]),
      );
      let area = "",
        end = null,
        next = null,
        lastEnd = null,
        distinct = 0,
        count = 0,
        pending = [];
      const eventKey = (e) => [id, e.area, e.time, e.phone, e.seq];
      const right = this.walk("locationEvents", [id], "ordered")[
        Symbol.asyncIterator
      ]();
      next = (await right.next()).value;
      const changeBatch = async (events, delta) => {
        if (!events.length) return;
        const tx = this.db.transaction("locationCounts", "readwrite"),
          table = tx.objectStore("locationCounts");
        const done = new Promise((ok, no) => {
          tx.oncomplete = ok;
          tx.onabort = tx.onerror = () =>
            no(tx.error || new Error("Storage failed"));
        });
        done.catch(() => {});
        for (const e of events) {
          D.check(signal);
          const key = [e.area, e.phone],
            old = await req(table.get([id, key])),
            n = (old?.count || 0) + delta;
          if (!old?.count && n) distinct++;
          if (old?.count && !n) distinct--;
          table.put({
            dataset_id: id,
            key,
            area: e.area,
            phone: e.phone,
            positive: n > 0 ? 1 : 0,
            count: n,
          });
          count += delta;
        }
        await done;
      };
      let pendingBytes = 0;
      for await (const start of this.walk("locationEvents", [id], "ordered")) {
        D.check(signal);
        if (start.area !== area) {
          await changeBatch(pending, -1);
          pending = [];
          pendingBytes = 0;
          area = start.area;
          end = null;
          lastEnd = null;
          distinct = 0;
          count = 0;
        }
        const limit = Date.parse(start.time + "Z") + 1800000;
        if (
          next &&
          next.area === area &&
          Date.parse(next.time + "Z") <= limit
        ) {
          await changeBatch(pending, -1);
          pending = [];
          pendingBytes = 0;
          let added = [],
            addedBytes = 0;
          while (
            next &&
            next.area === area &&
            Date.parse(next.time + "Z") <= limit
          ) {
            const compact = { area: next.area, phone: next.phone };
            if (added.length && addedBytes + size(compact) > 4194304) {
              await changeBatch(added, 1);
              added = [];
              addedBytes = 0;
            }
            added.push(compact);
            addedBytes += size(compact);
            end = next;
            next = (await right.next()).value;
            if (added.length === 1000) {
              await changeBatch(added, 1);
              added = [];
              addedBytes = 0;
            }
          }
          await changeBatch(added, 1);
        }
        if (distinct >= 2 && end.key !== lastEnd) {
          lastEnd = end.key;
          const key = ++total;
          const preview = await req(
            this.source("locationCounts", "positivePhones").getAll(
              range([id, area, 1]),
              previewLimit,
            ),
          );
          const match = {
            dataset_id: id,
            key,
            id: "multi-location-match-" + key,
            start_at: start.time,
            end_at: end.time,
            county: start.county,
            district: start.district,
            district_order: districtOrder.get(start.area),
            county_rank: counties.indexOf(start.county),
            area,
            start: eventKey(start),
            end: eventKey(end),
            occurrence_count: count,
            phone_count: distinct,
            phones: preview.map((row) => row.phone),
          };
          maxMatchBytes = Math.max(maxMatchBytes, size(match));
          await this.put("locationMatches", match);
        }
        const compact = { area: start.area, phone: start.phone };
        if (pending.length && pendingBytes + size(compact) > 4194304) {
          await changeBatch(pending, -1);
          pending = [];
          pendingBytes = 0;
        }
        pending.push(compact);
        pendingBytes += size(compact);
        if (pending.length === 1000) {
          await changeBatch(pending, -1);
          pending = [];
          pendingBytes = 0;
        }
      }
      await this.buildCards(id, { signal });
      await this.store.finish(id, {
        location_analysis: { total, excluded },
        max_location_match_bytes: maxMatchBytes,
      });
      return { total, excluded };
    }
    async match(id, matchId) {
      return this.get(
        "locationMatches",
        id,
        Number(String(matchId).replace("multi-location-match-", "")),
      );
    }
    async eventPage(id, m, page) {
      const p = [id, m.area];
      return this.page("locationEvents", p, {
        page,
        index: "ordered",
        bounds: IDBKeyRange.bound(m.start, m.end),
        project: (r) => r.source_record,
      });
    }
    async locationPage(id, page = 1) {
      page = Math.max(1, Number(page) || 1);
      const metadata = await this.store.metadata(id);
      const pageSize = Math.max(
        1,
        Math.min(
          500,
          Math.floor(4190208 / (metadata?.max_location_match_bytes || 4190208)),
        ),
      );
      const queryRange = range([id]);
      const total = await req(
        this.source("locationMatches", "ordered").count(queryRange),
      );
      const offset = (page - 1) * pageSize;
      const firstKey = await new Promise((ok, no) => {
        const q = this.source("locationMatches", "ordered").openKeyCursor(
          queryRange,
        );
        let advanced = false;
        q.onerror = () => no(q.error);
        q.onsuccess = () => {
          const cursor = q.result;
          if (!cursor) return ok(null);
          if (offset && !advanced) {
            advanced = true;
            cursor.advance(offset);
            return;
          }
          ok(cursor.key);
        };
      });
      const matches = firstKey
        ? await req(
            this.source("locationMatches", "ordered").getAll(
              IDBKeyRange.bound(firstKey, queryRange.upper),
              pageSize,
            ),
          )
        : [];
      const rows = matches.map((m) => ({
        id: m.id,
        start_at: m.start_at,
        end_at: m.end_at,
        county: m.county,
        district: m.district,
        occurrence_count: m.occurrence_count,
        phone_count: m.phone_count,
        phones: m.phones || [],
      }));
      return { rows, total, page, pageSize };
    }
    async locationDetails(id, matchId, page = 1) {
      const m = await this.match(id, matchId);
      return m
        ? this.eventPage(id, m, page)
        : { rows: [], total: 0, page, pageSize: 500 };
    }
    async locationPhones(id, matchId, page = 1, pageSize = 500) {
      const m = await this.match(id, matchId);
      if (!m) return { rows: [], total: 0, page, pageSize };
      if (!(await this.get("locationCounts", id, ["ready", m.key]))) {
        let rows = [];
        for await (const e of this.walk(
          "locationEvents",
          [id, m.area],
          "ordered",
          IDBKeyRange.bound(m.start, m.end),
        )) {
          const k = [id, e.area, e.time, e.phone, e.seq];
          if (indexedDB.cmp(k, m.start) < 0) continue;
          if (indexedDB.cmp(k, m.end) > 0) break;
          rows.push({
            dataset_id: id,
            key: ["phones", m.key, e.phone],
            phone: e.phone,
          });
          if (rows.length === 1000) {
            await this.putMany("locationCounts", rows);
            rows = [];
          }
        }
        await this.putMany("locationCounts", rows);
        await this.put("locationCounts", {
          dataset_id: id,
          key: ["ready", m.key],
        });
      }
      return this.page("locationCounts", [id, ["phones", m.key]], {
        page,
        pageSize,
        project: (r) => ({ phone: r.phone }),
      });
    }
    async buildCards(id, { signal } = {}) {
      const meta = await this.store.metadata(id);
      if (meta.cards_ready) return;
      let page = 1;
      const batches = { cardFields: [], cardSources: [] };
      const batchBytes = { cardFields: 0, cardSources: 0 };
      const enqueue = async (name, row) => {
        const bytes = size(row);
        if (
          batches[name].length &&
          (batches[name].length >= 1000 || batchBytes[name] + bytes > 4194304)
        ) {
          await this.putMany(name, batches[name]);
          batches[name] = [];
          batchBytes[name] = 0;
        }
        batches[name].push(row);
        batchBytes[name] += bytes;
      };
      while (true) {
        D.check(signal);
        const users = await this.store.users({ datasetId: id }, { page });
        for (const u of users.rows) {
          for (const [name, value0] of Object.entries({
            ...u.subject,
            ...(u.carrier ? { 電信業者: u.carrier } : {}),
          })) {
            const value = String(value0 || "");
            if (!value) continue;
            const key = [u.phone, name, await digest(value)];
            await enqueue("cardFields", {
              dataset_id: id,
              key,
              value,
            });
            const source = {
              source_file: u.source_file || "",
              source_sheet: u.source_sheet || "",
              row_number: u.row_number ?? "",
            };
            await enqueue("cardSources", {
              dataset_id: id,
              key: [...key, await digest(JSON.stringify(source))],
              ...source,
            });
          }
        }
        if (page * users.pageSize >= users.total) break;
        page++;
      }
      await this.putMany("cardFields", batches.cardFields);
      await this.putMany("cardSources", batches.cardSources);
      await this.store.finish(id, { cards_ready: true });
    }
    async phoneCards(id, search = "", page = 1) {
      const query = String(search).trim().toLowerCase();
      return this.page("targets", [id], {
        page,
        pageSize: 20,
        predicate: async (t) => {
          if (!query || t.phone.includes(query)) return true;
          for await (const f of this.walk("cardFields", [id, [t.phone]])) {
            if (
              /姓名|名稱/.test(f.key[1]) &&
              f.value.toLowerCase().includes(query)
            )
              return true;
          }
          return false;
        },
        project: async (t) => {
          const names = [],
            carriers = [];
          for await (const f of this.walk("cardFields", [id, [t.phone]])) {
            const a = /姓名|名稱/.test(f.key[1])
              ? names
              : /業者|電信/.test(f.key[1])
                ? carriers
                : null;
            const preview =
              f.value.length > 80 ? f.value.slice(0, 80) + "…" : f.value;
            if (a && !a.includes(preview) && a.length < 8) a.push(preview);
          }
          return {
            phone: t.phone,
            names,
            carriers,
            call_count: t.call_count,
            data_count: t.data_count,
          };
        },
      });
    }
    async phoneCardDetails(id, phone, page = 1) {
      return this.page("cardFields", [id, [phone]], {
        page,
        project: async (f) => {
          const sources = await this.phoneCardSources(
            id,
            phone,
            f.key[1],
            f.value,
            1,
            1,
          );
          const row = {
            name: f.key[1],
            value: f.value,
            sources: sources.rows,
            source_count: sources.total,
          };
          if (size(row) > 4194304) row.sources = [];
          return row;
        },
      });
    }
    async phoneCardSources(id, phone, name, value, page = 1, pageSize = 500) {
      return this.page(
        "cardSources",
        [id, [phone, name, await digest(value)]],
        {
          page,
          pageSize,
          project: (r) => ({
            source_file: r.source_file,
            source_sheet: r.source_sheet,
            row_number: r.row_number,
          }),
        },
      );
    }
  }
  return { LocationStore };
});
