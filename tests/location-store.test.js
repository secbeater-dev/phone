const test = require("node:test");
const assert = require("node:assert/strict");
require("fake-indexeddb/auto");
const datasets = require("../dataset-store");
const app = require("../app");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");
test("JSON empty reports retain normalized subject phone associations", async (t) => {
  const store = await datasets.open("json-empty-" + crypto.randomUUID());
  t.after(() => store.destroy());
  const w = workerHarness(store);
  w.setStore(store);
  const file = {
    name: "synthetic-empty.json",
    size: 100,
    text: async () =>
      JSON.stringify({
        case: { subject: { 電話號碼: "0900000001", 用戶名稱: "合成空報表" } },
        records: [],
        base_stations: [],
      }),
  };
  const result = await w.handle(
    "importLocations",
    { files: [file] },
    new AbortController().signal,
    () => {},
  );
  const users = await store.users({ datasetId: result.dataset.id });
  assert.equal(users.rows[0].phone, "0900000001");
  const l = new (require("../location-store").LocationStore)(store);
  assert.equal(
    (await l.phoneCards(result.dataset.id, "合成空報表")).rows[0].phone,
    "0900000001",
  );
});
test("opening a page of overlapping windows does not materialize phone memberships", async (t) => {
  const store = await datasets.open("lazy-overlap-" + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin("d", {});
  const station = { station_key: "s", address: "臺北市中正區合成路" };
  const rows = Array.from({ length: 560 }, (_, i) => ({
    target_phone: "090000000" + ((i % 2) + 1),
    occurred_at: new Date(Date.UTC(2026, 0, 1, 0, i))
      .toISOString()
      .slice(0, 19),
    stations: [station],
    base_refs: [{ station_key: "s" }],
  }));
  await store.appendRecords("d", rows);
  await store.finish("d");
  const l = new (require("../location-store").LocationStore)(store);
  const result = await l.analyze("d", {
    classify: app.classifyTaiwanAdministrativeArea,
  });
  assert.ok(result.total > 500);
  const page = await l.locationPage("d");
  assert.equal(page.rows.length, 500);
  assert.equal(page.rows[0].phones.length, 2);
  const counts = await new Promise((resolve) => {
    const r = store.db
      .transaction("locationCounts")
      .objectStore("locationCounts")
      .getAll();
    r.onsuccess = () => resolve(r.result);
  });
  assert.equal(
    counts.filter((r) => r.key[0] === "phones" || r.key[0] === "ready").length,
    0,
  );
  await l.locationPhones("d", page.rows[0].id);
  const later = await new Promise((resolve) => {
    const r = store.db
      .transaction("locationCounts")
      .objectStore("locationCounts")
      .getAll();
    r.onsuccess = () => resolve(r.result);
  });
  assert.equal(later.filter((r) => r.key[0] === "ready").length, 1);
});
function workerHarness(
  store,
  LocationStore = require("../location-store").LocationStore,
) {
  const context = {
    importScripts() {},
    self: {},
    navigator: {
      storage: { estimate: async () => ({ quota: 1024 ** 3, usage: 0 }) },
    },
    crypto,
    TextEncoder,
    AbortController,
    DOMException,
    PhoneDatasetStore: datasets,
    PhoneLocationStore: { LocationStore },
    PhoneWorkbench: app,
    PhoneCdrModel: require("../cdr-model"),
    PhoneStreamingXlsx: { inspect: async () => ({ supported: false }) },
  };
  vm.createContext(context);
  return vm.runInContext(
    fs.readFileSync(path.join(__dirname, "../dataset-worker.js"), "utf8") +
      "\n({setStore(value){store=value;},handle})",
    context,
  );
}
test("district order follows the existing locale oracle", async (t) => {
  const store = await datasets.open("districts-" + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin("d", {});
  const stations = ["中正區", "大安區", "萬華區", "士林區"].map(
    (district, i) => ({
      station_key: "s" + i,
      address: "臺北市" + district + "合成路",
    }),
  );
  const records = stations.flatMap((s) =>
    [1, 2].map((n) => ({
      target_phone: "090000000" + n,
      occurred_at: "2026-01-01T10:00:00",
      stations: [s],
      base_refs: [{ station_key: s.station_key }],
    })),
  );
  await store.appendRecords("d", records);
  await store.finish("d");
  const l = new (require("../location-store").LocationStore)(store);
  await l.analyze("d", { classify: app.classifyTaiwanAdministrativeArea });
  const expected = app
    .computeMultiNumberLocationMatches({
      case: { subject: {} },
      records,
      base_stations: stations,
    })
    .matches.map((m) => m.district);
  assert.deepEqual(
    (await l.locationPage("d")).rows.map((m) => m.district),
    expected,
  );
});
test("wide card values use bounded keys and preserve complete values", async (t) => {
  const store = await datasets.open("wide-cards-" + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin("d", {});
  const value = "x".repeat(2300000);
  await store.appendUsers("d", [
    {
      phone: "0900000001",
      subject: { 合成欄位: value },
      source_file: "synthetic.xml",
    },
  ]);
  await store.finish("d");
  const l = new (require("../location-store").LocationStore)(store);
  await l.buildCards("d");
  for (const name of ["cardFields", "cardSources"]) {
    const rows = await new Promise((resolve) => {
      const r = store.db.transaction(name).objectStore(name).getAll();
      r.onsuccess = () => resolve(r.result);
    });
    for (const row of rows)
      assert.ok(new TextEncoder().encode(JSON.stringify(row)).length < 4194304);
  }
  const fields = await l.phoneCardDetails("d", "0900000001");
  assert.equal(fields.rows[0].value, value);
  assert.ok(new TextEncoder().encode(JSON.stringify(fields)).length < 4194304);
  assert.equal(
    (await l.phoneCardSources("d", "0900000001", "合成欄位", value)).total,
    1,
  );
});
test("Worker imports a combined batch above 16MiB and atomically retains old on analysis quota failure", async (t) => {
  const store = await datasets.open("worker-" + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin("old", {});
  await store.finish("old");
  const file = (i) => ({
    name: "synthetic-" + i + ".json",
    size: 9 * 1024 * 1024,
    text: async () =>
      JSON.stringify({
        case: { subject: {} },
        records: [
          {
            target_phone: "090000000" + i,
            occurred_at: "2026-01-01T10:00:00",
            base_refs: [{ station_key: "s" }],
          },
        ],
        base_stations: [{ station_key: "s", address: "臺北市中正區合成路" }],
        subjects: [
          { phone: "090000000" + i, subject: { 用戶名稱: "合成人物" + i } },
        ],
      }),
  });
  const w = workerHarness(store);
  w.setStore(store);
  const result = await w.handle(
    "importLocations",
    { files: [file(1), file(2)] },
    new AbortController().signal,
    () => {},
  );
  assert.equal(result.analysis.total, 1);
  assert.equal((await store.users({ datasetId: result.dataset.id })).total, 2);
  assert.ok(await store.metadata("old"));
  class Failing extends require("../location-store").LocationStore {
    async analyze() {
      throw new DOMException("synthetic", "QuotaExceededError");
    }
  }
  const bad = workerHarness(store, Failing);
  bad.setStore(store);
  await assert.rejects(
    bad.handle(
      "importLocations",
      { files: [file(1)] },
      new AbortController().signal,
      () => {},
    ),
    { name: "QuotaExceededError" },
  );
  assert.ok(await store.metadata("old"));
  const ids = await new Promise((resolve) => {
    const r = store.db
      .transaction("datasets")
      .objectStore("datasets")
      .getAllKeys();
    r.onsuccess = () => resolve(r.result);
  });
  assert.deepEqual(new Set(ids), new Set(["old", result.dataset.id]));
});
test("disk location windows match inclusive maximal oracle and lazy details", async (t) => {
  const store = await datasets.open("locations-" + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin("d", {});
  const station = {
    station_key: "s",
    cell_id: "S",
    address: "臺北市中正區合成路",
  };
  const records = [0, 0, 30, 31, 60].map((m, i) => ({
    target_phone: "090000000" + ((i % 2) + 1),
    occurred_at: `2026-01-01T${m >= 60 ? "11" : "10"}:${String(m % 60).padStart(2, "0")}:00`,
    base_refs: [{ station_key: "s", role: "start" }],
    stations: [station],
    row_number: i + 1,
  }));
  await store.appendRecords("d", records);
  await store.finish("d");
  const { LocationStore } = require("../location-store");
  const locations = new LocationStore(store);
  const result = await locations.analyze("d", {
    classify: app.classifyTaiwanAdministrativeArea,
    normalizePhone: app.normalizePhoneText,
  });
  const oracle = app.computeMultiNumberLocationMatches({
    case: { subject: {} },
    records,
    base_stations: [station],
  });
  assert.equal(result.total, oracle.matches.length);
  const page = await locations.locationPage("d");
  assert.deepEqual(
    page.rows.map((r) => [r.start_at, r.end_at, r.occurrence_count]),
    oracle.matches.map((r) => [r.start_at, r.end_at, r.occurrence_count]),
  );
  assert.equal(
    (await locations.locationDetails("d", page.rows[0].id)).total,
    oracle.matches[0].occurrence_count,
  );
  assert.equal((await locations.locationPhones("d", page.rows[0].id)).total, 2);
});
test("dense windows keep only boundaries and paginate every event and phone", async (t) => {
  const store = await datasets.open("dense-" + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin("d", {});
  const station = { station_key: "s", address: "臺北市中正區合成路" };
  const records = Array.from({ length: 1003 }, (_, i) => ({
    target_phone: "09" + String(i).padStart(8, "0"),
    occurred_at: "2026-01-01T10:00:00",
    stations: [station],
    base_refs: [{ station_key: "s", role: "primary" }],
  }));
  await store.appendRecords("d", records);
  await store.finish("d");
  const { LocationStore } = require("../location-store");
  const l = new LocationStore(store);
  assert.equal(
    (await l.analyze("d", { classify: app.classifyTaiwanAdministrativeArea }))
      .total,
    1,
  );
  const m = (await l.locationPage("d")).rows[0];
  assert.equal(m.phones.length, 8);
  assert.equal(m.phone_count, 1003);
  assert.equal((await l.locationDetails("d", m.id, 3)).rows.length, 3);
  assert.equal((await l.locationPhones("d", m.id, 3)).rows.length, 3);
  const raw = await new Promise((resolve) => {
    const r = store.db
      .transaction("locationMatches")
      .objectStore("locationMatches")
      .get(["d", 1]);
    r.onsuccess = () => resolve(r.result);
  });
  assert.equal(raw.source_records, undefined);
  assert.ok(raw.phones.length <= 8);
});
test("cards retain conflicting fields and deduplicated provenance, including empty reports", async (t) => {
  const store = await datasets.open("cards-" + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin("d", {});
  const users = Array.from({ length: 23 }, (_, i) => ({
    phone: "090000" + String(i).padStart(4, "0"),
    subject: { 用戶名稱: "合成" + i },
    source_file: "synthetic-" + i + ".xml",
  }));
  users.push({
    phone: users[0].phone,
    subject: { 用戶名稱: "另一合成人物" },
    carrier: "合成業者",
    source_file: "synthetic-b.xml",
  });
  users.push({ ...users[0] });
  await store.appendUsers("d", users);
  await store.finish("d");
  const { LocationStore } = require("../location-store");
  const l = new LocationStore(store);
  await l.buildCards("d");
  assert.equal((await l.phoneCards("d", "", 1)).rows.length, 20);
  assert.equal((await l.phoneCards("d", "", 2)).rows.length, 3);
  assert.equal((await l.phoneCards("d", "另一合成")).total, 1);
  const details = await l.phoneCardDetails("d", users[0].phone);
  assert.equal(details.rows.filter((r) => r.name === "用戶名稱").length, 2);
  assert.equal(
    (await l.phoneCardSources("d", users[0].phone, "用戶名稱", "合成0")).total,
    1,
  );
  await store.remove("d");
  for (const name of [
    "locationEvents",
    "locationCounts",
    "locationMatches",
    "cardFields",
    "cardSources",
  ])
    assert.equal(
      await new Promise((resolve) => {
        const r = store.db.transaction(name).objectStore(name).count();
        r.onsuccess = () => resolve(r.result);
      }),
      0,
    );
});
test("exclusions match oracle, abort and quota errors allow associated staging cleanup", async (t) => {
  const store = await datasets.open("cleanup-" + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin("old", {});
  await store.finish("old");
  await store.begin("new", {});
  const records = [
    { target_phone: "", occurred_at: "invalid" },
    { target_phone: "0900000001", occurred_at: "invalid" },
    {
      target_phone: "0900000001",
      occurred_at: "2026-01-01T10:00:00",
      base_refs: [],
    },
  ];
  await store.appendRecords("new", records);
  await store.finish("new");
  const { LocationStore } = require("../location-store");
  const l = new LocationStore(store);
  assert.deepEqual(
    (await l.analyze("new", { classify: app.classifyTaiwanAdministrativeArea }))
      .excluded,
    { missing_phone: 1, invalid_time: 1, invalid_address: 1 },
  );
  const c = new AbortController();
  c.abort();
  await assert.rejects(
    l.analyze("new", {
      signal: c.signal,
      classify: app.classifyTaiwanAdministrativeArea,
    }),
    { name: "AbortError" },
  );
  l.buildCards = async () => {
    throw new DOMException("synthetic", "QuotaExceededError");
  };
  await assert.rejects(
    l.analyze("new", { classify: app.classifyTaiwanAdministrativeArea }),
    { name: "QuotaExceededError" },
  );
  await store.recover("old");
  assert.ok(await store.metadata("old"));
  assert.equal(await store.metadata("new"), undefined);
});
