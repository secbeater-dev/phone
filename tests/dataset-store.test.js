const test = require('node:test');
const assert = require('node:assert/strict');
require('fake-indexeddb/auto');
let api = {};
try { api = require('../dataset-store.js'); } catch (error) { if (error.code !== 'MODULE_NOT_FOUND') throw error; }

test('legacy ranks data records and finds middle hotspot times without reducing counts',async t=>{
  const store=await api.open('legacy-search-'+crypto.randomUUID()); t.after(()=>store.destroy());
  await store.begin('legacy',{ui_mode:'legacy'});
  await store.appendRecords('legacy',['2026-01-01T10:00:00','2026-01-02T11:00:00','2026-01-03T12:00:00','invalid'].map(occurred_at=>({occurred_at,target_phone:'0900000001',counterparty_phone:'0900000002',record_kind:'data',duration_seconds:10,stations:[{station_key:'SYN',address:'臺北市合成地址'}]})));
  await store.finish('legacy');
  const scope={datasetId:'legacy'},summary=await store.summarize(scope);
  const ranks=await store.aggregatePage(summary.queryKey,'phone:total');
  assert.equal(ranks.total,2); assert.ok(ranks.rows.every(row=>row.count===4 && row.seconds===40));
  assert.equal((await store.metadata('legacy')).invalid_dates,1);
  const dated={...scope,date:{active:true,start:'2026-01-01',end:'2026-01-03'}};
  const found=await store.hotspotPage(dated,{search:'01-02T11'});
  assert.equal(found.total,1); assert.equal(found.rows[0].count,3);
  assert.equal((await store.hotspotPage(dated,{search:'absent'})).total,0);
  assert.equal((await store.hotspotPage(dated,{search:'合成'})).rows[0].count,3);
  assert.equal((await store.metadata('legacy')).invalid_dates,1);
});

async function fixture(t) {
  assert.equal(typeof api.open, 'function', 'temporary dataset store must be implemented');
  const store = await api.open('test-' + crypto.randomUUID());
  t.after(() => store.destroy());
  await store.begin('d1', { source_files: ['synthetic.xlsx'] });
  await store.appendUsers('d1', [
    {phone:'0900000001',subject:{'用戶名稱':'合成人物甲'}},
    {phone:'0900000002',subject:{'用戶名稱':'合成人物乙'}}
  ]);
  const rows = [];
  for (let n=0;n<1203;n++) rows.push({target_phone:n===1202?'0900000002':'0900000001',record_kind:n===1201?'data':'call',occurred_at:n===1200?'invalid':`2026-01-${n<600?'01':'02'}T10:00:00`,counterparty_phone:'0900000003',direction:n%2?'inbound':'outbound',duration_seconds:n===1201?900:30,imei:'123456789012345',source_file:'synthetic.xlsx',source_sheet:'通聯紀錄',row_number:n+2,stations:[{station_key:'s1',address:'臺北市中正區合成地址',normalized_address:'臺北市中正區合成地址',is_virtual:false}],base_refs:[{station_key:'s1',role:'start'}]});
  for(let n=0;n<rows.length;n+=500) await store.appendRecords('d1', rows.slice(n,n+500));
  await store.finish('d1');
  return store;
}

test('queries exact target/kind/date pages without losing equal timestamps', async t => {
  const store=await fixture(t);
  const scope={datasetId:'d1',target:'0900000001',kind:'call'};
  const p=await store.page(scope,{page:1,pageSize:500});
  assert.equal(p.total,1201); assert.equal(p.rows.length,500);
  const last=await store.page(scope,{page:3,pageSize:500});
  assert.equal(last.rows.length,201);
  assert.equal(new Set([...p.rows,...last.rows].map(r=>r.seq)).size,701);
  const day=await store.page({...scope,date:{active:true,start:'2026-01-02',end:'2026-01-02'}},{page:1});
  assert.equal(day.total,600);
  const network=await store.page({...scope,kind:'data'},{page:1});
  assert.equal(network.total,1); assert.equal(network.rows[0].duration_seconds,900);
});

test('separates per-target users and full-scope call statistics from data sessions',async t=>{
  const store=await fixture(t);
  const scope={datasetId:'d1',target:'0900000001'};
  const profile=await store.users(scope);
  assert.deepEqual(profile.rows.map(x=>x.subject['用戶名稱']),['合成人物甲']);
  const summary=await store.summarize(scope);
  assert.equal(summary.call_count,1201);assert.equal(summary.data_count,1);
  assert.equal(summary.call_seconds,36030);assert.equal(summary.invalid_dates,1);
  assert.equal(summary.target_count,1);assert.equal(summary.counterparty_count,1);
  assert.equal(summary.total_duration_seconds,36930);
  const stats=await store.aggregatePage(summary.queryKey,'phone:total',{page:1});
  assert.equal(stats.rows.find(x=>x.key==='0900000003').count,1201);
  const hotspots=await store.aggregatePage(summary.queryKey,'hotspot',{page:1});
  assert.equal(hotspots.rows[0].count,1202);
  assert.equal('times' in hotspots.rows[0],false);
});

test('summary counts distinct targets and counterparties across call and data records', async t => {
  const store = await api.open('summary-counts-' + crypto.randomUUID()); t.after(() => store.destroy());
  await store.begin('counts', {});
  await store.appendRecords('counts', [
    { target_phone: '0900000001', counterparty_phone: '0900000011', record_kind: 'call', occurred_at: '2026-01-01T00:00:00', duration_seconds: 10 },
    { target_phone: '0900000002', counterparty_phone: '0900000011', record_kind: 'data', occurred_at: '2026-01-01T01:00:00', duration_seconds: 20 },
    { target_phone: '0900000002', counterparty_phone: '0900000012', record_kind: 'data', occurred_at: '2026-01-01T02:00:00', duration_seconds: 30 },
  ]);
  await store.finish('counts');
  const summary = await store.summarize({ datasetId: 'counts' });
  assert.equal(summary.target_count, 2);
  assert.equal(summary.counterparty_count, 2);
  assert.equal(summary.total_duration_seconds, 60);
});

test('legacy summary ranks both endpoints of data records while multi summary remains call-only', async t => {
  const store = await api.open('legacy-stats-' + crypto.randomUUID()); t.after(() => store.destroy());
  const row = { target_phone: '0900000001', counterparty_phone: '0900000002', record_kind: 'data', direction: 'data', occurred_at: '2026-01-01T00:00:00', duration_seconds: 45 };
  await store.begin('legacy', { ui_mode: 'legacy' });
  await store.appendRecords('legacy', [row]);
  await store.finish('legacy');
  const legacySummary = await store.summarize({ datasetId: 'legacy' });
  const legacyTotal = await store.aggregatePage(legacySummary.queryKey, 'phone:total');
  assert.deepEqual(legacyTotal.rows.map(item => [item.key, item.count, item.seconds]), [
    ['0900000001', 1, 45],
    ['0900000002', 1, 45],
  ]);

  await store.begin('multi', { ui_mode: 'multi' });
  await store.appendRecords('multi', [row]);
  await store.finish('multi');
  const multiSummary = await store.summarize({ datasetId: 'multi' });
  const multiTotal = await store.aggregatePage(multiSummary.queryKey, 'phone:total');
  assert.equal(multiTotal.total, 0);
});

test('dataset metadata retains full invalid-date count after a filtered summary', async t => {
  const store = await api.open('metadata-dates-' + crypto.randomUUID()); t.after(() => store.destroy());
  await store.begin('dates', { ui_mode: 'legacy' });
  await store.appendRecords('dates', [
    { target_phone: '0900000001', record_kind: 'call', occurred_at: 'invalid' },
    { target_phone: '0900000001', record_kind: 'call', occurred_at: '2026-01-01T00:00:00' },
  ]);
  await store.finish('dates');
  assert.equal((await store.metadata('dates')).invalid_dates, 1);
  const filtered = await store.summarize({ datasetId: 'dates', date: { active: true, start: '2026-01-01', end: '2026-01-01' } });
  assert.equal(filtered.invalid_dates, 0);
  assert.equal((await store.metadata('dates')).invalid_dates, 1);
});

test('searches and sorts across all pages and invalidates note-dependent searches',async t=>{
  const store=await fixture(t);
  const scope={datasetId:'d1',target:'0900000001',kind:'call'};
  const found=await store.page(scope,{search:'特定備註',notes:{'0900000003':'特定備註'},sort:{column:'duration_seconds',direction:'desc'},page:3});
  assert.equal(found.total,1201); assert.equal(found.rows.length,201);
  const none=await store.page(scope,{search:'特定備註',notes:{},page:1});
  assert.equal(none.total,0);
});

test('staging failure does not delete a completed dataset; deletion is isolated',async t=>{
  const store=await fixture(t);
  await store.begin('d2',{});
  await store.appendRecords('d2',[{target_phone:'0900000009',record_kind:'data',occurred_at:'2026-01-01T00:00:00'}]);
  await store.remove('d2');
  assert.equal((await store.page({datasetId:'d1'},{page:1})).total,1203);
  assert.equal(await store.metadata('d2'),undefined);
});

test('cancellation aborts scans and subsequent operations remain usable',async t=>{
  const store=await fixture(t);const abort=new AbortController();let seen=0;
  await assert.rejects(store.scan({datasetId:'d1'},async rows=>{seen+=rows.length;abort.abort();},{signal:abort.signal,batchSize:100}),{name:'AbortError'});
  assert.equal(seen,100);
  assert.equal((await store.page({datasetId:'d1',target:'0900000002'},{page:1})).total,1);
});

test('strings round-trip in batches and can be cleared without deleting records',async t=>{
  const store=await fixture(t);const strings=store.stringStore('d1');
  await strings.putMany([{id:1,value:'合成文字'},{id:2,value:'0900000001'}]);
  assert.deepEqual(await strings.getMany([2,1]),new Map([[2,'0900000001'],[1,'合成文字']]));
  await strings.clear();assert.equal((await strings.getMany([1])).get(1),undefined);
  assert.equal((await store.page({datasetId:'d1'},{page:1})).total,1203);
});

test('bounded export cursors cover every record once, including invalid dates',async t=>{
  const store=await fixture(t);let after=null;const seen=new Set();
  do {
    const chunk=await store.volume({datasetId:'d1'},{after,limit:137,maxBytes:200000});
    assert.ok(chunk.rows.length<=137);
    for(const row of chunk.rows){assert.equal(seen.has(row.seq),false);seen.add(row.seq);}
    after=chunk.after;
  } while(after);
  assert.equal(seen.size,1203);
});

test('worker recovery retains only the UI acknowledged dataset',async t=>{
  const store=await fixture(t);
  await store.begin('abandoned',{});await store.finish('abandoned');
  await store.begin('partial',{});
  await store.recover('d1');
  assert.equal(await store.metadata('abandoned'),undefined);
  assert.equal(await store.metadata('partial'),undefined);
  assert.equal((await store.metadata('d1')).total_records,1203);
});
test('wide records bound every read batch and page by bytes without skipping rows', async t => {
  const store = await api.open('wide-' + crypto.randomUUID()); t.after(() => store.destroy());
  await store.begin('wide', {});
  const row = { target_phone: '0900000001', record_kind: 'call', occurred_at: '2026-01-01T00:00:00', note: 'x'.repeat(600000) };
  await store.appendRecords('wide', Array.from({ length: 15 }, () => row)); await store.finish('wide');
  let count = 0;
  await store.scan({ datasetId: 'wide' }, rows => { assert.ok(Buffer.byteLength(JSON.stringify(rows)) <= 4 * 1024 * 1024); count += rows.length; });
  assert.equal(count, 15);
  let seen = [];
  for (let page = 1; ; page++) {
    const result = await store.page({ datasetId: 'wide' }, { page });
    assert.ok(Buffer.byteLength(JSON.stringify(result.rows)) <= 4 * 1024 * 1024);
    seen.push(...result.rows.map(row => row.seq));
    if (page * result.pageSize >= result.total) break;
  }
  assert.equal(new Set(seen).size, 15);
});
test('wide profiles and aggregate keys use bounded pages, and sort rows have compact keys', async t => {
  const store = await api.open('profiles-' + crypto.randomUUID()); t.after(() => store.destroy());
  await store.begin('wide', {});
  await store.appendUsers('wide', Array.from({length:15},(_,i)=>({phone:'0900000001',subject:{'用戶名稱':'合成姓名'.repeat(1000),'其他':String(i)+'x'.repeat(600000)}})));
  await store.appendRecords('wide', Array.from({length:15},(_,i)=>({target_phone:'0900000001',record_kind:'call',occurred_at:'2026-01-01T00:00:00',imei:String(i)+'x'.repeat(600000),note:'find'})));
  await store.finish('wide');
  const users = await store.users({datasetId:'wide'}); assert.ok(Buffer.byteLength(JSON.stringify(users.rows)) <= 4 * 1024 * 1024); assert.ok(users.pageSize < 15);
  const targets = await store.targets('wide'); assert.ok(targets.rows[0].names.every(name=>name.length<=81));
  const summary = await store.summarize({datasetId:'wide'});
  const aggregates = await store.aggregatePage(summary.queryKey,'imei'); assert.ok(Buffer.byteLength(JSON.stringify(aggregates.rows)) <= 4 * 1024 * 1024); assert.ok(aggregates.pageSize < 15);
  await store.page({datasetId:'wide'},{search:'find',notes:{'0900000001':'n'.repeat(100000)}});
  const sorted=await new Promise(resolve=>{const r=store.db.transaction('sorts').objectStore('sorts').getAll();r.onsuccess=()=>resolve(r.result);});
  assert.ok(sorted.every(row=>row.query.length<100));
});
