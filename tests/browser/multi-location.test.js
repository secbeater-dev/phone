const test = require('node:test');
const assert = require('node:assert/strict');
const { start, synthetic, waitForImport } = require('./helpers');

function input(name = 'synthetic-location.json', count = 23, padding = 0) {
  return { name, mimeType: 'application/json', buffer: Buffer.from(JSON.stringify({
    case: { source_file: name },
    records: Array.from({ length: count }, (_, i) => ({ target_phone: `0900${String(i + 1).padStart(6, '0')}`, occurred_at: '2026-01-01T10:00:00', record_kind: 'data', location_time_source: 'internet', base_refs: [{ station_key: 'SYN-1', role: 'primary' }] })),
    base_stations: [{ station_key: 'SYN-1', cell_id: 'SYN-1', address: '臺北市中正區合成地址' }],
    subjects: Array.from({ length: count }, (_, i) => ({ phone: `0900${String(i + 1).padStart(6, '0')}`, subject: { '用戶名稱': `合成人物 ${i + 1}`, '電信業者': '合成電信', '證件號碼': 'SYNTHETIC-ONLY' }, source_file: name })),
    padding: ' '.repeat(padding)
  })) };
}

test('independent bulk location imports, full-batch cards, pagination and failed replacement', { timeout: 120000 }, async t => {
  const { page, errors } = await start(t);
  await page.locator('#fileInput').setInputFiles(synthetic(3));
  await waitForImport(page);
  await page.locator('[data-view="multiLocation"]').click();
  assert.equal(await page.locator('#multiLocationFolderInput').getAttribute('webkitdirectory'), '');
  await page.locator('#multiLocationFileInput').setInputFiles([input('synthetic-a.json', 23, 9 * 1024 * 1024), input('synthetic-b.json', 23, 9 * 1024 * 1024)]);
  await page.waitForFunction(() => document.querySelector('#multiLocationImportStatus').textContent.includes('匯入完成'));
  await page.waitForFunction(() => document.querySelectorAll('.multi-location-match-row').length === 1);
  await page.waitForFunction(() => document.querySelectorAll('#multiLocationCards .phone-card').length === 20);
  assert.equal(await page.locator('.multi-location-match-row').count(), 1);
  assert.equal(await page.locator('#multiLocationCards .phone-card').count(), 20);
  await page.locator('#multiLocationCards [data-cards-page="next"]').click();
  await page.waitForFunction(() => document.querySelectorAll('#multiLocationCards .phone-card').length === 3);
  await page.locator('#multiLocationCards input[type="search"]').fill('合成人物 23');
  await page.waitForFunction(() => document.querySelectorAll('#multiLocationCards .phone-card').length === 1);
  await page.locator('#multiLocationCards [data-card-toggle]').click();
  await page.waitForFunction(() => document.querySelector('#multiLocationCards').textContent.includes('SYNTHETIC-ONLY'));
  await page.locator('[data-multi-location-detail]').click();
  await page.waitForFunction(() => document.querySelector('.multi-location-detail-row')?.textContent.includes('internet連線時間'));
  assert.match(await page.locator('.multi-location-detail-row').innerText(), /0900000023/);
  await page.locator('#multiLocationFileInput').setInputFiles({ name: 'synthetic-bad.xml', mimeType: 'text/xml', buffer: Buffer.from('<unknown/>') });
  await page.waitForFunction(() => document.querySelector('#multiLocationImportStatus').textContent.includes('失敗'));
  assert.equal(await page.locator('.multi-location-match-row').count(), 1);
  const locationWorker = page.workers().at(-1);
  await locationWorker.evaluate(() => { navigator.storage.estimate = async () => ({ quota: 1, usage: 1 }); });
  await page.locator('#multiLocationFileInput').setInputFiles(input('synthetic-quota.json', 2));
  await page.waitForFunction(() => document.querySelector('#multiLocationImportStatus').textContent.includes('空間不足'));
  assert.equal(await page.locator('.multi-location-match-row').count(), 1);
  await locationWorker.evaluate(() => { delete navigator.storage.estimate; });
  await page.locator('#multiLocationFileInput').setInputFiles(synthetic(50000));
  await page.locator('#multiLocationCancel').click();
  await page.waitForFunction(() => document.querySelector('#multiLocationCancel').hidden);
  assert.match(await page.locator('#multiLocationImportStatus').innerText(), /已取消/);
  assert.equal(await page.locator('.multi-location-match-row').count(), 1);
  await page.locator('[data-view="profile"]').click();
  await page.waitForFunction(() => document.querySelectorAll('#profileBatchCards .phone-card').length === 2);
  assert.match(await page.locator('#profileBatchCards').innerText(), /合成人物甲/);
  assert.doesNotMatch(await page.locator('#profileBatchCards').innerText(), /合成人物 23/);
  assert.deepEqual(errors, []);
});

test('new XML base clock, internet fallback, partial failure and JSON roundtrip', { timeout: 90000 }, async t => {
  const { page, errors } = await start(t);
  const xml = (phone, base, internet) => ({ name: `synthetic-${phone}.xml`, mimeType: 'text/xml', buffer: Buffer.from(`<查詢單><電信業者>合成電信</電信業者><通聯記錄查詢條件><電話號碼>${phone}</電話號碼></通聯記錄查詢條件><通聯資料組><通聯資料><用戶號碼>${phone}</用戶號碼><手機連到基地台的時間>${base}</手機連到基地台的時間><連到internet的時間>${internet}</連到internet的時間><基地台代碼>SYN-A</基地台代碼><基地台地址>臺北市中正區合成地址</基地台地址></通聯資料></通聯資料組></查詢單>`) });
  const XLSX = require('../../vendor/xlsx.full.min'), summary = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(summary, XLSX.utils.aoa_to_sheet([['時間區間','行政區','手機門號'],['合成時間','臺北市中正區','0900000001']]), 'summary');
  const files = [xml('0900000001', '2026/10/01 10:00:00', '2026/10/01 12:00:00'), xml('0900000002', '', '2026/10/01 10:30:00'), { name: 'synthetic-summary.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(XLSX.write(summary, { type: 'buffer', bookType: 'xlsx' })) }];
  await page.locator('[data-view="multiLocation"]').click();
  await page.locator('#multiLocationFileInput').setInputFiles(files);
  await page.waitForFunction(() => document.querySelector('#multiLocationImportStatus').textContent.includes('匯入完成'));
  await page.waitForFunction(() => document.querySelectorAll('.multi-location-match-row').length === 1);
  assert.match(await page.locator('#multiLocationImportResults').innerText(), /失敗/);
  assert.match(await page.locator('#multiLocationImportResults').innerText(), /摘要工作簿缺少原始紀錄欄位/);
  await page.locator('[data-multi-location-detail]').click();
  await page.waitForFunction(() => document.querySelector('.multi-location-detail-row')?.textContent.includes('internet連線時間'));
  assert.match(await page.locator('.multi-location-detail-row').innerText(), /基地台連線時間/);
  await page.locator('#fileInput').setInputFiles(files.slice(0, 2));
  await waitForImport(page);
  await page.locator('#dataExportViewButton').click();
  const downloadPromise = page.waitForEvent('download');
  await page.locator('#exportWorkspaceButton').click();
  const download = await downloadPromise;
  const bytes = require('node:fs').readFileSync(await download.path()), workspace = JSON.parse(bytes);
  assert.equal(workspace.records.find(row => row.target_phone === '0900000001').base_connected_at, '2026-10-01T10:00:00');
  assert.equal(workspace.records.find(row => row.target_phone === '0900000002').location_time_source, 'internet');
  assert.ok(workspace.subjects.some(row => row.phone === '0900000002'));
  await page.waitForFunction(() => document.querySelector('#datasetJsonCancel').hidden);
  await page.locator('#multiLocationFileInput').setInputFiles({ name: 'synthetic-roundtrip.json', mimeType: 'application/json', buffer: bytes });
  await page.waitForFunction(() => document.querySelector('#multiLocationCancel').hidden && document.querySelector('#multiLocationImportStatus').textContent.includes('匯入完成'));
  assert.equal(await page.locator('.multi-location-match-row').count(), 1);
  assert.deepEqual(errors, []);
});

test('ordinary user cards show the entire batch regardless of selected target', { timeout: 60000 }, async t => {
  const { page, errors } = await start(t);
  await page.locator('#fileInput').setInputFiles(input());
  await waitForImport(page);
  await page.locator('[data-view="profile"]').click();
  await page.waitForFunction(() => document.querySelectorAll('#profileBatchCards .phone-card').length === 20);
  await page.locator('#profileBatchCards [data-cards-page="next"]').click();
  await page.waitForFunction(() => document.querySelectorAll('#profileBatchCards .phone-card').length === 3);
  await page.locator('#profileBatchCards input[type="search"]').fill('0900000023');
  await page.waitForFunction(() => document.querySelectorAll('#profileBatchCards .phone-card').length === 1);
  await page.locator('#profileBatchCards [data-card-toggle]').click();
  await page.waitForFunction(() => document.querySelector('#profileBatchCards').textContent.includes('SYNTHETIC-ONLY'));
  assert.deepEqual(errors, []);
});

test('first-page failure and cancellation before adoption preserve the previous location batch', { timeout: 60000 }, async t => {
  const { page, errors } = await start(t);
  await page.locator('[data-view="multiLocation"]').click();
  await page.locator('#multiLocationFileInput').setInputFiles(input('synthetic-old.json', 3));
  await page.waitForFunction(() => document.querySelector('#multiLocationCancel').hidden && document.querySelectorAll('#multiLocationCards .phone-card').length === 3);
  const worker = page.workers().at(-1);
  await worker.evaluate(() => {
    const original = self.onmessage;
    self.onmessage = event => {
      if (event.data.op === 'locationPage') {
        self.onmessage = original;
        self.postMessage({ id: event.data.id, ok: false, name: 'Error', message: 'synthetic first-page failure' });
      } else original(event);
    };
  });
  await page.locator('#multiLocationFileInput').setInputFiles(input('synthetic-error.json', 4));
  await page.waitForFunction(() => document.querySelector('#multiLocationImportStatus').textContent.includes('synthetic first-page failure'));
  assert.equal(await page.locator('#multiLocationCards .phone-card').count(), 3);
  await worker.evaluate(() => {
    const original = self.onmessage;
    self.onmessage = event => {
      if (event.data.op === 'locationPage') {
        self.onmessage = original;
        self.postMessage({ id: event.data.id, progress: { stage: 'synthetic pending preview' } });
        setTimeout(() => original(event), 6000);
      } else original(event);
    };
  });
  await page.locator('#multiLocationFileInput').setInputFiles(input('synthetic-cancel.json', 5));
  await page.waitForFunction(() => document.querySelector('#multiLocationImportStatus').textContent.includes('synthetic pending preview'));
  await page.locator('#multiLocationCancel').click();
  await page.waitForFunction(() => document.querySelector('#multiLocationCancel').hidden, undefined, { timeout: 3000 });
  assert.match(await page.locator('#multiLocationImportStatus').innerText(), /已取消/);
  assert.equal(await page.locator('#multiLocationCards .phone-card').count(), 3);
  assert.equal(await page.locator('.multi-location-match-row').count(), 1);
  await page.locator('[data-multi-location-detail]').click();
  await page.waitForFunction(() => document.querySelector('.multi-location-detail-row')?.textContent.includes('synthetic-old.json'));
  assert.deepEqual(errors, []);
});
