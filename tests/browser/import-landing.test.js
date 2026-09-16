const test = require('node:test');
const assert = require('node:assert/strict');
const { start, synthetic } = require('./helpers');

const workspace = () => ({ name: 'synthetic-hours.json', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify({
  case: { source_file: 'synthetic' },
  records: [{ target_phone: '0900000001', counterparty_phone: '0900000002', occurred_at: '2026-01-01T15:00:00', call_type: '發話', duration_seconds: 30 }],
  base_stations: []
})) });

async function assertHours(page, hour, count) {
  assert.equal(await page.locator('#hoursView').isVisible(), true, 'successful import opens the time distribution');
  assert.equal(await page.locator('button[data-view="hours"]').getAttribute('aria-current'), 'page');
  await page.waitForFunction(() => document.querySelector('#hoursView').getAttribute('aria-busy') !== 'true');
  assert.equal(await page.locator('.hour-column').count(), 24);
  assert.equal(await page.locator('.hour-bar-count').nth(hour).innerText(), String(count));
}

test('ordinary and workspace imports immediately show fresh time distributions', { timeout: 60000 }, async t => {
  const { page, errors } = await start(t);
  await page.locator('[data-view="calls"]').click();
  await page.locator('#fileInput').setInputFiles(synthetic(3));
  await page.locator('#importProgressModal').waitFor({ state: 'hidden' });
  await assertHours(page, 10, 3);

  // Import while already on hours: replace the previous chart, including partial success.
  await page.locator('#fileInput').setInputFiles([workspace(), { name: 'bad.xml', mimeType: 'text/xml', buffer: Buffer.from('<unknown/>') }]);
  await page.locator('#importProgressModal').waitFor({ state: 'hidden' });
  await assertHours(page, 15, 1);
  assert.equal(await page.locator('.hour-bar-count').nth(10).innerText(), '0');
  assert.match(await page.locator('#importResults').innerText(), /失敗/);

  await page.locator('[data-view="export"]').click();
  await page.locator('#importWorkspaceInput').setInputFiles(workspace());
  await page.locator('#importProgressModal').waitFor({ state: 'hidden' });
  await assertHours(page, 15, 1);

  await page.locator('[data-view="profile"]').click();
  await page.locator('#fileInput').setInputFiles({ name: 'bad.xml', mimeType: 'text/xml', buffer: Buffer.from('<unknown/>') });
  await page.locator('#importProgressModal').waitFor({ state: 'hidden' });
  assert.equal(await page.locator('#profileView').isVisible(), true, 'failed import preserves the current view');
  assert.deepEqual(errors, []);
});

test('legacy fallback file and workspace imports also open the time distribution', { timeout: 60000 }, async t => {
  const { page, errors } = await start(t);
  await page.evaluate(() => { globalThis.PhoneDatasetUI = undefined; });
  const XLSX = require('../../vendor/xlsx.full.min');
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ['始話時間', '通話秒數', '調閱號碼', 'IMEI', '通話類別', '通話對象', '轉接電話', '基地台/交換機', '備註'],
    ['2026-01-01T10:00:00', 30, '0900000001', '', '發話', '0900000002', '', '', '']
  ]), '通聯紀錄');
  await page.locator('[data-view="calls"]').click();
  await page.locator('#fileInput').setInputFiles({ name: 'synthetic-legacy.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })) });
  await page.locator('#importProgressModal').waitFor({ state: 'hidden' });
  await assertHours(page, 10, 1);
  await page.locator('[data-view="export"]').click();
  await page.locator('#importWorkspaceInput').setInputFiles(workspace());
  await page.waitForFunction(() => document.querySelector('#exportMessage').textContent.includes('已匯入'));
  await assertHours(page, 15, 1);
  assert.deepEqual(errors, []);
});
