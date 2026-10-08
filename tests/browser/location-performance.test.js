const test = require('node:test');
const assert = require('node:assert/strict');
const { start } = require('./helpers');

test('large synthetic location batch keeps dense events on disk and UI pages bounded', { skip: !process.env.PHONE_LOCATION_PERF, timeout: 900000 }, async t => {
  const { page, errors } = await start(t, { args: ['--js-flags=--max-old-space-size=512'] });
  const files = Array.from({ length: 4 }, (_, part) => ({ name: `synthetic-dense-${part}.json`, mimeType: 'application/json', buffer: Buffer.from(JSON.stringify({
    case: { source_file: `synthetic-dense-${part}.json` },
    base_stations: [{ station_key: 'SYN', address: '臺北市中正區合成地址' }],
    records: Array.from({ length: 25000 }, (_, i) => ({ target_phone: i % 2 ? '0900000001' : '0900000002', occurred_at: '2026-01-01T10:00:00', record_kind: 'data', base_refs: [{ station_key: 'SYN', role: 'primary' }], row_number: i + 1 }))
  })) }));
  await page.locator('[data-view="multiLocation"]').click();
  await page.evaluate(() => { window.__lag = 0; let last = performance.now(); window.__ticker = setInterval(() => { const now = performance.now(); window.__lag = Math.max(window.__lag, now - last - 50); last = now; }, 50); });
  const began = Date.now();
  await page.locator('#multiLocationFileInput').setInputFiles(files);
  await page.waitForFunction(() => document.querySelector('#multiLocationCancel').hidden && document.querySelector('#multiLocationImportStatus').textContent.startsWith('匯入完成'), undefined, { timeout: 800000 });
  assert.equal(await page.locator('.multi-location-match-row').count(), 1);
  assert.equal(await page.locator('#multiLocationCards .phone-card').count(), 2);
  assert.match(await page.locator('[data-multi-location-detail]').innerText(), /100,000/);
  await page.locator('[data-multi-location-detail]').click();
  await page.waitForFunction(() => document.querySelectorAll('.multi-location-source-table tbody tr').length === 500, undefined, { timeout: 120000 });
  await page.locator('[data-location-kind="details"][data-location-page="2"]').click();
  await page.waitForFunction(() => document.querySelector('.multi-location-detail-row').textContent.includes('第 2 頁'), undefined, { timeout: 120000 });
  assert.equal(await page.locator('.multi-location-source-table tbody tr').count(), 500);
  const lag = await page.evaluate(() => { clearInterval(window.__ticker); return Math.round(window.__lag); });
  t.diagnostic(JSON.stringify({ syntheticLocationEvents: 100000, importAndDetailsSeconds: Math.round((Date.now() - began) / 1000), maxUiTimerLagMs: lag }));
  assert.deepEqual(errors, []);
});
