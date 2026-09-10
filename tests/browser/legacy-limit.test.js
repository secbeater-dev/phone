const test = require('node:test');
const assert = require('node:assert/strict');
const { start, synthetic } = require('./helpers');
const { legacyWorkbook } = require('../fixtures/legacy-workbook');

test('compressed legacy XLSX exceeding the expanded limit preserves the current dataset', { timeout: 120000 }, async t => {
  const { page, errors } = await start(t);
  await page.locator('#fileInput').setInputFiles(synthetic(3));
  await page.waitForFunction(() => document.querySelector('#datasetTarget')?.options.length === 2);
  const before = await page.locator('#datasetTarget').inputValue();
  const fixture = legacyWorkbook(65 * 1024 * 1024);
  await page.locator('#fileInput').setInputFiles({ name: 'synthetic-legacy.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: fixture.buffer });
  await page.waitForFunction(() => document.querySelector('#importStatus')?.textContent.includes('64 MiB'));
  assert.equal(await page.locator('#datasetTarget').inputValue(), before);
  await page.locator('button[data-view="network"]').click();
  await page.waitForFunction(() => document.querySelector('#datasetRows')?.textContent.includes('192.0.2.1'));
  assert.deepEqual(errors, []);
});
