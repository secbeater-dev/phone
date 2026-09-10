const test = require('node:test');
const assert = require('node:assert/strict');
const parser = require('../streaming-xlsx.js');
const { legacyWorkbook } = require('./fixtures/legacy-workbook');

test('inspection reports all expanded ZIP parts for unsupported legacy workbooks', async () => {
  const fixture = legacyWorkbook(65 * 1024 * 1024);
  assert.ok(fixture.buffer.length < 16 * 1024 * 1024);
  const info = await parser.inspect(new File([fixture.buffer], 'synthetic-legacy.xlsx'));
  assert.equal(info.supported, false);
  assert.equal(info.expandedBytes, fixture.expandedBytes);
  assert.ok(info.expandedBytes > 64 * 1024 * 1024);
});
