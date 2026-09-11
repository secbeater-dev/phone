const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '..', 'dataset-worker.js'), 'utf8');
const context = {
  importScripts() {},
  self: { onmessage: null },
  TextEncoder,
  AbortController,
  DOMException,
};
vm.createContext(context);
const helpers = vm.runInContext(`${source}\n({ sourceMetadata, mergedSourceMetadata, workspaceVolumeCase });`, context);

test('streaming imports select multi UI and retain their source format', () => {
  assert.deepEqual(
    structuredClone(helpers.sourceMetadata({ source_format: 'multi_phone_streaming_xlsx' })),
    { source_formats: ['multi_phone_streaming_xlsx'], ui_mode: 'multi' },
  );
});

test('legacy JSON uses raw case provenance while old workspace volumes remain legacy', () => {
  assert.deepEqual(
    structuredClone(helpers.sourceMetadata({ source_format: 'workspace_volume', source_formats: ['multi_phone_streaming_xlsx'] })),
    { source_formats: ['multi_phone_streaming_xlsx'], ui_mode: 'multi' },
  );
  assert.deepEqual(
    structuredClone(helpers.sourceMetadata({ source_format: 'workspace_volume' })),
    { source_formats: ['legacy'], ui_mode: 'legacy' },
  );
  assert.deepEqual(
    structuredClone(helpers.sourceMetadata({})),
    { source_formats: ['legacy'], ui_mode: 'legacy' },
  );
});

test('merged metadata is multi when any successful source is multi', () => {
  assert.deepEqual(
    structuredClone(helpers.mergedSourceMetadata([
      { source_formats: ['fet_order_cdr_xml'], ui_mode: 'legacy' },
      { source_formats: ['multi_phone_streaming_xlsx'], ui_mode: 'multi' },
    ])),
    { source_formats: ['fet_order_cdr_xml', 'multi_phone_streaming_xlsx'], ui_mode: 'multi' },
  );
});

test('workspace JSON volumes carry source provenance for reimport', () => {
  const volumeCase = structuredClone(helpers.workspaceVolumeCase({
    source_format: 'merged_datasets',
    source_formats: ['fet_order_cdr_xml', 'multi_phone_streaming_xlsx'],
    ui_mode: 'multi',
  }, 3, 27));
  assert.equal(volumeCase.source_format, 'workspace_volume');
  assert.deepEqual(volumeCase.source_formats, ['fet_order_cdr_xml', 'multi_phone_streaming_xlsx']);
  assert.equal(volumeCase.ui_mode, 'multi');
  assert.equal(volumeCase.volume, 3);
  assert.equal(volumeCase.total_records, 27);
});
