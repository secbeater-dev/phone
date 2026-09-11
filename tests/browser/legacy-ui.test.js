const test = require('node:test');
const assert = require('node:assert/strict');
const { start, synthetic } = require('./helpers');

const legacy = () => ({name:'synthetic-legacy.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({case:{source_file:'synthetic',subject:{'用戶名稱':'合成使用者'}},records:[
  {target_phone:'0900000001',counterparty_phone:'0900000003',occurred_at:'2026-01-01T10:00:00',call_type:'發話',direction:'outbound',duration_seconds:30},
  {target_phone:'0900000002',counterparty_phone:'0900000004',occurred_at:'2026-01-02T11:00:00',call_type:'受話',direction:'inbound',duration_seconds:45}
],base_stations:[]}))});
async function imported(page, file) {
  await page.locator('#fileInput').setInputFiles(file);
  await page.waitForFunction(()=>document.querySelector('#importProgressModal').hidden && document.querySelector('#importStatus').textContent.includes('匯入完成'));
  await page.waitForFunction(()=>document.querySelector('#callsView').getAttribute('aria-busy') !== 'true');
}
test('empty views keep original controls and never show multi-phone controls', {timeout:60000}, async t=>{
  const {page,errors}=await start(t);
  assert.equal(await page.locator('#datasetToolbar').isVisible(),false);
  assert.equal(await page.locator('[data-view="network"]').isVisible(),false);
  assert.equal(await page.locator('#hoursView').isVisible(),true);
  assert.equal(await page.locator('.hour-tile').count(),24);
  assert.equal(await page.locator('.hour-column').count(),24);
  await page.locator('[data-view="calls"]').click();
  assert.equal(await page.locator('#recordSearch').isVisible(),true);
  assert.match(await page.locator('#callRows').innerText(),/尚未匯入/);
  await page.locator('#attachmentExportButton').click();
  assert.equal(await page.locator('[data-attachment-pdf="network"]').isVisible(),false);
  assert.deepEqual(errors,[]);
});
test('legacy imports retain all targets, old date modal and original analysis layouts', {timeout:60000}, async t=>{
  const {page,errors}=await start(t);
  await imported(page,legacy());
  assert.equal(await page.locator('#datasetToolbar').isVisible(),false);
  assert.equal(await page.locator('[data-view="network"]').isVisible(),false);
  assert.equal(await page.locator('#callRows tr').count(),2);
  assert.match(await page.locator('#callRows').innerText(),/0900000001/);
  assert.match(await page.locator('#callRows').innerText(),/0900000002/);
  assert.equal(await page.locator('.call-column-resizer').count(),10);
  await page.locator('#dateFilterButton').click();
  await page.locator('#dateFilterStartInput').fill('2026-01-02');
  await page.locator('#dateFilterEndInput').fill('2026-01-02');
  await page.locator('#dateFilterApplyButton').click();
  await page.waitForFunction(()=>document.querySelector('#callRows').children.length===1);
  assert.match(await page.locator('#callRows').innerText(),/0900000002/);
  await page.locator('[data-view="profile"]').click();
  await page.waitForFunction(()=>document.querySelector('#profileView').getAttribute('aria-busy')==='false');
  assert.equal(await page.locator('#profileSummaryCards .metric-card').count(),6);
  assert.match(await page.locator('#profileContent').innerText(),/合成使用者/);
  await page.locator('[data-view="stats"]').click();
  await page.waitForFunction(()=>document.querySelectorAll('#statsContent .stats-card').length===3);
  await page.locator('[data-stats-rank-mode="seconds"]').click();
  assert.equal(await page.locator('[data-stats-rank-mode="seconds"]').getAttribute('aria-pressed'),'true');
  assert.deepEqual(errors,[]);
});
test('format switching only changes UI after adoption, including mixed and failed batches', {timeout:120000}, async t=>{
  const {page,errors}=await start(t);
  await imported(page,[legacy(),synthetic(3)]);
  assert.equal(await page.locator('#datasetToolbar').isVisible(),true);
  assert.equal(await page.locator('[data-view="network"]').isVisible(),true);
  await page.locator('#fileInput').setInputFiles({name:'bad.xml',mimeType:'text/xml',buffer:Buffer.from('<unknown/>')});
  await page.waitForFunction(()=>document.querySelector('#importStatus').textContent.includes('不支援') || document.querySelector('#importStatus').textContent.includes('失敗'));
  assert.equal(await page.locator('#datasetToolbar').isVisible(),true);
  await imported(page,legacy());
  assert.equal(await page.locator('#datasetToolbar').isVisible(),false);
  assert.equal(await page.locator('[data-view="network"]').isVisible(),false);
  await imported(page,synthetic(3));
  await page.locator('[data-view="submission"]').click();
  await page.waitForFunction(()=>document.querySelector('#ticketSourceInput').value.includes('0900000001'));
  await page.locator('[data-view="calls"]').click();
  await page.locator('#datasetClear').click();
  await page.waitForFunction(()=>!document.querySelector('#datasetToolbar').checkVisibility());
  assert.equal(await page.locator('#hoursView').isVisible(),true);
  assert.equal(await page.locator('.hour-column').count(),24);
  await page.locator('[data-view="submission"]').click();
  assert.equal(await page.locator('#ticketSourceInput').inputValue(),'');
  assert.deepEqual(errors,[]);
});
test('legacy hotspot time search and date exclusions stay accurate without visiting profile first', {timeout:60000}, async t=>{
  const {page,errors}=await start(t);
  const records=['2026-01-01T10:00:00','2026-01-02T11:00:00','2026-01-03T12:00:00','not-a-date'].map(occurred_at=>({occurred_at,target_phone:'0900000001',counterparty_phone:'0900000002',call_type:'GPRS',direction:'data',record_kind:'data',duration_seconds:10,base_refs:[{station_key:'SYN',role:'primary'}]}));
  await imported(page,{name:'synthetic-time.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({case:{source_file:'synthetic'},records,base_stations:[{station_key:'SYN',address:'臺北市中正區合成地址'}]}))});
  await page.locator('#dateFilterButton').click();
  await page.locator('#dateFilterApplyButton').click();
  await page.waitForFunction(()=>document.querySelector('#callsView').getAttribute('aria-busy')==='false');
  assert.match(await page.locator('#dateFilterSummary').innerText(),/排除 1 筆/);
  await page.locator('[data-view="hours"]').click();
  await page.waitForFunction(()=>document.querySelector('#hoursView').getAttribute('aria-busy')==='false');
  await page.locator('#hourHotspotSearch').fill('01-02T11');
  await page.waitForFunction(()=>document.querySelector('#hoursView').getAttribute('aria-busy')==='false');
  assert.equal(await page.locator('[data-hotspot-address]').count(),1);
  assert.match(await page.locator('[data-hotspot-address]').innerText(),/3 筆/);
  await page.locator('#hotspotCountyFilterButton').click();
  assert.equal(await page.locator('[data-county-filter]').count(),23);
  await page.locator('#hotspotCountyClearAllButton').click();
  assert.equal(await page.locator('#hotspotCountyFilterApplyButton').isDisabled(),true);
  await page.locator('[data-county-filter="臺北市"]').check();
  await page.locator('#hotspotCountyFilterApplyButton').click();
  await page.locator('[data-view="stats"]').click();
  await page.waitForFunction(()=>document.querySelector('#statsView').getAttribute('aria-busy')==='false');
  assert.equal(await page.locator('#statsContent .stats-card').last().locator('.stats-table-row:not(.stats-table-head)').count(),2);
  assert.deepEqual(errors,[]);
});

test('legacy merged profile values deduplicate and IMEIs retain identifier order', {timeout:60000},async t=>{
  const {page}=await start(t);
  const input={case:{source_file:'synthetic'},records:['100000000000001','200000000000001','200000000000001'].map(imei=>({imei,target_phone:'0900000001',occurred_at:'2026-01-01T10:00:00',call_type:'發話'})),subjects:[{phone:'0900000001',subject:{'用戶名稱':'合成甲、合成乙'}},{phone:'0900000001',subject:{'用戶名稱':'合成甲'}}],base_stations:[]};
  await imported(page,{name:'synthetic-subjects.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(input))});
  await page.locator('[data-view="profile"]').click();
  await page.waitForFunction(()=>document.querySelector('#profileView').getAttribute('aria-busy')==='false');
  assert.equal(await page.locator('#profileContent dd').innerText(),'合成甲、合成乙');
  assert.deepEqual(await page.locator('.imei-chip').allTextContents(),['100000000000001','200000000000001']);
});

test('multi JSON roundtrip and partial failures retain only successful format provenance', {timeout:90000},async t=>{
  const {page,errors}=await start(t);
  const bad={name:'bad.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from('invalid')};
  await imported(page,[synthetic(3),bad]);
  assert.equal(await page.locator('#datasetToolbar').isVisible(),true);
  await page.locator('[data-view="export"]').click();
  const downloads=[];page.on('download',item=>downloads.push(item));
  await page.locator('#exportWorkspaceButton').click();
  await page.waitForFunction(()=>document.querySelector('#datasetJsonCancel').hidden);
  const files=[];for(const download of downloads) files.push({name:download.suggestedFilename(),mimeType:'application/json',buffer:require('node:fs').readFileSync(await download.path())});
  assert.ok(files.length);
  await imported(page,legacy());assert.equal(await page.locator('#datasetToolbar').isVisible(),false);
  await imported(page,files);assert.equal(await page.locator('#datasetToolbar').isVisible(),true);
  await page.locator('[data-view="network"]').click();
  await page.waitForFunction(()=>document.querySelector('#networkView').getAttribute('aria-busy')==='false');
  assert.match(await page.locator('#datasetRows').innerText(),/192\.0\.2\.1/);
  await imported(page,[legacy(),bad]);assert.equal(await page.locator('#datasetToolbar').isVisible(),false);
  const attachments=[];const collect=item=>attachments.push(item);page.on('download',collect);
  await page.locator('#attachmentExportButton').click();await page.locator('#attachmentXlsxButton').click();
  await page.waitForFunction(()=>document.querySelector('#datasetExportCancel').hidden);
  assert.ok(attachments.length>=2);
  const ExcelJS=require('../../vendor/exceljs.min');
  for(const download of attachments) {
    const book=new ExcelJS.Workbook();await book.xlsx.load(require('node:fs').readFileSync(await download.path()));
    assert.equal(book.worksheets.length,6);assert.equal(book.getWorksheet('網路歷程'),undefined);
  }
  assert.deepEqual(errors,[]);
});
