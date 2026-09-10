const test = require('node:test');
const assert = require('node:assert/strict');
const app = require('../app');
const attachment = require('../attachment-export');
const ExcelJS = require('../vendor/exceljs.min');
let api = {};
try { api = require('../dataset-report'); } catch (error) { if (error.code !== 'MODULE_NOT_FOUND') throw error; }
const rows = [
  {target_phone:'0900000001',counterparty_phone:'0900000002',record_kind:'call',direction:'outbound',call_type:'發話',occurred_at:'2026-01-01T10:00:00',duration_seconds:30,imei:'123456789012345',source_file:'synthetic.xlsx',source_sheet:'通聯紀錄',row_number:2,base_refs:[],stations:[]},
  {target_phone:'0900000001',record_kind:'data',direction:'data',call_type:'數據',occurred_at:'2026-01-01T11:00:00',ended_at:'2026-01-01T11:10:00',duration_seconds:600,external_ipv4:'192.0.2.1',external_ipv6:'2001:db8::1',internal_ip:'10.0.0.1',imei:'123456789012345',source_file:'synthetic.xlsx',source_sheet:'網路歷程',row_number:2,base_refs:[],stations:[]}
];
test('volume report separates network detail and call statistics and declares volume scope',async()=>{
  assert.equal(typeof api.buildVolume,'function');
  const report=api.buildVolume({records:rows,users:[{phone:'0900000001',subject:{'用戶名稱':'合成人物'}}],metadata:{source_files:['synthetic.xlsx']},summary:{call_count:100,data_count:200},volume:2,label:'合成範圍',notes:{}},app);
  assert.equal(report.calls.length,1);assert.equal(report.network.length,1);
  assert.equal(report.profile.summary['通聯筆數'],1);
  assert.equal(report.profile.summary['通話總秒數'],30);
  assert.equal(report.profile.summary['連線總秒數'],600);
  assert.equal(report.network[0].external_ipv6,'2001:db8::1');
  assert.equal(report.stats.count.totalRows.find(row=>row.phone==='0900000001').seconds,30);
  assert.match(report.meta.scope_label,/第 2 卷/);assert.match(report.meta.scope_label,/100/);
  const bytes=await attachment.createAttachmentXlsx(report,ExcelJS);
  const book=new ExcelJS.Workbook();await book.xlsx.load(bytes);
  const sheet=book.getWorksheet('網路歷程');assert.ok(sheet);
  assert.match(JSON.stringify(sheet.getSheetValues()),/2001:db8::1/);
  assert.match(JSON.stringify(sheet.getSheetValues()),/11:10:00/);
  assert.ok(attachment.PDF_SECTIONS.some(section=>section.key==='network'));
});
test('new record kind and network/source fields survive legacy normalization',()=>{
  const workspace=app.normalizeWorkspace({case:{source_file:'synthetic.xlsx'},records:rows,base_stations:[]});
  assert.equal(workspace.records[1].record_kind,'data');
  assert.equal(workspace.records[1].external_ipv6,'2001:db8::1');
});
test('full-scope summary keeps global counts, percentages and rank offsets across volumes', () => {
  const summary = { call_count: 20, data_count: 10, call_seconds: 600, data_seconds: 100, invalid_dates: 1, first_seen: '2026-01-01T00:00:00', last_seen: '2026-01-02T00:00:00', hours: [{ hour: 0, label: '00-01', count: 29 }], counties: { '臺北市': 30 } };
  const report = api.buildSummary({ summary, metadata: { source_files: ['synthetic.xlsx'] }, label: '合成範圍', volume: 2, users: [], imeis: ['12345'], hotspots: [{ address: '合成地址', count: 15 }], stats: { count: { totalRows: [{ phone: '0900000001', count: 20 }] }, seconds: {} } });
  assert.equal(report.meta.rank_offset, 500);
  assert.equal(report.hours[0].percent, 100);
  assert.equal(report.hotspots[0].percent, 50);
  assert.equal(report.profile.summary['通聯筆數'], 20);
  assert.equal(report.stats.count.totalRows[0].count, 20);
});
