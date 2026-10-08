const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('../vendor/xlsx.full.min.js');
const app = require('../app.js');
const columns = ['資料來源','檔案序號','交換機代號','用戶號碼','用戶手機序號IMEI','手機連到基地台的時間','手機連到基地台的秒數','連到internet的時間','連到internet的秒數','用戶連線時被指配之內網IP','基地台代碼','基地台地址'];
const row = ['合成','1','交換機','0911111111','123456789012345','2026/10/01 10:00:00','60','2026/10/01 10:01:00','30','10.0.0.1','CELL-A','臺北市中正區合成路'];
function workbook(sheets) { const wb=XLSX.utils.book_new(); for(const [name,rows] of Object.entries(sheets)) XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),name); return XLSX.write(wb,{type:'buffer',bookType:'xlsx'}); }
function xmlRecord(values) { return '<通聯資料>'+columns.map((k,i)=>`<${k}>${values[i]}</${k}>`).join('')+'</通聯資料>'; }
test('Chinese network XML chooses base time and preserves both clocks',()=>{
 const xml='<查詢單><通聯記錄查詢條件><電話號碼>0911111111</電話號碼></通聯記錄查詢條件><通聯資料組>'+xmlRecord(row)+'</通聯資料組></查詢單>';
 const ws=app.parseImportFile('synthetic.xml',Buffer.from(xml));
 assert.equal(ws.case.source_format,'taiwan_mobile_network_xml'); assert.equal(ws.records.length,1);
 const r=ws.records[0]; assert.equal(r.record_kind,'data');assert.equal(r.location_time_source,'base');assert.equal(r.base_connected_at,'2026-10-01T10:00:00');assert.equal(r.internet_connected_at,'2026-10-01T10:01:00');assert.equal(r.base_duration_seconds,60);assert.equal(r.internet_duration_seconds,30);
 const restored=app.normalizeWorkspace(JSON.parse(JSON.stringify(ws)));
 for(const field of ['location_time_source','base_connected_at','internet_connected_at','base_duration_seconds','internet_duration_seconds']) assert.equal(restored.records[0][field],r[field]);
 assert.deepEqual(restored.subjects,ws.subjects);assert.equal(ws.subjects[0].phone,'0911111111');
});
test('network workbook supports context, repeated headers, multiple sheets and invalid time retention',()=>{
 const invalid=[...row]; invalid[3]='';invalid[5]='invalid-base';invalid[7]='2026/10/01 11:00:00';
 const bad=[...row];bad[5]='bad';bad[7]='also bad';
 const bytes=workbook({one:[['電話號碼：0922222222'],columns,invalid,columns,bad],two:[columns,row]});
 const ws=app.parseImportFile('synthetic.xlsx',bytes);assert.equal(ws.records.length,3);assert.equal(ws.records[0].target_phone,'0922222222');assert.equal(ws.records[0].location_time_source,'internet');assert.equal(ws.records[1].occurred_at,'bad');assert.equal(ws.records[1].base_connected_at,'bad');assert.equal(new Set(ws.subjects.map(user=>user.phone)).size,2);
});
test('expanded workbook restores source record identity and ignores analytical repetition',()=>{
 const headers=['來源 XML','原始紀錄序號',...columns,'匹配時間'];
 const values=['original.xml',7,...row,'ignored'];
 const ws=app.parseImportFile('synthetic.xlsx',workbook({detail:[headers,values,values]}));
 assert.equal(ws.records.length,1);assert.equal(ws.records[0].source_file,'original.xml');assert.equal(ws.records[0].row_number,7);assert.equal(ws.records[0].occurred_at,'2026-10-01T10:00:00');
});
test('CSP preserves customer PHONE association and record overrides including invalid records',()=>{
 const xml='<RESPONSE><CUSTOMERINFO><PHONE>0933333333</PHONE><NAME>合成姓名</NAME></CUSTOMERINFO><CELLINFO><STARTDT>invalid</STARTDT><CELLID>A</CELLID><CELLADDRESS>臺北市中正區合成路</CELLADDRESS></CELLINFO><CELLINFO><MSISDN>0944444444</MSISDN><STARTDT>20261001100000</STARTDT></CELLINFO></RESPONSE>';
 const ws=app.parseImportFile('synthetic.xml',Buffer.from(xml));assert.equal(ws.records.length,2);assert.equal(ws.records[0].target_phone,'0933333333');assert.equal(ws.records[1].target_phone,'0944444444');assert.equal(ws.subjects[0].phone,'0933333333');
 assert.deepEqual(ws.subjects.find(user=>user.phone==='0944444444').subject,{});
});
test('network workbook retains empty query subject before a later query',()=>{
 const values=[...row];values[3]='';
 const ws=app.parseImportFile('synthetic.xlsx',workbook({raw:[['電話號碼：0922222222'],['用戶名稱：空報表用戶'],columns,['電話號碼：0911111111'],columns,values]}));
 assert.equal(ws.records.length,1);assert.equal(ws.records[0].target_phone,'0911111111');
 assert.equal(ws.subjects.find(user=>user.phone==='0922222222').subject['用戶名稱'],'空報表用戶');
 assert.equal(ws.subjects.find(user=>user.phone==='0911111111').subject['用戶名稱'],undefined);
});
test('summary-only workbook is explicitly unsupported and empty XML retains subject',()=>{
 assert.throws(()=>app.parseImportFile('synthetic.xlsx',workbook({summary:[['門號','次數','時間'],['0911111111',2,'合成']]})),error=>/摘要/.test(error.message)&&error.code==='UNSUPPORTED_SUMMARY_WORKBOOK');
 const ws=app.parseImportFile('empty.xml',Buffer.from('<RESPONSE><CUSTOMERINFO><MSISDN>0911111111</MSISDN><NAME>合成</NAME></CUSTOMERINFO></RESPONSE>'));assert.equal(ws.subjects[0].phone,'0911111111');assert.equal(app.mergeWorkspaces([ws,ws]).subjects.length,1);
});
test('CSP query records keep distinct customers and empty report subjects',()=>{
 const xml='<ORDER><RECORD><RESPONSE><CUSTOMERINFO><PHONE>0911111111</PHONE><NAME>甲</NAME></CUSTOMERINFO><CELLINFO><STARTDT>2026/10/01 10:00:00</STARTDT></CELLINFO></RESPONSE></RECORD><RECORD><RESPONSE><CUSTOMERINFO><PHONE>0922222222</PHONE><NAME>乙</NAME></CUSTOMERINFO><CELLINFO><STARTDT>2026/10/01 11:00:00</STARTDT></CELLINFO></RESPONSE></RECORD><RECORD><RESPONSE><CUSTOMERINFO><PHONE>0933333333</PHONE></CUSTOMERINFO></RESPONSE></RECORD></ORDER>';
 const ws=app.parseImportFile('synthetic.xml',Buffer.from(xml));assert.deepEqual(ws.records.map(r=>r.target_phone),['0911111111','0922222222']);assert.equal(ws.subjects.length,3);
});
test('raw CSP workbook restores fields and rejects impossible base date in favor of internet time',()=>{
 const ws=app.parseImportFile('synthetic.xlsx',workbook({raw:[['STARTDT','DURATION','CELLADDRESS','CELLID','MSISDN','連到internet的時間'],['2026/02/31 10:00:00',20,'臺北市中正區合成路','A','0911111111','2026/03/01 10:00:00']]}));
 assert.equal(ws.records[0].location_time_source,'internet');assert.equal(ws.records[0].base_connected_at,'2026/02/31 10:00:00');
});
test('workspace normalization preserves embedded station references',()=>{
 const ws=app.normalizeWorkspace({case:{source_file:'synthetic.json',subject:{}},records:[{target_phone:'0911111111',stations:[{cell_id:'A',address:'臺北市中正區合成路'}]}]});assert.equal(ws.records[0].stations[0].cell_id,'A');
});
test('legacy merged subject phones preserve ambiguity without fabricating a phone',()=>{
 const subject={'申請號碼':'0911111111、0922222222','用戶名稱':'甲、乙'};
 for(const records of [[],[{target_phone:'0911111111'},{target_phone:'0922222222'}]]) {
  const ws=app.normalizeWorkspace({case:{source_file:'synthetic.json',subject},records});
  assert.deepEqual(ws.subjects.map(user=>user.phone).sort(),['','0911111111','0922222222']);
  assert.deepEqual(ws.subjects.find(user=>user.phone==='').subject,subject);
  assert.deepEqual(ws.subjects.filter(user=>user.phone).map(user=>user.subject),[{},{}]);
 }
});
test('legacy subject phone aliases retain individually evidenced empty report identity',()=>{
 for(const alias of ['電話號碼','申請號碼','設備號碼','調閱門號','調閱號碼','查詢項目']) {
  const subject={[alias]:'0911111111','用戶名稱':'甲'};
  const ws=app.normalizeWorkspace({case:{source_file:'synthetic.json',subject},records:[]});
  assert.deepEqual(ws.subjects.map(user=>user.phone),['0911111111']);assert.deepEqual(ws.subjects[0].subject,subject);
 }
});
