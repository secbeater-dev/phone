const test=require('node:test'),assert=require('node:assert/strict');
const XLSX=require('../vendor/xlsx.full.min.js'),app=require('../app.js');
test('flattened telecom XML XLSX preserves calls, station continuations, raw IMEI and metadata',()=>{
  const headers=['查詢日期','文號','電信業者','查詢狀態','備註','通聯類別','開始時間','結束時間','電話號碼','電話種類','用戶基本資料組','始話日期時間','通話時間','通話類別','目標電話','電話種類2','對象電話','電話種類3','電信業者4','IMEI','備註5','順序','基地台編號','基地台位置','轉接資訊組'];
  const first=['2026-01-01','SYNTHETIC','台灣大哥大','成功','','雙向','','','0900000001','','','2026-01-01T10:00:00',30,'發話','0900000001','','0900000002','','',123456789012345,'合成備註',1,'SYN-1','臺北市中正區合成地址',''];
  const next=[...first];next[21]=2;next[22]='SYN-2';next[23]='臺北市大同區合成地址';
  const invalid=[...first];invalid[11]='invalid-date';invalid[14]='0900000003';
  const wb=XLSX.utils.book_new(),sheet=XLSX.utils.aoa_to_sheet([headers,first,next,first,invalid]);sheet.T2.z='0.00E+00';
  XLSX.utils.book_append_sheet(wb,sheet,'合成資料');
  const result=app.parseImportFile('synthetic-flat.xlsx',XLSX.write(wb,{type:'buffer',bookType:'xlsx'}));
  assert.equal(result.case.source_format,'taiwan_mobile_flat_call_xlsx');
  assert.equal(result.records.length,3);
  assert.equal(result.records[0].base_refs.length,2);
  assert.equal(result.records[0].imei,'123456789012345');
  assert.equal(result.records[1].base_refs.length,1);
  assert.equal(result.records[2].occurred_at,'invalid-date');
  assert.equal(result.case.subject['文號'],'SYNTHETIC');
  assert.match(result.records[0].note,/合成備註/);
});
