'use strict';
const { crc32, deflateRawSync } = require('node:zlib');

function zipped(parts) {
  const local = [], central = []; let offset = 0;
  for (const [name, content] of Object.entries(parts)) {
    const filename = Buffer.from(name), raw = Buffer.from(content), body = deflateRawSync(raw);
    const crc = crc32(raw), header = Buffer.alloc(30), entry = Buffer.alloc(46);
    header.writeUInt32LE(0x04034b50); header.writeUInt16LE(20, 4); header.writeUInt16LE(8, 8);
    header.writeUInt32LE(crc, 14); header.writeUInt32LE(body.length, 18); header.writeUInt32LE(raw.length, 22); header.writeUInt16LE(filename.length, 26);
    entry.writeUInt32LE(0x02014b50); entry.writeUInt16LE(20, 4); entry.writeUInt16LE(20, 6); entry.writeUInt16LE(8, 10);
    entry.writeUInt32LE(crc, 16); entry.writeUInt32LE(body.length, 20); entry.writeUInt32LE(raw.length, 24); entry.writeUInt16LE(filename.length, 28); entry.writeUInt32LE(offset, 42);
    local.push(header, filename, body); central.push(entry, filename); offset += header.length + filename.length + body.length;
  }
  const directory = Buffer.concat(central), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50); end.writeUInt16LE(central.length / 2, 8); end.writeUInt16LE(central.length / 2, 10); end.writeUInt32LE(directory.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...local, directory, end]);
}

function legacyWorkbook(paddingBytes = 0) {
  const rows = [
    ['始話時間', '通話秒數', '調閱號碼', 'IMEI', '通話類別', '通話對象', '轉接電話', '基地台/交換機', '備註'],
    ['2026-09-01T01:02:03', '10', '0900000008', '', '發話', '0900000009', '', 'SYNTH-1', '']
  ];
  const sheet = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + rows.map((row, i) => `<row r="${i + 1}">${row.map((value, col) => `<c r="${String.fromCharCode(65 + col)}${i + 1}" t="inlineStr"><is><t>${value}</t></is></c>`).join('')}</row>`).join('') + ' '.repeat(paddingBytes) + '</sheetData></worksheet>';
  const parts = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Legacy" sheetId="1" r:id="r1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': sheet
  };
  return { buffer: zipped(parts), expandedBytes: Object.values(parts).reduce((sum, value) => sum + Buffer.byteLength(value), 0) };
}
module.exports = { legacyWorkbook };
