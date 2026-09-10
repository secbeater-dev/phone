const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { chromium } = require('playwright');
const XLSX = require('../../vendor/xlsx.full.min.js');
async function start(t, options = {}) {
  const root = path.resolve(__dirname, '../..');
  const server = http.createServer((req, res) => {
    const pathname = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
    const file = path.resolve(root, '.' + (pathname === '/' ? '/index.html' : pathname));
    if (!file.startsWith(root + path.sep) || !fs.existsSync(file)) { res.writeHead(404); res.end(); return; }
    res.setHeader('Content-Type', ({'.js':'application/javascript','.html':'text/html','.css':'text/css','.png':'image/png'})[path.extname(file)] || 'application/octet-stream');
    fs.createReadStream(file).pipe(res);
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise(resolve => { server.closeAllConnections(); server.close(resolve); }));
  const browser = await chromium.launch({ headless: true, channel: process.env.PHONE_BROWSER_CHANNEL || 'msedge', args: options.args || [] });
  t.after(() => browser.close());
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
  const requests = [];
  context.on('request', request => requests.push({ method: request.method(), url: request.url() }));
  const page = await context.newPage();
  const errors = []; page.on('pageerror', error => errors.push(error.message));
  await page.goto(process.env.PHONE_TEST_URL || `http://127.0.0.1:${server.address().port}/`);
  await page.waitForLoadState('networkidle');
  await page.locator('#noticeDismissButton').click();
  return { page, context, browser, errors, requests };
}
function synthetic(count = 1101) {
  const calls = [['通話類別','始話時間','查詢項目','調閱門號','imei','對象門號','通話期間','開始基地台編號','開始基地台','結束基地台編號','結束基地台']];
  for(let n=0;n<count;n++) calls.push([n%2?'受話':'發話',`2026-01-${n<600?'01':'02'} 10:00:00`,'0900000001','0900000001','123456789012345','0900000003',30,'SYN-1','臺北市中正區合成地址','','']);
  calls.push(['發話','2026-01-03 12:00:00','0900000002','0900000002','123456789012346','0900000004',45,'SYN-2','新北市板橋區合成地址','','']);
  const users = [['用戶名稱','用戶回應_用戶編號_帳號','查詢項目','身份識別碼','戶籍地址','帳寄地址'],['合成人物甲','0900000001','0900000001','','',''],['合成人物乙','0900000002','0900000002','','','']];
  const data = [['查詢項目','通聯回應_用戶編號_帳號','imei','開始時間','連線期間','結束時間','開始基地台編號','開始基地台','結束基地台編號','結束基地台','上網ipv4'],['0900000001','0900000001','123456789012345','2026-01-01 11:00:00',600,'2026-01-01 11:10:00','SYN-1','臺北市中正區合成地址','','','192.0.2.1']];
  const wb=XLSX.utils.book_new();
  for(const [name,rows] of Object.entries({'通聯紀錄':calls,'使用者資料':users,'網路歷程':data,'通聯整合歷程紀錄':calls})) XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),name);
  return { name: 'synthetic-multi.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(XLSX.write(wb,{bookType:'xlsx',type:'buffer',compression:true,bookSST:true})) };
}
module.exports={start,synthetic};
