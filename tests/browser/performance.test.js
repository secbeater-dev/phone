const test=require('node:test');
const assert=require('node:assert/strict');
const {start,synthetic,waitForImport}=require('./helpers');
test('synthetic large and double-size imports keep bounded UI and cancellable exports', {skip:!process.env.PHONE_PERF,timeout:900000}, async t=>{
  const {page}=await start(t,{args:['--js-flags=--max-old-space-size=512']});
  let file;
  const fs=require('node:fs'),os=require('node:os'),path=require('node:path');
  const directory=fs.mkdtempSync(path.join(os.tmpdir(),'phone-synthetic-perf-'));
  const filePath=path.join(directory,'synthetic-multi.xlsx');
  t.after(()=>{if(fs.existsSync(filePath))fs.unlinkSync(filePath);fs.rmdirSync(directory);});
  for(const count of [100000,200000]){
    file=synthetic(count);
    fs.writeFileSync(filePath,file.buffer);file=null;
    await page.evaluate(()=>{window.__lag=0;let last=performance.now();window.__timer=setInterval(()=>{const now=performance.now();window.__lag=Math.max(window.__lag,now-last-50);last=now;},50);});
    const started=Date.now();
    await page.locator('#fileInput').setInputFiles(filePath);
    await waitForImport(page,600000);
    assert.equal(await page.locator('.hour-bar-count').nth(10).innerText(),String(count));
    await page.locator('[data-view="calls"]').click();
    await page.waitForFunction(n=>document.querySelector('#importProgressModal')?.hidden && document.querySelector('#callPageSummary')?.textContent.includes(n.toLocaleString()),count,{timeout:600000});
    assert.equal(await page.locator('#callRows tr').count(),500);
    const lag=await page.evaluate(()=>{clearInterval(window.__timer);return Math.round(window.__lag);});
    t.diagnostic(JSON.stringify({syntheticCallRows:count,importSeconds:Math.round((Date.now()-started)/1000),maxUiTimerLagMs:lag}));
  }
  await page.locator('#attachmentExportButton').click();
  const downloaded=page.waitForEvent('download',{timeout:180000}); await page.locator('#attachmentXlsxButton').click(); await downloaded;
  const cancelStarted=Date.now();await page.locator('#datasetExportCancel').click();
  await page.waitForFunction(()=>document.querySelector('#datasetExportCancel')?.hidden);
  t.diagnostic(JSON.stringify({exportCancelMs:Date.now()-cancelStarted}));
  await page.locator('#attachmentExportCloseButton').click();
  await page.locator('button[data-view="network"]').click();
  await page.waitForFunction(()=>document.querySelector('#datasetRows')?.textContent.includes('192.0.2.1'));
});
