const test=require('node:test');
const assert=require('node:assert/strict');
const {start,synthetic}=require('./helpers');
test('page and worker requests stay on-origin and never submit imported data', {timeout:60000}, async t=>{
  const {page,requests,errors}=await start(t);
  await page.locator('#fileInput').setInputFiles(synthetic(3));
  await page.waitForFunction(()=>Boolean(document.querySelector('#datasetPageInfo')));
  await page.locator('#attachmentExportButton').click();
  const download=page.waitForEvent('download');await page.locator('#attachmentXlsxButton').click();await download;
  await page.waitForFunction(()=>!document.querySelector('#attachmentXlsxButton').disabled);
  const origin=new URL(page.url()).origin;
  assert.ok(requests.some(r=>r.url.includes('dataset-worker.js')));
  const network=requests.filter(r=>/^https?:/.test(r.url));
  assert.ok(network.every(r=>(r.method==='GET' && new URL(r.url).origin===origin) || (!r.status && /^(csp|net::ERR_BLOCKED_BY_CSP)$/.test(r.failed || ''))),'Only same-origin GET requests or requests prevented by CSP are allowed');
  assert.equal(await page.locator('iframe').count(),0,'Injected challenge scripts must not create frames');
  assert.deepEqual(errors,[]);
});
