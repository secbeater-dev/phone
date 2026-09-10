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
  assert.ok(network.every(r=>r.method==='GET' && new URL(r.url).origin===origin),'Only same-origin static GET requests are allowed');
  assert.deepEqual(errors,[]);
});
