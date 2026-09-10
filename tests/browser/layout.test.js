const test = require('node:test');
const assert = require('node:assert/strict');
const {start,synthetic}=require('./helpers');
test('synthetic desktop and mobile layout keeps controls reachable', {timeout:60000}, async t=>{
  const {page,errors}=await start(t);
  await page.locator('#fileInput').setInputFiles(synthetic(30));
  await page.waitForFunction(()=>Boolean(document.querySelector('#datasetPageInfo')));
  if(process.env.PHONE_VISUAL){require('node:fs').mkdirSync('test-results',{recursive:true});await page.screenshot({path:'test-results/synthetic-desktop.png',fullPage:true});}
  await page.locator('button[data-view="hours"]').click();
  await page.waitForFunction(()=>Boolean(document.querySelector('.dataset-hour')));
  await page.locator('[data-dataset-hour="10"]').click();
  await page.waitForFunction(()=>document.querySelector('[data-dataset-hour="10"]')?.getAttribute('aria-pressed')==='true');
  await page.locator('[data-dataset-hour="11"]').click();
  await page.waitForFunction(()=>document.querySelectorAll('.dataset-hour[aria-pressed="true"]').length===2);
  await page.locator('[data-hotspot]').first().click();
  await page.waitForFunction(()=>document.querySelector('#datasetHotspotRows tr'));
  await page.setViewportSize({width:430,height:932});
  assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth+1));
  if(process.env.PHONE_VISUAL)await page.screenshot({path:'test-results/synthetic-mobile.png',fullPage:true});
  assert.deepEqual(errors,[]);
});
