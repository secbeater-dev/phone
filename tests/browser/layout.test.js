const test = require('node:test');
const assert = require('node:assert/strict');
const {start,synthetic}=require('./helpers');
test('synthetic desktop and mobile layout keeps controls reachable', {timeout:60000}, async t=>{
  const {page,errors}=await start(t);
  await page.locator('#fileInput').setInputFiles(synthetic(30));
  await page.waitForFunction(()=>document.querySelector('#importProgressModal').hidden && document.querySelector('#importStatus').textContent.startsWith('匯入完成') && document.querySelector('#callsView').getAttribute('aria-busy')==='false');
  if(process.env.PHONE_VISUAL){require('node:fs').mkdirSync('test-results',{recursive:true});await page.screenshot({path:'test-results/synthetic-desktop.png',fullPage:true});}
  await page.locator('button[data-view="hours"]').click();
  await page.waitForFunction(()=>document.querySelector('#hoursView').getAttribute('aria-busy')==='false');
  await page.locator('#hourClearAll').click();
  await page.locator('[data-hour="10"]').click();
  await page.locator('[data-hour="11"]').click();
  assert.equal(await page.locator('.hour-tile.active').count(),2);
  await page.locator('#hourApplyButton').click();
  await page.waitForFunction(()=>document.querySelector('#hoursView').getAttribute('aria-busy')==='false');
  await page.locator('[data-hotspot-address]').first().click();
  await page.waitForFunction(()=>document.querySelector('#datasetHotspotRows li'));
  await page.setViewportSize({width:430,height:932});
  assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth+1));
  if(process.env.PHONE_VISUAL)await page.screenshot({path:'test-results/synthetic-mobile.png',fullPage:true});
  assert.deepEqual(errors,[]);
});
