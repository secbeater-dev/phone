const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const {execFileSync}=require('node:child_process');
const {start}=require('./helpers');

async function samePixels(page,left,right) {
  if(left.equals(right)) return true;
  // Permit one colour level plus <=0.005% edge pixels for Chromium raster rounding.
  return page.evaluate(async pair=>{
    const images=await Promise.all(pair.map(data=>new Promise(resolve=>{const img=new Image();img.onload=()=>{const canvas=document.createElement('canvas');canvas.width=img.width;canvas.height=img.height;const ctx=canvas.getContext('2d');ctx.drawImage(img,0,0);resolve({width:img.width,height:img.height,data:ctx.getImageData(0,0,img.width,img.height).data});};img.src='data:image/png;base64,'+data;})));
    if(images[0].width!==images[1].width || images[0].height!==images[1].height) return false;
    let edgePixels=0;for(let i=0;i<images[0].data.length;i+=4) if([0,1,2,3].some(c=>Math.abs(images[0].data[i+c]-images[1].data[i+c])>1)) edgePixels++;
    return edgePixels<=Math.max(1,images[0].width*images[0].height*0.00005);
  },[left.toString('base64'),right.toString('base64')]);
}

async function originalPage(context,url) {
  const baseline=await context.newPage(),files=new Map();
  for(const file of ['index.html','app.js','styles.css','attachment-export.js','import-parser.js']) files.set(file,execFileSync('git',['show',`b922521:${file}`]));
  await baseline.route('**/*',route=>{
    const pathname=new URL(route.request().url()).pathname,file=pathname==='/'?'index.html':pathname.slice(1);
    return files.has(file) ? route.fulfill({body:files.get(file),contentType:file.endsWith('.html')?'text/html':file.endsWith('.css')?'text/css':'application/javascript'}) : route.continue();
  });
  await baseline.goto(url);await baseline.waitForLoadState('networkidle');await baseline.locator('#noticeDismissButton').click();
  return baseline;
}

test('empty workbench pixels match the pre-multi interface across themes and viewports', {skip:!process.env.PHONE_VISUAL_BASELINE,timeout:180000}, async t=>{
  const {page,context}=await start(t);
  const baseline=await context.newPage(), files=new Map();
  for(const file of ['index.html','app.js','styles.css','attachment-export.js','import-parser.js']) files.set(file,execFileSync('git',['show',`b922521:${file}`]));
  await baseline.route('**/*',route=>{
    const pathname=new URL(route.request().url()).pathname, file=pathname==='/'?'index.html':pathname.slice(1);
    if(!files.has(file)) return route.continue();
    return route.fulfill({body:files.get(file),contentType:file.endsWith('.html')?'text/html':file.endsWith('.css')?'text/css':'application/javascript'});
  });
  await baseline.goto(page.url()); await baseline.waitForLoadState('networkidle'); await baseline.locator('#noticeDismissButton').click();
  fs.mkdirSync('test-results',{recursive:true});
  for(const [width,height,theme,collapsed] of [[1440,1000,'light',false],[1440,1000,'dark',false],[1440,1000,'light',true],[430,932,'light',false],[430,932,'dark',true]]) {
    for(const p of [baseline,page]) {
      await p.setViewportSize({width,height});
      if(await p.locator('body').getAttribute('data-theme')!==theme) await p.locator('#themeToggleButton').click();
      if(await p.locator('#sidebarCollapseButton').getAttribute('aria-pressed')!==String(collapsed)) await p.locator('#sidebarCollapseButton').click();
    }
    for(const view of ['hours','calls','profile','stats','multiLocation','submission','export']) {
      for(const p of [baseline,page]) await p.locator(`button[data-view="${view}"]`).click();
      const name=`${width}-${theme}-${collapsed}-${view}`;
      const old=await baseline.screenshot({animations:'disabled',fullPage:true,path:`test-results/baseline-${name}.png`});
      const now=await page.screenshot({animations:'disabled',fullPage:true,path:`test-results/restored-${name}.png`});
      assert.ok(await samePixels(page,old,now),`Original empty appearance differs for ${name}`);
    }
  }
});

test('legacy populated analysis pixels match the pre-multi interface', {skip:!process.env.PHONE_VISUAL_BASELINE,timeout:180000},async t=>{
  const {page,context}=await start(t),baseline=await originalPage(context,page.url());
  const records=[['2026-01-01T10:00:00','outbound','發話',30],['2026-01-02T11:00:00','inbound','受話',45],['2026-01-03T12:00:00','data','上網',60]].map(([occurred_at,direction,call_type,duration_seconds],i)=>({occurred_at,direction,call_type,duration_seconds,target_phone:`090000000${i%2+1}`,counterparty_phone:'0900000003',imei:'123456789012345',base_refs:i<2?[{station_key:'SYN',role:'primary'}]:[]}));
  const input={name:'synthetic-original.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({case:{source_file:'synthetic-original.json',subject:{'用戶名稱':'合成使用者'}},records,base_stations:[{station_key:'SYN',address:'臺北市中正區合成地址'}]}))};
  await baseline.locator('#importWorkspaceInput').setInputFiles(input);
  await baseline.waitForFunction(()=>document.querySelector('#exportMessage').textContent.includes('已匯入資料 JSON'));
  await baseline.locator('[data-view="calls"]').click();
  await page.locator('#importWorkspaceInput').setInputFiles(input);
  await page.waitForFunction(()=>document.querySelector('#importProgressModal').hidden && document.querySelector('#callRows').children.length===3);
  // Compare analysis panels with equal sidebar height; import progress wording is separate.
  for(const p of [baseline,page]) await p.evaluate(()=>{document.querySelector('#importStatus').textContent='合成驗證';document.querySelector('#importResults').innerHTML='';});
  fs.mkdirSync('test-results',{recursive:true});
  for(const [width,height,theme,collapsed] of [[1440,1000,'light',false],[1440,1000,'dark',false],[1440,1000,'light',true],[430,932,'light',false],[430,932,'dark',true]]) {
    for(const p of [baseline,page]) {
      await p.setViewportSize({width,height});
      if(await p.locator('body').getAttribute('data-theme')!==theme) await p.locator('#themeToggleButton').click();
      if(await p.locator('#sidebarCollapseButton').getAttribute('aria-pressed')!==String(collapsed)) await p.locator('#sidebarCollapseButton').click();
    }
    for(const view of ['calls','profile','stats','hours']) {
      for(const p of [baseline,page]) await p.locator(`button[data-view="${view}"]`).click();
      await page.waitForFunction(view=>document.querySelector(`#${view}View`).getAttribute('aria-busy')==='false',view);
      const name=`legacy-${width}-${theme}-${collapsed}-${view}`;
      const old=await baseline.locator(`#${view}View`).screenshot({animations:'disabled',path:`test-results/baseline-${name}.png`});
      const now=await page.locator(`#${view}View`).screenshot({animations:'disabled',path:`test-results/restored-${name}.png`});
      assert.ok(await samePixels(page,old,now),`Original populated appearance differs for ${name}`);
      if(view==='hours') {
        for(const p of [baseline,page]) await p.locator('[data-hotspot-address]').first().click();
        await page.locator('#datasetHotspotRows').waitFor();
        const oldDetail=await baseline.locator('#hoursView').screenshot({animations:'disabled',path:`test-results/baseline-${name}-expanded.png`});
        const newDetail=await page.locator('#hoursView').screenshot({animations:'disabled',path:`test-results/restored-${name}-expanded.png`});
        assert.ok(await samePixels(page,oldDetail,newDetail),`Original expanded hotspot appearance differs for ${name}`);
        for(const p of [baseline,page]) await p.locator('[data-hotspot-address]').first().click();
      }
    }
  }
});
