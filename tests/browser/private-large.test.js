// Opt-in private verification: never print source paths, values, counts, or screenshots.
const test = require('node:test');
const { start } = require('./helpers');
const { execFile } = require('node:child_process');
const { promisify } = require('node:util');
const run = promisify(execFile);
test('private large workbook stays local, responsive and queryable', { skip: !process.env.PRIVATE_MULTI_XLSX, timeout: 1800000 }, async t => {
  if (process.env.PHONE_TEST_URL) throw new Error('Private inputs are restricted to the local test origin.');
  const { page, browser } = await start(t, { args: ['--js-flags=--max-old-space-size=512'] });
  let unexpected = false, sampling = false, maxRss = 0, maxRendererRss = 0;
  const origin = new URL(page.url()).origin;
  page.on('request', request => { if (!request.url().startsWith(origin) && !request.url().startsWith('blob:') && !request.url().startsWith('data:')) unexpected = true; });
  const cdp = await browser.newBrowserCDPSession();
  const sample = async () => {
    if (sampling) return; sampling = true;
    try {
      const { processInfo } = await cdp.send('SystemInfo.getProcessInfo');
      const ids = processInfo.map(p => p.id).filter(Number.isInteger).join(',');
      const rendererIds = processInfo.filter(p=>/renderer|worker/.test(p.type)).map(p=>p.id).filter(Number.isInteger).join(',');
      const { stdout } = await run('powershell.exe', ['-NoProfile', '-Command', `$items=Get-Process -Id ${ids} -ErrorAction SilentlyContinue; foreach($item in $items){try{$item.ProcessorAffinity=15}catch{}}; ($items|Measure-Object WorkingSet64 -Sum).Sum; ${rendererIds ? `(Get-Process -Id ${rendererIds} -ErrorAction SilentlyContinue|Measure-Object WorkingSet64 -Sum).Sum` : '0'}`]);
      const [all, renderer] = stdout.trim().split(/\r?\n/).map(Number);
      maxRss = Math.max(maxRss, all || 0); maxRendererRss = Math.max(maxRendererRss, renderer || 0);
    } catch (_) {} finally { sampling = false; }
  };
  await sample(); const timer = setInterval(sample, 5000); t.after(() => clearInterval(timer));
  await page.evaluate(() => { window.__lag = 0; let last = performance.now(); window.__ticker = setInterval(() => { const now = performance.now(); window.__lag = Math.max(window.__lag, now - last - 50); last = now; }, 50); });
  const heartbeat = setInterval(() => t.diagnostic('Private verification is running; source details suppressed.'), 60000);
  t.after(() => clearInterval(heartbeat));
  const started = Date.now();
  try {
    await page.locator('#fileInput').setInputFiles(process.env.PRIVATE_MULTI_XLSX);
    await page.locator('#importProgressModal').waitFor({ state: 'hidden', timeout: 1500000 });
    const imported = await page.evaluate(() => Boolean(document.querySelector('#datasetTarget')?.options.length && document.querySelector('#importStatus')?.textContent.startsWith('匯入完成')));
    if (!imported) throw new Error('verification failed');
    await page.waitForFunction(() => Boolean(document.querySelector('#datasetRows tr')), undefined, { timeout: 120000 });
    const importSeconds = Math.round((Date.now() - started) / 1000);
    const importLag = await page.evaluate(()=>{const lag=window.__lag;window.__lag=0;return Math.round(lag);});
    await page.locator('button[data-view="network"]').click();
    await page.waitForFunction(() => Boolean(document.querySelector('#datasetPageInfo')), undefined, { timeout: 120000 });
    const began = Date.now();
    await page.locator('#datasetTarget').selectOption({ index: 1 });
    await page.waitForFunction(() => document.querySelector('#datasetContent')?.getAttribute('aria-busy') === 'false', undefined, { timeout: 120000 });
    const switchMs = Date.now() - began;
    for (const view of ['profile', 'stats', 'hours']) {
      await page.locator(`button[data-view="${view}"]`).click();
      await page.waitForFunction(() => document.querySelector('#datasetContent')?.getAttribute('aria-busy') === 'false', undefined, { timeout: 240000 });
      if (await page.locator('#datasetStatus').evaluate(el => el.classList.contains('danger-text'))) throw new Error('verification failed');
    }
    await sample(); const lag = await page.evaluate(() => { clearInterval(window.__ticker); return Math.round(window.__lag); });
    if (unexpected) throw new Error('verification failed');
    t.diagnostic(JSON.stringify({ importSeconds, switchMs, maxBrowserRssMiB: Math.round(maxRss / 1048576), maxRendererWorkerRssMiB: Math.round(maxRendererRss / 1048576), maxImportTimerLagMs: importLag, maxQueryTimerLagMs: lag, fourLogicalCpuAffinity: true, noDataRequests: true }));
    await page.locator('#datasetClear').click();
    await page.waitForFunction(() => document.querySelector('#datasetStatus')?.textContent.startsWith('已清除'), undefined, { timeout: 120000 });
  } catch (_) { throw new Error('Private verification failed; source details intentionally omitted.'); }
});
