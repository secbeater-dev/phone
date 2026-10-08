// Opt-in only. No source names, identifiers, counts, screenshots or output artifacts.
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const app = require('../../app');
const { start } = require('./helpers');

test('private network structures and entire XML folder import remain local', { skip: !process.env.PRIVATE_NETWORK_ROOT, timeout: 1800000 }, async t => {
  if (process.env.PHONE_TEST_URL) throw new Error('Private verification requires the local test origin.');
  const heartbeat = setInterval(() => t.diagnostic('Private verification running; source details suppressed.'), 50000);
  t.after(() => clearInterval(heartbeat));
  let stage = 'source-boundary';
  try {
    const base = fs.realpathSync(process.env.PRIVATE_NETWORK_ROOT), files = [];
    function visit(directory) {
      for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
        const file = fs.realpathSync(path.join(directory, entry.name));
        const relative = path.relative(base, file);
        if (relative.startsWith('..') || path.isAbsolute(relative)) throw new Error();
        if (entry.isDirectory()) visit(file);
        else if (/\.(xml|xlsx)$/i.test(entry.name)) files.push(file);
      }
    }
    visit(base);
    if (!files.length) throw new Error();
    stage = 'per-file-parsing';
    for (const file of files) {
      try {
        const workspace = app.parseImportFile('anonymous' + path.extname(file), fs.readFileSync(file));
        if (!Array.isArray(workspace.records) || !Array.isArray(workspace.subjects)) throw new Error();
      } catch (error) { if (!/摘要/.test(error.message)) throw new Error(); }
      await new Promise(resolve => setImmediate(resolve));
    }
    t.diagnostic('Private per-file parsing: PASS');
    stage = 'browser-start';
    const { page, errors, requests } = await start(t);
    const origin = new URL(page.url()).origin;
    await page.locator('[data-view="multiLocation"]').click();
    stage = 'folder-select';
    await page.locator('#multiLocationFolderInput').setInputFiles(process.env.PRIVATE_NETWORK_XML_FOLDER);
    stage = 'bulk-import';
    await page.waitForFunction(() => document.querySelector('#multiLocationImportStatus')?.textContent.startsWith('匯入完成'), undefined, { timeout: 1500000 });
    stage = 'per-file-results';
    if (await page.evaluate(() => Boolean(document.querySelector('#multiLocationImportResults .danger-text')))) throw new Error();
    stage = 'cards-and-results';
    await page.waitForFunction(() => document.querySelectorAll('#multiLocationCards .phone-card').length > 0);
    if (!await page.evaluate(() => document.querySelectorAll('.multi-location-match-row').length <= 500)) throw new Error();
    stage = 'local-only-requests';
    if (errors.length || requests.some(row => row.method !== 'GET' || new URL(row.url).origin !== origin)) throw new Error();
    t.diagnostic('Private bulk folder import, cards and local-only requests: PASS');
  } catch (_) {
    const error = new Error(`Private network verification ${stage}: FAIL (source details suppressed).`);
    error.stack = error.message;
    throw error;
  }
});
