// Opt-in compatibility verification. Never emit private paths, names, values, counts, or screenshots.
const test = require('node:test');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const { execFile, execFileSync } = require('node:child_process');
const { promisify } = require('node:util');
const { start } = require('./helpers');
const PhoneWorkbench = require('../../app.js');

function originalParser() {
  const context={module:{exports:{}},require:name=>require(path.resolve(__dirname,'../..',name)),TextDecoder,TextEncoder,Uint8Array,ArrayBuffer,Buffer};
  require('node:vm').runInNewContext(execFileSync('git',['show','b922521:app.js'],{encoding:'utf8'}),context);
  return context.module.exports;
}

const execFileAsync = promisify(execFile);
const MIME = {
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xml': 'text/xml',
};

function privateFailure(stage, sequence) {
  const error = new Error(`Private compatibility verification failed at ${stage} for anonymous file ${sequence}.`);
  error.stack = error.message;
  return error;
}

function diagnostic(t, sequence, stage, passed = true) {
  console.log(`anonymous file ${sequence}: ${stage} ${passed ? 'pass' : 'fail'}`);
}

function isWithin(root, candidate) {
  const relative = path.relative(root, candidate);
  return relative === '' || (!relative.startsWith(`..${path.sep}`) && relative !== '..' && !path.isAbsolute(relative));
}

function collectSourceFiles(root) {
  const files = [];
  const visit = directory => {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const candidate = path.resolve(directory, entry.name);
      const resolved = fs.realpathSync(candidate);
      if (!isWithin(root, resolved)) throw privateFailure('source-boundary', 0);
      if (entry.isDirectory()) visit(resolved);
      else if (entry.isFile() && ['.xlsx', '.xml', '.zip'].includes(path.extname(entry.name).toLowerCase())) files.push(resolved);
    }
  };
  visit(root);
  return files.sort((a, b) => a.localeCompare(b));
}

async function extractXmlAnonymously(zipPath, tempRoot, archiveSequence) {
  const archiveRoot = path.join(tempRoot, `archive-${archiveSequence}-${crypto.randomUUID()}`);
  fs.mkdirSync(archiveRoot, { recursive: false });
  const script = [
    'Add-Type -AssemblyName System.IO.Compression.FileSystem',
    '$archive=[System.IO.Compression.ZipFile]::OpenRead($env:PHONE_COMPAT_ARCHIVE)',
    'try {',
    '  $index=0',
    '  foreach($entry in $archive.Entries) {',
    "    if([System.IO.Path]::GetExtension($entry.FullName) -ieq '.xml') {",
    '      $index++',
    "      $target=[System.IO.Path]::Combine($env:PHONE_COMPAT_OUTPUT, ('entry-{0:D6}.xml' -f $index))",
    '      $zipInput=$entry.Open()',
    '      try { $zipOutput=[System.IO.File]::Create($target); try { $zipInput.CopyTo($zipOutput) } finally { $zipOutput.Dispose() } } finally { $zipInput.Dispose() }',
    '    }',
    '  }',
    '} finally { $archive.Dispose() }',
  ].join('; ');
  try {
    await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(script,'utf16le').toString('base64')], { windowsHide: true, env:{...process.env,PHONE_COMPAT_ARCHIVE:zipPath,PHONE_COMPAT_OUTPUT:archiveRoot} });
  } catch (_) {
    throw privateFailure('archive-extraction', archiveSequence);
  }
  return fs.readdirSync(archiveRoot)
    .filter(name => /^entry-\d{6}\.xml$/i.test(name))
    .map(name => fs.realpathSync(path.join(archiveRoot, name)))
    .filter(candidate => {
      if (!isWithin(tempRoot, candidate)) throw privateFailure('archive-boundary', archiveSequence);
      return true;
    });
}

function anonymousFile(bytes, sequence, extension) {
  return { name: `anonymous-${String(sequence).padStart(6, '0')}${extension}`, mimeType: MIME[extension], buffer: bytes };
}

function displayedRecord(record) {
  return [
    record.source_file || '', record.occurred_at || '', record.call_type || '',
    PhoneWorkbench.normalizePhoneText(record.target_phone), '',
    PhoneWorkbench.normalizePhoneText(record.counterparty_phone), '',
    String(record.duration_seconds ?? ''), record.imei || '', [record.external_ip,record.note].filter(Boolean).join('；'),
  ];
}

function sortedRecords(records, column = 'occurred_at', direction = 'asc') {
  const collator = new Intl.Collator('zh-Hant', { numeric: true });
  return records.slice().sort((a, b) => {
    const av = column === 'duration_seconds' ? Number(a.duration_seconds || 0) : (a[column] || '');
    const bv = column === 'duration_seconds' ? Number(b.duration_seconds || 0) : (b[column] || '');
    const compared = typeof av === 'number' && typeof bv === 'number' ? av - bv : collator.compare(String(av || ''), String(bv || ''));
    return direction === 'asc' ? compared : -compared;
  });
}

async function waitForImport(page) {
  await page.locator('#importProgressModal').waitFor({ state: 'hidden', timeout: 300000 });
  const adopted = await page.evaluate(() => {
    const status = document.querySelector('#importStatus');
    return Boolean(status && status.textContent.includes('匯入完成'));
  });
  if (!adopted) throw privateFailure('browser-import', 0);
  await page.locator('[data-view="calls"]').click();
  await page.waitForFunction(() => document.querySelector('#callsView')?.getAttribute('aria-busy') === 'false', undefined, { timeout: 120000 });
}

async function visibleCallRows(page) {
  return page.locator('#callRows tr').evaluateAll(rows => rows.map(row => Array.from(row.cells, cell => {
    const input = cell.querySelector('input');
    return input ? input.value : cell.textContent;
  })));
}

async function verifyAllCallPages(page, expected, column = 'occurred_at', direction = 'asc') {
  await page.waitForFunction(()=>document.querySelector('#callsView').getAttribute('aria-busy')==='false');
  if (!expected.length) return await page.locator('#callRows td[colspan="10"]').count()===1;
  const remaining = new Map();
  for (const row of expected) { const key=JSON.stringify(displayedRecord(row)); remaining.set(key,(remaining.get(key)||0)+1); }
  let offset = 0, previous;
  while (true) {
    const actual = await visibleCallRows(page);
    for (const row of actual) {
      const key=JSON.stringify(row), count=remaining.get(key)||0;
      if (!count) return false;
      if(count===1) remaining.delete(key); else remaining.set(key,count-1);
      const value=column==='duration_seconds' ? Number(row[7]||0) : (/^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}/.test(row[1]) && Number.isFinite(Date.parse(row[1])) ? row[1].slice(0,19).replace(' ','T') : '');
      if(previous!==undefined && (direction==='asc' ? value<previous : value>previous)) return false;
      previous=value;
    }
    offset += actual.length;
    if (await page.locator('#callNextPage').isDisabled()) break;
    await page.locator('#callNextPage').click();
    await page.waitForFunction(()=>document.querySelector('#callsView').getAttribute('aria-busy')==='false');
  }
  return offset === expected.length && remaining.size === 0;
}

async function verifyProfileStatsHours(page, workspace, sequence = 0) {
  await page.locator('[data-view="profile"]').click();
  await page.waitForFunction(() => document.querySelector('#profileView')?.getAttribute('aria-busy') === 'false');
  const records=workspace.records,dates=records.map(record=>require('../../dataset-store.js').dateKey(record.occurred_at)).filter(Boolean).sort();
  const cards=[records.length,new Set(records.map(r=>r.target_phone).filter(Boolean)).size,new Set(records.map(r=>r.counterparty_phone).filter(Boolean)).size,dates[0]||'-',dates.at(-1)||'-',records.reduce((sum,r)=>sum+Number(r.duration_seconds||0),0)].map(String);
  if(JSON.stringify(await page.locator('#profileSummaryCards strong').allTextContents())!==JSON.stringify(cards)) throw privateFailure('profile-cards',sequence);
  const subjects=new Map(),chips=new Set();
  while(true) {
    const values=await page.evaluate(()=>({entries:Array.from(document.querySelectorAll('#profileContent dt')).map((dt,index)=>[dt.textContent,document.querySelectorAll('#profileContent dd')[index]?.textContent || '']),chips:Array.from(document.querySelectorAll('.imei-chip'),node=>node.textContent)}));
    for(const [key,value] of values.entries) { if(key==='狀態' && value==='尚無用戶資料。') continue; const set=subjects.get(key)||new Set(); value.split('、').map(v=>v.trim()).filter(Boolean).forEach(v=>set.add(v));subjects.set(key,set); }
    values.chips.forEach(value=>chips.add(value));
    const next=page.locator('#dataset-profile-pages [data-dataset-page="next"]');
    if(!await next.count() || await next.isDisabled()) break;
    await next.click();await page.waitForFunction(()=>document.querySelector('#profileView').getAttribute('aria-busy')==='false');
  }
  const canonical=entries=>JSON.stringify(entries.map(([k,values])=>[k,[...values].sort()]).sort(([a],[b])=>a.localeCompare(b)));
  const expectedSubjects=Object.entries(workspace.case.subject || {}).map(([key,value])=>[key,new Set(String(value).split('、').map(v=>v.trim()).filter(Boolean))]).filter(([,values])=>values.size);
  if(canonical([...subjects])!==canonical(expectedSubjects) || JSON.stringify([...chips].sort())!==JSON.stringify(PhoneWorkbench.collectUniqueImeis(workspace.records).sort()) || await page.locator('#profileSummaryCards .metric-card').count()!==6) throw privateFailure('profile',sequence);

  await page.locator('[data-view="stats"]').click();
  await page.waitForFunction(() => document.querySelector('#statsView')?.getAttribute('aria-busy') === 'false');
  for (const mode of ['count', 'seconds']) {
    await page.locator(`[data-stats-rank-mode="${mode}"]`).click();
    await page.waitForFunction(()=>document.querySelector('#statsView').getAttribute('aria-busy')==='false');
    const expected = PhoneWorkbench.computePhoneStats(workspace.records, mode);
    const actual=[[],[],[]];
    while(true) {
      const rows=await page.evaluate(()=>Array.from(document.querySelectorAll('#statsContent .stats-card'),card=>Array.from(card.querySelectorAll('.stats-table-row:not(.stats-table-head)'),row=>Array.from(row.children,cell=>cell.querySelector('input')?.value ?? cell.textContent))));
      rows.forEach((group,index)=>actual[index].push(...group));
      const next=page.locator('#dataset-stats-pages [data-dataset-page="next"]');
      if(!await next.count() || await next.isDisabled()) break;
      await next.click();await page.waitForFunction(()=>document.querySelector('#statsView').getAttribute('aria-busy')==='false');
    }
    const wanted=[expected.inboundRows,expected.outboundRows,expected.totalRows].map(rows=>rows.map((row,index)=>[String(index+1),row.phone,'',String(row.count),String(row.seconds)]));
    if(JSON.stringify(actual)!==JSON.stringify(wanted)) throw privateFailure('rankings',sequence);
  }

  await page.locator('[data-view="hours"]').click();
  await page.waitForFunction(() => document.querySelector('#hoursView')?.getAttribute('aria-busy') === 'false');
  const buckets = PhoneWorkbench.computeHourBuckets(workspace.records);
  const hourOk = await page.evaluate(wanted => {
    const actual = Array.from(document.querySelectorAll('#hoursView .hour-column'), column => column.querySelector('.hour-bar-count')?.textContent || '');
    return JSON.stringify(actual) === JSON.stringify(wanted.map(item => String(item.count))) && Boolean(document.querySelector('#hourHotspotContent'));
  }, buckets);
  if(!hourOk) throw privateFailure('hour-buckets',sequence);
  const hotspots=PhoneWorkbench.computeAddressHotspots(workspace.records.filter(record=>require('../../dataset-store.js').dateKey(record.occurred_at)),workspace.base_stations);
  const remaining=new Map(hotspots.map(row=>[row.key,row]));
  let expanded=false;
  while(true) {
    const rows=await page.locator('[data-hotspot-address]').evaluateAll(nodes=>nodes.map(node=>({key:node.dataset.hotspotAddress,address:node.querySelector('strong').textContent,count:Number(node.querySelector('span').textContent.split(' ')[0].replace(/,/g,''))})));
    for(const row of rows) { const wanted=remaining.get(row.key);if(!wanted || wanted.address!==row.address || wanted.count!==row.count) throw privateFailure('hotspots',sequence);remaining.delete(row.key); }
    if(!expanded && rows.length) {
      expanded=true;
      const expected=hotspots.find(row=>row.key===rows[0].key).times.slice().sort(),times=[];
      await page.locator('[data-hotspot-address]').first().click();
      await page.locator('#datasetHotspotRows').waitFor();
      while(true) {
        times.push(...await page.locator('#datasetHotspotRows li').allTextContents());
        const next=page.locator('[data-hotspot-page]').last();
        if(!await next.count() || await next.isDisabled()) break;
        const before=await page.locator('#datasetHotspotRows').elementHandle();
        await next.click();
        await page.waitForFunction(previous=>document.querySelector('#datasetHotspotRows')!==previous,before);
        await before.dispose();
      }
      if(JSON.stringify(times.sort())!==JSON.stringify(expected)) throw privateFailure('hotspot-times',sequence);
    }
    const next=page.locator('#dataset-hours-pages [data-dataset-page="next"]');
    if(!await next.count() || await next.isDisabled()) break;
    await next.click();await page.waitForFunction(()=>document.querySelector('#hoursView').getAttribute('aria-busy')==='false');
  }
  if(remaining.size) throw privateFailure('hotspot-pages',sequence);
  return true;
}

async function verifyInteractions(page, workspace) {
  const records = workspace.records;
  const searchable = records.find(record => record.counterparty_phone || record.imei || record.occurred_at);
  if (searchable) {
    const query = String(searchable.counterparty_phone || searchable.imei || searchable.occurred_at);
    await page.locator('[data-view="calls"]').click();
    await page.locator('#recordSearch').fill(query);
    const expected = sortedRecords(records.filter(record => [record.occurred_at, record.source_file, record.call_type, record.target_phone, record.counterparty_phone, record.imei, record.external_ip, record.note].some(value => String(value || '').toLowerCase().includes(query.toLowerCase()))));
    if (!await verifyAllCallPages(page, expected)) return false;
    await page.locator('#recordSearch').fill('');
  }

  await page.locator('#callsView [data-call-sort="duration_seconds"]').click();
  if (!await verifyAllCallPages(page, records, 'duration_seconds', 'asc')) return false;
  await page.locator('#callsView [data-call-sort="duration_seconds"]').click();
  if (!await verifyAllCallPages(page, records, 'duration_seconds', 'desc')) return false;

  const dated = records.find(record => /^\d{4}-\d{2}-\d{2}/.test(record.occurred_at || ''));
  if (dated) {
    const day = dated.occurred_at.slice(0, 10);
    await page.locator('#dateFilterButton').click();
    await page.locator('#dateFilterStartInput').fill(day);
    await page.locator('#dateFilterEndInput').fill(day);
    await page.locator('#dateFilterApplyButton').click();
    const filtered = PhoneWorkbench.filterRecordsByDateRange(records, { start: day, end: day, active:true });
    if (!await verifyAllCallPages(page, filtered, 'duration_seconds', 'desc')) return false;
    await page.locator('#dateFilterButton').click();
    await page.locator('#dateFilterResetButton').click();
  }
  return true;
}

test('private legacy compatibility corpus', { skip: !process.env.PRIVATE_COMPAT_DIR, timeout: 1800000 }, async t => {
  if (process.env.PHONE_TEST_URL) throw privateFailure('local-origin', 0);
  let tempRoot;
  try {
    const requestedRoot = path.resolve(process.env.PRIVATE_COMPAT_DIR);
    const sourceRoot = fs.realpathSync(requestedRoot);
    if (!fs.statSync(sourceRoot).isDirectory()) throw privateFailure('source-directory', 0);
    tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'phone-private-compat-'));
    const sourceFiles = collectSourceFiles(sourceRoot);
    const inputs = [];
    let sequence = 0;
    let archiveSequence = 0;
    for (const source of sourceFiles) {
      const extension = path.extname(source).toLowerCase();
      if (extension === '.zip') {
        archiveSequence++;
        for (const extracted of await extractXmlAnonymously(source, tempRoot, archiveSequence)) {
          sequence++;
          inputs.push({ sequence, extension: '.xml', bytes: fs.readFileSync(extracted) });
        }
      } else {
        sequence++;
        inputs.push({ sequence, extension, bytes: fs.readFileSync(source) });
      }
    }
    if (!inputs.length) throw privateFailure('discovery', 0);

    const parsed = [], baseline = originalParser();
    for (const input of inputs) {
      try {
        const file = anonymousFile(input.bytes, input.sequence, input.extension);
        const workspace = PhoneWorkbench.parseImportFile(file.name, input.bytes);
        if(workspace.case.source_format!=='taiwan_mobile_flat_call_xlsx') {
          const previous=baseline.parseImportFile(file.name,input.bytes);
          const projection=value=>JSON.stringify([value.records.map(({external_ipv4,external_ipv6,record_kind,source_query,source_target,...row})=>row),value.base_stations,value.case.subject]);
          if(projection(previous)!==projection(workspace)) throw privateFailure('parser-baseline',input.sequence);
        } else {
          const XLSX=require('../../vendor/xlsx.full.min.js'),wb=XLSX.read(input.bytes,{type:'buffer'});
          let sourceCalls=0;
          for(const name of wb.SheetNames) {
            const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:true}),header=rows.findIndex(row=>row.includes('始話日期時間') && row.includes('順序'));
            if(header<0) continue;
            const h=rows[header],seq=h.indexOf('順序'),time=h.indexOf('始話日期時間');
            sourceCalls+=rows.slice(header+1).filter(row=>row[time] && row[time]!=='始話日期時間' && Number(row[seq])<=1).length;
          }
          if(sourceCalls!==workspace.records.length) throw privateFailure('source-reconciliation',input.sequence);
        }
        parsed.push({ input, file, workspace });
        diagnostic(t, input.sequence, 'parser');
      } catch (error) {
        if(error?.message?.startsWith('Private compatibility verification failed')) throw error;
        diagnostic(t, input.sequence, 'parser', false);
        throw privateFailure('parser', input.sequence);
      }
    }

    const { page, context, errors, requests } = await start(t);
    for (const item of parsed) {
      await page.locator('#fileInput').setInputFiles(item.file);
      await waitForImport(page);
      if (await page.locator('#datasetToolbar').isVisible() || await page.locator('[data-view="network"]').isVisible()) throw privateFailure('legacy-controls', item.input.sequence);
      if (!await verifyAllCallPages(page, sortedRecords(item.workspace.records))) throw privateFailure('rows', item.input.sequence);
      diagnostic(t, item.input.sequence, 'all-pages');
      if (!await verifyProfileStatsHours(page, item.workspace, item.input.sequence)) throw privateFailure('summaries', item.input.sequence);
      diagnostic(t, item.input.sequence, 'summaries');
      if (!await verifyInteractions(page, item.workspace)) throw privateFailure('interactions', item.input.sequence);
      diagnostic(t, item.input.sequence, 'query-controls');
    }

    const merged = PhoneWorkbench.mergeWorkspaces(parsed.map(item => item.workspace));
    await page.locator('#fileInput').setInputFiles(parsed.map(item => item.file));
    await waitForImport(page);
    if (!await verifyAllCallPages(page, sortedRecords(merged.records))) throw privateFailure('batch-merge', 0);
    if (!await verifyProfileStatsHours(page, merged)) throw privateFailure('batch-summaries', 0);
    if (!await verifyInteractions(page, merged)) throw privateFailure('batch-interactions', 0);
    diagnostic(t, 0, 'batch-merge');

    const origin = new URL(page.url()).origin;
    const networkSafe = requests.every(request => {
      if (!/^https?:/.test(request.url)) return true;
      return (request.method === 'GET' && new URL(request.url).origin === origin) || (!request.status && /^(csp|net::ERR_BLOCKED_BY_CSP)$/.test(request.failed || ''));
    });
    if (!networkSafe || errors.length) throw privateFailure('browser-audit', 0);
    diagnostic(t, 0, 'browser-audit');
    await context.close();
  } catch (error) {
    if (error?.message?.startsWith('Private compatibility verification failed')) throw error;
    throw privateFailure('harness', 0);
  } finally {
    if (tempRoot) {
      const resolvedTemp = path.resolve(tempRoot);
      const osTemp = fs.realpathSync(os.tmpdir());
      if (isWithin(osTemp, resolvedTemp) && path.basename(resolvedTemp).startsWith('phone-private-compat-')) fs.rmSync(resolvedTemp, { recursive: true, force: true });
    }
  }
});
