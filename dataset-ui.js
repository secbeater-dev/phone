(function (root) {
  'use strict';
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const number = value => Number(value || 0).toLocaleString();
  const state = { enabled: false, dataset: null, client: null, starting: null, target: '', view: 'calls', page: 1, targetPage: 1, search: '', sort: { column: 'occurred_at', direction: 'asc' }, date: { active: false }, hours: null, counties: null, mode: 'count', revision: 0, busy: false, exporting: false };
  let bridge;
  function init(api) {
    bridge = api;
    const toolbar = document.createElement('section');
    toolbar.id = 'datasetToolbar'; toolbar.className = 'panel dataset-toolbar'; toolbar.hidden = true;
    toolbar.innerHTML = `<div class="dataset-target-group"><label for="datasetTargetSearch">目標電話</label><input id="datasetTargetSearch" type="search" placeholder="搜尋電話或用戶名稱" autocomplete="off"><select id="datasetTarget" aria-label="選擇目標電話"></select><div class="dataset-target-pages"><button type="button" id="datasetTargetPrev" class="ghost-button compact-action">上一組</button><small id="datasetTargetCount"></small><button type="button" id="datasetTargetNext" class="ghost-button compact-action">下一組</button></div></div><div class="dataset-date-group"><label for="datasetStart">起始日期</label><input id="datasetStart" type="date"><label for="datasetEnd">結束日期</label><input id="datasetEnd" type="date"><button id="datasetDateApply" class="secondary-button" type="button">套用日期</button><button id="datasetDateReset" class="ghost-button" type="button">完整日期</button></div><div class="dataset-scope-line"><strong id="datasetScope"></strong><button id="datasetClear" class="ghost-button" type="button">清除本次資料</button></div><p id="datasetStatus" class="message" role="status" aria-live="polite"></p>`;
    document.querySelector('.topbar').after(toolbar);
    const section = document.createElement('section'); section.id = 'networkView'; section.className = 'view';
    section.innerHTML = '<section class="panel"><div id="datasetContent"></div></section>';
    document.querySelector('.workspace').append(section);
    const cancel = document.createElement('button'); cancel.id = 'datasetImportCancel'; cancel.type = 'button'; cancel.className = 'secondary-button'; cancel.textContent = '取消匯入';
    $('importProgressDetail').parentElement.append(cancel);
    cancel.addEventListener('click', () => { state.importCancelled = true; state.client?.cancel('import'); status('已取消匯入，正在清理未完成資料。'); });
    $('datasetClear').addEventListener('click', clear);
    let targetTimer;
    $('datasetTargetSearch').addEventListener('input', () => { clearTimeout(targetTimer); targetTimer = setTimeout(() => { state.targetPage = 1; loadTargets(false).catch(failure); }, 250); });
    $('datasetTargetPrev').addEventListener('click', () => { state.targetPage--; loadTargets(false).catch(failure); });
    $('datasetTargetNext').addEventListener('click', () => { state.targetPage++; loadTargets(false).catch(failure); });
    $('datasetTarget').addEventListener('change', () => { state.target = $('datasetTarget').value; resetPage(); render(state.view); });
    $('datasetDateApply').addEventListener('click', () => {
      const start = $('datasetStart').value, end = $('datasetEnd').value;
      if (!start || !end || start > end) { status('請設定有效日期，起始日期不可晚於結束日期。', true); return; }
      bridge.setDatasetDate({ active: true, start, end }); resetPage(); render(state.view);
    });
    $('datasetDateReset').addEventListener('click', () => { bridge.setDatasetDate({ active: false }); fillDates(); resetPage(); render(state.view); });
    $('datasetContent').addEventListener('click', click);
    for (const [view, container] of [['profile', 'profileView'], ['stats', 'statsView'], ['hours', 'hoursView']]) {
      const pages = document.createElement('div'); pages.id = `dataset-${view}-pages`; pages.hidden = true;
      $(container).querySelector('.panel').append(pages); pages.addEventListener('click', click);
    }
    $('hourHotspotContent').addEventListener('click', hotspotClick);
    $('hourHotspotContent').addEventListener('keydown', event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); hotspotClick(event); } });
    $('datasetContent').addEventListener('change', event => {
      if (event.target.id === 'datasetRankMode') { state.mode = event.target.value; state.page = 1; render(state.view); }
    });
    let searchTimer;
    $('datasetContent').addEventListener('input', event => {
      if (event.target.matches('[data-phone-note]')) bridge.updateNote(event.target.dataset.phoneNote, event.target.value, event.target);
      if (event.target.id === 'datasetSearch') {
        state.search = event.target.value;
        clearTimeout(searchTimer); searchTimer = setTimeout(() => { state.page = 1; render(state.view, true); }, 300);
      }
    });
    document.querySelector('.workspace').addEventListener('dragover', event => { if (event.dataTransfer?.types.includes('Files')) event.preventDefault(); });
    document.querySelector('.workspace').addEventListener('drop', event => { if (event.dataTransfer?.files.length) { event.preventDefault(); importFiles([...event.dataTransfer.files]); } });
    const pdf = document.createElement('button'); pdf.type = 'button'; pdf.className = 'secondary-button'; pdf.dataset.attachmentPdf = 'network'; pdf.textContent = '網路歷程 PDF'; pdf.hidden = true;
    pdf.addEventListener('click', () => exportAttachments('pdf', 'network'));
    document.querySelector('.attachment-pdf-grid').append(pdf);
    const exportCancel = document.createElement('button'); exportCancel.id = 'datasetExportCancel'; exportCancel.type = 'button'; exportCancel.className = 'ghost-button'; exportCancel.textContent = '取消匯出'; exportCancel.hidden = true;
    exportCancel.addEventListener('click', () => { state.exporting = false; state.resumeExport?.(); state.client?.cancel('export'); $('attachmentExportStatus').textContent = '已取消；先前已下載的分卷保留。'; });
    $('attachmentExportStatus').after(exportCancel);
    const jsonCancel = exportCancel.cloneNode(true); jsonCancel.id = 'datasetJsonCancel';
    jsonCancel.addEventListener('click', () => { state.exporting = false; state.resumeExport?.(); state.client?.cancel('export'); $('exportMessage').textContent = '已取消；先前已下載的分卷保留。'; });
    $('exportMessage').after(jsonCancel);
    root.addEventListener('pageshow', event => { if (event.persisted) root.location.reload(); });
    client().catch(failure);
  }
  function status(message, danger = false) { if ($('datasetStatus')) { $('datasetStatus').textContent = message; $('datasetStatus').classList.toggle('danger-text', danger); } }
  function failure(error) { if (error?.name !== 'AbortError') { const message = error.message || '處理失敗，請重新匯入。'; status(message, true); if (!multi()) $('importStatus').textContent = message; } }
  function multi() { return state.dataset?.ui_mode === 'multi'; }
  function resetPage() { state.page = 1; state.search = ''; state.hotspot = ''; bridge.setCallPage(1); $('recordSearch').value = ''; }
  function syncMode() {
    document.querySelector('[data-view="network"]').hidden = !multi();
    document.querySelector('[data-attachment-pdf="network"]').hidden = !multi();
    $('datasetToolbar').hidden = !multi() || ['multiLocation', 'submission'].includes(state.view);
    bridge.syncDateFilterPanel();
  }
  async function client() {
    if (state.client) return state.client;
    if (!state.starting) state.starting = root.PhoneDatasetClient.create({ onProgress(progress) {
      if (state.busy) {
        $('importProgressStage').textContent = `${progress.stage || '處理中'}${progress.fileCount ? ` · ${progress.fileIndex}/${progress.fileCount}` : ''}`;
        $('importProgressDetail').textContent = `${progress.fileName || ''} ${number(progress.processedRows)} 列已處理`;
        const percent = progress.totalBytes ? Math.min(99, 100 * progress.completedBytes / progress.totalBytes) : 8;
        $('importProgressFill').style.width = percent + '%';
      }
    } }).then(value => state.client = value).catch(error => { state.starting = null; throw error; });
    return state.starting;
  }
  async function importFiles(files) {
    if (!files?.length || state.busy || state.exporting) return;
    state.busy = true;
    state.importCancelled = false;
    $('importProgressModal').hidden = false; $('datasetImportCancel').hidden = false;
    $('importProgressStage').textContent = '準備匯入'; $('importProgressDetail').textContent = '檢查檔案與本機暫存空間。';
    $('importButton').disabled = true;
    try {
      const service = await client(); service.cancel('view');
      if (state.importCancelled) throw new DOMException('已取消', 'AbortError');
      const result = await service.request('importBatch', { files }, { channel: 'import' });
      const targets = result.dataset.ui_mode === 'multi' ? await service.request('targets', { datasetId: result.dataset.id, options: { page: 1 } }, { channel: 'import' }) : null;
      const old = state.dataset;
      state.dataset = result.dataset; state.page = 1; state.targetPage = 1; state.search = ''; state.date = { active: false }; state.hours = null; state.counties = null;
      service.activate(state.dataset.id);
      state.sort = { column: 'occurred_at', direction: 'asc' }; $('datasetTargetSearch').value = '';
      bridge.adoptDataset(state.dataset); state.hotSummary = null; state.summary = null; state.target = '';
      fillDates(); if (targets) showTargets(targets, true); else $('datasetTarget').innerHTML = '';
      syncMode();
      $('importResults').innerHTML = result.results.map(r => `<div class="import-result-item ${r.ok ? '' : 'danger-text'}"><strong>${esc(r.fileName)}</strong><span>${r.ok ? '完成' : esc(r.message)}</span></div>`).join('');
      $('importStatus').textContent = `匯入完成：通聯 ${number(state.dataset.call_count)} 筆，網路 ${number(state.dataset.data_count)} 筆。`;
      if (old) await service.request('remove', { datasetId: old.id });
      status('資料只在本次瀏覽器工作階段使用；重新開啟須再次匯入。');
      bridge.setView('hours');
    } catch (error) { $('importStatus').textContent = error.name === 'AbortError' ? '已取消匯入，原有資料未變更。' : error.message; failure(error); }
    finally { state.busy = false; $('importButton').disabled = false; $('importProgressModal').hidden = true; $('fileInput').value = ''; }
  }
  function fillDates() { $('datasetStart').value = state.dataset?.date_bounds?.start || ''; $('datasetEnd').value = state.dataset?.date_bounds?.end || ''; }
  async function loadTargets(selectFirst) {
    if (!state.dataset) return;
    const revision = (state.targetRevision || 0) + 1; state.targetRevision = revision;
    const result = await state.client.request('targets', { datasetId: state.dataset.id, options: { search: $('datasetTargetSearch').value, page: state.targetPage } });
    if (revision !== state.targetRevision) return;
    showTargets(result, selectFirst);
  }
  function showTargets(result, selectFirst) {
    const options = result.rows.map(row => `<option value="${esc(row.phone)}">${esc(row.phone || '未辨識目標')}${row.names?.length ? ' · ' + esc(row.names.join('、')) : ''}｜通聯 ${number(row.call_count)} · 網路 ${number(row.data_count)}</option>`);
    if (selectFirst) state.target = result.rows.find(row => row.call_count + row.data_count > 0)?.phone ?? result.rows[0]?.phone ?? '';
    if (!result.rows.some(row => row.phone === state.target)) options.unshift(`<option value="${esc(state.target)}">${esc(state.target || '未辨識目標')}（目前選取）</option>`);
    $('datasetTarget').innerHTML = options.join(''); $('datasetTarget').value = state.target;
    $('datasetTargetCount').textContent = `${number(result.total)} 個目標 · 第 ${state.targetPage} 組`;
    $('datasetTargetPrev').disabled = state.targetPage <= 1;
    $('datasetTargetNext').disabled = state.targetPage * 500 >= result.total;
  }
  function scope(kind) { return { datasetId: state.dataset.id, ...(multi() ? { target: state.target } : {}), kind, date: bridge.readAnalysis().date }; }
  function scopeLabel() { const date = bridge.readAnalysis().date; return `${multi() ? state.target || '未辨識目標' : '完整匯入資料'} · ${date.active ? `${date.start} 至 ${date.end}` : '完整日期'}`; }
  const note = phone => phone ? `<input class="phone-note-input" data-phone-note="${esc(phone)}" aria-label="電話備註 ${esc(phone)}" value="${esc(bridge.getNotes()[phone] || '')}">` : '';
  function pager(total, page = state.page, pageSize = 500, label = '') {
    return `<div class="dataset-pagination"><span id="datasetPageInfo">${label || `共 ${number(total)} 筆`} · 第 ${page} / ${Math.max(1, Math.ceil(total / pageSize))} 頁${label ? '' : ` · 每頁 ${pageSize} 筆`}</span><div><button type="button" class="ghost-button" data-dataset-page="prev" ${page <= 1 ? 'disabled' : ''}>上一頁</button><button type="button" class="secondary-button" data-dataset-page="next" ${page * pageSize >= total ? 'disabled' : ''}>下一頁</button></div></div>`;
  }
  function searchBox(placeholder) { return `<input id="datasetSearch" class="search-input dataset-search" type="search" placeholder="${esc(placeholder)}" value="${esc(state.search)}">`; }
  function table(headers, rows, id = 'datasetRows') { return `<div class="table-wrap dataset-table-wrap"><table class="${headers.length >= 8 ? 'dataset-record-table' : ''}"><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody id="${id}">${rows.map(cells => `<tr>${cells.map(v => `<td>${v}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`; }
  function heading(title, detail = '') { return `<div class="section-title"><div><h2>${esc(title)}</h2><p class="muted">${esc(detail)}</p></div></div>`; }
  function metrics(summary) { return `<div class="summary-grid">${[['通聯筆數', summary.call_count], ['網路歷程', summary.data_count], ['通話總秒數', summary.call_seconds], ['日期異常', summary.invalid_dates]].map(([k, v]) => `<div class="metric-card"><span>${k}</span><strong>${number(v)}</strong></div>`).join('')}</div>`; }
  function setPages(view, total, pageSize, label) {
    const node = $(`dataset-${view}-pages`);
    node.hidden = total <= pageSize;
    node.innerHTML = node.hidden ? '' : pager(total, state.page, pageSize, label);
  }
  function countyRows(keys) {
    const counts = state.hotSummary?.counties || {}, total = Object.values(counts).reduce((sum, n) => sum + n, 0);
    return keys.map(county => ({ county, count: counts[county] || 0, percent: total ? (counts[county] || 0) * 100 / total : 0 }));
  }
  function callRow(row) {
    const extra = [row.external_ip, row.note].filter(Boolean).join('；');
    const details = multi() ? `<details class="record-details"><summary>完整資訊</summary><dl class="detail-list">${[['IMSI',row.imsi],['IPv4',row.external_ipv4],['IPv6',row.external_ipv6],['內網 IP',row.internal_ip],['基地台',(row.stations || []).map(s=>s.address || s.cell_id).join('；')],['來源位置',`${row.source_sheet || ''} / 第 ${row.row_number || ''} 列`]].filter(([,v])=>v).map(([k,v])=>`<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join('')}</dl></details>` : '';
    return `<tr><td>${esc(row.source_file || '')}</td><td>${esc(row.occurred_at)}</td><td>${esc(row.call_type)}</td><td><span class="phone-value">${esc(row.target_phone)}</span></td><td>${note(row.target_phone)}</td><td><span class="phone-value">${esc(row.counterparty_phone)}</span></td><td>${note(row.counterparty_phone)}</td><td>${esc(row.duration_seconds ?? '')}</td><td>${esc(row.imei)}</td><td>${esc(extra)}${details}</td></tr>`;
  }
  async function render(view) {
    if (!state.dataset) return false;
    const changed = state.view !== view;
    if (changed) { state.page = 1; state.search = ''; state.hotspot = ''; state.sort = { column: 'occurred_at', direction: 'asc' }; }
    state.view = view; syncMode();
    const revision = ++state.revision; state.client.cancel('view');
    if (['multiLocation', 'submission', 'export'].includes(view)) {
      if (view === 'submission' && state.ticketDataset !== state.dataset.id) prefillTickets().catch(failure);
      return false;
    }
    $('datasetScope').textContent = scopeLabel();
    const content = $(`${view}View`); if (!content) return false;
    content.setAttribute('aria-busy', 'true');
    const filters = bridge.readAnalysis(), queryScope = scope();
    const filterKey = JSON.stringify([queryScope,filters.hours,filters.counties,filters.rank,$('hourHotspotSearch').value]);
    if (filterKey !== state.filterKey) { state.page = 1; state.hotspot = ''; state.filterKey = filterKey; }
    const query = async (op, payload) => {
      if (revision !== state.revision) throw new DOMException('已取消', 'AbortError');
      const result = await state.client.request(op, payload, { channel: 'view' });
      if (revision !== state.revision) throw new DOMException('已取消', 'AbortError');
      return result;
    };
    try {
      if (view === 'calls') {
        const result = await query('page', { scope: { ...queryScope, kind: multi() ? 'call' : undefined }, options: { page: filters.page, search: $('recordSearch').value, sort: filters.sort, notes: bridge.getNotes() } });
        $('callRows').innerHTML = result.rows.length ? result.rows.map(callRow).join('') : '<tr><td colspan="10">尚無符合資料。</td></tr>';
        bridge.renderCallPagination(result.total, (filters.page - 1) * result.pageSize, result.rows.length, Math.max(1, Math.ceil(result.total / result.pageSize)));
        bridge.applyCallColumnWidths();
      } else if (view === 'network') {
        const result = await query('page', { scope: { ...queryScope, kind: 'data' }, options: { page: state.page, search: state.search, sort: state.sort, notes: bridge.getNotes() } });
        const cols = [['occurred_at','開始時間'],['ended_at','結束時間'],['target_phone','目標電話'],['duration_seconds','連線秒數'],['external_ipv4','IPv4'],['external_ipv6','IPv6'],['internal_ip','內網 IP'],['imei','IMEI'],['imsi','IMSI'],['stations','基地台'],['note','備註'],['source_file','來源']];
        const focused = document.activeElement?.id === 'datasetSearch', cursor = $('datasetSearch')?.selectionStart;
        $('datasetContent').innerHTML = heading('網路歷程') + searchBox('搜尋時間、電話、IMEI、IP、備註或來源') + table(cols.map(([key,label])=>key==='stations'?label:`<button class="table-sort-button" data-dataset-sort="${key}" type="button">${label}</button>`),result.rows.map(row=>cols.map(([key])=>key==='stations'?esc((row.stations || []).map(s=>s.address || s.cell_id).join('；')):key==='source_file'?esc(`${row.source_file || ''} / ${row.source_sheet || ''} / 第 ${row.row_number || ''} 列`):esc(row[key] ?? '')))) + pager(result.total,state.page,result.pageSize);
        if (focused) { $('datasetSearch').focus(); if (cursor != null) $('datasetSearch').setSelectionRange(cursor,cursor); }
      } else if (view === 'profile') {
        const summary = await query('summary', { scope: queryScope });
        const users = await query('users', { scope: queryScope, options: { page: state.page } });
        const imeis = await query('aggregate', { queryKey: summary.queryKey, group: 'imei', options: { page: state.page, mode: 'key' } });
        state.summary = summary;
        $('profileSummaryCards').innerHTML = [['通聯筆數',summary.call_count + summary.data_count],['目標電話',summary.target_count],['對象電話',summary.counterparty_count],['第一筆時間',summary.first_seen || '-'],['最後時間',summary.last_seen || '-'],['總秒數',summary.call_seconds + summary.data_seconds]].map(([label,value])=>bridge.metricCard(label,value ?? 0)).join('');
        const subjects = new Map();
        for (const row of users.rows) for (const [key,value] of Object.entries(row.subject || {})) { const values = subjects.get(key) || new Set(); for(const item of String(value ?? '').split('、').map(v=>v.trim()).filter(Boolean)) values.add(item); if(values.size) subjects.set(key,values); }
        $('profileContent').innerHTML = subjects.size ? [...subjects].map(([key,values])=>`<dt>${esc(key)}</dt><dd>${esc([...values].join('、'))}</dd>`).join('') : '<dt>狀態</dt><dd>尚無用戶資料。</dd>';
        $('profileImeiList').innerHTML = imeis.rows.length ? imeis.rows.map(row=>`<span class="imei-chip">${esc(row.key)}</span>`).join('') : '<p class="muted">尚無 IMEI 資料。</p>';
        setPages('profile',Math.max(Math.ceil(users.total/users.pageSize),Math.ceil(imeis.total/imeis.pageSize)),1,`用戶 ${number(users.total)} 筆／IMEI ${number(imeis.total)} 項`);
      } else if (view === 'stats') {
        const summary = await query('summary', { scope: queryScope }), groups = [];
        for (const [key,label] of [['phone:inbound','來電排行'],['phone:outbound','去電排行'],['phone:total','完整排行']]) {
          const result = await query('aggregate', { queryKey: summary.queryKey, group: key, options: { page: state.page, mode: filters.rank } }); groups.push({label,...result});
        }
        $('statsContent').innerHTML = groups.map(group=>bridge.statsCard(group.label,group.rows.map((row,index)=>({...row,phone:row.key,rank:(state.page-1)*group.pageSize+index+1})))).join('');
        setPages('stats',Math.max(...groups.map(g=>Math.ceil(g.total/g.pageSize))),1,'完整電話排行');
      } else if (view === 'hours') {
        const summary = await query('summary', { scope: { ...queryScope, hours: filters.hours } });
        const hotspots = await query('hotspots', { scope: { ...queryScope, hours: filters.hours }, options:{page:state.page,pageSize:20,search:$('hourHotspotSearch').value,counties:filters.counties} });
        state.hotSummary = summary;
        bridge.renderHourBuckets(summary.hours);
        const total = summary.call_count + summary.data_count;
        $('hourHotspotContent').innerHTML = hotspots.rows.length ? hotspots.rows.map(row=>`<div class="hotspot-item" data-hotspot-address="${esc(row.key)}" role="button" tabindex="0" aria-expanded="false"><strong>${esc(row.address)}</strong><span>${row.count} 筆 / ${bridge.formatPercent(total ? 100*row.count/total : 0)}</span><small>${esc(row.first_seen || '-')} 至 ${esc(row.last_seen || '-')}</small></div>`).join('') : '<p class="muted">尚無符合資料。</p>';
        setPages('hours',hotspots.total,hotspots.pageSize);
      }
    } catch (error) { if (revision === state.revision) failure(error); }
    finally { if (revision === state.revision) { content.setAttribute('aria-busy', 'false'); bridge.syncDateFilterPanel(); } }
    return true;
  }

  async function prefillTickets() {
    const id = state.dataset.id;
    if (state.ticketLoading === id) return;
    state.ticketLoading = id;
    const phones = [];
    try {
      for (let page = 1; ; page++) {
        const result = await state.client.request('ticketPhones', { datasetId: id, page }, { channel: 'view' });
        if (state.dataset?.id !== id || state.view !== 'submission') return;
        phones.push(...result.phones);
        if (page * result.pageSize >= result.total) break;
      }
      bridge.setTicketPhones(phones); state.ticketDataset = id;
    } finally { state.ticketLoading = null; }
  }
  function click(event) {
    const button = event.target.closest('button'); if (!button || !state.dataset) return;
    if (button.dataset.datasetPage) { state.page += button.dataset.datasetPage === 'next' ? 1 : -1; render(state.view); }
    if (button.dataset.datasetSort) { const column = button.dataset.datasetSort; state.sort = { column, direction: state.sort.column === column && state.sort.direction === 'asc' ? 'desc' : 'asc' }; state.page = 1; render(state.view); }
  }
  function hotspotClick(event) {
    if (!state.dataset || state.view !== 'hours') return;
    const pageButton = event.target.closest('[data-hotspot-page]');
    if (pageButton) { showHotspot(state.hotspot,Number(pageButton.dataset.hotspotPage)).catch(failure); return; }
    if (event.target.closest('.hotspot-times')) return;
    const item = event.target.closest('[data-hotspot-address]'); if (!item) return;
    if (state.hotspot === item.dataset.hotspotAddress) { state.hotspot = ''; ++state.detailRevision; item.classList.remove('expanded'); item.setAttribute('aria-expanded','false'); item.querySelector('.hotspot-times')?.remove(); item.querySelector('.hotspot-detail-pagination')?.remove(); }
    else showHotspot(item.dataset.hotspotAddress,1).catch(failure);
  }
  async function showHotspot(address, page) {
    state.hotspot = address;
    const revision = state.revision, detailRevision = state.detailRevision = (state.detailRevision || 0) + 1;
    const result = await state.client.request('page', { scope: { ...scope(), hours: bridge.readAnalysis().hours }, options: { address, page } }, { channel: 'view' });
    if (revision !== state.revision || detailRevision !== state.detailRevision || address !== state.hotspot) return;
    for (const item of $('hourHotspotContent').querySelectorAll('[data-hotspot-address]')) {
      item.querySelector('.hotspot-times')?.remove(); item.querySelector('.hotspot-detail-pagination')?.remove(); const expanded = item.dataset.hotspotAddress === address;
      item.classList.toggle('expanded',expanded); item.setAttribute('aria-expanded',String(expanded));
      if (!expanded) continue;
      const node = document.createElement('ol'); node.className = 'hotspot-times'; node.id = 'datasetHotspotRows';
      node.innerHTML = result.rows.map(r=>`<li>${esc(r.occurred_at)}${multi()?` · ${esc(r.record_kind==='data'?'網路':r.call_type)} · ${esc(r.target_phone)} · ${esc(r.source_sheet)} / 第 ${esc(r.row_number)} 列`:''}</li>`).join('');
      item.append(node);
      item.querySelector('.hotspot-detail-pagination')?.remove();
      if(result.total > result.pageSize) {
        const pages=document.createElement('p');pages.className='hotspot-detail-pagination';
        pages.innerHTML=`共 ${number(result.total)} 筆 · 第 ${page} 頁 <button class="secondary-button compact-action" data-hotspot-page="${page - 1}" ${page===1?'disabled':''}>上一頁</button><button class="secondary-button compact-action" data-hotspot-page="${page + 1}" ${page*result.pageSize>=result.total?'disabled':''}>下一頁</button>`;item.append(pages);
      }
    }
  }
  async function clear() {
    if (state.busy || state.exporting) return;
    ++state.revision; state.client?.cancel('view');
    try {
      if (state.dataset) await state.client.request('remove', { datasetId: state.dataset.id });
      state.dataset = null; state.target = ''; $('datasetTarget').innerHTML = ''; $('datasetScope').textContent = '';
      state.client?.activate(null);
      bridge.adoptDataset(null); state.summary = null; state.hotSummary = null;
      document.querySelectorAll('.view').forEach(node=>node.removeAttribute('aria-busy'));
      for (const view of ['profile','stats','hours']) { const node=$(`dataset-${view}-pages`); node.innerHTML=''; node.hidden=true; }
      $('datasetContent').innerHTML = ''; syncMode();
      $('importResults').innerHTML = ''; $('importStatus').textContent = '本次匯入資料已清除。'; fillDates();
      status('已清除本次資料；電話備註與介面設定仍保留。'); bridge.setView('hours');
    } catch (error) { failure(error); }
  }
  function showExport() {
    $('attachmentExportModal').hidden = false;
    $('attachmentExportStatus').textContent = state.dataset ? `匯出範圍：${scopeLabel()}。每卷最多 10,000 筆，另附完整摘要。請允許此網站下載多個檔案並保留所有分卷。` : '尚未匯入資料。';
    document.querySelectorAll('#attachmentXlsxButton,[data-attachment-pdf]').forEach(button => button.disabled = !state.dataset || state.exporting);
  }
  function download(name, bytes, type) {
    const blob = new Blob([bytes], { type }), url = URL.createObjectURL(blob), link = document.createElement('a');
    link.href = url; link.download = name; link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  function downloadPause() {
    return new Promise(resolve => {
      const timer = setTimeout(() => { state.resumeExport = null; resolve(); }, 1100);
      state.resumeExport = () => { clearTimeout(timer); state.resumeExport = null; resolve(); };
    });
  }
  async function exportAttachments(format, section) {
    if (!state.dataset || state.exporting || state.busy) return;
    state.exporting = true; const exportScope = scope(), label = scopeLabel(), notes = { ...bridge.getNotes() }, stamp = new Date().toISOString().replace(/[:.]/g, '-');
    state.client.cancel('view'); $('datasetExportCancel').hidden = false;
    document.querySelectorAll('#attachmentXlsxButton,[data-attachment-pdf]').forEach(b => b.disabled = true);
    try {
      let after = null, volume = 1;
      const service = state.client;
      do {
        if (!state.exporting) break;
        $('attachmentExportStatus').textContent = `正在產生第 ${volume} 卷，請保留此分頁。`;
        const result = await service.request('exportVolume', { scope: exportScope, after, volume, format, section, label, notes, stamp }, { channel: 'export' });
        if (!state.exporting) break;
        if (result.bytes) download(`附卷-${section || '完整'}-${stamp}-第${String(volume).padStart(3, '0')}卷.${format}`, result.bytes, result.mime);
        after = result.after; volume++;
        await downloadPause();
      } while (after && state.exporting);
      let summaryVolume = 0, more = true;
      if (format === 'xlsx' || !['calls', 'network'].includes(section)) while (more && state.exporting) {
        summaryVolume++;
        $('attachmentExportStatus').textContent = `正在產生完整範圍摘要第 ${summaryVolume} 卷。`;
        const result = await service.request('exportSummary', { scope: exportScope, volume: summaryVolume, format, section, label, notes }, { channel: 'export' });
        if (!state.exporting) break;
        download(`附卷-完整摘要-${section || '完整'}-${stamp}-第${String(summaryVolume).padStart(3, '0')}卷.${format}`, result.bytes, result.mime);
        more = result.more; await downloadPause();
      }
      if (state.exporting) $('attachmentExportStatus').textContent = `已完成明細 ${volume - 1} 卷、完整範圍摘要 ${summaryVolume} 卷。請允許此網站下載多個檔案。`;
    } catch (error) { if (error.name !== 'AbortError') $('attachmentExportStatus').textContent = error.message; }
    finally { state.exporting = false; $('datasetExportCancel').hidden = true; document.querySelectorAll('#attachmentXlsxButton,[data-attachment-pdf]').forEach(b => b.disabled = !state.dataset); }
  }
  async function exportJson() {
    if (!state.dataset || state.exporting || state.busy) return;
    state.exporting = true;
    $('datasetJsonCancel').hidden = false;
    state.client.cancel('view');
    const exportScope = { datasetId: state.dataset.id }, stamp = new Date().toISOString().replace(/[:.]/g, '-');
    try {
      let after = null, volume = 1;
      do {
        const result = await state.client.request('jsonVolume', { scope: exportScope, after, volume }, { channel: 'export' });
        if (!state.exporting) break;
        if (result.bytes) download(`phone-workspace-${stamp}-part-${volume}.json`, result.bytes, 'application/json');
        after = result.after; volume++; $('exportMessage').textContent = `已匯出 ${volume - 1} 卷資料 JSON。`;
        await downloadPause();
      } while (after && state.exporting);
    } catch (error) { $('exportMessage').textContent = error.message; }
    finally { state.exporting = false; $('datasetJsonCancel').hidden = true; }
  }
  root.PhoneDatasetUI = { init, active: () => Boolean(state.dataset), multi, invalidDates: () => state.dataset?.invalid_dates || 0, countyRows, render, importFiles, clear, showExport, exportAttachments, exportJson };
})(globalThis);
