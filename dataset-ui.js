(function (root) {
  'use strict';
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const number = value => Number(value || 0).toLocaleString();
  const state = { enabled: false, dataset: null, client: null, starting: null, target: '', view: 'calls', page: 1, targetPage: 1, search: '', sort: { column: 'occurred_at', direction: 'asc' }, date: { active: false }, hours: null, counties: null, mode: 'count', revision: 0, busy: false, exporting: false };
  let bridge;
  function init(api) {
    bridge = api;
    state.enabled = true;
    const toolbar = document.createElement('section');
    toolbar.id = 'datasetToolbar'; toolbar.className = 'panel dataset-toolbar'; toolbar.hidden = true;
    toolbar.innerHTML = `<div class="dataset-target-group"><label for="datasetTargetSearch">目標電話</label><input id="datasetTargetSearch" type="search" placeholder="搜尋電話或用戶名稱" autocomplete="off"><select id="datasetTarget" aria-label="選擇目標電話"></select><div class="dataset-target-pages"><button type="button" id="datasetTargetPrev" class="ghost-button compact-action">上一組</button><small id="datasetTargetCount"></small><button type="button" id="datasetTargetNext" class="ghost-button compact-action">下一組</button></div></div><div class="dataset-date-group"><label for="datasetStart">起始日期</label><input id="datasetStart" type="date"><label for="datasetEnd">結束日期</label><input id="datasetEnd" type="date"><button id="datasetDateApply" class="secondary-button" type="button">套用日期</button><button id="datasetDateReset" class="ghost-button" type="button">完整日期</button></div><div class="dataset-scope-line"><strong id="datasetScope"></strong><button id="datasetClear" class="ghost-button" type="button">清除本次資料</button></div><p id="datasetStatus" class="message" role="status" aria-live="polite"></p>`;
    document.querySelector('.topbar').after(toolbar);
    const section = document.createElement('section'); section.id = 'datasetView'; section.className = 'view';
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
    $('datasetTarget').addEventListener('change', () => { state.target = $('datasetTarget').value; state.page = 1; state.search = ''; render(state.view); });
    $('datasetDateApply').addEventListener('click', () => {
      const start = $('datasetStart').value, end = $('datasetEnd').value;
      if (!start || !end || start > end) { status('請設定有效日期，起始日期不可晚於結束日期。', true); return; }
      state.date = { active: true, start, end }; state.page = 1; render(state.view);
    });
    $('datasetDateReset').addEventListener('click', () => { state.date = { active: false }; fillDates(); state.page = 1; render(state.view); });
    $('datasetContent').addEventListener('click', click);
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
    const pdf = document.createElement('button'); pdf.type = 'button'; pdf.className = 'secondary-button'; pdf.dataset.attachmentPdf = 'network'; pdf.textContent = '網路歷程 PDF';
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
  function failure(error) { if (error?.name !== 'AbortError') status(error.message || '處理失敗，請重新匯入。', true); }
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
    state.busy = true; state.enabled = true; $('datasetToolbar').hidden = false;
    state.importCancelled = false;
    $('importProgressModal').hidden = false; $('datasetImportCancel').hidden = false;
    $('importProgressStage').textContent = '準備匯入'; $('importProgressDetail').textContent = '檢查檔案與本機暫存空間。';
    $('importButton').disabled = true;
    try {
      const service = await client(); service.cancel('view');
      if (state.importCancelled) throw new DOMException('已取消', 'AbortError');
      const result = await service.request('importBatch', { files }, { channel: 'import' });
      const old = state.dataset;
      state.dataset = result.dataset; state.page = 1; state.targetPage = 1; state.search = ''; state.date = { active: false }; state.hours = null; state.counties = null;
      service.activate(state.dataset.id);
      state.sort = { column: 'occurred_at', direction: 'asc' }; $('datasetTargetSearch').value = '';
      bridge.clearLegacy(); fillDates(); await loadTargets(true);
      $('importResults').innerHTML = result.results.map(r => `<div class="import-result-item ${r.ok ? '' : 'danger-text'}"><strong>${esc(r.fileName)}</strong><span>${r.ok ? '完成' : esc(r.message)}</span></div>`).join('');
      $('importStatus').textContent = `匯入完成：通聯 ${number(state.dataset.call_count)} 筆，網路 ${number(state.dataset.data_count)} 筆。`;
      if (old) await service.request('remove', { datasetId: old.id });
      status('資料只在本次瀏覽器工作階段使用；重新開啟須再次匯入。');
      bridge.setView('calls');
    } catch (error) { $('importStatus').textContent = error.name === 'AbortError' ? '已取消匯入，原有資料未變更。' : error.message; failure(error); }
    finally { state.busy = false; $('importButton').disabled = false; $('importProgressModal').hidden = true; $('fileInput').value = ''; }
  }
  function fillDates() { $('datasetStart').value = state.dataset?.date_bounds?.start || ''; $('datasetEnd').value = state.dataset?.date_bounds?.end || ''; }
  async function loadTargets(selectFirst) {
    if (!state.dataset) return;
    const revision = (state.targetRevision || 0) + 1; state.targetRevision = revision;
    const result = await state.client.request('targets', { datasetId: state.dataset.id, options: { search: $('datasetTargetSearch').value, page: state.targetPage } });
    if (revision !== state.targetRevision) return;
    const options = result.rows.map(row => `<option value="${esc(row.phone)}">${esc(row.phone || '未辨識目標')}${row.names?.length ? ' · ' + esc(row.names.join('、')) : ''}｜通聯 ${number(row.call_count)} · 網路 ${number(row.data_count)}</option>`);
    if (selectFirst) state.target = result.rows.find(row => row.call_count + row.data_count > 0)?.phone ?? result.rows[0]?.phone ?? '';
    if (!result.rows.some(row => row.phone === state.target)) options.unshift(`<option value="${esc(state.target)}">${esc(state.target || '未辨識目標')}（目前選取）</option>`);
    $('datasetTarget').innerHTML = options.join(''); $('datasetTarget').value = state.target;
    $('datasetTargetCount').textContent = `${number(result.total)} 個目標 · 第 ${state.targetPage} 組`;
    $('datasetTargetPrev').disabled = state.targetPage <= 1;
    $('datasetTargetNext').disabled = state.targetPage * 500 >= result.total;
  }
  function scope(kind) { return { datasetId: state.dataset.id, target: state.target, kind, date: { ...state.date } }; }
  function scopeLabel() { return `${state.target || '未辨識目標'} · ${state.date.active ? `${state.date.start} 至 ${state.date.end}` : '完整日期'}`; }
  const note = phone => phone ? `<input class="phone-note-input" data-phone-note="${esc(phone)}" aria-label="電話備註 ${esc(phone)}" value="${esc(bridge.getNotes()[phone] || '')}">` : '';
  function pager(total, page = state.page, pageSize = 500, label = '') {
    return `<div class="dataset-pagination"><span id="datasetPageInfo">${label || `共 ${number(total)} 筆`} · 第 ${page} / ${Math.max(1, Math.ceil(total / pageSize))} 頁${label ? '' : ` · 每頁 ${pageSize} 筆`}</span><div><button type="button" class="ghost-button" data-dataset-page="prev" ${page <= 1 ? 'disabled' : ''}>上一頁</button><button type="button" class="secondary-button" data-dataset-page="next" ${page * pageSize >= total ? 'disabled' : ''}>下一頁</button></div></div>`;
  }
  function searchBox(placeholder) { return `<input id="datasetSearch" class="dataset-search" type="search" placeholder="${esc(placeholder)}" value="${esc(state.search)}">`; }
  function table(headers, rows, id = 'datasetRows') { return `<div class="table-wrap dataset-table-wrap"><table class="${headers.length >= 8 ? 'dataset-record-table' : ''}"><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody id="${id}">${rows.map(cells => `<tr>${cells.map(v => `<td>${v}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`; }
  function heading(title, detail = '') { return `<div class="section-title"><div><h2>${esc(title)}</h2><p class="muted">${esc(detail)}</p></div></div>`; }
  function metrics(summary) { return `<div class="summary-grid">${[['通聯筆數', summary.call_count], ['網路歷程', summary.data_count], ['通話總秒數', summary.call_seconds], ['日期異常', summary.invalid_dates]].map(([k, v]) => `<div class="metric-card"><span>${k}</span><strong>${number(v)}</strong></div>`).join('')}</div>`; }
  async function render(view, preserveSearch = false) {
    if (state.view !== view) { state.page = 1; state.search = ''; state.sort = { column: 'occurred_at', direction: 'asc' }; }
    state.view = view;
    if (!state.enabled) return false;
    $('datasetToolbar').hidden = ['multiLocation', 'submission'].includes(view);
    $('dateFilterPanel').hidden = true;
    if (['multiLocation', 'submission', 'export'].includes(view)) {
      $('datasetView').classList.remove('active-view');
      if (view === 'submission' && state.dataset && state.ticketDataset !== state.dataset.id) prefillTickets().catch(failure);
      return false;
    }
    document.querySelectorAll('.view').forEach(node => node.classList.remove('active-view')); $('datasetView').classList.add('active-view');
    const revision = ++state.revision; state.client?.cancel('view');
    if (!state.dataset) { $('datasetContent').innerHTML = heading('尚未匯入資料', '選擇一份多電話 Excel，即可切換通聯與網路歷程。') + '<table><tbody id="datasetRows"></tbody></table>'; return true; }
    $('datasetScope').textContent = scopeLabel();
    const content = $('datasetContent'); content.setAttribute('aria-busy', 'true');
    if (!preserveSearch) content.innerHTML = heading('載入中…', scopeLabel());
    const query = (op, payload) => state.client.request(op, payload, { channel: 'view' });
    try {
      let html;
      if (view === 'calls' || view === 'network') {
        const kind = view === 'network' ? 'data' : 'call';
        const result = await query('page', { scope: scope(kind), options: { page: state.page, search: state.search, sort: state.sort, notes: bridge.getNotes() } });
        const cols = kind === 'call' ? [['occurred_at','時間'],['call_type','類型'],['target_phone','目標電話'],['target_note','目標備註'],['counterparty_phone','對象電話'],['counterparty_note','對象備註'],['duration_seconds','秒數'],['imei','IMEI'],['imsi','IMSI'],['external_ip','外部 IP'],['internal_ip','內網 IP'],['stations','基地台'],['note','備註'],['source_file','來源']] : [['occurred_at','開始時間'],['ended_at','結束時間'],['target_phone','目標電話'],['duration_seconds','連線秒數'],['external_ipv4','IPv4'],['external_ipv6','IPv6'],['internal_ip','內網 IP'],['imei','IMEI'],['imsi','IMSI'],['stations','基地台'],['note','備註'],['source_file','來源']];
        html = heading(kind === 'call' ? '雙向通聯' : '網路歷程', scopeLabel()) + searchBox('搜尋時間、電話、IMEI、IP、備註或來源') + pager(result.total, state.page, result.pageSize) + table(cols.map(([key, label]) => key === 'stations' ? label : `<button class="dataset-sort" data-dataset-sort="${key}" type="button">${label}${state.sort.column === key ? state.sort.direction === 'asc' ? ' ↑' : ' ↓' : ''}</button>`), result.rows.map(row => cols.map(([key]) => key === 'target_note' ? note(row.target_phone) : key === 'counterparty_note' ? note(row.counterparty_phone) : key === 'source_file' ? esc(`${row.source_file || ''} / ${row.source_sheet || ''} / 第 ${row.row_number || ''} 列`) : key === 'stations' ? esc((row.stations || []).map(s => s.address || s.cell_id).join('；')) : esc(row[key] ?? ''))));
      } else if (view === 'profile') {
        const summary = await query('summary', { scope: scope() });
        const users = await query('users', { scope: scope(), options: { page: state.page } });
        const imeis = await query('aggregate', { queryKey: summary.queryKey, group: 'imei', options: { page: state.page } });
        html = heading('用戶資料', scopeLabel()) + metrics(summary) + `<p>第一筆：${esc(summary.first_seen || '—')}　最後一筆：${esc(summary.last_seen || '—')}</p>` + table(['資料來源','欄位','內容'], users.rows.flatMap(row => Object.entries(row.subject || {}).map(([key, value]) => [esc(`${row.source_sheet || ''} / ${row.row_number || ''}`), esc(key), esc(value)])), 'datasetUsers') + '<h3>IMEI</h3>' + table(['IMEI','出現筆數'], imeis.rows.map(row => [esc(row.key), number(row.count)]), 'datasetImeis') + pager(Math.max(Math.ceil(users.total / users.pageSize), Math.ceil(imeis.total / imeis.pageSize)), state.page, 1, `用戶 ${number(users.total)} 筆／IMEI ${number(imeis.total)} 項`);
      } else if (view === 'stats') {
        const summary = await query('summary', { scope: scope() });
        const groups = [];
        for (const [key, label] of [['phone:inbound','來電／收訊'],['phone:outbound','去電／發訊'],['phone:total','全部電話']]) {
          const result = await query('aggregate', { queryKey: summary.queryKey, group: key, options: { page: state.page, mode: state.mode } });
          groups.push({ label, ...result });
        }
        html = heading('電話統計', '依所選電話與日期統計通聯；網路連線不計入通話。') + `<label>排序 <select id="datasetRankMode"><option value="count" ${state.mode === 'count' ? 'selected' : ''}>次數</option><option value="seconds" ${state.mode === 'seconds' ? 'selected' : ''}>秒數</option></select></label>` + groups.map(group => `<h3>${group.label} · ${number(group.total)} 個電話</h3>` + table(['電話','次數','秒數','備註'], group.rows.map(row => [esc(row.key), number(row.count), number(row.seconds), note(row.key)]), `dataset${group.label}`)).join('') + pager(Math.max(...groups.map(g => Math.ceil(g.total / g.pageSize))), state.page, 1, '完整電話排行');
      } else {
        const summary = await query('summary', { scope: scope() });
        const hotSummary = state.hours ? await query('summary', { scope: { ...scope(), hours: state.hours } }) : summary;
        const hotspots = await query('aggregate', { queryKey: hotSummary.queryKey, group: 'hotspot', options: { page: state.page, search: state.search, counties: state.counties } });
        const max = Math.max(1, ...summary.hours.map(h => h.count));
        const chart = summary.hours.map(h => `<button type="button" class="dataset-hour ${!state.hours || state.hours.includes(h.hour) ? 'selected' : ''}" data-dataset-hour="${h.hour}" aria-pressed="${!state.hours || state.hours.includes(h.hour)}"><span>${number(h.count)}</span><i style="height:${Math.round(80 * h.count / max)}px"></i><small>${String(h.hour).padStart(2, '0')}</small></button>`).join('');
        html = heading('時間分布與基地台熱點', '以通聯及網路歷程的起始時間計算；點選時段可篩選熱點。') + metrics(summary) + `<div class="dataset-hours">${chart}</div><button class="ghost-button" data-dataset-hour="all" type="button">全部時段</button><div class="dataset-counties">${Object.entries(hotSummary.counties).map(([county, count]) => `<label><input type="checkbox" data-county="${esc(county)}" ${!state.counties || state.counties.includes(county) ? 'checked' : ''}>${esc(county)} ${number(count)}</label>`).join('')}<button type="button" data-county-action="all" class="ghost-button">全選</button><button type="button" data-county-action="clear" class="ghost-button">全部取消</button><button type="button" data-county-action="apply" class="secondary-button">套用縣市</button></div>` + searchBox('搜尋基地台地址') + pager(hotspots.total, state.page, hotspots.pageSize) + table(['基地台地址','次數','首次時間','末次時間','明細'], hotspots.rows.map(row => [esc(row.address), number(row.count), esc(row.first_seen), esc(row.last_seen), `<button type="button" class="ghost-button" data-hotspot="${esc(row.key)}">查看</button>`]));
      }
      if (revision !== state.revision) return true;
      const cursor = preserveSearch ? $('datasetSearch')?.selectionStart : null;
      content.innerHTML = html;
      if (preserveSearch) { $('datasetSearch')?.focus(); if (cursor != null) try { $('datasetSearch').setSelectionRange(cursor, cursor); } catch (_) {} }
    } catch (error) { if (revision === state.revision) failure(error); }
    finally { if (revision === state.revision) content.setAttribute('aria-busy', 'false'); }
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
    if (button.dataset.datasetHour) {
      const hour = button.dataset.datasetHour;
      if (hour === 'all') state.hours = null;
      else if (!state.hours) state.hours = [Number(hour)];
      else state.hours = state.hours.includes(Number(hour)) ? state.hours.filter(h => h !== Number(hour)) : [...state.hours, Number(hour)].sort((a, b) => a - b);
      state.page = 1; render(state.view);
    }
    if (button.dataset.countyAction) {
      if (button.dataset.countyAction === 'apply') { state.counties = [...document.querySelectorAll('[data-county]:checked')].map(node => node.dataset.county); state.page = 1; render(state.view); }
      else document.querySelectorAll('[data-county]').forEach(node => node.checked = button.dataset.countyAction === 'all');
    }
    if (button.dataset.hotspot) showHotspot(button.dataset.hotspot, 1).catch(failure);
    if (button.dataset.hotspotPage) showHotspot(state.hotspot, Number(button.dataset.hotspotPage)).catch(failure);
  }
  async function showHotspot(address, page) {
    state.hotspot = address;
    const revision = state.revision;
    const result = await state.client.request('page', { scope: { ...scope(), hours: state.hours }, options: { address, page } }, { channel: 'view' });
    if (revision !== state.revision) return;
    $('datasetHotspotDetail')?.remove();
    const node = document.createElement('section'); node.id = 'datasetHotspotDetail';
    node.innerHTML = `<h3>基地台來源明細</h3>` + table(['時間','類型','目標','來源'], result.rows.map(r => [esc(r.occurred_at), esc(r.record_kind === 'data' ? '網路' : r.call_type), esc(r.target_phone), esc(`${r.source_sheet} / 第 ${r.row_number} 列`)]), 'datasetHotspotRows') + `<p>共 ${number(result.total)} 筆 · 第 ${page} 頁 <button class="ghost-button" data-hotspot-page="${page - 1}" ${page === 1 ? 'disabled' : ''}>上一頁</button><button class="ghost-button" data-hotspot-page="${page + 1}" ${page * result.pageSize >= result.total ? 'disabled' : ''}>下一頁</button></p>`;
    $('datasetContent').append(node);
  }
  async function clear() {
    if (state.busy || state.exporting) return;
    ++state.revision; state.client?.cancel('view');
    try {
      if (state.dataset) await state.client.request('remove', { datasetId: state.dataset.id });
      state.dataset = null; state.target = ''; $('datasetTarget').innerHTML = ''; $('datasetScope').textContent = '';
      state.client?.activate(null);
      $('importResults').innerHTML = ''; $('importStatus').textContent = '本次匯入資料已清除。'; fillDates();
      status('已清除本次資料；電話備註與介面設定仍保留。'); render(state.view);
    } catch (error) { failure(error); }
  }
  function showExport() {
    $('attachmentExportModal').hidden = false;
    $('attachmentXlsxButton').textContent = '下載 XLSX（大型資料自動分卷）';
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
  root.PhoneDatasetUI = { init, active: () => state.enabled, render, importFiles, clear, showExport, exportAttachments, exportJson };
})(globalThis);
