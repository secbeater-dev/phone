(function (root) {
  'use strict';
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const number = value => Number(value || 0).toLocaleString();
  const state = { client: null, starting: null, dataset: null, page: 1, busy: false, revision: 0, expanded: new Map() };
  let bridge, cards;
  const timeSource = value => ({ base: '基地台連線時間', internet: 'internet連線時間（備援）', record: '通聯時間' })[value] || '通聯時間';
  function init(options) {
    bridge = options;
    cards = root.PhoneCardsUI.create($('multiLocationCards'));
    $('multiLocationImportButton').addEventListener('click', () => $('multiLocationFileInput').click());
    $('multiLocationFolderButton').addEventListener('click', () => $('multiLocationFolderInput').click());
    for (const id of ['multiLocationFileInput', 'multiLocationFolderInput']) $(id).addEventListener('change', event => importFiles([...event.target.files]));
    $('multiLocationCancel').addEventListener('click', () => { state.cancelled = true; state.client?.cancel('import'); });
    $('multiLocationPrevPage').addEventListener('click', () => changePage(-1));
    $('multiLocationNextPage').addEventListener('click', () => changePage(1));
    $('multiLocationRows').addEventListener('input', event => { if (event.target.matches('[data-phone-note]')) bridge.updateNote(event.target.dataset.phoneNote, event.target.value, event.target); });
    $('multiLocationRows').addEventListener('click', event => {
      const toggle = event.target.closest('[data-multi-location-detail]');
      if (toggle) { const id = toggle.dataset.multiLocationDetail; if (state.expanded.has(id)) state.expanded.delete(id); else state.expanded.set(id, { details: 1, phones: 1 }); render().catch(fail); }
      const pager = event.target.closest('[data-location-page]');
      if (pager && !pager.disabled) {
        const id = pager.closest('[data-match-id]').dataset.matchId, pages = state.expanded.get(id);
        pages[pager.dataset.locationKind] = Number(pager.dataset.locationPage);
        loadDetails(id).catch(fail);
      }
    });
    cards.render();
  }
  function fail(error) { if (error.name !== 'AbortError') $('multiLocationImportStatus').textContent = error.message || '位置資料讀取失敗。'; }
  function client() {
    if (!state.starting) state.starting = root.PhoneDatasetClient.create({ onProgress(progress) {
      if (state.busy) $('multiLocationImportStatus').textContent = `${progress.stage || '處理中'}${progress.fileCount ? ` · ${progress.fileIndex}/${progress.fileCount}` : ''}${progress.fileName ? ' · ' + progress.fileName : ''} · ${number(progress.processedRows)} 列`;
    } }).then(value => state.client = value).catch(error => { state.starting = null; throw error; });
    return state.starting;
  }
  async function importFiles(inputs) {
    if (!inputs.length || state.busy) return;
    const files = inputs.filter(file => /\.(xml|xlsx|json)$/i.test(file.name)), skipped = inputs.length - files.length;
    if (!files.length) { $('multiLocationImportStatus').textContent = '沒有支援的檔案；請選擇 XML、XLSX 或 workspace JSON。'; return; }
    state.busy = true; state.cancelled = false;
    for (const id of ['multiLocationImportButton','multiLocationFolderButton']) $(id).disabled = true;
    $('multiLocationCancel').hidden = false;
    $('multiLocationImportStatus').textContent = '準備匯入與位置比對…';
    $('multiLocationImportResults').innerHTML = '';
    let service, stagedId, adopted = false;
    try {
      service = await client();
      if (state.cancelled) throw new DOMException('已取消', 'AbortError');
      const result = await service.request('importLocations', { files }, { channel: 'import' });
      stagedId = result.dataset.id;
      $('multiLocationImportStatus').textContent = '比對完成，載入列表與門號…';
      const firstPage = await service.request('locationPage', { datasetId: stagedId, page: 1 }, { channel: 'import' });
      const firstCards = await service.request('phoneCards', { datasetId: stagedId, search: '', page: 1 }, { channel: 'import' });
      if (state.cancelled) throw new DOMException('已取消', 'AbortError');
      // The import is cancellable until every initial query has succeeded.
      $('multiLocationCancel').hidden = true;
      const old = state.dataset;
      ++state.revision; state.dataset = result.dataset; state.analysis = result.analysis; state.page = 1; state.expanded.clear();
      service.activate(result.dataset.id); cards.setDataset(service, result.dataset.id);
      adopted = true;
      $('multiLocationImportResults').innerHTML = result.results.map(row => `<div class="import-result-item ${row.ok ? '' : 'danger-text'}"><strong>${esc(row.fileName)}</strong><span>${row.ok ? '完成' : '失敗：' + esc(row.message)}</span></div>`).join('');
      await render(firstPage); await cards.render(firstCards);
      $('multiLocationImportStatus').textContent = `匯入完成：${number(files.length)} 個檔案，符合 ${number(result.analysis.total)} 筆。${skipped ? `略過 ${number(skipped)} 個不支援的檔案。` : ''}`;
      if (old) await service.request('remove', { datasetId: old.id });
    } catch (error) {
      if (!adopted && stagedId && service) await service.request('remove', { datasetId: stagedId }).catch(() => {});
      $('multiLocationImportStatus').textContent = adopted ? `資料已匯入；畫面或暫存清理失敗：${error.message}` : error.name === 'AbortError' ? '已取消匯入，原有資料未變更。' : `匯入失敗：${error.message} 原有資料未變更。`;
    } finally {
      state.busy = false;
      for (const id of ['multiLocationImportButton','multiLocationFolderButton']) $(id).disabled = false;
      $('multiLocationCancel').hidden = true; $('multiLocationFileInput').value = ''; $('multiLocationFolderInput').value = '';
    }
  }
  async function changePage(delta) { state.page += delta; state.expanded.clear(); await render(); }
  async function render(prefetched) {
    if (!state.dataset) { cards?.render(); return; }
    const token = ++state.revision;
    const result = prefetched || await state.client.request('locationPage', { datasetId: state.dataset.id, page: state.page }, { channel: 'locations' });
    if (token !== state.revision) return;
    state.page = result.page;
    const excluded = state.analysis?.excluded || {}, totalExcluded = Object.values(excluded).reduce((sum, value) => sum + Number(value || 0), 0);
    $('multiLocationSummary').textContent = `符合 ${number(result.total)} 筆｜排除 ${number(totalExcluded)} 項無法比對資料`;
    $('multiLocationSummary').title = `缺少門號 ${number(excluded.missing_phone)}；無效時間 ${number(excluded.invalid_time)}；無效地址 ${number(excluded.invalid_address)}`;
    const notes = bridge.getNotes();
    $('multiLocationRows').innerHTML = result.rows.length ? result.rows.map(row => {
      const id = String(row.id), open = state.expanded.has(id);
      return `<tr class="multi-location-match-row"><td>${esc(row.start_at.replace('T',' '))} ～ ${esc(row.end_at.replace('T',' '))}</td><td>${esc(row.county)}</td><td>${esc(row.district)}</td><td><div class="multi-location-phone-list">${row.phones.map(phone => `<span class="phone-value">${esc(phone)}</span>`).join('')}</div><span class="muted">共 ${number(row.phone_count)} 個門號</span></td><td><div class="multi-location-note-list">${row.phones.map(phone => `<label><span>${esc(phone)}</span><input class="phone-note-input" data-phone-note="${esc(phone)}" value="${esc(notes[phone] || '')}" aria-label="備註(只存瀏覽器) ${esc(phone)}" /></label>`).join('')}</div></td><td><button type="button" class="multi-location-detail-toggle" data-multi-location-detail="${esc(id)}" aria-expanded="${open}" aria-controls="multiLocationDetail-${esc(id)}">${open ? '收合' : `展開（${number(row.occurrence_count)} 筆）`}</button></td></tr>${open ? `<tr class="multi-location-detail-row" id="multiLocationDetail-${esc(id)}" data-match-id="${esc(id)}"><td colspan="6">讀取原始明細…</td></tr>` : ''}`;
    }).join('') : '<tr><td colspan="6" class="muted">目前沒有不同門號在同一行政區 30 分鐘內出現的結果。</td></tr>';
    $('multiLocationPageSummary').textContent = `第 ${state.page} / ${Math.max(1, Math.ceil(result.total/result.pageSize))} 頁，共 ${number(result.total)} 筆`;
    $('multiLocationPrevPage').disabled = state.page <= 1; $('multiLocationNextPage').disabled = state.page * result.pageSize >= result.total;
    for (const id of state.expanded.keys()) await loadDetails(id, token);
  }
  function pager(result, kind) {
    return `<div class="call-pagination"><span>共 ${number(result.total)} ${kind === 'phones' ? '個門號' : '筆明細'} · 第 ${result.page} 頁</span><div class="inline-actions"><button class="secondary-button compact-action" data-location-kind="${kind}" data-location-page="${result.page - 1}" ${result.page <= 1 ? 'disabled' : ''}>上一頁</button><button class="secondary-button compact-action" data-location-kind="${kind}" data-location-page="${result.page + 1}" ${result.page * result.pageSize >= result.total ? 'disabled' : ''}>下一頁</button></div></div>`;
  }
  async function loadDetails(id, token = state.revision) {
    const pages = state.expanded.get(id); if (!pages) return;
    const payload = { datasetId: state.dataset.id, matchId: id };
    const details = await state.client.request('locationDetails', { ...payload, page: pages.details }, { channel: 'location-details' });
    const phones = await state.client.request('locationPhones', { ...payload, page: pages.phones }, { channel: 'location-details' });
    if (token !== state.revision || !state.expanded.has(id)) return;
    const node = $('multiLocationDetail-' + id); if (!node) return;
    node.firstElementChild.innerHTML = `<h3>完整門號</h3><div class="multi-location-phone-list">${phones.rows.map(row => `<span class="phone-value">${esc(row.phone)}</span>`).join('')}</div>${pager(phones,'phones')}<h3>原始明細</h3><div class="multi-location-source-wrap"><table class="multi-location-source-table"><thead><tr><th>來源位置</th><th>比對時間與來源</th><th>調閱門號</th><th>通話對象</th><th>基地台</th></tr></thead><tbody>${details.rows.map(row => `<tr><td>${esc(row.source_file)}<br>${esc(row.source_sheet)} / 第 ${esc(row.row_number)} 列</td><td>${esc(row.occurred_at)}<br>${esc(timeSource(row.time_source))}</td><td>${esc(row.target_phone)}</td><td>${esc(row.counterparty_phone)}</td><td>${(row.matched_stations || []).map(station => `${esc(station.role)}：${esc(station.cell_id)}｜${esc(station.address)}`).join('<br>')}</td></tr>`).join('')}</tbody></table></div>${pager(details,'details')}`;
  }
  root.PhoneMultiLocationUI = { init, render: () => render().catch(fail), importFiles };
})(globalThis);
