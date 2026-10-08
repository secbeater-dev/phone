(function (root) {
  'use strict';
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const number = value => Number(value || 0).toLocaleString();

  function create(container) {
    let client, datasetId, revision = 0, page = 1, search = '', timer;
    const expanded = new Map();
    container.innerHTML = '<div class="section-title"><div><h2>整批門號資訊</h2><p class="muted">列出這一批匯入的全部門號，每頁 20 張；不受目前選取門號或日期篩選影響。</p></div></div><label class="phone-card-search">搜尋門號或用戶名稱<input type="search" placeholder="輸入門號或姓名" /></label><p class="phone-card-status muted" role="status"></p><div class="phone-card-grid"></div><div class="call-pagination"><span class="phone-card-page-info muted"></span><div class="inline-actions"><button type="button" class="secondary-button compact-action" data-cards-page="prev">上一頁</button><button type="button" class="secondary-button compact-action" data-cards-page="next">下一頁</button></div></div>';
    const grid = container.querySelector('.phone-card-grid'), status = container.querySelector('.phone-card-status');
    const channel = 'cards-' + container.id;
    function fail(error) { if (error.name !== 'AbortError') status.textContent = error.message || '讀取門號資訊失敗。'; }
    function setDataset(service, id) {
      if (datasetId === id && client === service) return;
      ++revision; clearTimeout(timer); client?.cancel(channel);
      client = service; datasetId = id; page = 1; search = ''; expanded.clear();
      container.querySelector('input').value = ''; grid.innerHTML = ''; status.textContent = '';
    }
    async function render(prefetched) {
      const token = ++revision;
      if (!datasetId) {
        grid.innerHTML = ''; status.textContent = '尚未匯入門號資料。';
        container.querySelector('.phone-card-page-info').textContent = '共 0 個門號';
        container.querySelectorAll('[data-cards-page]').forEach(button => button.disabled = true);
        return;
      }
      status.textContent = '讀取門號資訊…';
      try {
        const result = prefetched || await client.request('phoneCards', { datasetId, search, page }, { channel });
        if (token !== revision) return;
        page = result.page; expanded.clear();
        grid.innerHTML = result.rows.map((row, index) => `<article class="phone-card"><div class="phone-card-heading"><h3>${esc(row.phone || '未辨識門號')}</h3><span class="phone-card-carrier">${esc((row.carriers || []).join('、') || '業者未提供')}</span></div><p class="phone-card-name">${esc((row.names || []).join('、') || '姓名未提供')}</p><p class="muted">通聯 ${number(row.call_count)} 筆 · 網路 ${number(row.data_count)} 筆</p><button type="button" class="secondary-button compact-action" data-card-toggle="${esc(row.phone)}" aria-expanded="false" aria-controls="${container.id}-card-${index}">展開用戶與來源資訊</button><div id="${container.id}-card-${index}" class="phone-card-details" hidden></div></article>`).join('');
        status.textContent = result.total ? '' : '沒有符合的門號。';
        container.querySelector('.phone-card-page-info').textContent = `共 ${number(result.total)} 個門號 · 第 ${page} / ${Math.max(1, Math.ceil(result.total / result.pageSize))} 頁`;
        container.querySelector('[data-cards-page="prev"]').disabled = page <= 1;
        container.querySelector('[data-cards-page="next"]').disabled = page * result.pageSize >= result.total;
      } catch (error) { if (token === revision) fail(error); }
    }
    async function detail(phone, node, detailPage = 1) {
      const token = revision, id = datasetId;
      node.textContent = '讀取詳細資訊…';
      try {
        const result = await client.request('phoneCardDetails', { datasetId: id, phone, page: detailPage }, { channel });
        if (token !== revision || !node.isConnected) return;
        expanded.set(phone, result.rows);
        node.innerHTML = result.rows.length ? `<dl class="detail-list">${result.rows.map((row, index) => `<dt>${esc(row.name)}</dt><dd>${esc(row.value)}<div class="phone-card-sources muted">${sourcesText(row.sources || [])}</div>${row.source_count > (row.sources || []).length ? `<button type="button" class="secondary-button compact-action" data-card-sources="${index}">查看全部 ${number(row.source_count)} 個來源</button><div class="phone-card-source-page"></div>` : ''}</dd>`).join('')}</dl>` : '<p class="muted">來源未提供用戶資訊。</p>';
        if (result.total > result.pageSize) node.insertAdjacentHTML('beforeend', `<p>用戶欄位共 ${number(result.total)} 項 · 第 ${detailPage} 頁 <button class="secondary-button compact-action" data-card-detail-page="${detailPage - 1}" ${detailPage <= 1 ? 'disabled' : ''}>上一頁</button> <button class="secondary-button compact-action" data-card-detail-page="${detailPage + 1}" ${detailPage * result.pageSize >= result.total ? 'disabled' : ''}>下一頁</button></p>`);
      } catch (error) { if (token === revision) node.textContent = error.name === 'AbortError' ? '' : error.message; }
    }
    function sourcesText(rows) { return rows.map(source => esc([source.source_file, source.source_sheet, source.row_number ? '第 ' + source.row_number + ' 列' : ''].filter(Boolean).join(' / '))).join('<br>'); }
    async function loadSources(button, sourcePage = 1) {
      const card = button.closest('.phone-card'), phone = card.querySelector('[data-card-toggle]').dataset.cardToggle;
      const index = Number(button.dataset.cardSources ?? button.closest('dd').querySelector('[data-card-sources]').dataset.cardSources);
      const field = expanded.get(phone)?.[index]; if (!field) return;
      const token = revision, node = button.closest('dd').querySelector('.phone-card-source-page');
      try {
        const result = await client.request('phoneCardSources', { datasetId, phone, name: field.name, value: field.value, page: sourcePage }, { channel });
        if (token !== revision || !node.isConnected) return;
        node.innerHTML = `<div class="phone-card-sources muted">${sourcesText(result.rows)}</div><p>第 ${result.page} 頁 · 共 ${number(result.total)} 個來源 <button type="button" class="secondary-button compact-action" data-card-source-page="${result.page - 1}" ${result.page <= 1 ? 'disabled' : ''}>上一頁</button> <button type="button" class="secondary-button compact-action" data-card-source-page="${result.page + 1}" ${result.page * result.pageSize >= result.total ? 'disabled' : ''}>下一頁</button></p>`;
      } catch (error) { if (token === revision && error.name !== 'AbortError') node.textContent = error.message; }
    }
    container.querySelector('input').addEventListener('input', event => {
      search = event.target.value; page = 1; ++revision; clearTimeout(timer); client?.cancel(channel);
      timer = setTimeout(() => render().catch(fail), 180);
    });
    container.addEventListener('click', event => {
      const pager = event.target.closest('[data-cards-page]');
      if (pager && !pager.disabled) { page += pager.dataset.cardsPage === 'next' ? 1 : -1; render().catch(fail); }
      const toggle = event.target.closest('[data-card-toggle]');
      if (toggle) {
        const node = document.getElementById(toggle.getAttribute('aria-controls'));
        node.hidden = !node.hidden; toggle.setAttribute('aria-expanded', String(!node.hidden));
        toggle.textContent = node.hidden ? '展開用戶與來源資訊' : '收合資訊';
        if (!node.hidden) detail(toggle.dataset.cardToggle, node).catch(fail);
      }
      const next = event.target.closest('[data-card-detail-page]');
      if (next && !next.disabled) { const card = next.closest('.phone-card'); detail(card.querySelector('[data-card-toggle]').dataset.cardToggle, card.querySelector('.phone-card-details'), Number(next.dataset.cardDetailPage)).catch(fail); }
      const sources = event.target.closest('[data-card-sources]');
      if (sources) loadSources(sources).catch(fail);
      const sourcePage = event.target.closest('[data-card-source-page]');
      if (sourcePage && !sourcePage.disabled) loadSources(sourcePage, Number(sourcePage.dataset.cardSourcePage)).catch(fail);
    });
    return { setDataset, render };
  }
  root.PhoneCardsUI = { create };
})(globalThis);
