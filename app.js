(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(root.XLSX || require("./vendor/xlsx.full.min.js"));
  } else {
    root.PhoneWorkbench = factory(root.XLSX);
  }
})(typeof globalThis !== "undefined" ? globalThis : window, function (XLSX) {
  "use strict";

  const STORAGE_KEYS = {
    localSettings: "phone-workbench-local-settings-v1",
    statsRankMode: "phone-workbench-stats-rank-mode",
    hourSelection: "phone-workbench-hour-selection",
    phoneNotes: "phone-workbench-phone-notes-v1",
    callColumnWidths: "phone-workbench-call-column-widths-v1",
    theme: "phone-workbench-theme",
    sidebarCollapsed: "phone-workbench-sidebar-collapsed",
  };
  const SUPPORT_PASSWORD_SHA256 = "44d987773df84d3bdb849615d9f4d37567d46759e2fbeff2809ffe403af04aef";
  const LOCAL_EXPORT_VERSION = "phone-workbench-local-settings-v1";
  const HOUR_LABELS = Array.from({ length: 24 }, (_, hour) => `${String(hour).padStart(2, "0")}-${String(hour + 1).padStart(2, "0")}`);
  const CALL_COLUMNS = [
    { key: "occurred_at", width: 168, min: 120 },
    { key: "call_type", width: 110, min: 82 },
    { key: "target_phone", width: 150, min: 120 },
    { key: "target_note", width: 210, min: 150 },
    { key: "counterparty_phone", width: 150, min: 120 },
    { key: "counterparty_note", width: 210, min: 150 },
    { key: "duration_seconds", width: 90, min: 72 },
    { key: "imei", width: 160, min: 120 },
    { key: "note", width: 220, min: 150 },
  ];
  const VIEW_TITLES = {
    calls: "通聯列表",
    profile: "用戶資料",
    stats: "電話統計",
    hours: "時間分布圖",
    submission: "電話投單",
    export: "資料匯出",
  };

  const FET_HEADERS = ["通聯起始時間", "通聯時間(秒)", "通聯結束時間", "起始基地台編號", "離開基地台編號", "起始基地台地址", "離開基地台地址"];
  const TWM_HEADERS = ["通話類別", "目標電話", "對象電話", "始話日期時間", "通話時間(秒)", "基地台編號1/位置1"];
  const TWM_DATA_HEADERS = ["進入基地台時間", "基地台停留時間", "離開基地台時間", "離開基地台編號", "離開基地台地址", "上傳使用量(Byte)", "下載使用量(Byte)", "全部使用量(Byte)", "IMEI", "備註"];
  const CONVERTED_USER_HEADERS = ["用戶名稱", "用戶回應-用戶編號/帳號", "查詢項目", "身份識別碼", "戶籍地址", "帳寄地址"];
  const CONVERTED_CALL_HEADERS = ["通話類別", "始話時間", "調閱門號", "對象門號", "通話期間"];
  const CONVERTED_DATA_HEADERS = ["查詢項目", "開始時間", "連線期間", "開始基地台編號", "開始基地台", "結束基地台編號", "結束基地台"];
  const COMPACT_HEADERS = ["時間", "通話秒數", "調閱號碼", "IMEI", "通話類別", "基地台", "迄基地台"];
  const FET_WEB_HEADERS = ["啟始時間", "通聯時間(秒)", "結束時間", "上行用量", "下行用量", "全部用量", "基地台 ID", "最終基地台 ID", "基地台位址", "最終基地台位址", "上網IPv4", "上網IPv6", "IMEI", "IMSI", "MSISDN", "備註"];
  const TWM_CALL_PIG_HEADERS = ["通話類別", "始話時間", "調閱門號", "對象門號", "通話期間", "開始基地台編號", "開始基地台"];

  const state = {
    view: "hours",
    cases: [],
    currentWorkspace: null,
    callRecords: [],
    callSort: { column: "occurred_at", direction: "asc" },
    phoneStatsRankMode: "count",
    hourSelection: new Set(Array.from({ length: 24 }, (_, index) => index)),
    appliedHourSelection: new Set(Array.from({ length: 24 }, (_, index) => index)),
    expandedHotspotAddress: "",
    phoneNotes: {},
    callColumnWidths: {},
    theme: "light",
    sidebarCollapsed: false,
    activeColumnResize: null,
  };

  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", init);
  }

  function init() {
    restoreSettings();
    syncTheme();
    syncSidebarCollapsed();
    bindEvents();
    setSubmissionDefaults();
    renderHourTiles();
    renderAllViews();
    initCallColumnResize();
    maybeShowUsageNotice();
  }

  function $(id) {
    return document.getElementById(id);
  }

  function bindEvents() {
    document.querySelectorAll(".nav-item").forEach((button) => {
      button.addEventListener("click", () => setView(button.dataset.view));
    });
    $("importButton")?.addEventListener("click", () => $("fileInput")?.click());
    $("fileInput")?.addEventListener("change", handleFileImport);
    $("recordSearch")?.addEventListener("input", renderTwoWayCalls);
    document.querySelector("#callsView thead")?.addEventListener("click", (event) => {
      if (event.target.closest(".call-column-resizer")) return;
      const button = event.target.closest("[data-call-sort]");
      if (button) toggleCallSort(button.dataset.callSort);
    });
    $("callRows")?.addEventListener("input", (event) => {
      const input = event.target.closest("[data-phone-note]");
      if (input) updatePhoneNote(input.dataset.phoneNote, input.value, input);
    });
    document.querySelectorAll("[data-stats-rank-mode]").forEach((button) => {
      button.addEventListener("click", () => setPhoneStatsRankMode(button.dataset.statsRankMode));
    });
    $("statsContent")?.addEventListener("input", (event) => {
      const input = event.target.closest("[data-phone-note]");
      if (input) updatePhoneNote(input.dataset.phoneNote, input.value, input);
    });
    $("hourHotspotSearch")?.addEventListener("input", () => renderHourHotspots(filteredHourRecords()));
    $("hourSelectAll")?.addEventListener("click", () => {
      state.hourSelection = new Set(Array.from({ length: 24 }, (_, index) => index));
      renderHourTiles();
    });
    $("hourClearAll")?.addEventListener("click", () => {
      state.hourSelection = new Set();
      renderHourTiles();
    });
    $("hourApplyButton")?.addEventListener("click", applyHourSelection);
    $("submissionPreviewButton")?.addEventListener("click", renderSubmissionPreview);
    $("submissionDownloadButton")?.addEventListener("click", downloadSubmissionCsv);
    ["submissionPhonesInput", "submissionStartInput", "submissionEndInput"].forEach((id) => $(id)?.addEventListener("input", renderSubmissionPreview));
    $("exportWorkspaceButton")?.addEventListener("click", exportWorkspaceJson);
    $("importWorkspaceInput")?.addEventListener("change", importWorkspaceJson);
    $("exportLocalSettingsButton")?.addEventListener("click", exportLocalSettings);
    $("importLocalSettingsInput")?.addEventListener("change", importLocalSettings);
    $("sidebarCollapseButton")?.addEventListener("click", toggleSidebarCollapsed);
    $("themeToggleButton")?.addEventListener("click", toggleTheme);
    $("noticeDismissButton")?.addEventListener("click", hideUsageNotice);
    $("supportTypesButton")?.addEventListener("click", verifySupportPassword);
    $("supportPasswordInput")?.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        verifySupportPassword();
      }
    });
    $("oneClickUpdateButton")?.addEventListener("click", runOneClickUpdate);
    $("usageNoticeModal")?.addEventListener("click", (event) => {
      if (event.target === $("usageNoticeModal")) hideUsageNotice();
    });
    document.addEventListener("pointermove", handleCallColumnResizeMove);
    document.addEventListener("pointerup", handleCallColumnResizeEnd);
  }

  function syncTheme() {
    if (typeof document === "undefined") return;
    document.body.dataset.theme = state.theme;
    const button = $("themeToggleButton");
    if (button) {
      const dark = state.theme === "dark";
      button.setAttribute("aria-pressed", dark ? "true" : "false");
      button.title = dark ? "切換淺色模式" : "深色模式";
      const label = button.querySelector(".tool-label");
      if (label) label.textContent = dark ? "淺色模式" : "深色模式";
    }
  }

  function toggleTheme() {
    state.theme = state.theme === "dark" ? "light" : "dark";
    localStorage.setItem(STORAGE_KEYS.theme, state.theme);
    syncTheme();
  }

  function syncSidebarCollapsed() {
    if (typeof document === "undefined") return;
    document.body.classList.toggle("sidebar-collapsed", state.sidebarCollapsed);
    const button = $("sidebarCollapseButton");
    if (button) {
      button.setAttribute("aria-pressed", state.sidebarCollapsed ? "true" : "false");
      button.title = state.sidebarCollapsed ? "展開側邊欄" : "縮放側邊欄";
    }
  }

  function toggleSidebarCollapsed() {
    state.sidebarCollapsed = !state.sidebarCollapsed;
    localStorage.setItem(STORAGE_KEYS.sidebarCollapsed, state.sidebarCollapsed ? "true" : "false");
    syncSidebarCollapsed();
  }

  function maybeShowUsageNotice() {
    showUsageNotice();
  }

  function showUsageNotice() {
    const modal = $("usageNoticeModal");
    if (!modal) return;
    modal.hidden = false;
    $("noticeDismissButton")?.focus();
  }

  function hideUsageNotice() {
    const modal = $("usageNoticeModal");
    if (modal) modal.hidden = true;
  }

  async function verifySupportPassword() {
    const panel = $("supportTypesPanel");
    const button = $("supportTypesButton");
    const input = $("supportPasswordInput");
    const message = $("supportPasswordMessage");
    if (!panel) return;
    if (!panel.hidden) {
      panel.hidden = true;
      if (button) button.textContent = "顯示支援檔案類型";
      if (message) message.textContent = "";
      return;
    }
    const password = input?.value || "";
    const digest = await sha256Hex(password);
    if (digest === SUPPORT_PASSWORD_SHA256) {
      panel.hidden = false;
      if (button) button.textContent = "隱藏支援檔案類型";
      if (message) {
        message.classList.remove("danger-text");
        message.classList.add("success-text");
        message.textContent = "已解鎖支援檔案類型。";
      }
    } else if (message) {
      const supportHelpText = "請找 Telegram 管理員領取";
      panel.hidden = true;
      message.classList.remove("success-text");
      message.classList.add("danger-text");
      message.innerHTML = `${supportHelpText.replace("Telegram", '<a href="https://t.me/secbeater" target="_blank" rel="noopener noreferrer">Telegram</a>')}。`;
    }
  }

  async function sha256Hex(value) {
    if (!globalThis.crypto?.subtle || typeof TextEncoder === "undefined") return "";
    const bytes = new TextEncoder().encode(String(value || ""));
    const hash = await globalThis.crypto.subtle.digest("SHA-256", bytes);
    return Array.from(new Uint8Array(hash)).map((byte) => byte.toString(16).padStart(2, "0")).join("");
  }

  function runOneClickUpdate() {
    try {
      localStorage.clear();
      sessionStorage.clear();
    } catch (_error) {
      // Some privacy modes can block storage clearing; reload still helps fetch fresh assets.
    }
    const url = new URL(window.location.href);
    url.searchParams.set("refresh", String(Date.now()));
    window.location.replace(url.toString());
  }

  function setView(view) {
    if (!VIEW_TITLES[view]) view = "hours";
    state.view = view;
    document.querySelectorAll(".view").forEach((section) => section.classList.remove("active-view"));
    $(`${view}View`)?.classList.add("active-view");
    document.querySelectorAll(".nav-item").forEach((button) => button.classList.toggle("active", button.dataset.view === view));
    $("pageTitle").textContent = VIEW_TITLES[view];
    if (view === "calls") renderTwoWayCalls();
    if (view === "profile") renderProfileView();
    if (view === "stats") renderStatsView();
    if (view === "hours") renderHoursView();
  }

  function restoreSettings() {
    state.phoneStatsRankMode = localStorage.getItem(STORAGE_KEYS.statsRankMode) === "seconds" ? "seconds" : "count";
    state.phoneNotes = normalizePhoneNotes(readJson(localStorage.getItem(STORAGE_KEYS.phoneNotes), {}));
    state.callColumnWidths = normalizeColumnWidths(readJson(localStorage.getItem(STORAGE_KEYS.callColumnWidths), {}));
    state.theme = localStorage.getItem(STORAGE_KEYS.theme) === "dark" ? "dark" : "light";
    state.sidebarCollapsed = localStorage.getItem(STORAGE_KEYS.sidebarCollapsed) === "true";
    const storedHours = readJson(localStorage.getItem(STORAGE_KEYS.hourSelection), null);
    if (Array.isArray(storedHours)) {
      state.hourSelection = new Set(storedHours.map(Number).filter((hour) => hour >= 0 && hour <= 23));
      state.appliedHourSelection = new Set(state.hourSelection);
    }
    syncPhoneStatsRankModeUI();
  }

  function setPhoneStatsRankMode(mode) {
    state.phoneStatsRankMode = mode === "seconds" ? "seconds" : "count";
    localStorage.setItem(STORAGE_KEYS.statsRankMode, state.phoneStatsRankMode);
    syncPhoneStatsRankModeUI();
    renderStatsView();
  }

  function syncPhoneStatsRankModeUI() {
    if (typeof document === "undefined") return;
    document.querySelectorAll("[data-stats-rank-mode]").forEach((button) => {
      const active = button.dataset.statsRankMode === state.phoneStatsRankMode;
      button.classList.toggle("active", active);
      button.setAttribute("aria-pressed", active ? "true" : "false");
    });
  }

  async function handleFileImport(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;
    $("importStatus").textContent = `匯入 ${files.length} 個檔案中...`;
    $("importResults").innerHTML = "";
    for (const file of files) {
      try {
        const content = await readFileArrayBuffer(file);
        const workspace = parseImportFile(file.name, content);
        applyWorkspace(workspace);
        appendImportResult(workspace.case);
      } catch (error) {
        appendImportError(file.name, error);
      }
    }
    $("importStatus").textContent = "匯入完成。";
    $("fileInput").value = "";
  }

  function readFileArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error || new Error("檔案讀取失敗"));
      reader.readAsArrayBuffer(file);
    });
  }

  function appendImportResult(item) {
    const div = document.createElement("div");
    div.className = "import-result-item";
    div.innerHTML = `<strong>${escapeHtml(item.source_file || "匯入檔案")}</strong><span>${escapeHtml(item.source_format || "")}</span><span>${Number(item.total_records || 0).toLocaleString()} 筆</span>`;
    $("importResults").appendChild(div);
  }

  function appendImportError(name, error) {
    const div = document.createElement("div");
    div.className = "import-result-item danger-text";
    div.textContent = `${name} 匯入失敗：${error.message}`;
    $("importResults").appendChild(div);
  }

  function applyWorkspace(workspace) {
    state.currentWorkspace = workspace || null;
    state.cases.push(workspace?.case || {});
    state.callRecords = workspace?.records || [];
    state.expandedHotspotAddress = "";
    prefillSubmissionPhones(state.callRecords);
    renderAllViews();
  }

  function renderAllViews() {
    renderTwoWayCalls();
    renderProfileView();
    renderStatsView();
    renderHoursView();
    renderSubmissionPreview();
  }

  function renderProfileSummaryCards() {
    const summary = state.currentWorkspace?.case?.summary || {};
    const target = $("profileSummaryCards");
    if (!target) return;
    target.innerHTML = [
      metricCard("通聯筆數", summary.records || 0),
      metricCard("目標電話", summary.target_phones || 0),
      metricCard("對象電話", summary.counterparty_phones || 0),
      metricCard("第一筆時間", summary.first_seen || "-"),
      metricCard("最後時間", summary.last_seen || "-"),
      metricCard("總秒數", summary.total_duration_seconds || 0),
    ].join("");
  }

  function metricCard(label, value) {
    return `<div class="metric-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(String(value))}</strong></div>`;
  }

  function callPhoneLink(phone) {
    const normalized = normalizePhoneText(phone);
    if (!normalized) return "";
    return `<a class="phone-link" href="${tellowsUrl(normalized)}" title="點我查詢" target="_blank" rel="noopener noreferrer">${escapeHtml(normalized)}</a>`;
  }

  function callPhoneNoteInput(phone, label) {
    const normalized = normalizePhoneText(phone);
    if (!normalized) return "";
    return `<input class="phone-note-input call-note-input" data-phone-note="${escapeHtml(normalized)}" value="${escapeHtml(phoneNote(normalized))}" aria-label="${escapeHtml(label)} ${escapeHtml(normalized)}" />`;
  }

  function renderTwoWayCalls() {
    const query = ($("recordSearch")?.value || "").trim().toLowerCase();
    const rows = sortedCallRecords().filter((record) => {
      if (!query) return true;
      return [
        record.occurred_at,
        record.call_type,
        record.target_phone,
        record.counterparty_phone,
        phoneNote(record.target_phone),
        phoneNote(record.counterparty_phone),
        record.imei,
        record.external_ip,
        record.note,
      ]
        .some((value) => String(value || "").toLowerCase().includes(query));
    });
    $("callRows").innerHTML = rows.length
      ? rows.slice(0, 5000).map((record) => `<tr>
          <td>${escapeHtml(record.occurred_at || "")}</td>
          <td>${escapeHtml(record.call_type || "")}</td>
          <td>${callPhoneLink(record.target_phone)}</td>
          <td>${callPhoneNoteInput(record.target_phone, "目標電話備註(只存瀏覽器)")}</td>
          <td>${callPhoneLink(record.counterparty_phone)}</td>
          <td>${callPhoneNoteInput(record.counterparty_phone, "對象電話備註(只存瀏覽器)")}</td>
          <td>${escapeHtml(String(record.duration_seconds ?? ""))}</td>
          <td>${escapeHtml(record.imei || "")}</td>
          <td>${escapeHtml(record.external_ip || record.note || "")}</td>
        </tr>`).join("")
      : `<tr><td colspan="9">尚未匯入資料。</td></tr>`;
    applyCallColumnWidths();
  }

  function sortedCallRecords() {
    const { column, direction } = state.callSort;
    return [...state.callRecords].sort((a, b) => {
      const av = sortValue(a, column);
      const bv = sortValue(b, column);
      const compare = typeof av === "number" && typeof bv === "number"
        ? av - bv
        : String(av || "").localeCompare(String(bv || ""), "zh-Hant", { numeric: true });
      return direction === "asc" ? compare : -compare;
    });
  }

  function sortValue(record, column) {
    if (column === "duration_seconds") return Number(record.duration_seconds || 0);
    return record[column] || "";
  }

  function toggleCallSort(column) {
    if (state.callSort.column === column) {
      state.callSort.direction = state.callSort.direction === "asc" ? "desc" : "asc";
    } else {
      state.callSort = { column, direction: "asc" };
    }
    renderTwoWayCalls();
  }

  function initCallColumnResize() {
    const headers = Array.from(document.querySelectorAll("#callsView th"));
    headers.forEach((th, index) => {
      const column = CALL_COLUMNS[index];
      if (!column) return;
      th.dataset.callColumn = column.key;
      if (!th.querySelector(".call-column-resizer")) {
        const handle = document.createElement("span");
        handle.className = "call-column-resizer";
        handle.setAttribute("role", "separator");
        handle.setAttribute("aria-orientation", "vertical");
        handle.setAttribute("aria-label", "調整欄寬");
        handle.dataset.callColumn = column.key;
        handle.addEventListener("pointerdown", (event) => {
          event.preventDefault();
          event.stopPropagation();
          const width = columnWidth(column);
          state.activeColumnResize = {
            key: column.key,
            startX: event.clientX,
            startWidth: width,
            min: column.min,
          };
          handle.setPointerCapture?.(event.pointerId);
          document.body.classList.add("resizing-call-column");
        });
        th.appendChild(handle);
      }
    });
    applyCallColumnWidths();
  }

  function handleCallColumnResizeMove(event) {
    if (!state.activeColumnResize) return;
    resizeCallColumn(event.clientX);
  }

  function handleCallColumnResizeEnd() {
    if (!state.activeColumnResize) return;
    persistCallColumnWidths();
    state.activeColumnResize = null;
    document.body.classList.remove("resizing-call-column");
  }

  function resizeCallColumn(clientX) {
    const resize = state.activeColumnResize;
    if (!resize) return;
    const width = Math.max(resize.min, Math.round(resize.startWidth + clientX - resize.startX));
    state.callColumnWidths[resize.key] = width;
    applyCallColumnWidths();
  }

  function applyCallColumnWidths() {
    const table = document.querySelector("#callsView table");
    if (!table) return;
    const widths = CALL_COLUMNS.map((column) => columnWidth(column));
    table.style.minWidth = `${widths.reduce((sum, width) => sum + width, 0)}px`;
    document.querySelectorAll("#callsView tr").forEach((row) => {
      Array.from(row.children).forEach((cell, index) => {
        const width = widths[index];
        if (!width) return;
        cell.style.width = `${width}px`;
        cell.style.minWidth = `${width}px`;
        cell.style.maxWidth = `${width}px`;
      });
    });
  }

  function columnWidth(column) {
    const stored = Number(state.callColumnWidths[column.key]);
    return Number.isFinite(stored) && stored >= column.min ? stored : column.width;
  }

  function persistCallColumnWidths() {
    if (typeof localStorage === "undefined") return;
    localStorage.setItem(STORAGE_KEYS.callColumnWidths, JSON.stringify(state.callColumnWidths));
  }

  function normalizeColumnWidths(widths) {
    const normalized = {};
    Object.entries(widths || {}).forEach(([key, value]) => {
      const column = CALL_COLUMNS.find((item) => item.key === key);
      const width = Number(value);
      if (column && Number.isFinite(width) && width >= column.min) normalized[key] = Math.round(width);
    });
    return normalized;
  }

  function renderProfileView() {
    renderProfileSummaryCards();
    const subject = state.currentWorkspace?.case?.subject || {};
    const entries = Object.entries(subject);
    $("profileContent").innerHTML = entries.length
      ? entries.map(([key, value]) => `<dt>${escapeHtml(key)}</dt><dd>${escapeHtml(String(value || ""))}</dd>`).join("")
      : `<dt>狀態</dt><dd>尚未匯入資料。</dd>`;
    renderProfileImeiList();
  }

  function renderProfileImeiList() {
    const target = $("profileImeiList");
    if (!target) return;
    const imeis = collectUniqueImeis(state.callRecords);
    target.innerHTML = imeis.length
      ? imeis.map((imei) => `<span class="imei-chip">${escapeHtml(imei)}</span>`).join("")
      : `<p class="muted">尚無 IMEI 資料。</p>`;
  }

  function collectUniqueImeis(records) {
    return Array.from(new Set((records || [])
      .map((record) => cellText(record?.imei))
      .filter(Boolean)))
      .sort((a, b) => a.localeCompare(b, "zh-Hant", { numeric: true }));
  }

  function renderStatsView() {
    const stats = computePhoneStats(state.callRecords, state.phoneStatsRankMode);
    $("statsContent").innerHTML = [
      statsCard("來電排行", stats.inboundRows),
      statsCard("去電排行", stats.outboundRows),
      statsCard("完整排行", stats.totalRows),
    ].join("");
  }

  function statsCard(title, rows) {
    return `<section class="stats-card"><h3>${escapeHtml(title)}</h3>${
      rows.length
        ? `<div class="stats-table" role="table">
            <div class="stats-table-row stats-table-head" role="row"><span>#</span><span>電話</span><span>備註(只存瀏覽器)</span><span>次數</span><span>秒數</span></div>
            ${rows.slice(0, 20).map((row, index) => `<div class="stats-table-row" role="row">
              <span>${index + 1}</span>
              <span><a class="phone-link" href="${tellowsUrl(row.phone)}" title="點我查詢" target="_blank" rel="noopener noreferrer">${escapeHtml(row.phone)}</a></span>
              <span><input class="phone-note-input" data-phone-note="${escapeHtml(row.phone)}" value="${escapeHtml(phoneNote(row.phone))}" aria-label="備註(只存瀏覽器) ${escapeHtml(row.phone)}" /></span>
              <span>${row.count}</span>
              <span>${row.seconds}</span>
            </div>`).join("")}
          </div>`
        : `<p class="muted">尚無資料</p>`
    }</section>`;
  }

  function phoneNote(phone) {
    return state.phoneNotes[normalizePhoneText(phone)] || "";
  }

  function updatePhoneNote(phone, note, sourceInput) {
    const key = normalizePhoneText(phone);
    if (!key) return;
    const value = String(note || "");
    if (value) state.phoneNotes[key] = value;
    else delete state.phoneNotes[key];
    persistPhoneNotes();
    syncPhoneNoteInputs(key, value, sourceInput);
  }

  function persistPhoneNotes() {
    if (typeof localStorage === "undefined") return;
    localStorage.setItem(STORAGE_KEYS.phoneNotes, JSON.stringify(state.phoneNotes));
  }

  function syncPhoneNoteInputs(phone, value, sourceInput) {
    if (typeof document === "undefined") return;
    document.querySelectorAll("[data-phone-note]").forEach((input) => {
      if (input === sourceInput) return;
      if (normalizePhoneText(input.dataset.phoneNote) === phone) input.value = value;
    });
  }

  function normalizePhoneNotes(notes) {
    const normalized = {};
    Object.entries(notes || {}).forEach(([phone, note]) => {
      const key = normalizePhoneText(phone);
      if (key && note) normalized[key] = String(note);
    });
    return normalized;
  }

  function tellowsUrl(phone) {
    const normalized = normalizePhoneText(phone);
    if (!normalized) return "https://www.tellows.tw/";
    const international = normalized.startsWith("0") ? `886${normalized.slice(1)}` : normalized;
    return `https://www.tellows.tw/num/%2B${international}`;
  }

  function renderHourTiles() {
    if (!$("hourGrid")) return;
    $("hourGrid").innerHTML = HOUR_LABELS.map((label, hour) => {
      const active = state.hourSelection.has(hour);
      const tone = isNightHour(hour) ? "night-hour" : "day-hour";
      return `<button class="hour-tile ${tone} ${active ? "active" : "inactive"}" data-hour="${hour}" type="button" aria-label="${label}">${hourButtonLabel(hour)}</button>`;
    }).join("");
    $("hourGrid").querySelectorAll("[data-hour]").forEach((button) => {
      button.addEventListener("click", () => {
        const hour = Number(button.dataset.hour);
        if (state.hourSelection.has(hour)) state.hourSelection.delete(hour);
        else state.hourSelection.add(hour);
        renderHourTiles();
      });
    });
  }

  function hourButtonLabel(hour) {
    return String(hour).padStart(2, "0");
  }

  function isNightHour(hour) {
    return hour < 6 || hour >= 18;
  }

  function formatPercent(value) {
    if (!Number.isFinite(value) || value <= 0) return "0%";
    if (value >= 10) return `${Math.round(value)}%`;
    return `${value.toFixed(1)}%`;
  }

  function applyHourSelection() {
    state.appliedHourSelection = new Set(state.hourSelection);
    localStorage.setItem(STORAGE_KEYS.hourSelection, JSON.stringify(Array.from(state.hourSelection).sort((a, b) => a - b)));
    renderHoursView();
  }

  function renderHoursView() {
    const records = filteredHourRecords();
    const buckets = computeHourBuckets(records);
    const max = Math.max(1, ...buckets.map((item) => item.count));
    const total = buckets.reduce((sum, item) => sum + item.count, 0);
    $("hourChart").className = "hour-chart hour-chart-vertical";
    $("hourChart").innerHTML = buckets.map((item) => {
      const height = item.count ? (item.count / max) * 100 : 0;
      const percent = total ? (item.count / total) * 100 : 0;
      return `<div class="hour-column" style="--bar-height:${height}%">
        <div class="hour-bar-count">${item.count}</div>
        <div class="hour-bar-frame">
          <div class="hour-bar-fill" style="height:${height}%"></div>
          <span class="hour-bar-percent">${formatPercent(percent)}</span>
        </div>
        <div class="hour-label-full">${item.label}</div>
      </div>`;
    }).join("");
    renderHourHotspots(records);
  }

  function filteredHourRecords() {
    const allowed = state.appliedHourSelection;
    return state.callRecords.filter((record) => {
      const hour = hourFromIso(record.occurred_at);
      return hour >= 0 && allowed.has(hour);
    });
  }

  function renderHourHotspots(records) {
    const query = ($("hourHotspotSearch")?.value || "").trim().toLowerCase();
    const hotspots = computeAddressHotspots(records, state.currentWorkspace?.base_stations || []);
    const filtered = query
      ? hotspots.filter((item) => [item.address, item.first_seen, item.last_seen, ...item.times].some((value) => String(value || "").toLowerCase().includes(query)))
      : hotspots;
    $("hourHotspotContent").innerHTML = filtered.length
      ? filtered.slice(0, 20).map((item) => hotspotItemHtml(item)).join("")
      : `<p class="muted">尚無符合資料。</p>`;
    $("hourHotspotContent").querySelectorAll("[data-hotspot-address]").forEach((button) => {
      const toggle = () => {
        const key = button.dataset.hotspotAddress || "";
        state.expandedHotspotAddress = state.expandedHotspotAddress === key ? "" : key;
        renderHourHotspots(filteredHourRecords());
      };
      button.addEventListener("click", toggle);
      button.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          toggle();
        }
      });
    });
  }

  function hotspotItemHtml(item) {
    const expanded = state.expandedHotspotAddress === item.key;
    const times = expanded
      ? `<ol class="hotspot-times">${item.times.map((time) => `<li>${escapeHtml(time)}</li>`).join("")}</ol>`
      : "";
    return `<div class="hotspot-item ${expanded ? "expanded" : ""}" data-hotspot-address="${escapeHtml(item.key)}" role="button" tabindex="0">
      <strong>${escapeHtml(item.address)}</strong>
      <span>${item.count} 筆 / ${formatPercent(item.percent)}</span>
      <small>${escapeHtml(item.first_seen || "-")} 至 ${escapeHtml(item.last_seen || "-")}</small>
      ${times}
    </div>`;
  }

  function setSubmissionDefaults() {
    if (!$("submissionStartInput")) return;
    const now = new Date();
    const end = new Date(now.getTime() + 86400000);
    $("submissionStartInput").value = toLocalDatetimeValue(now);
    $("submissionEndInput").value = toLocalDatetimeValue(end);
  }

  function renderSubmissionPreview() {
    if (!$("submissionSummary")) return;
    const rows = normalizeSubmissionPhones($("submissionPhonesInput").value);
    $("submissionSummary").innerHTML = `<span>有效 ${rows.valid.length}</span><span>錯誤 ${rows.invalid.length}</span><span>重複 ${rows.duplicates.length}</span>`;
    $("submissionPreviewRows").innerHTML = [
      ...rows.valid.map((phone) => `<div class="submission-row"><span>${escapeHtml(phone)}</span><strong>有效</strong></div>`),
      ...rows.invalid.map((phone) => `<div class="submission-row danger-text"><span>${escapeHtml(phone)}</span><strong>錯誤</strong></div>`),
    ].join("");
  }

  function downloadSubmissionCsv() {
    const csv = buildSubmissionCsv({
      phones: $("submissionPhonesInput").value,
      start: $("submissionStartInput").value,
      end: $("submissionEndInput").value,
    });
    downloadText(`phone-submission-${new Date().toISOString().slice(0, 10)}.csv`, csv, "text/csv;charset=utf-8");
    $("submissionStatus").textContent = "已下載 CSV。";
  }

  function prefillSubmissionPhones(records) {
    const input = $("submissionPhonesInput");
    if (!input) return [];
    const phones = collectSubmissionPhones(records);
    input.value = phones.join("\n");
    const status = $("submissionStatus");
    if (status) status.textContent = phones.length ? `已自匯入資料填入 ${phones.length} 支不重複電話。` : "";
    return phones;
  }

  function collectSubmissionPhones(records) {
    const fields = [
      "target_phone",
      "counterparty_phone",
      "msisdn",
      "phone",
      "subject_phone",
      "subscriber_phone",
      "application_phone",
      "query_phone",
    ];
    const seen = new Set();
    const phones = [];
    (records || []).forEach((record) => {
      fields.forEach((field) => {
        const phone = normalizePhoneText(record?.[field]);
        if (phone && !seen.has(phone)) {
          seen.add(phone);
          phones.push(phone);
        }
      });
    });
    return phones;
  }

  function buildSubmissionCsv({ phones, start, end }) {
    const rows = normalizeSubmissionPhones(phones).valid;
    const header = ["phone", "start_at", "end_at"];
    return [header, ...rows.map((phone) => [phone, start || "", end || ""])].map(csvLine).join("\r\n");
  }

  function normalizeSubmissionPhones(text) {
    const valid = [];
    const invalid = [];
    const duplicates = [];
    const seen = new Set();
    String(text || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean).forEach((line) => {
      const phone = normalizePhoneText(line);
      if (!phone) invalid.push(line);
      else if (seen.has(phone)) duplicates.push(phone);
      else {
        seen.add(phone);
        valid.push(phone);
      }
    });
    return { valid, invalid, duplicates };
  }

  function exportWorkspaceJson() {
    if (!state.currentWorkspace) {
      $("exportMessage").textContent = "尚未匯入資料。";
      return;
    }
    downloadText(`phone-workspace-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(state.currentWorkspace, null, 2), "application/json;charset=utf-8");
    $("exportMessage").textContent = "已匯出目前資料。";
  }

  async function importWorkspaceJson(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const payload = JSON.parse(text);
      if (!payload || !Array.isArray(payload.records) || !payload.case) throw new Error("不是支援的資料 JSON");
      applyWorkspace(payload);
      $("exportMessage").textContent = "已匯入資料 JSON。";
    } catch (error) {
      $("exportMessage").textContent = `匯入失敗：${error.message}`;
    } finally {
      event.target.value = "";
    }
  }

  function exportLocalSettings() {
    const payload = {
      version: LOCAL_EXPORT_VERSION,
      exported_at: new Date().toISOString(),
      phoneStatsRankMode: state.phoneStatsRankMode,
      hourSelection: Array.from(state.hourSelection).sort((a, b) => a - b),
      phoneNotes: state.phoneNotes,
      callColumnWidths: state.callColumnWidths,
      theme: state.theme,
      sidebarCollapsed: state.sidebarCollapsed,
    };
    downloadText(`phone-settings-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(payload, null, 2), "application/json;charset=utf-8");
    $("exportMessage").textContent = "已匯出本機設定。";
  }

  async function importLocalSettings(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const payload = JSON.parse(await file.text());
      if (payload.phoneStatsRankMode) setPhoneStatsRankMode(payload.phoneStatsRankMode);
      if (Array.isArray(payload.hourSelection)) {
        state.hourSelection = new Set(payload.hourSelection.map(Number).filter((hour) => hour >= 0 && hour <= 23));
        state.appliedHourSelection = new Set(state.hourSelection);
        renderHourTiles();
        applyHourSelection();
      }
      if (payload.phoneNotes && typeof payload.phoneNotes === "object") {
        state.phoneNotes = normalizePhoneNotes(payload.phoneNotes);
        persistPhoneNotes();
        renderStatsView();
        renderTwoWayCalls();
      }
      if (payload.callColumnWidths && typeof payload.callColumnWidths === "object") {
        state.callColumnWidths = normalizeColumnWidths(payload.callColumnWidths);
        persistCallColumnWidths();
        applyCallColumnWidths();
      }
      if (payload.theme === "dark" || payload.theme === "light") {
        state.theme = payload.theme;
        localStorage.setItem(STORAGE_KEYS.theme, state.theme);
        syncTheme();
      }
      if (typeof payload.sidebarCollapsed === "boolean") {
        state.sidebarCollapsed = payload.sidebarCollapsed;
        localStorage.setItem(STORAGE_KEYS.sidebarCollapsed, state.sidebarCollapsed ? "true" : "false");
        syncSidebarCollapsed();
      }
      $("exportMessage").textContent = "已匯入本機設定。";
    } catch (error) {
      $("exportMessage").textContent = `匯入失敗：${error.message}`;
    } finally {
      event.target.value = "";
    }
  }

  function parseImportFile(fileName, inputBytes) {
    const bytes = toUint8Array(inputBytes);
    const lower = String(fileName || "").toLowerCase();
    if (lower.endsWith(".xml") || looksLikeXml(bytes)) {
      return makeWorkspace(parseXmlWorkbook(fileName, decodeUtf8(bytes)));
    }
    if (!XLSX) throw new Error("XLSX parser is not available");
    return makeWorkspace(parseXlsxWorkbook(fileName, bytes));
  }

  function parseXlsxWorkbook(fileName, bytes) {
    const workbook = XLSX.read(bytes, { type: "array", cellDates: false });
    const sheets = workbook.SheetNames.map((title) => ({
      title,
      rows: XLSX.utils.sheet_to_json(workbook.Sheets[title], { header: 1, defval: "", raw: false, blankrows: false })
        .map((values, index) => ({ rowNumber: index + 1, values: values.map(cellText) })),
    }));
    const converted = convertedMultiSheetHeaders(sheets);
    if (converted) return parseConvertedMultiSheet(fileName, sheets, converted);
    const twmCallPig = twmCallPigHeaders(sheets);
    if (twmCallPig) return parseTwmCallPigXlsx(fileName, sheets, twmCallPig);
    if (imeiLookupHeaders(sheets)) return parseImeiLookup(fileName, sheets);
    const compact = compactHeaders(sheets);
    if (compact) return parseTwmCompact(fileName, sheets, compact);
    const fetWeb = fetWebHeaders(sheets);
    if (fetWeb) return parseFetWeb(fileName, fetWeb);
    const sheet = sheets[0] || { title: "工作表1", rows: [] };
    if (!sheet.rows.length) throw new Error("檔案沒有可讀取資料列");
    const detected = detectHeader(sheet.rows);
    if (detected.sourceFormat === "fet_data_session") return parseFet(fileName, sheet.title, sheet.rows, detected);
    if (detected.sourceFormat === "taiwan_mobile_call") return parseTwm(fileName, sheet.title, sheet.rows, detected);
    if (detected.sourceFormat === "taiwan_mobile_data_session") return parseTwmDataSession(fileName, sheet.title, sheet.rows, detected);
    throw new Error("找不到支援的標題列");
  }

  function parseXmlWorkbook(fileName, xml) {
    if (isTwmCallXml(xml)) return parseTwmCallXml(fileName, xml);
    return parseTwmCspXml(fileName, xml);
  }

  function makeParsed({ fileName, carrier, sourceFormat, sheetName, headerRow, totalSourceRows, subject }) {
    return {
      file_name: fileName,
      carrier,
      source_format: sourceFormat,
      sheet_name: sheetName,
      header_row: headerRow,
      total_source_rows: totalSourceRows,
      subject: subject || {},
      records: [],
      base_stations: [],
      warnings: [],
    };
  }

  function makeWorkspace(parsed) {
    const stations = parsed.base_stations || [];
    const records = (parsed.records || []).map((record) => parsedRecordPayload(record));
    const summary = parsedSummary(records, stations);
    return {
      case: {
        source_file: parsed.file_name,
        carrier: parsed.carrier,
        source_format: parsed.source_format,
        sheet_name: parsed.sheet_name,
        header_row: parsed.header_row,
        total_source_rows: parsed.total_source_rows,
        total_records: records.length,
        subject: parsed.subject || {},
        summary,
        parse_warnings: parsed.warnings || [],
      },
      records,
      base_stations: stations,
      parse_warnings: parsed.warnings || [],
    };
  }

  function parsedRecordPayload(record) {
    const target = normalizePhoneText(record.target_phone);
    const counterparty = normalizePhoneText(record.counterparty_phone);
    return {
      row_number: record.row_number,
      occurred_at: record.occurred_at || "",
      ended_at: record.ended_at || "",
      duration_seconds: toInt(record.duration_seconds),
      call_type: record.call_type || "",
      direction: directionLabel(record.call_type),
      target_phone: target,
      counterparty_phone: counterparty,
      imei: record.imei || "",
      imsi: record.imsi || "",
      external_ip: record.external_ip || "",
      internal_ip: record.internal_ip || "",
      upload_bytes: toInt(record.upload_bytes),
      download_bytes: toInt(record.download_bytes),
      total_bytes: toInt(record.total_bytes),
      note: record.note || "",
      base_refs: record.base_refs || [],
    };
  }

  function parsedSummary(records, stations) {
    const targetPhones = new Set(records.map((record) => record.target_phone).filter(Boolean));
    const counterpartyPhones = new Set(records.map((record) => record.counterparty_phone).filter(Boolean));
    const occurred = records.map((record) => record.occurred_at).filter(Boolean).sort();
    const callCounts = new Map();
    records.forEach((record) => {
      const label = record.call_type || "未知";
      callCounts.set(label, (callCounts.get(label) || 0) + 1);
    });
    return {
      records: records.length,
      target_phones: targetPhones.size,
      counterparty_phones: counterpartyPhones.size,
      first_seen: occurred[0] || "",
      last_seen: occurred[occurred.length - 1] || "",
      total_duration_seconds: records.reduce((sum, record) => sum + Number(record.duration_seconds || 0), 0),
      total_bytes: records.reduce((sum, record) => sum + Number(record.total_bytes || 0), 0),
      station_count: stations.length,
      call_types: Array.from(callCounts, ([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value),
    };
  }

  function parseTwmCallXml(fileName, xml) {
    const subject = twmCallXmlSubject(xml);
    const parsed = makeParsed({
      fileName,
      carrier: firstXmlText(xml, "電信業者") || "台灣大哥大",
      sourceFormat: "taiwan_mobile_call_xml",
      sheetName: "XML",
      headerRow: 0,
      totalSourceRows: iterXmlBlocks(xml, "通聯資料").length,
      subject,
    });
    const stationMap = new Map();
    const targetFallback = normalizePhoneText(subject["申請號碼"] || subject["電話號碼"]);
    iterXmlBlocks(xml, "通聯資料").forEach((block, index) => {
      const data = xmlChildren(block);
      const occurredAt = normalizeDatetime(data["始話日期時間"] || data["始話時間"]);
      if (!occurredAt) return;
      const record = baseRecord({
        row_number: index + 1,
        call_type: data["通話類別"] || "",
        occurred_at: occurredAt,
        duration_seconds: toInt(data["通話時間"] || data["通話期間"]),
        target_phone: normalizePhoneText(data["目標電話"] || data["調閱門號"]) || targetFallback,
        counterparty_phone: normalizePhoneText(data["對象電話"] || data["對象門號"]),
        imei: data["IMEI"] || "",
        imsi: data["IMSI"] || "",
        external_ip: data["IP"] || "",
        note: data["備註"] || data["IP"] || "",
      });
      iterXmlBlocks(block, "基地台資訊").forEach((stationBlock) => {
        const stationData = xmlChildren(stationBlock);
        addStationRef(record, stationMap, "start", stationData["基地台編號"] || "", stationData["基地台位置"] || "");
      });
      parsed.records.push(record);
    });
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function twmCallXmlSubject(xml) {
    const subject = {};
    iterXmlBlocks(xml, "用戶基本資料").forEach((block) => {
      const data = xmlChildren(block);
      [
        ["用戶名稱", data["姓名"]],
        ["申請號碼", data["電話號碼"]],
        ["電話號碼", data["電話號碼"]],
        ["身份證字號", data["身份證號碼"]],
        ["生日", data["出生日期"]],
        ["帳寄地址", data["帳寄地址"]],
        ["戶籍地址", data["戶籍地址"]],
        ["申裝地址", data["申裝地址"]],
        ["電信業者", data["電信業者"]],
      ].forEach(([key, value]) => {
        const text = cellText(value);
        if (text && !subject[key]) subject[key] = text;
      });
    });
    const phone = firstXmlText(xml, "電話號碼");
    if (phone && !subject["申請號碼"]) subject["申請號碼"] = normalizePhoneText(phone);
    return subject;
  }

  function parseTwmCspXml(fileName, xml) {
    const customer = xmlChildren(firstXmlBlock(xml, "CUSTOMERINFO") || "");
    const subject = compactObject({
      "用戶名稱": customer["NAME"] || "",
      "申請號碼": normalizePhoneText(customer["MSISDN"] || ""),
      "身份證字號": customer["ID"] || "",
      "生日": normalizeDate(customer["BIRTHDAY"] || ""),
      "帳寄地址": customer["BILLINGADDRESS"] || "",
      "戶籍地址": customer["CUSTOMERADDRESS"] || "",
    });
    const parsed = makeParsed({
      fileName,
      carrier: customer["CSPNAME"] || "台灣大哥大",
      sourceFormat: "taiwan_mobile_csp_xml_data_session",
      sheetName: "XML",
      headerRow: 0,
      totalSourceRows: iterXmlBlocks(xml, "CELLINFO").length,
      subject,
    });
    const stationMap = new Map();
    const targetPhone = subject["申請號碼"] || "";
    iterXmlBlocks(xml, "CELLINFO").forEach((block, index) => {
      const data = xmlChildren(block);
      const occurredAt = normalizeDatetime(data["STARTDT"]);
      if (!occurredAt) return;
      const record = baseRecord({
        row_number: index + 1,
        call_type: "數據",
        occurred_at: occurredAt,
        ended_at: normalizeDatetime(data["ENDDT"], false),
        duration_seconds: toInt(data["DURATION"]),
        target_phone: normalizePhoneText(data["MSISDN"]) || targetPhone,
        imei: data["IMEI"] || "",
        imsi: data["IMSI"] || "",
        external_ip: data["INTERNETREALIP"] || data["INTERNETREALIPV6"] || data["NOTE"] || "",
        internal_ip: data["USERINTRAIP"] || "",
        upload_bytes: toInt(data["UPLOAD"]),
        download_bytes: toInt(data["DOWNLOAD"]),
        total_bytes: toInt(data["TOTAL"]),
        note: data["NOTE"] || "",
      });
      addStationRef(record, stationMap, "primary", data["CELLENDID"] || data["CELLID"] || "", data["CELLENDADDRESS"] || data["CELLADDRESS"] || "");
      parsed.records.push(record);
    });
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseTwmCallPigXlsx(fileName, sheets, header) {
    const subject = twmCallPigSubject(sheets);
    const carrier = subject._carrier || subject["電信業者"] || "台灣大哥大";
    delete subject._carrier;
    const parsed = makeParsed({ fileName, carrier, sourceFormat: "taiwan_mobile_call_pig_xlsx", sheetName: header.sheet.title, headerRow: header.rowNumber, totalSourceRows: header.sheet.rows.length, subject });
    const stationMap = new Map();
    const targetFallback = normalizePhoneText(subject["申請號碼"]);
    forRowsAfter(header, (rowNumber, data) => {
      const occurredAt = normalizeDatetime(data["始話日期時間"] || data["始話時間"]);
      if (!occurredAt) return;
      const record = baseRecord({
        row_number: rowNumber,
        call_type: data["通話類別"] || "",
        occurred_at: occurredAt,
        duration_seconds: toInt(data["通話時間"] || data["通話期間"]),
        target_phone: normalizePhoneText(data["目標電話"] || data["調閱門號"] || data["查詢項目"]) || targetFallback,
        counterparty_phone: normalizePhoneText(data["對象電話"] || data["對象門號"]),
        imei: data["IMEI"] || "",
        imsi: data["IMSI"] || "",
        external_ip: data["IP"] || "",
        note: data["備註"] || data["IP"] || "",
      });
      addStationRef(record, stationMap, "start", data["開始基地台編號"] || "", data["開始基地台"] || "");
      addStationRef(record, stationMap, "end", data["結束基地台編號"] || "", data["結束基地台"] || "");
      parsed.records.push(record);
    });
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseTwmCompact(fileName, sheets, header) {
    const subject = compactSubject(sheets);
    const carrier = subject._carrier || "台灣大哥大";
    delete subject._carrier;
    const parsed = makeParsed({ fileName, carrier, sourceFormat: "taiwan_mobile_compact_data_session_xlsx", sheetName: header.sheet.title, headerRow: header.rowNumber, totalSourceRows: sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0), subject });
    const stationMap = new Map();
    forRowsAfter(header, (rowNumber, data) => {
      const occurredAt = normalizeDatetime(data["時間"]);
      if (!occurredAt) return;
      const record = baseRecord({
        row_number: rowNumber,
        call_type: data["通話類別"] === "行動上網" ? "數據" : data["通話類別"] || "",
        occurred_at: occurredAt,
        duration_seconds: toInt(data["通話秒數"]),
        target_phone: normalizePhoneText(data["調閱號碼"]),
        counterparty_phone: normalizePhoneText(data["通話對象"]),
        imei: data["IMEI"] || "",
        note: data["備註"] || "",
      });
      addStationCompoundRef(record, stationMap, "primary", data["迄基地台"] || data["基地台"]);
      parsed.records.push(record);
    });
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseFetWeb(fileName, header) {
    const subject = fetWebSubject(header.sheet.rows, header.rowNumber);
    const parsed = makeParsed({ fileName, carrier: "遠傳電信", sourceFormat: "fet_web_data_session_xlsx", sheetName: header.sheet.title, headerRow: header.rowNumber, totalSourceRows: header.sheet.rows.length, subject });
    const stationMap = new Map();
    forRowsAfter(header, (rowNumber, data) => {
      const occurredAt = normalizeDatetime(data["啟始時間"]);
      if (!occurredAt) return;
      const record = baseRecord({
        row_number: rowNumber,
        call_type: "數據",
        occurred_at: occurredAt,
        ended_at: normalizeDatetime(data["結束時間"], false),
        duration_seconds: toInt(data["通聯時間(秒)"]),
        target_phone: normalizePhoneText(data["MSISDN"]) || subject["申請號碼"] || "",
        imei: data["IMEI"] || "",
        imsi: data["IMSI"] || "",
        external_ip: data["上網IPv4"] || data["上網IPv6"] || data["備註"] || "",
        upload_bytes: toInt(data["上行用量"]),
        download_bytes: toInt(data["下行用量"]),
        total_bytes: toInt(data["全部用量"]),
        note: data["備註"] || "",
      });
      addStationRef(record, stationMap, "start", cleanCellId(data["基地台 ID"] || ""), data["基地台位址"] || "");
      addStationRef(record, stationMap, record.base_refs.length ? "end" : "primary", cleanCellId(data["最終基地台 ID"] || ""), data["最終基地台位址"] || "");
      parsed.records.push(record);
    });
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseImeiLookup(fileName, sheets) {
    const parsed = makeParsed({ fileName, carrier: "遠傳電信", sourceFormat: "imei_call_lookup_xlsx", sheetName: "IMEI", headerRow: 1, totalSourceRows: sheets.filter((sheet) => sheet.title.toUpperCase().startsWith("IMEI")).reduce((sum, sheet) => sum + sheet.rows.length, 0), subject: {} });
    const stationMap = new Map();
    sheets.filter((sheet) => sheet.title.toUpperCase().startsWith("IMEI")).forEach((sheet) => {
      const header = findHeaderInSheet(sheet, COMPACT_HEADERS);
      if (!header) return;
      forRowsAfter(header, (rowNumber, data) => {
        const occurredAt = normalizeDatetime(data["時間"]);
        if (!occurredAt) return;
        const record = baseRecord({
          row_number: rowNumber,
          call_type: data["通話類別"] || "",
          occurred_at: occurredAt,
          duration_seconds: toInt(data["通話秒數"]),
          target_phone: normalizePhoneText(data["調閱號碼"]),
          counterparty_phone: normalizePhoneText(data["通話對象"]),
          imei: data["IMEI"] || "",
          note: data["備註"] || "",
        });
        addStationCompoundRef(record, stationMap, "start", data["基地台"]);
        addStationCompoundRef(record, stationMap, "end", data["迄基地台"]);
        parsed.records.push(record);
      });
    });
    parsed.records.sort((a, b) => String(a.occurred_at || "").localeCompare(String(b.occurred_at || "")) || String(a.imei || "").localeCompare(String(b.imei || "")) || Number(a.row_number || 0) - Number(b.row_number || 0));
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseConvertedMultiSheet(fileName, sheets, headersByType) {
    const userRows = convertedRows(headersByType.user);
    const subject = convertedSubject(userRows);
    const carrier = convertedFirstUnique(userRows, "電信業者") || "多電信業者";
    const parsed = makeParsed({ fileName, carrier, sourceFormat: "converted_multi_sheet_cdr", sheetName: "多工作表", headerRow: 1, totalSourceRows: sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0), subject });
    const stationMap = new Map();
    const targetFallback = convertedFirstUnique(userRows, "查詢項目", "用戶回應-用戶編號/帳號");
    forRowsAfter(headersByType.calls, (rowNumber, data) => {
      const occurredAt = normalizeDatetime(data["始話時間"]);
      if (!occurredAt) return;
      const record = baseRecord({
        row_number: rowNumber,
        call_type: data["通話類別"] || "",
        occurred_at: occurredAt,
        duration_seconds: toInt(data["通話期間"]),
        target_phone: data["調閱門號"] || data["查詢項目"] || targetFallback,
        counterparty_phone: data["對象門號"] || "",
        imei: data["IMEI"] || "",
        imsi: data["IMSI"] || data["Q"] || "",
        external_ip: data["IP"] || "",
        note: data["備註"] || data["其他"] || "",
      });
      addStationRef(record, stationMap, "start", data["開始基地台編號"] || "", data["開始基地台"] || "");
      addStationRef(record, stationMap, "end", data["結束基地台編號"] || "", data["結束基地台"] || "");
      parsed.records.push(record);
    });
    forRowsAfter(headersByType.data, (rowNumber, data) => {
      const occurredAt = normalizeDatetime(data["開始時間"]);
      if (!occurredAt) return;
      const record = baseRecord({
        row_number: rowNumber,
        call_type: "數據",
        occurred_at: occurredAt,
        ended_at: normalizeDatetime(data["結束時間"], false),
        duration_seconds: toInt(data["連線期間"]),
        target_phone: data["查詢項目"] || data["通聯回應-用戶編號/帳號"] || targetFallback,
        imei: data["IMEI"] || "",
        imsi: data["IMSI"] || "",
        external_ip: data["上網IPv4"] || data["上網IPv6"] || data["使用者內網IP"] || "",
        internal_ip: data["使用者內網IP"] || "",
        note: [data["上網IPv4"], data["上網IPv6"], data["使用者內網IP"], data["備註"]].filter(Boolean).join(" "),
      });
      addStationRef(record, stationMap, "start", data["開始基地台編號"] || "", data["開始基地台"] || "");
      addStationRef(record, stationMap, "end", data["結束基地台編號"] || "", data["結束基地台"] || "");
      parsed.records.push(record);
    });
    parsed.records.sort((a, b) => String(a.occurred_at || "").localeCompare(String(b.occurred_at || "")) || Number(a.row_number || 0) - Number(b.row_number || 0));
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseFet(fileName, sheetName, rows, detected) {
    const subject = metadataBefore(rows, detected.rowNumber);
    const parsed = makeParsed({ fileName, carrier: "遠傳電信", sourceFormat: "fet_data_session", sheetName, headerRow: detected.rowNumber, totalSourceRows: rows.length, subject });
    const stationMap = new Map();
    forRowsAfter(detected, (rowNumber, data) => {
      const occurredAt = normalizeDatetime(data["通聯起始時間"]);
      if (!occurredAt) return;
      const record = baseRecord({
        row_number: rowNumber,
        call_type: "數據",
        occurred_at: occurredAt,
        ended_at: normalizeDatetime(data["通聯結束時間"], false) || data["通聯結束時間"],
        duration_seconds: toInt(data["通聯時間(秒)"]),
        target_phone: data["移動用戶的ISDN號碼"] || "",
        imei: data["手機序號"] || "",
        imsi: data["國際移動使用者識別碼"] || "",
        external_ip: data["外部IP位址"] || "",
        internal_ip: data["內部IP位址"] || "",
        upload_bytes: toInt(data["上傳使用量(Byte)"]),
        download_bytes: toInt(data["下載使用量(Byte)"]),
        total_bytes: toInt(data["全部使用量(Byte)"]),
        note: data["備註"] || "",
      });
      addStationRef(record, stationMap, "start", data["起始基地台編號"] || "", data["起始基地台地址"] || "");
      addStationRef(record, stationMap, "end", data["離開基地台編號"] || "", data["離開基地台地址"] || "");
      parsed.records.push(record);
    });
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseTwm(fileName, sheetName, rows, detected) {
    const subject = metadataBefore(rows, detected.rowNumber);
    const parsed = makeParsed({ fileName, carrier: "台灣大哥大", sourceFormat: "taiwan_mobile_call", sheetName, headerRow: detected.rowNumber, totalSourceRows: rows.length, subject });
    const stationMap = new Map();
    forRowsAfter(detected, (rowNumber, data) => {
      if (!data["始話日期時間"]) return;
      const record = baseRecord({
        row_number: rowNumber,
        call_type: data["通話類別"] || "",
        occurred_at: normalizeDatetime(data["始話日期時間"]) || data["始話日期時間"],
        duration_seconds: toInt(data["通話時間(秒)"]),
        target_phone: data["目標電話"] || "",
        counterparty_phone: data["對象電話"] || "",
        imei: data["IMEI別"] || "",
        note: data["轉接"] || "",
      });
      addStationCompoundRef(record, stationMap, "primary", data["基地台編號1/位置1"]);
      parsed.records.push(record);
    });
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseTwmDataSession(fileName, sheetName, rows, detected) {
    const subject = metadataBefore(rows, detected.rowNumber);
    const sectionSubject = { ...subject };
    let targetPhone = targetPhoneFromSubject(sectionSubject);
    const parsed = makeParsed({ fileName, carrier: "台灣大哥大", sourceFormat: "taiwan_mobile_data_session", sheetName, headerRow: detected.rowNumber, totalSourceRows: rows.length, subject });
    const stationMap = new Map();
    for (const row of rows) {
      if (row.rowNumber <= detected.rowNumber || !row.values.some(Boolean)) continue;
      const metadata = metadataFromValues(row.values);
      if (Object.keys(metadata).length) {
        Object.assign(sectionSubject, metadata);
        targetPhone = targetPhoneFromSubject(sectionSubject);
      }
      const data = rowDict(detected.headers, row.values);
      if (!isIsoDatetime(data["進入基地台時間"])) continue;
      const record = baseRecord({
        row_number: row.rowNumber,
        call_type: "數據",
        occurred_at: data["進入基地台時間"] || "",
        ended_at: data["離開基地台時間"] || "",
        duration_seconds: toInt(data["基地台停留時間"]),
        target_phone: targetPhone,
        imei: data["IMEI"] || "",
        external_ip: data["備註"] || "",
        upload_bytes: toInt(data["上傳使用量(Byte)"]),
        download_bytes: toInt(data["下載使用量(Byte)"]),
        total_bytes: toInt(data["全部使用量(Byte)"]),
        note: data["備註"] || "",
      });
      addStationRef(record, stationMap, "primary", data["離開基地台編號"] || "", data["離開基地台地址"] || "");
      parsed.records.push(record);
    }
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function baseRecord(values) {
    return {
      row_number: values.row_number,
      call_type: values.call_type || "",
      occurred_at: values.occurred_at || "",
      ended_at: values.ended_at || "",
      duration_seconds: values.duration_seconds ?? null,
      target_phone: values.target_phone || "",
      counterparty_phone: values.counterparty_phone || "",
      imei: values.imei || "",
      imsi: values.imsi || "",
      external_ip: values.external_ip || "",
      internal_ip: values.internal_ip || "",
      upload_bytes: values.upload_bytes ?? null,
      download_bytes: values.download_bytes ?? null,
      total_bytes: values.total_bytes ?? null,
      note: values.note || "",
      raw: values.raw || {},
      base_refs: [],
    };
  }

  function detectHeader(rows) {
    for (const row of rows.slice(0, 80)) {
      if (hasHeaders(row.values, FET_HEADERS)) return { sheet: { rows }, rowNumber: row.rowNumber, headers: row.values, sourceFormat: "fet_data_session" };
      if (hasHeaders(row.values, TWM_HEADERS)) return { sheet: { rows }, rowNumber: row.rowNumber, headers: row.values, sourceFormat: "taiwan_mobile_call" };
      if (hasHeaders(row.values, TWM_DATA_HEADERS)) return { sheet: { rows }, rowNumber: row.rowNumber, headers: row.values, sourceFormat: "taiwan_mobile_data_session" };
    }
    throw new Error("找不到支援的標題列");
  }

  function convertedMultiSheetHeaders(sheets) {
    const user = findHeaderSheet(sheets, CONVERTED_USER_HEADERS);
    const calls = findHeaderSheet(sheets, CONVERTED_CALL_HEADERS);
    const data = findHeaderSheet(sheets, CONVERTED_DATA_HEADERS);
    return user && calls && data ? { user, calls, data } : null;
  }

  function findHeaderSheet(sheets, required) {
    for (const sheet of sheets) {
      const found = findHeaderInSheet(sheet, required);
      if (found) return found;
    }
    return null;
  }

  function findHeaderInSheet(sheet, required, scanRows = 20) {
    for (const row of sheet.rows.slice(0, scanRows)) {
      if (hasHeaders(row.values, required)) return { sheet, rowNumber: row.rowNumber, headers: row.values };
    }
    return null;
  }

  function imeiLookupHeaders(sheets) {
    return sheets.some((sheet) => sheet.title.toUpperCase().startsWith("IMEI") && findHeaderInSheet(sheet, COMPACT_HEADERS));
  }

  function compactHeaders(sheets) {
    for (const sheet of sheets) {
      if (sheet.title.toUpperCase().startsWith("IMEI")) continue;
      const found = findHeaderInSheet(sheet, COMPACT_HEADERS);
      if (found) return found;
    }
    return null;
  }

  function fetWebHeaders(sheets) {
    for (const sheet of sheets) {
      const found = findHeaderInSheet(sheet, FET_WEB_HEADERS, 100);
      if (found) return found;
    }
    return null;
  }

  function twmCallPigHeaders(sheets) {
    const preferred = sheets.filter((sheet) => sheet.title === "通聯紀錄");
    for (const sheet of [...preferred, ...sheets.filter((sheet) => !preferred.includes(sheet))]) {
      const found = findHeaderInSheet(sheet, TWM_CALL_PIG_HEADERS, 20);
      if (found) return found;
    }
    return null;
  }

  function hasHeaders(values, required) {
    const present = new Set(values.map(cellText).filter(Boolean));
    return required.every((header) => present.has(header));
  }

  function forRowsAfter(header, callback) {
    for (const row of header.sheet.rows) {
      if (row.rowNumber <= header.rowNumber || !row.values.some((value) => cellText(value))) continue;
      callback(row.rowNumber, rowDict(header.headers, row.values), row.values);
    }
  }

  function rowDict(headers, values) {
    const result = {};
    headers.forEach((header, index) => {
      const key = cellText(header);
      if (key && result[key] === undefined) result[key] = cellText(values[index]);
    });
    return result;
  }

  function convertedRows(header) {
    const result = [];
    forRowsAfter(header, (_rowNumber, data) => result.push(data));
    return result;
  }

  function convertedSubject(rows) {
    const startTimes = rows.map((row) => normalizeDatetime(row["區段時間-開始日期時間"])).filter(Boolean);
    const endTimes = rows.map((row) => normalizeDatetime(row["區段時間-結束日期時間"])).filter(Boolean);
    const allTimes = [...startTimes, ...endTimes].sort();
    return compactObject({
      "用戶名稱": convertedFirstUnique(rows, "用戶名稱"),
      "申請號碼": convertedFirstUnique(rows, "查詢項目", "用戶回應-用戶編號/帳號"),
      "身份證字號": convertedFirstUnique(rows, "身份識別碼", "第二身分識別碼"),
      "生日": convertedFirstUnique(rows, "生日"),
      "帳寄地址": convertedFirstUnique(rows, "帳寄地址"),
      "戶籍地址": convertedFirstUnique(rows, "戶籍地址"),
      "時間區間": allTimes.length ? `${allTimes[0]} ~ ${allTimes[allTimes.length - 1]}` : "",
    });
  }

  function convertedFirstUnique(rows, ...keys) {
    const values = [];
    const seen = new Set();
    keys.forEach((key) => {
      rows.forEach((row) => {
        const value = cellText(row[key]);
        if (value && !seen.has(value)) {
          seen.add(value);
          values.push(value);
        }
      });
    });
    return values.join("、");
  }

  function twmCallPigSubject(sheets) {
    const subject = {};
    sheets.filter((sheet) => sheet.title === "使用者資料").forEach((sheet) => {
      const found = findHeaderInSheet(sheet, ["用戶名稱", "查詢項目", "電信業者"], 10);
      if (!found) return;
      forRowsAfter(found, (_rowNumber, data) => {
        [
          ["用戶名稱", data["用戶名稱"]],
          ["申請號碼", data["查詢項目"] || data["用戶回應-用戶編號/帳號"]],
          ["身份證字號", data["身份識別碼"] || data["第二身分識別碼"]],
          ["帳寄地址", data["帳寄地址"]],
          ["戶籍地址", data["戶籍地址"]],
          ["申裝地址", data["申裝地址"]],
          ["電信業者", data["電信業者"]],
          ["_carrier", data["電信業者"]],
        ].forEach(([key, value]) => {
          const text = cellText(value);
          if (text && !subject[key]) subject[key] = text;
        });
      });
    });
    return subject;
  }

  function compactSubject(sheets) {
    for (const sheet of sheets) {
      if (sheet.title !== "基本人資") continue;
      const found = findHeaderInSheet(sheet, ["電話號碼", "姓名", "身份證號碼"], 5);
      if (!found) continue;
      let subject = {};
      forRowsAfter(found, (_rowNumber, data) => {
        if (Object.keys(subject).length) return;
        subject = compactObject({
          "用戶名稱": data["姓名"] || "",
          "申請號碼": normalizePhoneText(data["電話號碼"]),
          "身份證字號": data["身份證號碼"] || "",
          "生日": normalizeDate(data["出生日期"] || ""),
          "帳寄地址": data["帳寄地址"] || "",
          "戶籍地址": data["戶籍地址"] || "",
          "_carrier": data["調閱電信業者"] || "",
        });
      });
      return subject;
    }
    return {};
  }

  function fetWebSubject(rows, headerRow) {
    const subject = {};
    for (const row of rows) {
      if (row.rowNumber >= headerRow) break;
      row.values.forEach((value, index) => {
        const text = cellText(value);
        if (!text) return;
        if (text.startsWith("調閱目標")) subject["申請號碼"] = normalizePhoneText(valueAfterLabel(text) || nextValue(row.values, index));
        else if (text.startsWith("用戶名稱")) subject["用戶名稱"] = valueAfterLabel(text);
        else if (text.startsWith("出生日期")) subject["生日"] = normalizeDate(valueAfterLabel(text));
        else if (text.startsWith("帳寄地址")) subject["帳寄地址"] = valueAfterLabel(text);
        else if (text.startsWith("戶籍地址")) subject["戶籍地址"] = valueAfterLabel(text);
        else if (text === "身分證號碼") subject["身份證字號"] = nextValue(row.values, index);
      });
    }
    return compactObject(subject);
  }

  function metadataBefore(rows, headerRow) {
    const metadata = {};
    rows.forEach((row) => {
      if (row.rowNumber < headerRow) Object.assign(metadata, metadataFromValues(row.values));
    });
    return metadata;
  }

  function metadataFromValues(values) {
    const metadata = {};
    values.forEach((value) => {
      const text = cellText(value);
      if (!text || (!text.includes(":") && !text.includes("："))) return;
      const [key, ...rest] = text.split(/[:：]/);
      metadata[key.trim()] = rest.join(":").trim();
    });
    return metadata;
  }

  function targetPhoneFromSubject(subject) {
    return cellText(subject["申請號碼"] || subject["調閱目標"]);
  }

  function isTwmCallXml(xml) {
    return Boolean(firstXmlBlock(xml, "通聯資料")) || /<[^>]*查詢單[\s>]/.test(xml) && (Boolean(firstXmlBlock(xml, "通聯記錄查詢條件")) || Boolean(firstXmlBlock(xml, "電話號碼")) || Boolean(firstXmlBlock(xml, "電信業者")));
  }

  function iterXmlBlocks(xml, name) {
    const escaped = escapeRegExp(name);
    const re = new RegExp(`<(?:[\\w.-]+:)?${escaped}(?=[\\s>/])[^>]*>[\\s\\S]*?<\\/(?:[\\w.-]+:)?${escaped}>`, "g");
    return String(xml || "").match(re) || [];
  }

  function firstXmlBlock(xml, name) {
    return iterXmlBlocks(xml, name)[0] || "";
  }

  function firstXmlText(xml, name) {
    const block = firstXmlBlock(xml, name);
    return block ? stripXmlTags(block).trim() : "";
  }

  function xmlChildren(block) {
    const result = {};
    const inner = innerXml(block);
    const re = /<([^\s>/]+)(?=[\s>/])[^>]*>([\s\S]*?)<\/\1>/g;
    let match;
    while ((match = re.exec(inner))) {
      const name = match[1].split(":").pop();
      if (!result[name]) result[name] = xmlDecode(stripXmlTags(match[2]).trim());
    }
    return result;
  }

  function innerXml(block) {
    return String(block || "").replace(/^<[^>]+>/, "").replace(/<\/[^>]+>\s*$/, "");
  }

  function stripXmlTags(text) {
    return String(text || "").replace(/<[^>]+>/g, "");
  }

  function xmlDecode(text) {
    return String(text || "")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }

  function addStationCompoundRef(record, stationMap, role, value) {
    const station = stationFromCompound(value);
    if (!station) return;
    const key = stationKey(station);
    station.station_key = key;
    stationMap.set(key, station);
    record.base_refs.push({ role, station_key: key });
  }

  function addStationRef(record, stationMap, role, cellId, address) {
    const station = stationFromParts(cellId, address);
    if (!station) return;
    const key = stationKey(station);
    station.station_key = key;
    stationMap.set(key, station);
    record.base_refs.push({ role, station_key: key });
  }

  function stationFromCompound(value) {
    const text = cellText(value);
    if (!text || text.startsWith("路由:")) return null;
    if (isVirtualStationText(text)) return { cell_id: "VOWIFI", address: text, normalized_address: "VOWIFI", status: "not_applicable", is_virtual: true };
    if (text.includes("/")) {
      const [cellId, ...rest] = text.split("/");
      return stationFromParts(cellId, rest.join("/"));
    }
    const match = text.match(/^([0-9A-Za-z.+\-Ee]+)\s+(.+)$/);
    if (match && /\d/.test(match[1])) return stationFromParts(cleanCellId(match[1]), match[2]);
    return stationFromParts(text, "");
  }

  function stationFromParts(cellId, address) {
    const cleanId = cellText(cellId);
    const cleanAddress = cellText(address);
    if (!cleanId && !cleanAddress) return null;
    if (isVirtualStationText(cleanId) || isVirtualStationText(cleanAddress)) {
      return { cell_id: "VOWIFI", address: cleanAddress || cleanId || "VOWIFI", normalized_address: "VOWIFI", status: "not_applicable", is_virtual: true };
    }
    return {
      cell_id: cleanId,
      address: cleanAddress,
      normalized_address: normalizeAddress(cleanAddress),
      status: cleanAddress ? "pending" : "missing_address",
      is_virtual: false,
    };
  }

  function stationKey(station) {
    return `${station.cell_id || ""}|${station.normalized_address || ""}`;
  }

  function isVirtualStationText(value) {
    const text = cellText(value).replace(/\s+/g, "").toUpperCase();
    return text.includes("VOWIFI") || text.includes("WIFI熱點") || (text.includes("WI-FI") && text.includes("通話"));
  }

  function normalizeAddress(address) {
    return cellText(address).replace(/\s+/g, "").replace(/[（(][2345]G[)）]/gi, "");
  }

  function cleanCellId(value) {
    const text = cellText(value);
    if (!text) return "";
    if (/^\d+(?:\.\d+)?(?:[Ee][+\-]?\d+)?$/.test(text)) {
      const number = Number(text);
      if (Number.isFinite(number)) return String(Math.round(number));
    }
    return text;
  }

  function normalizeDatetime(value, allowEpoch = true) {
    let text = cellText(value).replace(/\u00a0/g, " ");
    if (!text) return "";
    if (/^\d+(?:\.\d+)?$/.test(text)) {
      const number = Number(text);
      if (number >= 20000 && number <= 60000) {
        const normalized = excelSerialDatetime(number);
        if (!allowEpoch && normalized === "1970-01-01T00:00:00") return "";
        return normalized;
      }
      return "";
    }
    text = text.replace(/\//g, "-").replace(/\s+/g, " ").trim();
    let match = text.match(/^(\d{4}-\d{1,2}-\d{1,2})[ T](\d{1,2}:\d{1,2}:\d{1,2})$/);
    if (!match) match = text.match(/^(\d{4}-\d{1,2}-\d{1,2})(\d{1,2}:\d{1,2}:\d{1,2})$/);
    if (!match) return isIsoDatetime(text) ? text : "";
    const [year, month, day] = match[1].split("-").map(Number);
    const [hour, minute, second] = match[2].split(":").map(Number);
    const normalized = `${pad(year, 4)}-${pad(month)}-${pad(day)}T${pad(hour)}:${pad(minute)}:${pad(second)}`;
    if (!allowEpoch && normalized === "1970-01-01T00:00:00") return "";
    return normalized;
  }

  function excelSerialDatetime(value) {
    const ms = Math.round((Number(value) - 25569) * 86400000);
    const date = new Date(ms);
    return `${pad(date.getUTCFullYear(), 4)}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}T${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}:${pad(date.getUTCSeconds())}`;
  }

  function normalizeDate(value) {
    const normalized = normalizeDatetime(value);
    return normalized ? normalized.slice(0, 10) : cellText(value).replace(/\//g, "-");
  }

  function isIsoDatetime(value) {
    return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(cellText(value));
  }

  function normalizePhoneText(value) {
    const text = cellText(value).split("(", 1)[0];
    const digits = text.replace(/\D+/g, "");
    if (!digits) return "";
    if (digits.startsWith("886") && digits.length >= 11) return `0${digits.slice(3)}`;
    if (digits.length === 9 && digits.startsWith("9")) return `0${digits}`;
    if (digits.length === 8 && "2345678".includes(digits[0])) return `0${digits}`;
    return digits;
  }

  function directionLabel(callType) {
    const text = cellText(callType);
    if (text.includes("受") || text.includes("收")) return "inbound";
    if (text.includes("發") || text.includes("撥") || text.includes("去")) return "outbound";
    if (text.includes("數據") || text.includes("上網")) return "data";
    return "other";
  }

  function computePhoneStats(records, mode = "count") {
    const inbound = new Map();
    const outbound = new Map();
    const total = new Map();
    records.forEach((record) => {
      const seconds = Number(record.duration_seconds || 0);
      const direction = directionLabel(record.call_type || record.direction);
      if (record.counterparty_phone) addPhoneStat(total, record.counterparty_phone, "對象", seconds);
      if (record.target_phone) addPhoneStat(total, record.target_phone, "目標", seconds);
      if (direction === "inbound" && record.counterparty_phone) addPhoneStat(inbound, record.counterparty_phone, "來電/收訊", seconds);
      if (direction === "outbound" && record.counterparty_phone) addPhoneStat(outbound, record.counterparty_phone, "去電/發訊", seconds);
    });
    return {
      inboundRows: sortedStats(inbound, mode),
      outboundRows: sortedStats(outbound, mode),
      totalRows: sortedStats(total, mode),
    };
  }

  function addPhoneStat(map, phone, role, seconds) {
    const key = normalizePhoneText(phone);
    if (!key) return;
    const item = map.get(key) || { phone: key, role, count: 0, seconds: 0 };
    item.count += 1;
    item.seconds += Number(seconds || 0);
    map.set(key, item);
  }

  function sortedStats(map, mode) {
    return Array.from(map.values()).sort((a, b) => {
      if (mode === "seconds") return b.seconds - a.seconds || b.count - a.count || a.phone.localeCompare(b.phone);
      return b.count - a.count || b.seconds - a.seconds || a.phone.localeCompare(b.phone);
    });
  }

  function computeHourBuckets(records) {
    const buckets = HOUR_LABELS.map((label, hour) => ({ hour, label, count: 0 }));
    records.forEach((record) => {
      const hour = hourFromIso(record.occurred_at);
      if (hour >= 0) buckets[hour].count += 1;
    });
    return buckets;
  }

  function computeAddressHotspots(records, stations) {
    const stationMap = new Map();
    (stations || []).forEach((station) => {
      stationMap.set(station.station_key || stationKey(station), station);
    });
    const groups = new Map();
    (records || []).forEach((record) => {
      const seenInRecord = new Set();
      (record.base_refs || []).forEach((ref) => {
        const station = stationMap.get(ref.station_key);
        if (!station || station.is_virtual) return;
        const address = cellText(station.address);
        const normalized = station.normalized_address || normalizeAddress(address);
        if (!address || !normalized || seenInRecord.has(normalized)) return;
        seenInRecord.add(normalized);
        const group = groups.get(normalized) || {
          key: normalized,
          address,
          normalized_address: normalized,
          count: 0,
          percent: 0,
          times: [],
          first_seen: "",
          last_seen: "",
        };
        group.count += 1;
        if (record.occurred_at) group.times.push(record.occurred_at);
        groups.set(normalized, group);
      });
    });
    const total = Math.max(1, (records || []).length);
    return Array.from(groups.values()).map((group) => {
      group.times.sort();
      group.first_seen = group.times[0] || "";
      group.last_seen = group.times[group.times.length - 1] || "";
      group.percent = (group.count / total) * 100;
      return group;
    }).sort((a, b) => b.count - a.count || a.address.localeCompare(b.address, "zh-Hant", { numeric: true }));
  }

  function hourFromIso(value) {
    const match = cellText(value).match(/T(\d{2}):/);
    return match ? Number(match[1]) : -1;
  }

  function toUint8Array(input) {
    if (input instanceof Uint8Array) return input;
    if (typeof Buffer !== "undefined" && Buffer.isBuffer(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    if (input instanceof ArrayBuffer) return new Uint8Array(input);
    if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    throw new Error("Unsupported file content");
  }

  function looksLikeXml(bytes) {
    const slice = bytes.slice(0, 32);
    return Array.from(slice).some((byte) => byte > 32) && decodeUtf8(slice).trimStart().startsWith("<");
  }

  function decodeUtf8(bytes) {
    if (typeof TextDecoder !== "undefined") return new TextDecoder("utf-8").decode(bytes);
    return Buffer.from(bytes).toString("utf8");
  }

  function cellText(value) {
    return String(value ?? "").replace(/\u00a0/g, " ").trim();
  }

  function toInt(value) {
    const text = cellText(value).replace(/,/g, "").replace(/\s*BYTE$/i, "");
    if (!text) return null;
    const number = Number.parseFloat(text);
    return Number.isFinite(number) ? Math.trunc(number) : null;
  }

  function compactObject(object) {
    return Object.fromEntries(Object.entries(object).filter(([, value]) => cellText(value)));
  }

  function valueAfterLabel(text) {
    const parts = cellText(text).split(/[:：]/);
    return parts.length > 1 ? parts.slice(1).join(":").trim() : "";
  }

  function nextValue(values, index) {
    for (const value of values.slice(index + 1)) {
      const text = cellText(value);
      if (text) return text;
    }
    return "";
  }

  function readJson(text, fallback) {
    try {
      return text ? JSON.parse(text) : fallback;
    } catch (_error) {
      return fallback;
    }
  }

  function csvLine(values) {
    return values.map((value) => `"${String(value ?? "").replace(/"/g, '""')}"`).join(",");
  }

  function downloadText(fileName, text, type) {
    const blob = new Blob([text], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  function toLocalDatetimeValue(date) {
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
  }

  function pad(value, size = 2) {
    return String(value).padStart(size, "0");
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  return {
    parseImportFile,
    computePhoneStats,
    computeHourBuckets,
    computeAddressHotspots,
    buildSubmissionCsv,
    collectSubmissionPhones,
    collectUniqueImeis,
    normalizePhoneText,
    tellowsUrl,
    hourButtonLabel,
  };
});
