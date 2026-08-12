(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(root.XLSX || require("./vendor/xlsx.full.min.js"), require("./attachment-export.js"));
  } else {
    root.PhoneWorkbench = factory(root.XLSX, root.PhoneAttachmentExport);
  }
})(typeof globalThis !== "undefined" ? globalThis : window, function (XLSX, AttachmentExport) {
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
  const LOCAL_EXPORT_VERSION = "phone-workbench-local-settings-v1";
  const CALL_PAGE_SIZE = 500;
  const MULTI_LOCATION_PAGE_SIZE = 500;
  const MULTI_LOCATION_WINDOW_MINUTES = 30;
  const ATTACHMENT_ASSETS = {
    exceljs: { src: "./vendor/exceljs.min.js?v=20260812-multi-number-source-detail-v1", integrity: "sha384-Pqp51FUN2/qzfxZxBCtF0stpc9ONI6MYZpVqmo8m20SoaQCzf+arZvACkLkirlPz" },
    pdfLib: { src: "./vendor/pdf-lib.min.js?v=20260812-multi-number-source-detail-v1", integrity: "sha384-weMABwrltA6jWR8DDe9Jp5blk+tZQh7ugpCsF3JwSA53WZM9/14PjS5LAJNHNjAI" },
    fontkit: { src: "./vendor/fontkit.umd.min.js?v=20260812-multi-number-source-detail-v1", integrity: "sha384-2p6U+1mmqF10USehFeRiyG2ESG9FwIqN+jxULn5w9jjQIihSn9Pt13dVCn/Hawjn" },
    fontData: { src: "./vendor/open-huninn-data.js?v=20260812-multi-number-source-detail-v1", integrity: "sha384-upBq5rvuXmWYAJi6vO2VylcS6jMVjb7GMuvCJguhimt6kQ2uYG8eZz4GfqsI4Hou" },
  };
  const loadedAttachmentAssets = new Map();
  const HOUR_LABELS = Array.from({ length: 24 }, (_, hour) => `${String(hour).padStart(2, "0")}-${String(hour + 1).padStart(2, "0")}`);
  const CALL_COLUMNS = [
    { key: "source_file", width: 180, min: 130 },
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
    multiLocation: "多門號位置",
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
  const CHT_PROSECUTOR_HEADERS = ["CDR類別", "主叫號碼", "查詢狀態", "受叫號碼", "始話日期時間", "通話秒數", "IMEI", "指定轉接", "起始基地台-地址/終止基地台-地址"];
  const FET_PROSECUTOR_CALL_HEADERS = ["始話時間", "通話秒數", "調閱號碼", "IMEI", "通話類別", "通話對象", "轉接電話", "基地台/交換機", "備註"];
  const FET_PROSECUTOR_METADATA_KEYS = new Set(["文號", "查詢日期", "電信業者", "通聯類別", "查詢狀態", "區段時間", "備註", "電話號碼"]);
  const FET_ORDER_HEADERS = [
    "DocNo", "Seq", "Status2", "QueType", "QueObject", "QueDirection", "CallDirection",
    "CallingNumber", "CalledNumber", "CallStartTimeStamp", "Duration", "IMEI",
  ];
  const FET_ORDER_CALL_TYPES = {
    O: "Original發話",
    T: "Terminal受話",
    I: "Incoming Gateway進來CDR",
    1: "SMS 系統發訊",
    2: "SMS 手機發訊",
    9: "SMS 收訊",
    S: "雙號共振",
    M: "多媒體簡訊",
  };
  const FET_ORDER_QUERY_TYPES = {
    1: "電話號碼", 2: "手機序號", 3: "身份證ID", 4: "SIM卡", 6: "IMSI",
    10: "護照號碼", 11: "統一證號", 12: "公話編號",
  };
  const FET_ORDER_QUERY_DIRECTIONS = { 1: "發話", 2: "受話", 3: "雙向" };
  const FET_ORDER_RESULT_TYPES = { 1: "用戶資料與使用記錄", 2: "使用記錄", 3: "用戶資料" };
  const TAIWAN_COUNTIES = [
    "臺北市", "新北市", "桃園市", "臺中市", "臺南市", "高雄市", "基隆市", "新竹市", "嘉義市", "新竹縣", "苗栗縣",
    "彰化縣", "南投縣", "雲林縣", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣", "臺東縣", "澎湖縣", "金門縣", "連江縣",
  ];
  const UNKNOWN_COUNTY = "未辨識";
  const ALL_COUNTY_FILTER_KEYS = [...TAIWAN_COUNTIES, UNKNOWN_COUNTY];
  const TAIWAN_ADMINISTRATIVE_DISTRICTS = {
    "臺北市": "松山區|信義區|大安區|中山區|中正區|大同區|萬華區|文山區|南港區|內湖區|士林區|北投區".split("|"),
    "新北市": "板橋區|三重區|中和區|永和區|新莊區|新店區|樹林區|鶯歌區|三峽區|淡水區|汐止區|瑞芳區|土城區|蘆洲區|五股區|泰山區|林口區|深坑區|石碇區|坪林區|三芝區|石門區|八里區|平溪區|雙溪區|貢寮區|金山區|萬里區|烏來區".split("|"),
    "桃園市": "桃園區|中壢區|大溪區|楊梅區|蘆竹區|大園區|龜山區|八德區|龍潭區|平鎮區|新屋區|觀音區|復興區".split("|"),
    "臺中市": "中區|東區|南區|西區|北區|西屯區|南屯區|北屯區|豐原區|東勢區|大甲區|清水區|沙鹿區|梧棲區|后里區|神岡區|潭子區|大雅區|新社區|石岡區|外埔區|大安區|烏日區|大肚區|龍井區|霧峰區|太平區|大里區|和平區".split("|"),
    "臺南市": "新營區|鹽水區|白河區|柳營區|後壁區|東山區|麻豆區|下營區|六甲區|官田區|大內區|佳里區|學甲區|西港區|七股區|將軍區|北門區|新化區|善化區|新市區|安定區|山上區|玉井區|楠西區|南化區|左鎮區|仁德區|歸仁區|關廟區|龍崎區|永康區|東區|南區|北區|安南區|安平區|中西區".split("|"),
    "高雄市": "鹽埕區|鼓山區|左營區|楠梓區|三民區|新興區|前金區|苓雅區|前鎮區|旗津區|小港區|鳳山區|林園區|大寮區|大樹區|大社區|仁武區|鳥松區|岡山區|橋頭區|燕巢區|田寮區|阿蓮區|路竹區|湖內區|茄萣區|永安區|彌陀區|梓官區|旗山區|美濃區|六龜區|甲仙區|杉林區|內門區|茂林區|桃源區|那瑪夏區".split("|"),
    "基隆市": "中正區|七堵區|暖暖區|仁愛區|中山區|安樂區|信義區".split("|"),
    "新竹市": "東區|北區|香山區".split("|"),
    "嘉義市": "東區|西區".split("|"),
    "新竹縣": "竹北市|竹東鎮|新埔鎮|關西鎮|湖口鄉|新豐鄉|芎林鄉|橫山鄉|北埔鄉|寶山鄉|峨眉鄉|尖石鄉|五峰鄉".split("|"),
    "苗栗縣": "苗栗市|苑裡鎮|通霄鎮|竹南鎮|頭份市|後龍鎮|卓蘭鎮|大湖鄉|公館鄉|銅鑼鄉|南庄鄉|頭屋鄉|三義鄉|西湖鄉|造橋鄉|三灣鄉|獅潭鄉|泰安鄉".split("|"),
    "彰化縣": "彰化市|鹿港鎮|和美鎮|線西鄉|伸港鄉|福興鄉|秀水鄉|花壇鄉|芬園鄉|員林市|溪湖鎮|田中鎮|大村鄉|埔鹽鄉|埔心鄉|永靖鄉|社頭鄉|二水鄉|北斗鎮|二林鎮|田尾鄉|埤頭鄉|芳苑鄉|大城鄉|竹塘鄉|溪州鄉".split("|"),
    "南投縣": "南投市|埔里鎮|草屯鎮|竹山鎮|集集鎮|名間鄉|鹿谷鄉|中寮鄉|魚池鄉|國姓鄉|水里鄉|信義鄉|仁愛鄉".split("|"),
    "雲林縣": "斗六市|斗南鎮|虎尾鎮|西螺鎮|土庫鎮|北港鎮|古坑鄉|大埤鄉|莿桐鄉|林內鄉|二崙鄉|崙背鄉|麥寮鄉|東勢鄉|褒忠鄉|臺西鄉|元長鄉|四湖鄉|口湖鄉|水林鄉".split("|"),
    "嘉義縣": "太保市|朴子市|布袋鎮|大林鎮|民雄鄉|溪口鄉|新港鄉|六腳鄉|東石鄉|義竹鄉|鹿草鄉|水上鄉|中埔鄉|竹崎鄉|梅山鄉|番路鄉|大埔鄉|阿里山鄉".split("|"),
    "屏東縣": "屏東市|潮州鎮|東港鎮|恆春鎮|萬丹鄉|長治鄉|麟洛鄉|九如鄉|里港鄉|鹽埔鄉|高樹鄉|萬巒鄉|內埔鄉|竹田鄉|新埤鄉|枋寮鄉|新園鄉|崁頂鄉|林邊鄉|南州鄉|佳冬鄉|琉球鄉|車城鄉|滿州鄉|枋山鄉|三地門鄉|霧臺鄉|瑪家鄉|泰武鄉|來義鄉|春日鄉|獅子鄉|牡丹鄉".split("|"),
    "宜蘭縣": "宜蘭市|羅東鎮|蘇澳鎮|頭城鎮|礁溪鄉|壯圍鄉|員山鄉|冬山鄉|五結鄉|三星鄉|大同鄉|南澳鄉".split("|"),
    "花蓮縣": "花蓮市|鳳林鎮|玉里鎮|新城鄉|吉安鄉|壽豐鄉|光復鄉|豐濱鄉|瑞穗鄉|富里鄉|秀林鄉|萬榮鄉|卓溪鄉".split("|"),
    "臺東縣": "臺東市|成功鎮|關山鎮|卑南鄉|鹿野鄉|池上鄉|東河鄉|長濱鄉|太麻里鄉|大武鄉|綠島鄉|海端鄉|延平鄉|金峰鄉|達仁鄉|蘭嶼鄉".split("|"),
    "澎湖縣": "馬公市|湖西鄉|白沙鄉|西嶼鄉|望安鄉|七美鄉".split("|"),
    "金門縣": "金城鎮|金沙鎮|金湖鎮|金寧鄉|烈嶼鄉|烏坵鄉".split("|"),
    "連江縣": "南竿鄉|北竿鄉|莒光鄉|東引鄉".split("|"),
  };

  const state = {
    view: "hours",
    cases: [],
    currentWorkspace: null,
    callRecords: [],
    callSort: { column: "occurred_at", direction: "asc" },
    callPage: 1,
    phoneStatsRankMode: "count",
    hourSelection: new Set(Array.from({ length: 24 }, (_, index) => index)),
    appliedHourSelection: new Set(Array.from({ length: 24 }, (_, index) => index)),
    expandedHotspotAddress: "",
    hotspotCountySelection: new Set(ALL_COUNTY_FILTER_KEYS),
    hotspotCountyDraft: new Set(ALL_COUNTY_FILTER_KEYS),
    dateRangeBounds: { start: "", end: "" },
    dateRange: { start: "", end: "", active: false },
    dateRangeDraft: { start: "", end: "" },
    multiLocationWorkspace: null,
    multiLocationMatches: [],
    multiLocationExcluded: { missing_phone: 0, invalid_time: 0, invalid_address: 0 },
    multiLocationPage: 1,
    multiLocationExpandedMatches: new Set(),
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
    syncHotspotCountyFilterButton();
    syncDateFilterPanel();
    initCallColumnResize();
    maybeShowUsageNotice();
  }

  function $(id) {
    return document.getElementById(id);
  }

  function bindEvents() {
    document.querySelectorAll("button[data-view]").forEach((button) => {
      button.addEventListener("click", () => setView(button.dataset.view));
    });
    $("importButton")?.addEventListener("click", () => $("fileInput")?.click());
    $("fileInput")?.addEventListener("change", handleFileImport);
    $("multiLocationImportButton")?.addEventListener("click", () => $("multiLocationFileInput")?.click());
    $("multiLocationFileInput")?.addEventListener("change", handleMultiLocationImport);
    $("multiLocationPrevPage")?.addEventListener("click", () => changeMultiLocationPage(-1));
    $("multiLocationNextPage")?.addEventListener("click", () => changeMultiLocationPage(1));
    $("multiLocationRows")?.addEventListener("input", (event) => {
      const input = event.target.closest("[data-phone-note]");
      if (input) updatePhoneNote(input.dataset.phoneNote, input.value, input);
    });
    $("multiLocationRows")?.addEventListener("click", (event) => {
      const button = event.target.closest("[data-multi-location-detail]");
      if (button) toggleMultiLocationDetail(button.dataset.multiLocationDetail);
    });
    $("dateFilterButton")?.addEventListener("click", showDateFilterModal);
    $("dateFilterCloseButton")?.addEventListener("click", hideDateFilterModal);
    $("dateFilterCancelButton")?.addEventListener("click", hideDateFilterModal);
    $("dateFilterResetButton")?.addEventListener("click", resetDateFilter);
    $("dateFilterApplyButton")?.addEventListener("click", applyDateFilter);
    ["dateFilterStartInput", "dateFilterEndInput"].forEach((id) => $(id)?.addEventListener("input", syncDateFilterDraftUi));
    $("dateFilterModal")?.addEventListener("click", (event) => {
      if (event.target === $("dateFilterModal")) hideDateFilterModal();
    });
    $("recordSearch")?.addEventListener("input", () => {
      state.callPage = 1;
      renderTwoWayCalls();
    });
    $("callPrevPage")?.addEventListener("click", () => changeCallPage(-1));
    $("callNextPage")?.addEventListener("click", () => changeCallPage(1));
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
    $("hotspotCountyFilterButton")?.addEventListener("click", showHotspotCountyFilterModal);
    $("hotspotCountyFilterCloseButton")?.addEventListener("click", hideHotspotCountyFilterModal);
    $("hotspotCountyFilterCancelButton")?.addEventListener("click", hideHotspotCountyFilterModal);
    $("hotspotCountySelectAllButton")?.addEventListener("click", selectAllHotspotCountyDraft);
    $("hotspotCountyClearAllButton")?.addEventListener("click", clearAllHotspotCountyDraft);
    $("hotspotCountyFilterApplyButton")?.addEventListener("click", applyHotspotCountyFilter);
    $("hotspotCountyFilterList")?.addEventListener("change", handleHotspotCountyDraftChange);
    $("hotspotCountyFilterModal")?.addEventListener("click", (event) => {
      if (event.target === $("hotspotCountyFilterModal")) hideHotspotCountyFilterModal();
    });
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
    $("attachmentExportButton")?.addEventListener("click", showAttachmentExportModal);
    $("attachmentExportCloseButton")?.addEventListener("click", hideAttachmentExportModal);
    $("attachmentXlsxButton")?.addEventListener("click", downloadAttachmentXlsx);
    document.querySelectorAll("[data-attachment-pdf]").forEach((button) => {
      button.addEventListener("click", () => downloadAttachmentPdf(button.dataset.attachmentPdf));
    });
    $("attachmentExportModal")?.addEventListener("click", (event) => {
      if (event.target === $("attachmentExportModal")) hideAttachmentExportModal();
    });
    $("noticeDismissButton")?.addEventListener("click", hideUsageNotice);
    $("oneClickUpdateButton")?.addEventListener("click", runOneClickUpdate);
    $("usageNoticeModal")?.addEventListener("click", (event) => {
      if (event.target === $("usageNoticeModal")) hideUsageNotice();
    });
    document.addEventListener("keydown", (event) => {
      if (event.key !== "Escape") return;
      if (!$("dateFilterModal")?.hidden) hideDateFilterModal();
      else if (!$("hotspotCountyFilterModal")?.hidden) hideHotspotCountyFilterModal();
      else if (!$("attachmentExportModal")?.hidden) hideAttachmentExportModal();
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
    document.querySelectorAll("button[data-view]").forEach((button) => {
      const active = button.dataset.view === view;
      button.classList.toggle("active", active);
      if (active) button.setAttribute("aria-current", "page");
      else button.removeAttribute("aria-current");
    });
    $("pageTitle").textContent = VIEW_TITLES[view];
    syncPrimarySidebarPanels();
    if (view === "calls") renderTwoWayCalls();
    if (view === "profile") renderProfileView();
    if (view === "stats") renderStatsView();
    if (view === "hours") renderHoursView();
    if (view === "multiLocation") renderMultiLocationView();
  }

  function syncPrimarySidebarPanels() {
    if (typeof document === "undefined") return;
    const multiLocationActive = state.view === "multiLocation";
    const importPanel = $("mainImportPanel");
    if (importPanel) importPanel.hidden = multiLocationActive;
    if (multiLocationActive) {
      const datePanel = $("dateFilterPanel");
      if (datePanel) datePanel.hidden = true;
    } else {
      syncDateFilterPanel();
    }
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
    const workspaces = [];
    let failures = 0;
    for (const file of files) {
      try {
        const content = await readFileArrayBuffer(file);
        const workspace = parseImportFile(file.name, content);
        workspaces.push(workspace);
        appendImportResult(workspace.case);
      } catch (error) {
        failures += 1;
        appendImportError(file.name, error);
      }
    }
    if (workspaces.length) {
      applyWorkspace(mergeWorkspaces(workspaces), workspaces.map((workspace) => workspace.case));
      setView("calls");
      $("importStatus").textContent = failures ? "部分檔案匯入完成；失敗項目請見下方。" : "匯入完成。";
    } else {
      $("importStatus").textContent = "所有檔案均匯入失敗，原資料未變更。";
    }
    $("fileInput").value = "";
  }

  async function handleMultiLocationImport(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;
    $("multiLocationImportStatus").textContent = `匯入 ${files.length} 個檔案中...`;
    $("multiLocationImportResults").innerHTML = "";
    const workspaces = [];
    let failures = 0;
    for (const file of files) {
      try {
        const content = await readFileArrayBuffer(file);
        const workspace = parseImportFile(file.name, content);
        workspaces.push(workspace);
        appendMultiLocationImportResult(workspace.case);
      } catch (error) {
        failures += 1;
        appendMultiLocationImportError(file.name, error);
      }
    }
    if (workspaces.length) {
      const workspace = mergeWorkspaces(workspaces);
      const analysis = computeMultiNumberLocationMatches(workspace, { windowMinutes: MULTI_LOCATION_WINDOW_MINUTES });
      state.multiLocationWorkspace = workspace;
      state.multiLocationMatches = analysis.matches;
      state.multiLocationExcluded = analysis.excluded;
      state.multiLocationPage = 1;
      state.multiLocationExpandedMatches = new Set();
      renderMultiLocationView();
      $("multiLocationImportStatus").textContent = failures
        ? "部分檔案匯入完成；失敗項目請見下方。"
        : "匯入與位置比對完成。";
    } else {
      $("multiLocationImportStatus").textContent = "所有檔案均匯入失敗，原多門號位置資料未變更。";
    }
    $("multiLocationFileInput").value = "";
  }

  function appendMultiLocationImportResult(item) {
    const div = document.createElement("div");
    div.className = "import-result-item";
    div.innerHTML = `<strong>${escapeHtml(item.source_file || "匯入檔案")}</strong><span>${escapeHtml(item.source_format || "")}</span><span>解析完成</span>`;
    $("multiLocationImportResults").appendChild(div);
  }

  function appendMultiLocationImportError(name, error) {
    const div = document.createElement("div");
    div.className = "import-result-item danger-text";
    div.textContent = `${name} 匯入失敗：${error.message}`;
    $("multiLocationImportResults").appendChild(div);
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

  function applyWorkspace(workspace, cases) {
    const normalized = normalizeWorkspace(workspace);
    state.currentWorkspace = normalized;
    state.cases = Array.isArray(cases) && cases.length ? cases : normalized ? [normalized.case] : [];
    state.callRecords = normalized?.records || [];
    state.callPage = 1;
    state.expandedHotspotAddress = "";
    state.hotspotCountySelection = new Set(ALL_COUNTY_FILTER_KEYS);
    state.hotspotCountyDraft = new Set(ALL_COUNTY_FILTER_KEYS);
    resetDateRangeState(state.callRecords);
    syncHotspotCountyFilterButton();
    syncDateFilterPanel();
    prefillSubmissionPhones(state.callRecords);
    renderAllViews();
  }

  function renderAllViews() {
    renderTwoWayCalls();
    renderProfileView();
    renderStatsView();
    renderHoursView();
    renderSubmissionPreview();
    syncDateFilterPanel();
  }

  function resetDateRangeState(records) {
    state.dateRangeBounds = computeDateRangeBounds(records);
    state.dateRange = { ...state.dateRangeBounds, active: false };
    state.dateRangeDraft = { ...state.dateRangeBounds };
  }

  function computeDateRangeBounds(records) {
    const dates = (records || []).map((record) => validRecordDate(record?.occurred_at)).filter(Boolean).sort();
    return { start: dates[0] || "", end: dates[dates.length - 1] || "" };
  }

  function filterRecordsByDateRange(records, range) {
    const input = Array.isArray(records) ? records : [];
    if (!range?.active) return [...input];
    const start = validDateText(range.start);
    const end = validDateText(range.end);
    if (!start || !end || start > end) return [];
    return input.filter((record) => {
      const date = validRecordDate(record?.occurred_at);
      return date && date >= start && date <= end;
    });
  }

  function validRecordDate(value) {
    const normalized = normalizeDatetime(value);
    return normalized ? validDateText(normalized.slice(0, 10)) : "";
  }

  function validDateText(value) {
    const text = cellText(value);
    const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return "";
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    return date.getUTCFullYear() === year && date.getUTCMonth() + 1 === month && date.getUTCDate() === day ? text : "";
  }

  function analysisRecords() {
    return filterRecordsByDateRange(state.callRecords, state.dateRange);
  }

  function analysisWorkspace() {
    if (!state.currentWorkspace) return null;
    if (!state.dateRange.active) return state.currentWorkspace;
    const records = analysisRecords();
    const stationKeys = new Set(records.flatMap((record) => (record.base_refs || []).map((ref) => ref.station_key).filter(Boolean)));
    const stations = (state.currentWorkspace.base_stations || []).filter((station) => stationKeys.has(station.station_key || stationKey(station)));
    return {
      ...state.currentWorkspace,
      case: {
        ...state.currentWorkspace.case,
        total_records: records.length,
        summary: parsedSummary(records, stations),
      },
      records,
      base_stations: stations,
    };
  }

  function invalidDateRecordCount(records) {
    return (records || []).reduce((count, record) => count + (validRecordDate(record?.occurred_at) ? 0 : 1), 0);
  }

  function syncDateFilterPanel() {
    if (typeof document === "undefined") return;
    const panel = $("dateFilterPanel");
    if (!panel) return;
    const hasWorkspace = Boolean(state.currentWorkspace);
    panel.hidden = !hasWorkspace;
    if (!hasWorkspace) return;
    const hasBounds = Boolean(state.dateRangeBounds.start && state.dateRangeBounds.end);
    const button = $("dateFilterButton");
    if (button) button.disabled = !hasBounds;
    let label = hasBounds ? `${state.dateRangeBounds.start} 至 ${state.dateRangeBounds.end}` : "沒有可解析日期";
    if (state.dateRange.active) {
      const excluded = invalidDateRecordCount(state.callRecords);
      label = `${state.dateRange.start} 至 ${state.dateRange.end}${excluded ? `；排除 ${excluded.toLocaleString()} 筆日期異常資料` : ""}`;
    }
    if ($("dateFilterSummary")) $("dateFilterSummary").textContent = label;
    if (button) {
      button.classList.toggle("active", state.dateRange.active);
      button.title = state.dateRange.active ? `時間篩選：${label}` : "時間篩選";
      button.setAttribute("aria-label", state.dateRange.active ? `時間篩選已套用，${label}` : `時間篩選，完整範圍 ${label}`);
    }
  }

  function showDateFilterModal() {
    const modal = $("dateFilterModal");
    if (!modal || !state.currentWorkspace || !state.dateRangeBounds.start || !state.dateRangeBounds.end) return;
    state.dateRangeDraft = {
      start: state.dateRange.active ? state.dateRange.start : state.dateRangeBounds.start,
      end: state.dateRange.active ? state.dateRange.end : state.dateRangeBounds.end,
    };
    const startInput = $("dateFilterStartInput");
    const endInput = $("dateFilterEndInput");
    if (startInput) {
      startInput.min = state.dateRangeBounds.start;
      startInput.max = state.dateRangeBounds.end;
      startInput.value = state.dateRangeDraft.start;
    }
    if (endInput) {
      endInput.min = state.dateRangeBounds.start;
      endInput.max = state.dateRangeBounds.end;
      endInput.value = state.dateRangeDraft.end;
    }
    syncDateFilterDraftUi();
    modal.hidden = false;
    $("dateFilterButton")?.setAttribute("aria-expanded", "true");
    startInput?.focus();
  }

  function hideDateFilterModal() {
    const modal = $("dateFilterModal");
    if (!modal || modal.hidden) return;
    state.dateRangeDraft = {
      start: state.dateRange.active ? state.dateRange.start : state.dateRangeBounds.start,
      end: state.dateRange.active ? state.dateRange.end : state.dateRangeBounds.end,
    };
    modal.hidden = true;
    $("dateFilterButton")?.setAttribute("aria-expanded", "false");
    $("dateFilterButton")?.focus();
  }

  function syncDateFilterDraftUi() {
    const start = cellText($("dateFilterStartInput")?.value);
    const end = cellText($("dateFilterEndInput")?.value);
    state.dateRangeDraft = { start, end };
    const withinBounds = (!state.dateRangeBounds.start || start >= state.dateRangeBounds.start)
      && (!state.dateRangeBounds.end || end <= state.dateRangeBounds.end);
    const valid = Boolean(validDateText(start) && validDateText(end) && start <= end && withinBounds);
    if ($("dateFilterApplyButton")) $("dateFilterApplyButton").disabled = !valid;
    if ($("dateFilterStatus")) {
      $("dateFilterStatus").textContent = !start || !end
        ? "請選擇起始與結束日期。"
        : start > end
          ? "起始日期不得晚於結束日期。"
          : !withinBounds
            ? `日期須介於 ${state.dateRangeBounds.start} 與 ${state.dateRangeBounds.end}。`
            : "日期範圍包含起始及結束整日。";
      $("dateFilterStatus").classList.toggle("danger-text", !valid);
    }
  }

  function applyDateFilter() {
    const start = validDateText(state.dateRangeDraft.start);
    const end = validDateText(state.dateRangeDraft.end);
    if (!start || !end || start > end || start < state.dateRangeBounds.start || end > state.dateRangeBounds.end) return;
    state.dateRange = { start, end, active: true };
    state.callPage = 1;
    state.expandedHotspotAddress = "";
    const modal = $("dateFilterModal");
    if (modal) modal.hidden = true;
    $("dateFilterButton")?.setAttribute("aria-expanded", "false");
    renderAllViews();
    $("dateFilterButton")?.focus();
  }

  function resetDateFilter() {
    state.dateRange = { ...state.dateRangeBounds, active: false };
    state.dateRangeDraft = { ...state.dateRangeBounds };
    state.callPage = 1;
    state.expandedHotspotAddress = "";
    const modal = $("dateFilterModal");
    if (modal) modal.hidden = true;
    $("dateFilterButton")?.setAttribute("aria-expanded", "false");
    renderAllViews();
    $("dateFilterButton")?.focus();
  }

  function renderProfileSummaryCards() {
    const summary = analysisWorkspace()?.case?.summary || {};
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
    return `<span class="phone-value">${escapeHtml(normalized)}</span>`;
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
        record.source_file,
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
    const pageCount = Math.max(1, Math.ceil(rows.length / CALL_PAGE_SIZE));
    state.callPage = Math.min(Math.max(1, state.callPage), pageCount);
    const start = (state.callPage - 1) * CALL_PAGE_SIZE;
    const pageRows = rows.slice(start, start + CALL_PAGE_SIZE);
    $("callRows").innerHTML = rows.length
      ? pageRows.map((record) => `<tr>
          <td>${escapeHtml(record.source_file || "")}</td>
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
      : `<tr><td colspan="10">尚未匯入資料。</td></tr>`;
    renderCallPagination(rows.length, start, pageRows.length, pageCount);
    applyCallColumnWidths();
  }

  function renderCallPagination(total, start, pageLength, pageCount) {
    const summary = $("callPageSummary");
    if (summary) {
      summary.textContent = total
        ? `第 ${(start + 1).toLocaleString()}-${(start + pageLength).toLocaleString()} 筆，共 ${total.toLocaleString()} 筆（第 ${state.callPage}/${pageCount} 頁）`
        : "共 0 筆";
    }
    if ($("callPrevPage")) $("callPrevPage").disabled = state.callPage <= 1;
    if ($("callNextPage")) $("callNextPage").disabled = state.callPage >= pageCount;
  }

  function changeCallPage(delta) {
    state.callPage = Math.max(1, state.callPage + Number(delta || 0));
    renderTwoWayCalls();
  }

  function sortedCallRecords() {
    const { column, direction } = state.callSort;
    return analysisRecords().sort((a, b) => {
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
    state.callPage = 1;
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
    const imeis = collectUniqueImeis(analysisRecords());
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
    const stats = computePhoneStats(analysisRecords(), state.phoneStatsRankMode);
    $("statsContent").innerHTML = [
      statsCard("來電排行", stats.inboundRows),
      statsCard("去電排行", stats.outboundRows),
      statsCard("完整排行", stats.totalRows),
    ].join("");
  }

  function renderMultiLocationView() {
    const rowsTarget = $("multiLocationRows");
    if (!rowsTarget) return;
    const matches = state.multiLocationMatches || [];
    const totalPages = Math.max(1, Math.ceil(matches.length / MULTI_LOCATION_PAGE_SIZE));
    state.multiLocationPage = Math.min(Math.max(1, state.multiLocationPage), totalPages);
    const startIndex = (state.multiLocationPage - 1) * MULTI_LOCATION_PAGE_SIZE;
    const pageRows = matches.slice(startIndex, startIndex + MULTI_LOCATION_PAGE_SIZE);
    const excluded = Object.values(state.multiLocationExcluded || {}).reduce((sum, value) => sum + Number(value || 0), 0);
    const summary = $("multiLocationSummary");
    if (summary) summary.textContent = excluded
      ? `符合 ${matches.length.toLocaleString()} 筆｜排除 ${excluded.toLocaleString()} 項無法比對資料`
      : `符合 ${matches.length.toLocaleString()} 筆`;
    if (!state.multiLocationWorkspace) {
      rowsTarget.innerHTML = `<tr><td colspan="6" class="muted">尚未匯入資料。</td></tr>`;
    } else if (!pageRows.length) {
      rowsTarget.innerHTML = `<tr><td colspan="6" class="muted">目前沒有不同門號在同一行政區 30 分鐘內出現的結果。</td></tr>`;
    } else {
      rowsTarget.innerHTML = pageRows.map(renderMultiLocationMatchRows).join("");
    }
    const pageSummary = $("multiLocationPageSummary");
    if (pageSummary) pageSummary.textContent = matches.length
      ? `第 ${state.multiLocationPage.toLocaleString()} / ${totalPages.toLocaleString()} 頁，共 ${matches.length.toLocaleString()} 筆`
      : "共 0 筆";
    if ($("multiLocationPrevPage")) $("multiLocationPrevPage").disabled = state.multiLocationPage <= 1;
    if ($("multiLocationNextPage")) $("multiLocationNextPage").disabled = state.multiLocationPage >= totalPages;
  }

  function changeMultiLocationPage(delta) {
    state.multiLocationPage += Number(delta || 0);
    renderMultiLocationView();
  }

  function toggleMultiLocationDetail(matchId) {
    if (!matchId) return;
    if (state.multiLocationExpandedMatches.has(matchId)) state.multiLocationExpandedMatches.delete(matchId);
    else state.multiLocationExpandedMatches.add(matchId);
    renderMultiLocationView();
    $(`multiLocationDetailToggle-${matchId}`)?.focus();
  }

  function renderMultiLocationMatchRows(match) {
    const matchId = cellText(match.id);
    const sourceRecords = Array.isArray(match.source_records) ? match.source_records : [];
    const expanded = state.multiLocationExpandedMatches.has(matchId);
    const detailId = `multiLocationDetail-${matchId}`;
    const toggleId = `multiLocationDetailToggle-${matchId}`;
    const summaryRow = `<tr class="multi-location-match-row">
      <td>${escapeHtml(formatMultiLocationTimeRange(match.start_at, match.end_at))}</td>
      <td>${escapeHtml(match.county)}</td>
      <td>${escapeHtml(match.district)}</td>
      <td><div class="multi-location-phone-list">${match.phones.map((phone) => `<span class="phone-value">${escapeHtml(phone)}</span>`).join("")}</div></td>
      <td><div class="multi-location-note-list">${match.phones.map((phone) => `<label><span>${escapeHtml(phone)}</span><input class="phone-note-input" data-phone-note="${escapeHtml(phone)}" value="${escapeHtml(phoneNote(phone))}" aria-label="備註(只存瀏覽器) ${escapeHtml(phone)}" /></label>`).join("")}</div></td>
      <td><button id="${escapeHtml(toggleId)}" class="multi-location-detail-toggle" data-multi-location-detail="${escapeHtml(matchId)}" type="button" aria-expanded="${expanded ? "true" : "false"}" aria-controls="${escapeHtml(detailId)}">${expanded ? "收合" : `展開（${sourceRecords.length.toLocaleString()} 筆）`}</button></td>
    </tr>`;
    if (!expanded) return summaryRow;
    return `${summaryRow}<tr id="${escapeHtml(detailId)}" class="multi-location-detail-row"><td colspan="6">${renderMultiLocationSourceRecords(sourceRecords)}</td></tr>`;
  }

  function renderMultiLocationSourceRecords(records) {
    if (!records.length) return `<p class="muted">沒有可顯示的原始資料。</p>`;
    return `<div class="multi-location-source-wrap">
      <table class="multi-location-source-table">
        <thead><tr><th>來源位置</th><th>發生時間</th><th>調閱門號</th><th>通話對象</th><th>基地台</th></tr></thead>
        <tbody>${records.map((record) => `<tr>
          <td>${renderMultiLocationSourcePosition(record)}</td>
          <td>${escapeHtml(record.occurred_at || "")}</td>
          <td><span class="phone-value">${escapeHtml(record.target_phone || "")}</span></td>
          <td><span class="phone-value">${escapeHtml(record.counterparty_phone || "")}</span></td>
          <td><div class="multi-location-source-stations">${(record.matched_stations || []).map((station) => `<span><strong>${escapeHtml(multiLocationStationRoleLabel(station.role))}</strong>${escapeHtml(station.cell_id || "無代碼")}｜${escapeHtml(station.address || "")}</span>`).join("")}</div></td>
        </tr>`).join("")}</tbody>
      </table>
    </div>`;
  }

  function renderMultiLocationSourcePosition(record) {
    const file = cellText(record.source_file) || "未標示來源檔案";
    const sheet = cellText(record.source_sheet) || "未標示工作表";
    const row = cellText(record.row_number);
    return `<div class="multi-location-source-position"><strong>${escapeHtml(file)}</strong><span>${escapeHtml(sheet)}${row ? `｜第 ${escapeHtml(row)} 列` : ""}</span></div>`;
  }

  function multiLocationStationRoleLabel(role) {
    const labels = { start: "起始：", end: "終止：", primary: "基地台：" };
    return labels[cellText(role).toLowerCase()] || (cellText(role) ? `${cellText(role)}：` : "基地台：");
  }

  function formatMultiLocationTimeRange(start, end) {
    const startText = cellText(start).replace("T", " ");
    const endText = cellText(end).replace("T", " ");
    return startText === endText ? startText : `${startText} ～ ${endText}`;
  }

  function statsCard(title, rows) {
    return `<section class="stats-card"><h3>${escapeHtml(title)}</h3>${
      rows.length
        ? `<div class="stats-table" role="table">
            <div class="stats-table-row stats-table-head" role="row"><span>#</span><span>電話</span><span>備註(只存瀏覽器)</span><span>次數</span><span>秒數</span></div>
            ${rows.map((row, index) => `<div class="stats-table-row" role="row">
              <span>${index + 1}</span>
              <span><span class="phone-value">${escapeHtml(row.phone)}</span></span>
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
    return analysisRecords().filter((record) => {
      const hour = hourFromIso(record.occurred_at);
      return hour >= 0 && allowed.has(hour);
    });
  }

  function renderHourHotspots(records) {
    const query = ($("hourHotspotSearch")?.value || "").trim().toLowerCase();
    const hotspots = computeAddressHotspots(records, analysisWorkspace()?.base_stations || []);
    const countyFiltered = hotspots.filter((item) => state.hotspotCountySelection.has(classifyTaiwanCounty(item.address)));
    const filtered = query
      ? countyFiltered.filter((item) => [item.address, item.first_seen, item.last_seen, ...item.times].some((value) => String(value || "").toLowerCase().includes(query)))
      : countyFiltered;
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

  function showHotspotCountyFilterModal() {
    const modal = $("hotspotCountyFilterModal");
    if (!modal) return;
    state.hotspotCountyDraft = new Set(state.hotspotCountySelection);
    renderHotspotCountyFilterModal();
    modal.hidden = false;
    $("hotspotCountyFilterButton")?.setAttribute("aria-expanded", "true");
    $("hotspotCountyFilterCloseButton")?.focus();
  }

  function hideHotspotCountyFilterModal() {
    const modal = $("hotspotCountyFilterModal");
    if (!modal || modal.hidden) return;
    state.hotspotCountyDraft = new Set(state.hotspotCountySelection);
    modal.hidden = true;
    $("hotspotCountyFilterButton")?.setAttribute("aria-expanded", "false");
    $("hotspotCountyFilterButton")?.focus();
  }

  function renderHotspotCountyFilterModal() {
    const list = $("hotspotCountyFilterList");
    if (!list) return;
    const hotspots = computeAddressHotspots(filteredHourRecords(), analysisWorkspace()?.base_stations || []);
    const rows = computeTaiwanCountyStats(hotspots);
    list.innerHTML = rows.map((row) => `<label class="county-filter-row">
      <span class="county-filter-name"><input type="checkbox" data-county-filter="${escapeHtml(row.county)}" ${state.hotspotCountyDraft.has(row.county) ? "checked" : ""} /><span>${escapeHtml(row.county)}</span></span>
      <span class="county-filter-count">${row.count.toLocaleString()}</span>
      <span class="county-filter-percent">${formatPercent(row.percent)}</span>
    </label>`).join("");
    syncHotspotCountyFilterDraftUi();
  }

  function handleHotspotCountyDraftChange(event) {
    const checkbox = event.target.closest("[data-county-filter]");
    if (!checkbox) return;
    if (checkbox.checked) state.hotspotCountyDraft.add(checkbox.dataset.countyFilter);
    else state.hotspotCountyDraft.delete(checkbox.dataset.countyFilter);
    syncHotspotCountyFilterDraftUi();
  }

  function syncHotspotCountyFilterDraftUi() {
    const selected = state.hotspotCountyDraft.size;
    const applyButton = $("hotspotCountyFilterApplyButton");
    if (applyButton) applyButton.disabled = selected === 0;
    const status = $("hotspotCountyFilterStatus");
    if (status) {
      status.textContent = selected ? `已選取 ${selected} 個分類。` : "請至少選擇一個縣市或未辨識。";
      status.classList.toggle("danger-text", selected === 0);
    }
  }

  function selectAllHotspotCountyDraft() {
    state.hotspotCountyDraft = new Set(ALL_COUNTY_FILTER_KEYS);
    renderHotspotCountyFilterModal();
  }

  function clearAllHotspotCountyDraft() {
    state.hotspotCountyDraft = new Set();
    renderHotspotCountyFilterModal();
  }

  function applyHotspotCountyFilter() {
    if (!state.hotspotCountyDraft.size) return;
    state.hotspotCountySelection = new Set(state.hotspotCountyDraft);
    state.expandedHotspotAddress = "";
    renderHourHotspots(filteredHourRecords());
    syncHotspotCountyFilterButton();
    hideHotspotCountyFilterModal();
  }

  function syncHotspotCountyFilterButton() {
    const button = $("hotspotCountyFilterButton");
    if (!button) return;
    const selected = state.hotspotCountySelection.size;
    const allSelected = selected === ALL_COUNTY_FILTER_KEYS.length;
    button.classList.toggle("filtered", !allSelected);
    button.setAttribute("aria-label", allSelected ? "縣市篩選，已選取全部分類" : `縣市篩選，已選取 ${selected} 個分類`);
    button.title = allSelected ? "縣市篩選（全部）" : `縣市篩選（已選 ${selected}）`;
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
        renderMultiLocationView();
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

  function showAttachmentExportModal() {
    const modal = $("attachmentExportModal");
    if (!modal) return;
    modal.hidden = false;
    const hasData = Boolean(analysisWorkspace()?.records?.length);
    $("attachmentXlsxButton").disabled = !hasData;
    document.querySelectorAll("[data-attachment-pdf]").forEach((button) => { button.disabled = !hasData; });
    setAttachmentExportStatus(hasData ? "請選擇要下載的附卷格式。" : "尚未匯入資料，無法產生附卷檔案。", !hasData);
    $("attachmentExportCloseButton")?.focus();
  }

  function hideAttachmentExportModal() {
    const modal = $("attachmentExportModal");
    if (modal) modal.hidden = true;
  }

  async function downloadAttachmentXlsx() {
    const workspace = analysisWorkspace();
    if (!workspace?.records?.length || !AttachmentExport) return;
    setAttachmentExportBusy(true);
    setAttachmentExportStatus("正在產生六分頁 XLSX，資料只在瀏覽器本機處理。", false);
    try {
      await loadAttachmentAsset("exceljs");
      const report = buildAttachmentReport(workspace, state.phoneNotes, new Date().toISOString(), attachmentReportOptions());
      const chartDataUrl = createHourChartDataUrl(report.hours);
      const bytes = await AttachmentExport.createAttachmentXlsx(report, globalThis.ExcelJS, chartDataUrl);
      downloadBlob(`附卷檔案-${attachmentFileStamp()}.xlsx`, bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      setAttachmentExportStatus("XLSX 已下載。", false);
    } catch (error) {
      setAttachmentExportStatus(`XLSX 產生失敗：${error.message}`, true);
    } finally {
      setAttachmentExportBusy(false);
    }
  }

  async function downloadAttachmentPdf(sectionKey) {
    const workspace = analysisWorkspace();
    if (!workspace?.records?.length || !AttachmentExport) return;
    const section = AttachmentExport.PDF_SECTIONS.find((item) => item.key === sectionKey);
    if (!section) return;
    setAttachmentExportBusy(true);
    setAttachmentExportStatus(`正在產生「${section.label}」PDF，資料只在瀏覽器本機處理。`, false);
    try {
      await loadAttachmentAsset("pdfLib");
      await loadAttachmentAsset("fontkit");
      await loadAttachmentAsset("fontData");
      const report = buildAttachmentReport(workspace, state.phoneNotes, new Date().toISOString(), attachmentReportOptions());
      const fontBytes = base64ToBytes(globalThis.PhoneExportFontBase64);
      const bytes = await AttachmentExport.createAttachmentPdf(report, sectionKey, globalThis.PDFLib, globalThis.fontkit, fontBytes);
      downloadBlob(`附卷-${section.label}-${attachmentFileStamp()}.pdf`, bytes, "application/pdf");
      setAttachmentExportStatus(`「${section.label}」PDF 已下載。`, false);
    } catch (error) {
      setAttachmentExportStatus(`PDF 產生失敗：${error.message}`, true);
    } finally {
      setAttachmentExportBusy(false);
    }
  }

  function setAttachmentExportBusy(busy) {
    if ($("attachmentXlsxButton")) $("attachmentXlsxButton").disabled = busy;
    document.querySelectorAll("[data-attachment-pdf]").forEach((button) => { button.disabled = busy; });
    if ($("attachmentExportCloseButton")) $("attachmentExportCloseButton").disabled = busy;
    $("attachmentExportModal")?.setAttribute("aria-busy", busy ? "true" : "false");
  }

  function attachmentReportOptions() {
    if (!state.dateRange.active) return {};
    return {
      scope: "date_filter",
      scope_label: `日期篩選：${state.dateRange.start} 至 ${state.dateRange.end}`,
      date_range: { start: state.dateRange.start, end: state.dateRange.end },
    };
  }

  function setAttachmentExportStatus(message, danger) {
    const target = $("attachmentExportStatus");
    if (!target) return;
    target.textContent = message;
    target.classList.toggle("danger-text", Boolean(danger));
    target.classList.toggle("success-text", !danger);
  }

  function loadAttachmentAsset(key) {
    if (loadedAttachmentAssets.has(key)) return loadedAttachmentAssets.get(key);
    const asset = ATTACHMENT_ASSETS[key];
    if (!asset) return Promise.reject(new Error("未知的本地匯出資產"));
    const promise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = asset.src;
      script.integrity = asset.integrity;
      script.crossOrigin = "anonymous";
      script.onload = () => resolve();
      script.onerror = () => {
        loadedAttachmentAssets.delete(key);
        reject(new Error("本地匯出資產載入失敗"));
      };
      document.head.appendChild(script);
    });
    loadedAttachmentAssets.set(key, promise);
    return promise;
  }

  function createHourChartDataUrl(hours) {
    const canvas = document.createElement("canvas");
    canvas.width = 1400;
    canvas.height = 520;
    const context = canvas.getContext("2d");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.fillStyle = "#171717";
    context.font = "bold 28px sans-serif";
    context.fillText("24 小時通聯分布", 54, 48);
    const max = Math.max(1, ...hours.map((item) => item.count));
    const chartLeft = 54;
    const chartTop = 86;
    const chartHeight = 340;
    const gap = 10;
    const barWidth = (canvas.width - chartLeft * 2 - gap * 23) / 24;
    context.textAlign = "center";
    hours.forEach((item, index) => {
      const height = item.count ? (item.count / max) * chartHeight : 0;
      const x = chartLeft + index * (barWidth + gap);
      context.fillStyle = "#171717";
      context.fillRect(x, chartTop + chartHeight - height, barWidth, height);
      context.fillStyle = "#60646c";
      context.font = "18px sans-serif";
      context.fillText(String(index).padStart(2, "0"), x + barWidth / 2, chartTop + chartHeight + 28);
      context.fillStyle = "#171717";
      context.font = "bold 16px sans-serif";
      context.fillText(String(item.count), x + barWidth / 2, chartTop + chartHeight - height - 8);
    });
    context.textAlign = "left";
    return canvas.toDataURL("image/png");
  }

  function base64ToBytes(value) {
    const binary = atob(String(value || ""));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }

  function attachmentFileStamp(date = new Date()) {
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}`;
  }

  function parseImportFile(fileName, inputBytes) {
    const bytes = toUint8Array(inputBytes);
    const lower = String(fileName || "").toLowerCase();
    if (lower.endsWith(".xml") || looksLikeXml(bytes)) {
      return makeWorkspace(parseXmlWorkbook(fileName, decodeXmlBytes(bytes)));
    }
    if (!XLSX) throw new Error("XLSX parser is not available");
    return makeWorkspace(parseXlsxWorkbook(fileName, bytes));
  }

  function parseXlsxWorkbook(fileName, bytes) {
    const workbook = XLSX.read(bytes, { type: "array", cellDates: false });
    const sheets = workbook.SheetNames.map((title) => {
      const displayRows = XLSX.utils.sheet_to_json(workbook.Sheets[title], { header: 1, defval: "", raw: false, blankrows: false });
      const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[title], { header: 1, defval: "", raw: true, blankrows: false });
      return {
        title,
        rows: displayRows.map((values, index) => ({
          rowNumber: index + 1,
          values: values.map(cellText),
          rawValues: rawRows[index] || values,
        })),
      };
    });
    const fetOrder = fetOrderHeaders(sheets);
    if (fetOrder.length) return parseFetOrderXlsx(fileName, sheets, fetOrder);
    const chtProsecutor = chtProsecutorHeaders(sheets);
    if (chtProsecutor) return parseChtProsecutor(fileName, sheets, chtProsecutor);
    const fetProsecutorCall = fetProsecutorCallHeaders(sheets);
    if (fetProsecutorCall) return parseFetProsecutorCall(fileName, sheets, fetProsecutorCall);
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

  function parseChtProsecutor(fileName, sheets, header) {
    const subject = metadataBefore(header.sheet.rows, header.rowNumber);
    const queryPhone = normalizePhoneText(subject["電話號碼"] || subject["設備號碼"]);
    const parsed = makeParsed({
      fileName,
      carrier: "中華電信",
      sourceFormat: "chunghwa_prosecutor_cdr_xlsx",
      sheetName: header.sheet.title,
      headerRow: header.rowNumber,
      totalSourceRows: sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0),
      subject,
    });
    const stationMap = new Map();
    forRowsAfter(header, (rowNumber, data) => {
      const rawCallType = cellText(data["CDR類別"]);
      const rawOccurredAt = cellText(data["始話日期時間"]);
      const caller = normalizePhoneText(data["主叫號碼"]);
      const callee = normalizePhoneText(data["受叫號碼"]);
      if (!rawCallType && !rawOccurredAt && !caller && !callee) return;
      if (canonicalHeaderText(rawCallType) === "CDR類別") return;

      const occurredAt = normalizeDatetime(rawOccurredAt);
      if (rawOccurredAt && !occurredAt) parsed.warnings.push(`第 ${rowNumber} 列的始話日期時間無法正規化，已保留原文。`);

      const rawDirection = directionLabel(rawCallType);
      let direction = rawDirection;
      let targetPhone = "";
      let counterpartyPhone = "";
      if (queryPhone && caller === queryPhone && callee !== queryPhone) {
        direction = "outbound";
        targetPhone = queryPhone;
        counterpartyPhone = callee;
      } else if (queryPhone && callee === queryPhone && caller !== queryPhone) {
        direction = "inbound";
        targetPhone = queryPhone;
        counterpartyPhone = caller;
      } else if (rawDirection === "inbound") {
        targetPhone = callee || queryPhone || caller;
        counterpartyPhone = caller && caller !== targetPhone ? caller : "";
      } else {
        targetPhone = caller || queryPhone || callee;
        counterpartyPhone = callee && callee !== targetPhone ? callee : "";
      }

      const note = [
        cellText(data["查詢狀態"]) ? `查詢狀態：${cellText(data["查詢狀態"])}` : "",
        cellText(data["指定轉接"]) ? `指定轉接：${cellText(data["指定轉接"])}` : "",
      ].filter(Boolean).join("；");
      const record = baseRecord({
        row_number: rowNumber,
        call_type: rawCallType,
        direction,
        occurred_at: occurredAt || rawOccurredAt,
        duration_seconds: toInt(data["通話秒數"]),
        target_phone: targetPhone,
        counterparty_phone: counterpartyPhone,
        imei: cellText(data["IMEI"]),
        note,
      });
      addChtStationRefs(record, stationMap, data["起始基地台-地址/終止基地台-地址"]);
      parsed.records.push(record);
    });
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseFetProsecutorCall(fileName, sheets, firstHeader) {
    const parsed = makeParsed({
      fileName,
      carrier: "遠傳電信",
      sourceFormat: "fet_prosecutor_call_xlsx",
      sheetName: firstHeader.sheet.title,
      headerRow: firstHeader.rowNumber,
      totalSourceRows: sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0),
      subject: {},
    });
    const subjectValues = new Map();
    const stationMap = new Map();

    sheets.forEach((sheet) => {
      let activeHeaders = null;
      let lastRecord = null;
      sheet.rows.forEach((row) => {
        const canonicalValues = row.values.map(canonicalHeaderText);
        if (hasHeaders(canonicalValues, FET_PROSECUTOR_CALL_HEADERS)) {
          activeHeaders = canonicalValues;
          lastRecord = null;
          return;
        }

        if (isFetProsecutorMetadataRow(row.values)) {
          mergeSubjectMetadata(subjectValues, metadataFromValuesWithAdjacent(row.values, FET_PROSECUTOR_METADATA_KEYS));
          lastRecord = null;
          return;
        }
        if (!activeHeaders) return;

        const data = rowDict(activeHeaders, row.values);
        const rawOccurredAt = cellText(data["始話時間"]);
        const rawCallType = cellText(data["通話類別"]);
        const hasPrimaryRecordData = [
          rawOccurredAt,
          data["通話秒數"],
          data["調閱號碼"],
          data["IMEI"],
          rawCallType,
          data["通話對象"],
          data["轉接電話"],
          data["備註"],
        ].some((value) => cellText(value));
        const stationValue = cellText(data["基地台/交換機"]);
        if (!hasPrimaryRecordData) {
          if (stationValue && lastRecord) addStationCompoundRef(lastRecord, stationMap, "primary", stationValue);
          else if (stationValue) parsed.warnings.push(`第 ${row.rowNumber} 列的基地台續行沒有可附加的通聯列，已略過。`);
          return;
        }

        const occurredAt = normalizeDatetime(rawOccurredAt);
        if (rawOccurredAt && !occurredAt) parsed.warnings.push(`第 ${row.rowNumber} 列的始話時間無法正規化，已保留原文。`);
        const transferPhone = cellText(data["轉接電話"]);
        const note = [
          transferPhone ? `轉接電話：${transferPhone}` : "",
          cellText(data["備註"]),
        ].filter(Boolean).join("；");
        const record = baseRecord({
          row_number: row.rowNumber,
          source_sheet: sheet.title,
          call_type: rawCallType,
          direction: directionLabel(rawCallType),
          occurred_at: occurredAt || rawOccurredAt,
          duration_seconds: toInt(data["通話秒數"]),
          target_phone: normalizePhoneText(data["調閱號碼"]),
          counterparty_phone: normalizePhoneText(data["通話對象"]),
          imei: cellText(data["IMEI"]),
          note,
        });
        addStationCompoundRef(record, stationMap, "primary", data["基地台/交換機"]);
        parsed.records.push(record);
        lastRecord = record;
      });
    });

    parsed.subject = subjectFromValueSets(subjectValues);
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function parseFetOrderXlsx(fileName, sheets, headers) {
    const parsed = makeParsed({
      fileName,
      carrier: "遠傳電信",
      sourceFormat: "fet_order_cdr_xlsx",
      sheetName: headers.map((header) => header.sheet.title).join("、"),
      headerRow: headers[0].rowNumber,
      totalSourceRows: headers.reduce((sum, header) => sum + header.sheet.rows.length, 0),
      subject: {},
    });
    const subjectValues = new Map();
    const stationMap = new Map();
    headers.forEach((header) => {
      header.sheet.rows.forEach((row) => {
        if (row.rowNumber <= header.rowNumber) return;
        const display = rowDict(header.headers, row.values);
        const raw = rowDict(header.headers, row.rawValues || row.values);
        if (hasHeaders(row.values, FET_ORDER_HEADERS)) return;
        mergeFetOrderSubject(subjectValues, raw);
        if (!fetOrderRowHasCdr(raw)) return;
        const record = fetOrderRecord(raw, row.rowNumber, header.sheet.title, parsed.warnings);
        addStationRef(record, stationMap, "start", cleanCellId(raw.CellStartID), display.CellStarted || raw.CellStarted);
        addStationRef(record, stationMap, "end", cleanCellId(raw.CellEndID), display.CellEnded || raw.CellEnded);
        parsed.records.push(record);
      });
    });
    parsed.subject = subjectFromValueSets(subjectValues);
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
  }

  function fetOrderRowHasCdr(data) {
    return ["CallStartTimeStamp", "Duration", "CallDirection", "CallingNumber", "CalledNumber", "IMEI", "CallFwd", "CellStartID", "CellEndID"]
      .some((key) => cellText(data[key]));
  }

  function fetOrderRecord(data, rowNumber, sourceSheet, warnings) {
    const rawOccurredAt = cellText(data.CallStartTimeStamp);
    const occurredAt = normalizeDatetime(rawOccurredAt);
    if (rawOccurredAt && !occurredAt) warnings.push(`第 ${rowNumber} 列的通聯時間無法正規化，已保留原文。`);
    const mapped = fetOrderPhonesAndDirection(data);
    const callDirection = cellText(data.CallDirection).toUpperCase();
    const callType = FET_ORDER_CALL_TYPES[callDirection] || (callDirection ? `其他（${callDirection}）` : "");
    const callFwd = cellText(data.CallFwd);
    return baseRecord({
      row_number: rowNumber,
      source_sheet: sourceSheet,
      occurred_at: occurredAt || rawOccurredAt,
      duration_seconds: toInt(data.Duration),
      call_type: callType,
      direction: mapped.direction,
      target_phone: mapped.targetPhone,
      counterparty_phone: mapped.counterpartyPhone,
      imei: exactNumericText(data.IMEI),
      note: callFwd ? `指定轉接：${callFwd}` : "",
    });
  }

  function fetOrderPhonesAndDirection(data) {
    const caller = normalizePhoneText(data.CallingNumber);
    const callee = normalizePhoneText(data.CalledNumber);
    const queryType = cellText(data.QueType);
    const queryPhone = queryType === "1" ? normalizePhoneText(data.QueObject) : "";
    if (queryPhone && caller === queryPhone && callee !== queryPhone) {
      return { direction: "outbound", targetPhone: caller, counterpartyPhone: callee };
    }
    if (queryPhone && callee === queryPhone && caller !== queryPhone) {
      return { direction: "inbound", targetPhone: callee, counterpartyPhone: caller };
    }
    const code = cellText(data.CallDirection).toUpperCase();
    const direction = ["O", "2"].includes(code) ? "outbound" : ["T", "I", "1", "9"].includes(code) ? "inbound" : "other";
    if (direction === "inbound") return { direction, targetPhone: callee || queryPhone || caller, counterpartyPhone: caller && caller !== callee ? caller : "" };
    return { direction, targetPhone: caller || queryPhone || callee, counterpartyPhone: callee && callee !== caller ? callee : "" };
  }

  function mergeFetOrderSubject(subjectValues, data) {
    const start = normalizeOrderRangeEndpoint(data.StartDate, data.StartTime);
    const end = normalizeOrderRangeEndpoint(data.EndDate, data.EndTime);
    const queryType = cellText(data.QueType);
    const entries = {
      "文號": data.DocNo,
      "案件狀態": data.Status,
      "案件備註": data.Comment,
      "查詢序號": data.Seq,
      "查詢狀態": data.Status2,
      "調閱類型": FET_ORDER_QUERY_TYPES[queryType] || data.QueType,
      "調閱目標": queryType === "1" ? normalizePhoneText(data.QueObject) : exactNumericText(data.QueObject),
      "查詢方向": FET_ORDER_QUERY_DIRECTIONS[cellText(data.QueDirection)] || data.QueDirection,
      "查詢起始": start,
      "查詢結束": end,
      "結果類型": FET_ORDER_RESULT_TYPES[cellText(data.ResultType)] || data.ResultType,
      "查詢備註": data.Memo,
    };
    mergeSubjectMetadata(subjectValues, entries);
  }

  function normalizeOrderRangeEndpoint(dateValue, timeValue) {
    const dateNumber = Number(dateValue);
    const timeNumber = Number(timeValue);
    if (Number.isFinite(dateNumber) && dateNumber >= 20000 && dateNumber <= 60000) {
      return excelSerialDatetime(dateNumber + (Number.isFinite(timeNumber) && timeNumber >= 0 && timeNumber < 1 ? timeNumber : 0));
    }
    const dateText = cellText(dateValue);
    const timeText = cellText(timeValue);
    return normalizeDatetime(`${dateText}${dateText && timeText ? " " : ""}${timeText}`) || [dateText, timeText].filter(Boolean).join(" ");
  }

  function exactNumericText(value) {
    const text = cellText(value);
    if (!text) return "";
    if (/^[+-]?\d+(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(text)) {
      const number = Number(text);
      if (Number.isFinite(number) && Number.isInteger(number)) return number.toLocaleString("fullwide", { useGrouping: false, maximumFractionDigits: 0 });
    }
    return text;
  }

  function parseXmlWorkbook(fileName, xml) {
    if (isFetOrderXml(xml)) return parseFetOrderXml(fileName, xml);
    if (isTwmCallXml(xml)) return parseTwmCallXml(fileName, xml);
    if (firstXmlBlock(xml, "CUSTOMERINFO") || firstXmlBlock(xml, "CELLINFO")) return parseTwmCspXml(fileName, xml);
    throw new Error("找不到支援的 XML 結構");
  }

  function parseFetOrderXml(fileName, xml) {
    const summary = xmlChildren(firstXmlBlock(xml, "SummaryInfo"));
    const parsed = makeParsed({
      fileName,
      carrier: "遠傳電信",
      sourceFormat: "fet_order_cdr_xml",
      sheetName: "XML",
      headerRow: 0,
      totalSourceRows: iterXmlBlocks(xml, "CDRInfo").length,
      subject: {},
    });
    const subjectValues = new Map();
    mergeFetOrderSubject(subjectValues, summary);
    const stationMap = new Map();
    let rowNumber = 0;
    iterXmlBlocks(xml, "Record").forEach((recordBlock) => {
      const query = xmlChildren(firstXmlBlock(recordBlock, "QueryInfo"));
      mergeFetOrderSubject(subjectValues, query);
      iterXmlBlocks(recordBlock, "CDRInfo").forEach((cdrBlock) => {
        rowNumber += 1;
        const cdr = { ...query, ...xmlChildren(cdrBlock) };
        if (!fetOrderRowHasCdr(cdr)) return;
        const record = fetOrderRecord(cdr, rowNumber, "XML", parsed.warnings);
        addStationRef(record, stationMap, "start", cleanCellId(cdr.CellStartID), cdr.CellStarted);
        addStationRef(record, stationMap, "end", cleanCellId(cdr.CellEndID), cdr.CellEnded);
        parsed.records.push(record);
      });
    });
    parsed.subject = subjectFromValueSets(subjectValues);
    parsed.base_stations = Array.from(stationMap.values());
    return parsed;
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
    const records = (parsed.records || []).map((record) => parsedRecordPayload(record, parsed.file_name, parsed.sheet_name));
    const summary = parsedSummary(records, stations);
    return {
      case: {
        source_file: parsed.file_name,
        source_files: parsed.file_name ? [parsed.file_name] : [],
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

  function parsedRecordPayload(record, sourceFile, sourceSheet) {
    const target = normalizePhoneText(record.target_phone);
    const counterparty = normalizePhoneText(record.counterparty_phone);
    return {
      row_number: record.row_number,
      source_file: record.source_file || sourceFile || "",
      source_sheet: record.source_sheet || sourceSheet || "",
      occurred_at: record.occurred_at || "",
      ended_at: record.ended_at || "",
      duration_seconds: toInt(record.duration_seconds),
      call_type: record.call_type || "",
      direction: normalizeDirection(record.direction, record.call_type),
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

  function normalizeWorkspace(workspace) {
    if (!workspace || !workspace.case || !Array.isArray(workspace.records)) return null;
    const sourceFiles = uniqueTextValues(workspace.case.source_files || [workspace.case.source_file]);
    const records = workspace.records.map((record) => parsedRecordPayload(
      record,
      record.source_file || workspace.case.source_file,
      record.source_sheet || workspace.case.sheet_name,
    ));
    const stations = dedupeStations(workspace.base_stations || []);
    const warnings = uniqueTextValues([...(workspace.parse_warnings || []), ...(workspace.case.parse_warnings || [])]);
    return {
      case: {
        ...workspace.case,
        source_file: workspace.case.source_file || sourceFiles[0] || "",
        source_files: sourceFiles,
        total_records: records.length,
        summary: parsedSummary(records, stations),
        parse_warnings: warnings,
      },
      records,
      base_stations: stations,
      parse_warnings: warnings,
    };
  }

  function mergeWorkspaces(workspaces) {
    const normalized = (workspaces || []).map(normalizeWorkspace).filter(Boolean);
    if (!normalized.length) return null;
    if (normalized.length === 1) return normalized[0];
    const records = normalized.flatMap((workspace) => workspace.records);
    const stations = dedupeStations(normalized.flatMap((workspace) => workspace.base_stations));
    const sourceFiles = uniqueTextValues(normalized.flatMap((workspace) => workspace.case.source_files || [workspace.case.source_file]));
    const warnings = uniqueTextValues(normalized.flatMap((workspace) => workspace.parse_warnings || []));
    const subject = mergeSubjects(normalized.map((workspace) => workspace.case.subject || {}));
    const carrierValues = uniqueTextValues(normalized.map((workspace) => workspace.case.carrier));
    const formatValues = uniqueTextValues(normalized.map((workspace) => workspace.case.source_format));
    const sheetValues = uniqueTextValues(normalized.map((workspace) => workspace.case.sheet_name));
    return {
      case: {
        source_file: sourceFiles[0] || "",
        source_files: sourceFiles,
        carrier: carrierValues.join("、"),
        source_format: formatValues.join("、"),
        sheet_name: sheetValues.join("、"),
        header_row: null,
        total_source_rows: normalized.reduce((sum, workspace) => sum + Number(workspace.case.total_source_rows || 0), 0),
        total_records: records.length,
        subject,
        summary: parsedSummary(records, stations),
        parse_warnings: warnings,
      },
      records,
      base_stations: stations,
      parse_warnings: warnings,
    };
  }

  function mergeSubjects(subjects) {
    const values = new Map();
    (subjects || []).forEach((subject) => {
      Object.entries(subject || {}).forEach(([key, value]) => {
        const set = values.get(key) || new Set();
        uniqueTextValues(String(value ?? "").split("、")).forEach((item) => set.add(item));
        values.set(key, set);
      });
    });
    return Object.fromEntries(Array.from(values, ([key, set]) => [key, Array.from(set).join("、")]));
  }

  function dedupeStations(stations) {
    const map = new Map();
    (stations || []).forEach((station) => {
      if (!station) return;
      const key = station.station_key || stationKey(station);
      if (key && !map.has(key)) map.set(key, { ...station, station_key: key });
    });
    return Array.from(map.values());
  }

  function uniqueTextValues(values) {
    return Array.from(new Set((values || []).map(cellText).filter(Boolean)));
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
      source_file: values.source_file || "",
      source_sheet: values.source_sheet || "",
      call_type: values.call_type || "",
      direction: values.direction || "",
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

  function chtProsecutorHeaders(sheets) {
    for (const sheet of sheets) {
      for (const row of sheet.rows.slice(0, 80)) {
        const headers = row.values.map(canonicalHeaderText);
        if (hasHeaders(headers, CHT_PROSECUTOR_HEADERS)) {
          return { sheet, rowNumber: row.rowNumber, headers };
        }
      }
    }
    return null;
  }

  function fetProsecutorCallHeaders(sheets) {
    for (const sheet of sheets) {
      for (const row of sheet.rows.slice(0, 80)) {
        const headers = row.values.map(canonicalHeaderText);
        if (hasHeaders(headers, FET_PROSECUTOR_CALL_HEADERS)) {
          return { sheet, rowNumber: row.rowNumber, headers };
        }
      }
    }
    return null;
  }

  function fetOrderHeaders(sheets) {
    const headers = [];
    sheets.forEach((sheet) => {
      for (const row of sheet.rows.slice(0, 100)) {
        if (!hasHeaders(row.values, FET_ORDER_HEADERS)) continue;
        headers.push({ sheet, rowNumber: row.rowNumber, headers: row.values });
        break;
      }
    });
    return headers;
  }

  function canonicalHeaderText(value) {
    return cellText(value).replace(/\s+/g, "");
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

  function isFetProsecutorMetadataRow(values) {
    return values.slice(0, 3).some((value) => {
      const match = cellText(value).match(/^([^:：]+)\s*[:：]/);
      return Boolean(match && FET_PROSECUTOR_METADATA_KEYS.has(canonicalHeaderText(match[1])));
    });
  }

  function metadataFromValuesWithAdjacent(values, allowedKeys) {
    const metadata = {};
    values.forEach((value, index) => {
      const text = cellText(value);
      const match = text.match(/^([^:：]+)\s*[:：]\s*(.*)$/);
      if (!match) return;
      const key = canonicalHeaderText(match[1]);
      if (!allowedKeys.has(key)) return;
      let metadataValue = cellText(match[2]);
      if (!metadataValue) {
        for (const candidate of values.slice(index + 1)) {
          const candidateText = cellText(candidate);
          if (!candidateText) continue;
          const candidateMatch = candidateText.match(/^([^:：]+)\s*[:：]/);
          if (candidateMatch && allowedKeys.has(canonicalHeaderText(candidateMatch[1]))) break;
          metadataValue = candidateText;
          break;
        }
      }
      if (metadataValue) metadata[key] = metadataValue;
    });
    return metadata;
  }

  function mergeSubjectMetadata(subjectValues, metadata) {
    Object.entries(metadata).forEach(([key, value]) => {
      const text = cellText(value);
      if (!text) return;
      const values = subjectValues.get(key) || new Set();
      values.add(text);
      subjectValues.set(key, values);
    });
  }

  function subjectFromValueSets(subjectValues) {
    return Object.fromEntries(Array.from(subjectValues, ([key, values]) => [key, Array.from(values).join("、")]));
  }

  function targetPhoneFromSubject(subject) {
    return cellText(subject["申請號碼"] || subject["調閱目標"]);
  }

  function isTwmCallXml(xml) {
    return Boolean(firstXmlBlock(xml, "通聯資料")) || /<[^>]*查詢單[\s>]/.test(xml) && (Boolean(firstXmlBlock(xml, "通聯記錄查詢條件")) || Boolean(firstXmlBlock(xml, "電話號碼")) || Boolean(firstXmlBlock(xml, "電信業者")));
  }

  function isFetOrderXml(xml) {
    return Boolean(firstXmlBlock(xml, "Order"))
      && Boolean(firstXmlBlock(xml, "SummaryInfo"))
      && Boolean(firstXmlBlock(xml, "QueryInfo"))
      && Boolean(firstXmlBlock(xml, "CDRInfo"));
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
    if (!record.base_refs.some((ref) => ref.station_key === key)) record.base_refs.push({ role, station_key: key });
  }

  function addStationRef(record, stationMap, role, cellId, address) {
    const station = stationFromParts(cellId, address);
    if (!station) return;
    const key = stationKey(station);
    station.station_key = key;
    stationMap.set(key, station);
    if (!record.base_refs.some((ref) => ref.station_key === key)) record.base_refs.push({ role, station_key: key });
  }

  function addChtStationRefs(record, stationMap, value) {
    const text = cellText(value);
    if (!text) return;
    const separator = text.indexOf("/");
    if (separator < 0) {
      addChtStationEndpoint(record, stationMap, "primary", text);
      return;
    }
    addChtStationEndpoint(record, stationMap, "start", text.slice(0, separator));
    addChtStationEndpoint(record, stationMap, "end", text.slice(separator + 1));
  }

  function addChtStationEndpoint(record, stationMap, role, value) {
    const text = cellText(value);
    if (!text) return;
    if (isVirtualStationText(text)) {
      addStationRef(record, stationMap, role, "VOWIFI", text);
      return;
    }
    const match = text.match(/^([^-]+)-(.*)$/);
    if (match) {
      addStationRef(record, stationMap, role, cleanCellId(match[1]), match[2]);
      return;
    }
    addStationCompoundRef(record, stationMap, role, text);
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
    let match = text.match(/^(\d{4}-\d{1,2}-\d{1,2})[ T](\d{1,2}:\d{1,2}:\d{1,2})(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$/);
    if (!match) match = text.match(/^(\d{4}-\d{1,2}-\d{1,2})(\d{1,2}:\d{1,2}:\d{1,2})(?:\.\d+)?$/);
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
    if (text.includes("受") || text.includes("收") || text.includes("進")) return "inbound";
    if (text.includes("發") || text.includes("撥") || text.includes("去")) return "outbound";
    if (text.includes("數據") || text.includes("上網")) return "data";
    return "other";
  }

  function normalizeDirection(value, fallback) {
    const direct = cellText(value).toLowerCase();
    if (["inbound", "outbound", "data", "other"].includes(direct)) return direct;
    return directionLabel(fallback || value);
  }

  function computePhoneStats(records, mode = "count") {
    const inbound = new Map();
    const outbound = new Map();
    const total = new Map();
    records.forEach((record) => {
      const seconds = Number(record.duration_seconds || 0);
      const direction = normalizeDirection(record.direction, record.call_type);
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

  function classifyTaiwanCounty(address) {
    const normalized = cellText(address).replace(/\s+/g, "").replace(/台/g, "臺");
    return TAIWAN_COUNTIES.find((county) => normalized.includes(county)) || UNKNOWN_COUNTY;
  }

  function classifyTaiwanAdministrativeArea(address) {
    const normalized = cellText(address).replace(/\s+/g, "").replace(/台/g, "臺");
    const county = TAIWAN_COUNTIES.find((item) => normalized.includes(item));
    if (!county) return null;
    const countyIndex = normalized.indexOf(county);
    const tail = normalized.slice(countyIndex + county.length);
    const districts = TAIWAN_ADMINISTRATIVE_DISTRICTS[county] || [];
    const district = [...districts].sort((a, b) => b.length - a.length).find((item) => tail.startsWith(item));
    return district ? { county, district } : null;
  }

  function computeMultiNumberLocationMatches(workspace, options = {}) {
    const normalized = normalizeWorkspace(workspace);
    const windowMinutes = Math.max(1, Number(options.windowMinutes || MULTI_LOCATION_WINDOW_MINUTES));
    const windowMs = windowMinutes * 60 * 1000;
    const excluded = { missing_phone: 0, invalid_time: 0, invalid_address: 0 };
    if (!normalized) return { matches: [], excluded };
    const stationMap = new Map(normalized.base_stations.map((station) => [station.station_key || stationKey(station), station]));
    const groups = new Map();
    let occurrenceSequence = 0;
    normalized.records.forEach((record) => {
      const phone = normalizePhoneText(record.target_phone);
      if (!phone) {
        excluded.missing_phone += 1;
        return;
      }
      const occurredAt = normalizeDatetime(record.occurred_at);
      const timeMs = multiLocationDatetimeMs(occurredAt);
      if (!occurredAt || !Number.isFinite(timeMs)) {
        excluded.invalid_time += 1;
        return;
      }
      const refs = Array.isArray(record.base_refs) ? record.base_refs : [];
      if (!refs.length) excluded.invalid_address += 1;
      const areaOccurrences = new Map();
      refs.forEach((ref) => {
        const station = stationMap.get(ref.station_key);
        if (!station || station.is_virtual) {
          excluded.invalid_address += 1;
          return;
        }
        const area = classifyTaiwanAdministrativeArea(station.address);
        if (!area) {
          excluded.invalid_address += 1;
          return;
        }
        const areaKey = `${area.county}\u0000${area.district}`;
        const areaOccurrence = areaOccurrences.get(areaKey) || {
          county: area.county,
          district: area.district,
          matched_stations: [],
          matched_station_keys: new Set(),
        };
        const matchedStationKey = `${cellText(ref.role)}\u0000${station.station_key || stationKey(station)}`;
        if (!areaOccurrence.matched_station_keys.has(matchedStationKey)) {
          areaOccurrence.matched_station_keys.add(matchedStationKey);
          areaOccurrence.matched_stations.push({
            role: cellText(ref.role),
            cell_id: cellText(station.cell_id),
            address: cellText(station.address),
          });
        }
        areaOccurrences.set(areaKey, areaOccurrence);
      });
      areaOccurrences.forEach((areaOccurrence, areaKey) => {
        const occurrence = {
          id: `location-${occurrenceSequence += 1}`,
          phone,
          occurred_at: occurredAt,
          time_ms: timeMs,
          county: areaOccurrence.county,
          district: areaOccurrence.district,
          source_record: {
            source_file: cellText(record.source_file),
            source_sheet: cellText(record.source_sheet),
            row_number: record.row_number ?? "",
            occurred_at: occurredAt,
            target_phone: phone,
            counterparty_phone: normalizePhoneText(record.counterparty_phone),
            matched_stations: areaOccurrence.matched_stations,
          },
        };
        const rows = groups.get(areaKey) || [];
        rows.push(occurrence);
        groups.set(areaKey, rows);
      });
    });

    const matches = [];
    groups.forEach((events) => {
      events.sort((a, b) => a.time_ms - b.time_ms || a.phone.localeCompare(b.phone, "zh-Hant", { numeric: true }) || a.id.localeCompare(b.id));
      let endIndex = -1;
      let lastAcceptedEnd = -1;
      for (let startIndex = 0; startIndex < events.length; startIndex += 1) {
        if (endIndex < startIndex) endIndex = startIndex;
        const limit = events[startIndex].time_ms + windowMs;
        while (endIndex + 1 < events.length && events[endIndex + 1].time_ms <= limit) endIndex += 1;
        const windowEvents = events.slice(startIndex, endIndex + 1);
        const phones = Array.from(new Set(windowEvents.map((event) => event.phone)))
          .sort((a, b) => a.localeCompare(b, "zh-Hant", { numeric: true }));
        if (phones.length < 2 || endIndex === lastAcceptedEnd) continue;
        lastAcceptedEnd = endIndex;
        matches.push({
          start_at: windowEvents[0].occurred_at,
          end_at: windowEvents[windowEvents.length - 1].occurred_at,
          county: windowEvents[0].county,
          district: windowEvents[0].district,
          phones,
          occurrence_count: windowEvents.length,
          source_records: windowEvents.map((event) => event.source_record),
        });
      }
    });
    matches.sort((a, b) => String(a.start_at).localeCompare(String(b.start_at))
      || TAIWAN_COUNTIES.indexOf(a.county) - TAIWAN_COUNTIES.indexOf(b.county)
      || a.district.localeCompare(b.district, "zh-Hant", { numeric: true }));
    matches.forEach((match, index) => {
      match.id = `multi-location-match-${index + 1}`;
    });
    return { matches, excluded };
  }

  function multiLocationDatetimeMs(value) {
    const match = cellText(value).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);
    if (!match) return NaN;
    const parts = match.slice(1).map(Number);
    const time = Date.UTC(parts[0], parts[1] - 1, parts[2], parts[3], parts[4], parts[5]);
    const date = new Date(time);
    const valid = date.getUTCFullYear() === parts[0]
      && date.getUTCMonth() + 1 === parts[1]
      && date.getUTCDate() === parts[2]
      && date.getUTCHours() === parts[3]
      && date.getUTCMinutes() === parts[4]
      && date.getUTCSeconds() === parts[5];
    return valid ? time : NaN;
  }

  function computeTaiwanCountyStats(hotspots) {
    const counts = new Map(ALL_COUNTY_FILTER_KEYS.map((county) => [county, 0]));
    (hotspots || []).forEach((hotspot) => {
      const county = classifyTaiwanCounty(hotspot?.address);
      const count = Number(hotspot?.count || 0);
      if (Number.isFinite(count) && count > 0) counts.set(county, counts.get(county) + count);
    });
    const total = Array.from(counts.values()).reduce((sum, count) => sum + count, 0);
    return ALL_COUNTY_FILTER_KEYS.map((county) => ({
      county,
      count: counts.get(county) || 0,
      percent: total ? ((counts.get(county) || 0) / total) * 100 : 0,
    }));
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

  function buildAttachmentReport(workspace, phoneNotes = {}, exportedAt = new Date().toISOString(), options = {}) {
    const normalized = normalizeWorkspace(workspace);
    if (!normalized) throw new Error("沒有可匯出的 workspace");
    const notes = normalizePhoneNotes(phoneNotes);
    const stationMap = new Map(normalized.base_stations.map((station) => [station.station_key || stationKey(station), station]));
    const records = [...normalized.records].sort((a, b) => {
      return String(a.occurred_at || "").localeCompare(String(b.occurred_at || ""), "zh-Hant", { numeric: true })
        || String(a.source_file || "").localeCompare(String(b.source_file || ""), "zh-Hant", { numeric: true })
        || Number(a.row_number || 0) - Number(b.row_number || 0);
    });
    const hours = computeHourBuckets(records);
    const hourTotal = hours.reduce((sum, item) => sum + item.count, 0);
    const addNotes = (rows) => rows.map((row) => ({ ...row, note: notes[normalizePhoneText(row.phone)] || "" }));
    const statsCount = computePhoneStats(records, "count");
    const statsSeconds = computePhoneStats(records, "seconds");
    const summary = normalized.case.summary || parsedSummary(records, normalized.base_stations);
    return {
      meta: {
        exported_at: exportedAt,
        scope: options.scope || "complete_import",
        scope_label: options.scope_label || "完整匯入資料",
        date_range: options.date_range && typeof options.date_range === "object"
          ? { start: cellText(options.date_range.start), end: cellText(options.date_range.end) }
          : null,
        source_files: normalized.case.source_files || uniqueTextValues(records.map((record) => record.source_file)),
      },
      hours: hours.map((item) => ({ ...item, percent: hourTotal ? (item.count / hourTotal) * 100 : 0 })),
      hotspots: computeAddressHotspots(records, normalized.base_stations),
      calls: records.map((record) => ({
        ...record,
        target_note: notes[normalizePhoneText(record.target_phone)] || "",
        counterparty_note: notes[normalizePhoneText(record.counterparty_phone)] || "",
        base_stations: recordStationLabels(record, stationMap).join("；"),
      })),
      profile: {
        summary: {
          "通聯筆數": summary.records || 0,
          "目標電話數": summary.target_phones || 0,
          "對象電話數": summary.counterparty_phones || 0,
          "第一筆時間": summary.first_seen || "-",
          "最後時間": summary.last_seen || "-",
          "總秒數": summary.total_duration_seconds || 0,
          "基地台數": summary.station_count || 0,
        },
        subject: Object.entries(normalized.case.subject || {}).map(([key, value]) => ({ key, value: String(value ?? "") })),
        imeis: collectUniqueImeis(records),
      },
      stats: {
        count: {
          inboundRows: addNotes(statsCount.inboundRows),
          outboundRows: addNotes(statsCount.outboundRows),
          totalRows: addNotes(statsCount.totalRows),
        },
        seconds: {
          inboundRows: addNotes(statsSeconds.inboundRows),
          outboundRows: addNotes(statsSeconds.outboundRows),
          totalRows: addNotes(statsSeconds.totalRows),
        },
      },
    };
  }

  function recordStationLabels(record, stationMap) {
    const labels = [];
    const seen = new Set();
    (record.base_refs || []).forEach((ref) => {
      const station = stationMap.get(ref.station_key);
      if (!station) return;
      const label = cellText(station.address || station.cell_id || (station.is_virtual ? "VOWIFI" : ""));
      if (label && !seen.has(label)) {
        seen.add(label);
        labels.push(label);
      }
    });
    return labels;
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

  function decodeXmlBytes(bytes) {
    const input = toUint8Array(bytes);
    let encoding = "utf-8";
    let offset = 0;
    if (input[0] === 0xef && input[1] === 0xbb && input[2] === 0xbf) offset = 3;
    else if (input[0] === 0xff && input[1] === 0xfe) { encoding = "utf-16le"; offset = 2; }
    else if (input[0] === 0xfe && input[1] === 0xff) { encoding = "utf-16be"; offset = 2; }
    else if (input[0] === 0x3c && input[1] === 0x00) encoding = "utf-16le";
    else if (input[0] === 0x00 && input[1] === 0x3c) encoding = "utf-16be";
    else {
      const declaration = Array.from(input.slice(0, 512), (byte) => byte < 128 ? String.fromCharCode(byte) : " ").join("");
      const match = declaration.match(/<\?xml[^>]*encoding\s*=\s*["']\s*([^"']+)\s*["']/i);
      const declared = cellText(match?.[1]).toLowerCase().replace(/_/g, "-");
      if (["big5", "big-5", "cp950", "windows-950"].includes(declared)) encoding = "big5";
      else if (["utf-16", "utf-16le", "unicode"].includes(declared)) encoding = "utf-16le";
      else if (declared === "utf-16be") encoding = "utf-16be";
    }
    try {
      if (typeof TextDecoder !== "undefined") return new TextDecoder(encoding).decode(input.slice(offset));
    } catch (_error) {
      throw new Error("XML 編碼不受此瀏覽器支援");
    }
    if (encoding !== "utf-8") throw new Error("XML 編碼不受此環境支援");
    return Buffer.from(input.slice(offset)).toString("utf8");
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

  function downloadBlob(fileName, bytes, type) {
    const blob = new Blob([bytes], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
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
    mergeWorkspaces,
    normalizeWorkspace,
    buildAttachmentReport,
    computePhoneStats,
    computeHourBuckets,
    computeAddressHotspots,
    classifyTaiwanCounty,
    classifyTaiwanAdministrativeArea,
    computeTaiwanCountyStats,
    computeMultiNumberLocationMatches,
    computeDateRangeBounds,
    filterRecordsByDateRange,
    buildSubmissionCsv,
    collectSubmissionPhones,
    collectUniqueImeis,
    normalizePhoneText,
    tellowsUrl,
    hourButtonLabel,
  };
});
