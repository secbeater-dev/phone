# Phone Workbench 完整架構

> 維護規則：任何修改前必須完整閱讀本文件。架構、資料介面、匯入格式、畫面、儲存、隱私控制、測試或部署有變動時，必須同步更新本文件與「異動紀錄」。本文件只使用欄位名稱與合成範例，不記錄真實通聯值、筆數或識別資訊。

## 1. 專案定位

Phone Workbench 是部署在 GitHub Pages 的純前端通聯資料分析工具。使用者選取的 XLSX/XML 由瀏覽器內的 `FileReader`、本地 SheetJS 與 `app.js` 處理；沒有後端、API、資料庫、Service Worker 或上傳端點。

功能包括通聯列表、用戶資料、電話統計、時間分布、基地台熱點、電話投單，以及 workspace、CSV 與本機設定匯出。

## 2. 目錄與責任

```text
.
├─ .github/workflows/pages.yml  # 測試、敏感檔拒絕、白名單 Pages 部署
├─ .gitignore                   # 排除通聯檔與本機匯出物
├─ .nojekyll                    # 停用 Jekyll
├─ AGENTS.md                    # 維護與個資規則
├─ 404.html                     # 無腳本維護頁
├─ admin.html                   # 無腳本管理提示頁
├─ app.js                       # 解析、狀態、統計、畫面與匯出
├─ assets/                      # 站內圖片
├─ fullview.md                  # 本文件
├─ index.html                   # 主頁 DOM、CSP 與 SRI
├─ README.md                    # 使用、隱私及維護提示
├─ styles.css                   # 版面、主題與響應式樣式
├─ tests/app.test.js            # Node 內建測試與合成 XLSX
└─ vendor/xlsx.full.min.js      # 本地 SheetJS
```

專案沒有套件安裝或建置步驟；測試直接使用 Node 與 repository 內的 SheetJS。

## 3. 載入順序

1. `index.html` 建立側欄、六個 view、匯入控制與使用提醒。
2. 瀏覽器驗證 `vendor/xlsx.full.min.js` 與 `app.js` 的 SHA-384 SRI 後依序執行；CSP 不允許其他 script。
3. `app.js` 以 UMD 包裝：瀏覽器掛載 `window.PhoneWorkbench`；Node CommonJS 載入本地 SheetJS 並輸出測試介面。
4. `DOMContentLoaded` 執行 `init()`，還原偏好、綁定事件並渲染所有 view。

## 4. 狀態模型

單一 `state` 保存目前頁面狀態：

- `view`：目前功能頁。
- `cases`、`currentWorkspace`、`callRecords`：案件、workspace 與標準化記錄。
- `callSort`、`callColumnWidths`：列表排序及欄寬。
- `phoneStatsRankMode`、`phoneNotes`：排行模式與電話備註。
- `hourSelection`、`appliedHourSelection`、`expandedHotspotAddress`：時段與熱點狀態。
- `theme`、`sidebarCollapsed`：介面偏好。

## 5. 匯入與格式判定

```text
file input
  → FileReader.readAsArrayBuffer
  → parseImportFile
  ├─ XML → parseXmlWorkbook → XML 格式 parser
  └─ XLSX → XLSX.read → parseXlsxWorkbook
       → 格式偵測 → 格式 parser
  → makeWorkspace → applyWorkspace → renderAllViews
```

XLSX 會先檢查所有工作表前 80 列是否包含中華電信地檢新版標題；只有這個格式會先移除標題內空白後再比對。識別欄位為 `CDR類別`、`主叫號碼`、`查詢狀態`、`受叫號碼`、`始話日期時間`、`通話秒數`、`IMEI`、`指定轉接`、`起始基地台-地址/終止基地台-地址`。命中後回傳：

```js
{ carrier: "中華電信", sourceFormat: "chunghwa_prosecutor_cdr_xlsx" }
```

標題列前的「名稱：值」完整收入 `subject`。調閱門號優先採 `電話號碼`，其次為 `設備號碼`。方向規則：調閱門號在主叫側為 `outbound`、在受叫側為 `inbound`；不在該列時才依 `受話／進來CDR` 或 `發話` 衍生。列表保留原始 `CDR類別`。

日期可正規化時轉為 `YYYY-MM-DDTHH:mm:ss`；非空且無法正規化時保留原文並加入只含列號的警告，不丟棄資料列。查詢狀態及指定轉接放入 `note`，IMEI 放入對應欄位。基地台欄的合成結構如下：

```text
起始代碼-合成起始地址/終止代碼-合成終止地址
```

解析成 `start`、`end` 兩個 `base_refs` 並加入去重後的 `base_stations`。其他既有 XLSX/XML 格式沿用原有精確標題判定與正規化流程。

## 6. 標準化介面

### Workspace

```js
{
  case: {
    source_file, carrier, source_format, sheet_name, header_row,
    total_source_rows, total_records, subject, summary, parse_warnings
  },
  records: [],
  base_stations: [],
  parse_warnings: []
}
```

### Record

```js
{
  row_number, occurred_at, ended_at, duration_seconds,
  call_type, direction, target_phone, counterparty_phone,
  imei, imsi, external_ip, internal_ip,
  upload_bytes, download_bytes, total_bytes,
  note, base_refs
}
```

`direction` 可明確指定 `inbound` 或 `outbound`。`computePhoneStats` 優先使用明確方向；舊格式沒有方向時仍由既有 `call_type` 邏輯推導。電話正規化為數字並處理 `886` 國碼。

### Base station

基地台以 `station_key` 去重，保存 `cell_id`、原始/正規化地址、狀態與虛擬標記；記錄透過帶角色的 `base_refs` 參照。

## 7. 畫面與本機輸出

- `renderTwoWayCalls`：通聯搜尋、排序、欄寬、備註與最多 5,000 筆顯示。
- `renderProfileView`：案件摘要、subject 與 IMEI 清單。
- `renderStatsView`：依明確或衍生方向統計來電、去電及完整排行。
- `renderHoursView`：24 小時分布、時段篩選及基地台熱點。
- `renderSubmissionPreview`：電話驗證、去重、投單預覽與 CSV。
- `renderExportView`：workspace JSON 與本機設定匯出。

匯入值進入 HTML 字串前由 `escapeHtml` 處理。電話只以 `.phone-value` 純文字顯示，不產生外部查詢連結。下載由 `Blob`、`URL.createObjectURL` 與 `download` 屬性在本機完成。

## 8. localStorage

同源 `localStorage` 保存排行模式、選取時段、電話備註、通聯欄寬、主題與側欄狀態。完整 workspace 與通聯記錄不會自動保存；只有使用者主動匯出才會產生本機下載檔。

## 9. 隱私與信任邊界

- 三個 HTML 都設定 `no-referrer`；已移除 Google Analytics。
- `index.html` 的 CSP 將 `connect-src`、`object-src`、`form-action`、`frame-src`、`worker-src`、`media-src` 與 `manifest-src` 設為 `none`，圖片與字型限同源/data，script 只允許兩個具正確 SRI 的本機檔案；`style-src-attr` 只為既有圖表高度與可調欄寬保留 inline CSS。
- `404.html` 與 `admin.html` 使用 `script-src 'none'`。
- 沒有 `fetch`、XHR、WebSocket、EventSource、Beacon、表單提交或檔案上傳路徑。
- 電話不再連往 Tellows；Cloudflare/其他注入 script 不在 CSP 許可清單內。
- `.gitignore`、CI 敏感副檔名檢查及部署白名單形成三層防護。

任何真實通聯檔、內容、衍生識別資訊、workspace、投單 CSV、設定匯出、含個資截圖或日誌，都不得加入 Git、Actions artifact、文件或任何外部服務。真實檔只能從 repository 外在本機記憶體中驗證，測試輸出限通過/失敗與檔名。

## 10. 測試

`node --test tests/app.test.js` 使用記憶體內合成 XLSX，涵蓋：空白標題、用戶資料、主叫/受叫/進來 CDR、調閱門號不在該列、無效日期、IMEI、指定轉接與雙基地台；另有既有台灣大哥大解析及電話方向統計回歸測試，並驗證 HTML 無 Google/Tellows 且 SRI 與檔案雜湊一致。

設定 `PRIVATE_CDR_XLSX` 時會啟用 repository 外真實檔整合測試，只斷言格式、有效來源列數相等與方向合法，不輸出資料內容。瀏覽器煙霧測試也只使用合成 XLSX，檢查六個 view、純文字電話與零第三方資產。

## 11. GitHub Pages 部署

`.github/workflows/pages.yml` 在 `main` push 或手動觸發時：

1. checkout。
2. 拒絕 Git 追蹤的 XLSX/XLS/XML/CSV/TSV、私密資料目錄及本機匯出檔。
3. 使用 Node 22 執行測試。
4. 只把 `.nojekyll`、三個 HTML、`app.js`、`styles.css`、`assets/` 與 `vendor/` 複製到 `_site`。
5. 上傳 `_site` 並部署 Pages。

自訂網域 `phone.secbeater.com` 保留在既有 Pages/Cloudflare 設定。正式站必須驗證 HTTP、資產版本、CSP/SRI、功能及第三方 script 未執行；若 Cloudflare challenge 或 Browser Insights 仍執行或破壞頁面，應停止宣告完成並由站方停用相關功能或改為 DNS-only。

## 12. Node 匯出介面

`app.js` 輸出：`parseImportFile`、`computePhoneStats`、`computeHourBuckets`、`computeAddressHotspots`、`buildSubmissionCsv`、`collectSubmissionPhones`、`collectUniqueImeis`、`normalizePhoneText`、`tellowsUrl`（僅相容舊程式碼，畫面不使用）、`hourButtonLabel`。

## 13. 維護檢查表

修改前：完整閱讀本文件、確認私密檔在 repository 外、檢查工作樹。修改後：更新本文件、以合成資料測試、以無內容輸出的方式驗證真實檔、檢查 Git index/歷史/部署白名單與網路請求，並新增異動紀錄。

## 14. 異動紀錄

- 2026-07-23：依未修改的 `main`（commit `13485b4`）建立基準架構文件。
- 2026-07-23：新增中華電信地檢新版 XLSX parser、明確方向統計、日期警告及雙基地台解析；移除 Google Analytics/Tellows，加入 CSP/SRI、忽略規則、合成/私密隔離測試與 Pages 白名單部署。
