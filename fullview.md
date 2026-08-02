# Phone Workbench 完整架構

> 維護規則：任何修改前必須完整閱讀本文件。架構、資料介面、匯入格式、畫面、儲存、隱私控制、測試或部署有變動時，必須同步更新本文件與「異動紀錄」。本文件只使用欄位名稱與合成範例，不記錄真實通聯值、筆數或識別資訊。

## 1. 專案定位

Phone Workbench 是部署在 GitHub Pages 的純前端通聯資料分析工具。使用者選取的 XLSX/XML 由瀏覽器內的 `FileReader`、本地 SheetJS 與 `app.js` 處理；附卷由本地 ExcelJS/pdf-lib 產生；沒有後端、API、資料庫、Service Worker 或上傳端點。

功能包括同批多檔合併、通聯列表、用戶資料、電話統計、時間分布、基地台熱點、電話投單，以及附卷 XLSX/PDF、workspace、CSV 與本機設定匯出。

## 2. 目錄與責任

```text
.
├─ .github/workflows/pages.yml  # 測試、敏感檔拒絕、白名單 Pages 部署
├─ .gitattributes               # 固定網頁文字資產為 LF，確保 SRI 可重現
├─ .gitignore                   # 排除通聯檔與本機匯出物
├─ .nojekyll                    # 停用 Jekyll
├─ AGENTS.md                    # 維護與個資規則
├─ 404.html                     # 無腳本維護頁
├─ admin.html                   # 無腳本管理提示頁
├─ app.js                       # 解析、多檔合併、狀態、統計、畫面與匯出協調
├─ attachment-export.js         # 六分頁 XLSX 與六類文字 PDF 產生器
├─ assets/                      # 站內圖片
├─ fullview.md                  # 本文件
├─ index.html                   # 主頁 DOM、CSP 與 SRI
├─ README.md                    # 使用、隱私及維護提示
├─ styles.css                   # 版面、主題與響應式樣式
├─ tests/app.test.js            # Node 內建測試與合成 XLSX
└─ vendor/
   ├─ xlsx.full.min.js          # 匯入用本地 SheetJS
   ├─ exceljs.min.js            # 延遲載入的附卷 XLSX 產生器
   ├─ pdf-lib.min.js            # 延遲載入的 PDF 產生器
   ├─ fontkit.umd.min.js         # 延遲載入的中文字型嵌入器
   ├─ open-huninn-data.js       # 延遲載入的 jf open 粉圓繁中字型資料
   └─ licenses/                 # 第三方程式與字型授權
```

專案沒有套件安裝或建置步驟；測試直接使用 Node 與 repository 內固定版本的瀏覽器套件。

## 3. 載入順序

1. `index.html` 建立側欄、六個 view、匯入控制、全域日期篩選視窗、熱點縣市篩選視窗、附卷匯出視窗與使用提醒；提醒中的「今日重點」列出目前發布的使用者可見更新，Gemini 區塊使用原生 `details` 顯示家庭分享教學。
2. 瀏覽器透過固定發布版本查詢字串載入 `styles.css`、`vendor/xlsx.full.min.js`、`attachment-export.js` 與 `app.js`；三個 script 驗證 SHA-384 SRI 後依序執行。版本字串避免舊快取與新版資產衝突，CSP 不允許其他 script。
3. `app.js` 與 `attachment-export.js` 均以 UMD 包裝，可供瀏覽器及 Node 測試使用。只有使用者按下附卷下載時，`app.js` 才以固定版本路徑與 SRI 延遲載入 ExcelJS，或依序載入 pdf-lib、fontkit 與字型資料。
4. `DOMContentLoaded` 執行 `init()`，還原偏好、綁定事件並渲染所有 view。

## 4. 狀態模型

單一 `state` 保存目前頁面狀態：

- `view`：目前功能頁。
- `cases`、`currentWorkspace`、`callRecords`：案件、workspace 與標準化記錄。
- `callSort`、`callPage`、`callColumnWidths`：列表排序、500 筆分頁及欄寬。
- `phoneStatsRankMode`、`phoneNotes`：排行模式與電話備註。
- `hourSelection`、`appliedHourSelection`、`expandedHotspotAddress`：時段與展開中熱點狀態。
- `hotspotCountySelection`、`hotspotCountyDraft`：目前套用及彈窗草稿的縣市條件；預設包含現行 22 縣市及「未辨識」，只存在記憶體且新 workspace 匯入時重設。
- `dateRangeBounds`、`dateRange`、`dateRangeDraft`：全部可解析通聯的日期界線、目前套用日期及彈窗草稿；只存在記憶體，新 workspace 成功匯入時重設為完整資料。
- `theme`、`sidebarCollapsed`：介面偏好。

## 5. 匯入與格式判定

```text
一批 file input
  → 逐檔 FileReader.readAsArrayBuffer → parseImportFile
  ├─ XML → 依 BOM／encoding 宣告解碼 → parseXmlWorkbook → XML 格式 parser
  └─ XLSX → XLSX.read → parseXlsxWorkbook
       → 格式偵測 → 格式 parser
  → 成功 workspace 集合 → mergeWorkspaces（一次）
  → applyWorkspace → 切換通聯列表 → renderAllViews
```

同一次選取的所有成功檔案合併成一個 workspace；部分失敗仍保留成功項目並顯示失敗狀態，全部失敗則不改變現有資料。下一批只要至少一檔成功，就以該批合併結果取代目前 workspace。合併時重算摘要、電話統計、時間分布及基地台；同名 `subject` 欄位的不同非空值去重後以 `、` 合併。

XLSX 會先檢查所有工作表前 80 列是否包含中華電信地檢新版標題；只有這個格式會先移除標題內空白後再比對。識別欄位為 `CDR類別`、`主叫號碼`、`查詢狀態`、`受叫號碼`、`始話日期時間`、`通話秒數`、`IMEI`、`指定轉接`、`起始基地台-地址/終止基地台-地址`。命中後回傳：

```js
{ carrier: "中華電信", sourceFormat: "chunghwa_prosecutor_cdr_xlsx" }
```

標題列前的「名稱：值」完整收入 `subject`。調閱門號優先採 `電話號碼`，其次為 `設備號碼`。方向規則：調閱門號在主叫側為 `outbound`、在受叫側為 `inbound`；不在該列時才依 `受話／進來CDR` 或 `發話` 衍生。列表保留原始 `CDR類別`。

日期可正規化時轉為 `YYYY-MM-DDTHH:mm:ss`；非空且無法正規化時保留原文並加入只含列號的警告，不丟棄資料列。查詢狀態及指定轉接放入 `note`，IMEI 放入對應欄位。基地台欄的合成結構如下：

```text
起始代碼-合成起始地址/終止代碼-合成終止地址
```

解析成 `start`、`end` 兩個 `base_refs` 並加入去重後的 `base_stations`。

遠傳地檢新版通聯 XLSX 以 `始話時間`、`通話秒數`、`調閱號碼`、`IMEI`、`通話類別`、`通話對象`、`轉接電話`、`基地台/交換機`、`備註` 識別，允許欄名空白差異及標題中的空白佔位欄。命中後回傳：

```js
{ carrier: "遠傳電信", sourceFormat: "fet_prosecutor_call_xlsx" }
```

解析器會逐列掃描整份工作表；遇到重複標題即更新目前欄位索引，因此同一檔案的單一或多個查詢區段都會匯入。案件資料支援 `名稱：值` 及 `名稱：` 後接相鄰儲存格，跨區段同名值去重後以 `、` 合併至 `subject`。已識別的案件資料列、空白列與重複標題不會建立 record。

`始話時間`、`通話秒數`、`調閱號碼`、`通話對象`、`IMEI` 分別映射至標準欄位；`通話類別`保留原文並衍生 `direction`；`轉接電話`與`備註`合併到 `note`；`基地台/交換機`建立 `primary` 基地台參照。只有基地台欄有值的續行列會附加至同區段上一筆有效通聯，同一基地台只附加一次且不建立空白 record；沒有上一筆可附加時，只加入不含儲存格內容的列號警告。非空且無法正規化的日期保留原文並加入只含列號的警告。其他既有 XLSX/XML 格式沿用原有精確標題判定與正規化流程。

遠傳 Order XLSX 以 `DocNo`、`Seq`、`Status2`、`QueType`、`QueObject`、`QueDirection`、`CallDirection`、`CallingNumber`、`CalledNumber`、`CallStartTimeStamp`、`Duration`、`IMEI` 等英文原始欄位識別，命中後回傳：

```js
{ carrier: "遠傳電信", sourceFormat: "fet_order_cdr_xlsx" }
```

解析器只處理符合英文欄位的原始資料工作表，忽略同檔中文展示頁，避免重複匯入。每列同時保存格式化顯示值與對齊的 raw cell 值；Excel serial 日期、數值秒數、基地台代碼及科學記號顯示的 IMEI 由 raw 值正規化。`CallStartTimeStamp`、`Duration`、主受叫號、`IMEI`、`CallFwd` 與起訖基地台分別映射至標準欄位；SummaryInfo/QueryInfo 對應的文號、狀態、調閱類型、調閱目標、查詢方向、起訖區間及備註以去重值合併至 `subject`。

遠傳 Order XML 以 `Order/SummaryInfo/Record/QueryInfo/CDRInfo` 結構識別，命中後格式為 `fet_order_cdr_xml`。匯入前依 UTF-8 BOM、UTF-16 BOM/位元組排列或 XML `encoding` 宣告選擇 UTF-8、Big5、UTF-16LE/BE 解碼；未知 XML 結構直接回報不支援，不再靜默交給台哥大 parser。來源 XSL 只負責原檔顯示，不是匯入資產，也不進入 repository。

Order 通聯類型依來源 XSL 語意顯示：`O`、`T`、`I`、`1`、`2`、`9`、`S`、`M` 分別為發話、受話、進來 CDR、系統發訊、手機發訊、收訊、雙號共振及多媒體簡訊。方向中 `O`/`2` 為 outbound，`T`/`I`/`1`/`9` 為 inbound，`S`/`M`/未知為 other；若調閱類型是電話且能比對主受叫側，實際側別優先。含毫秒日期正規化到秒；非空無法正規化的日期保留原文並加入只含列號的警告。

## 6. 標準化介面

### Workspace

```js
{
  case: {
    source_file, source_files, carrier, source_format, sheet_name, header_row,
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
  row_number, source_file?, source_sheet?, occurred_at, ended_at, duration_seconds,
  call_type, direction, target_phone, counterparty_phone,
  imei, imsi, external_ip, internal_ip,
  upload_bytes, download_bytes, total_bytes,
  note, base_refs
}
```

`source_files` 是本批所有來源檔名的去重清單；每筆新 record 以可選的 `source_file`、`source_sheet` 保留來源。載入舊 workspace JSON 時若沒有這些欄位，會從 `case.source_file` 與 `case.sheet_name` 補齊，因此介面向後相容。`direction` 可明確指定 `inbound` 或 `outbound`。`computePhoneStats` 優先使用明確方向；舊格式沒有方向時仍由既有 `call_type` 邏輯推導。電話正規化為數字並處理 `886` 國碼。

### Base station

基地台以 `station_key` 去重，保存 `cell_id`、原始/正規化地址、狀態與虛擬標記；記錄透過帶角色的 `base_refs` 參照。

## 7. 畫面與本機輸出

- `renderTwoWayCalls`：顯示來源檔案、通聯搜尋、全資料排序、欄寬、備註與每頁 500 筆分頁；先套用全域日期，再對結果搜尋、排序，條件改變時回到第一頁，沒有靜默筆數截斷。
- `renderProfileView`：subject 保留完整案件欄位，摘要與 IMEI 清單依全域日期重新計算。
- `renderStatsView`：依全域日期及明確/衍生方向統計來電、去電及完整排行；三種排行均顯示全部電話，不截斷前 20 名。
- `renderHoursView`：全域日期先套用，再與 24 小時選擇、地址搜尋及縣市條件取交集。熱點搜尋右側的綠色齒輪開啟縣市篩選；使用者可在現行 22 縣市及「未辨識」間多選，並以「全選」或「全部取消」快速調整草稿。按套用後才過濾熱點列表；取消、背景點擊或 Escape 會放棄草稿。全部取消時套用維持停用，至少重新選一項後才能套用。
- `renderSubmissionPreview`：電話驗證、去重、投單預覽與 CSV。
- `renderExportView`：workspace JSON 與本機設定匯出；入口位於側欄下方工具區，切換後顯示選取狀態。
- 主功能導覽「附卷檔案匯出」：開啟含進度與錯誤狀態的視窗，不切換目前 view；沒有 workspace/records 時停用下載。
- 「資料匯入」下方的「時間篩選」：只有 workspace 成功匯入後顯示；縮合側欄只保留日期圖示。黑白彈窗以上下排列的起訖 `date` 欄位設定包含首末整日的範圍；起始晚於結束時停用篩選。取消、背景或 Escape 放棄草稿；「回復預設」立即恢復完整資料、重繪並關閉彈窗。
- 使用提醒：顯示不含特定電信商名稱的多檔案匯入說明、當日更新、Gemini 優惠與可展開家庭分享教學；支援格式區提供作者 Telegram 與更多類型 Notion 頁面。

匯入值進入 HTML 字串前由 `escapeHtml` 處理。電話只以 `.phone-value` 純文字顯示，不產生外部查詢連結。下載由 `Blob`、`URL.createObjectURL` 與 `download` 屬性在本機完成，觸發後撤銷 Object URL。

熱點縣市分類會先移除地址空白並將「台」統一為「臺」，再依固定現行縣市名稱判定；非空、非虛擬且無法判定者歸入「未辨識」，畫面仍顯示來源地址原文。縣市統計沿用 `computeAddressHotspots` 的逐通聯地址去重結果，比例分母為目前已套用時段內所有縣市及未辨識的地址出現次數；統計基準不讀取地址搜尋字串或已套用縣市條件，因此篩選後總數與比例不會自行改變。地址文字搜尋與縣市條件在呈現熱點清單時取交集。

### 附卷資料流

`buildAttachmentReport` 由目前全域日期範圍的 records 及實際參照基地台建立報表模型，不讀取通聯搜尋、頁碼、熱點搜尋、縣市條件或 24 小時選擇。報表模型包含範圍內 24 小時桶、基地台熱點及其發生時間、通聯與來源、合併後用戶欄位與範圍內 IMEI，以及附帶瀏覽器電話備註的次數/秒數排行。`report.meta` 以向後相容的 `scope_label`、`date_range` 標示完整匯入或自訂日期；各 XLSX/PDF 頁首顯示相同範圍。

- XLSX：ExcelJS 產生 `時間分布圖`、`熱點時間摘要`、`通聯列表`、`用戶資料`、`電話統計-次數版`、`電話統計-秒數版` 六張工作表。使用凍結表頭、自動篩選、欄寬、換行及列印設定，時間頁另嵌入瀏覽器 canvas 產生的長條圖。電話、IMEI、來源及備註都以 ExcelJS 純字串型別寫入；即使內容以公式符號開頭也不建立 formula，避免前導零遺失與公式注入。
- PDF：pdf-lib、fontkit 與 OFL 授權的 jf open 粉圓產生六份獨立檔案。表格可跨頁並在新頁重複表頭；每頁有頁碼。中文與識別欄位使用嵌入字型的文字指令，因此可搜尋與複製，不是畫面截圖。pdf-lib/fontkit 對繁中 TTF 啟用子集化時會造成部分檢視器字形映射毀損，因此完整嵌入較精簡的字型檔，以跨檢視器可讀性優先。
- 檔名只有「附卷」、區段名稱及本機時間戳，不含姓名、電話、案號或來源檔名。

## 8. localStorage

同源 `localStorage` 保存排行模式、24 小時選取、電話備註、通聯欄寬、主題與側欄狀態。全域日期、完整 workspace 與通聯記錄不會自動保存；只有使用者主動匯出才會產生本機下載檔。

## 9. 隱私與信任邊界

- 三個 HTML 都設定 `no-referrer`；已移除 Google Analytics。
- `index.html` 的 CSP 將 `connect-src`、`object-src`、`form-action`、`frame-src`、`worker-src`、`media-src` 與 `manifest-src` 設為 `none`，圖片與字型限同源/data，script 只允許對應目前位元組的 SHA-384；靜態入口與四個延遲載入匯出資產都同時受 CSP hash、SRI 與固定版本查詢字串保護。`style-src-attr` 只為既有圖表高度與可調欄寬保留 inline CSS。
- `404.html` 與 `admin.html` 使用 `script-src 'none'`。
- 沒有 `fetch`、XHR、WebSocket、EventSource、Beacon、表單提交或檔案上傳路徑；匯出套件與字型皆從同源固定檔案載入，資料只在瀏覽器記憶體中交給 Blob。
- 電話不再連往 Tellows；Cloudflare/其他注入 script 不在 CSP 許可清單內。
- `.gitignore`、CI 敏感副檔名檢查及部署白名單形成三層防護。
- 完整支援格式清單與發布日誌只保存在 repository 上一層的本機 `supported-formats.md`、`update.md`；它們不屬於 Git 工作樹，也不會進入 Pages artifact。

任何真實通聯檔、內容、衍生識別資訊、workspace、投單 CSV、設定匯出、含個資截圖或日誌，都不得加入 Git、Actions artifact、文件或任何外部服務。真實檔只能從 repository 外在本機記憶體中驗證，測試輸出限各驗證階段的通過/失敗，不輸出檔名、內容或筆數。

## 10. 測試

`node --test tests/app.test.js` 使用記憶體內合成 XLSX/XML，涵蓋：中華電信空白標題、用戶資料、主叫/受叫/進來 CDR、調閱門號不在該列、無效日期、IMEI、指定轉接與雙基地台；遠傳地檢新版另涵蓋九欄、空白佔位欄、重複查詢區段、內嵌/相鄰案件資料、缺少通話對象、轉接、備註、基地台續行與重複續行；遠傳 Order 涵蓋雙工作表/英文單工作表、raw Excel serial、數值 IMEI、UTF-8/Big5 宣告 XML、毫秒日期、方向代碼、轉接及起訖基地台。另測試多檔合併、來源追蹤、subject 去重、舊 workspace 相容、既有台灣大哥大解析、日期界線/單日/錯置/無效日期、電話方向、完整排行、22 縣市與台/臺分類、未辨識、零資料、逐通聯地址去重、縣市比例、縣市全選/全部取消控制、日期範圍附卷 metadata、六張 XLSX 工作表、文字/公式安全、六種 PDF、彈窗文字、指定連結、HTML 無 Google/Tellows，以及靜態/延遲載入資產的 SRI 與檔案雜湊一致。

設定 `PRIVATE_CDR_XLSX` 時會啟用 repository 外中華電信真實檔整合測試；設定 `PRIVATE_FET_XLSX` 時會啟用 repository 外遠傳地檢真實檔整合測試。遠傳 Order 的兩份 XLSX 與一份 XML 另由 repository 外的隔離驗證器逐檔檢查格式、有效來源列完整性、日期、方向、識別欄位、基地台、六分頁 XLSX 與六份 PDF；所有真實檔驗證只回報檔案類別及通過/失敗，不輸出檔名、內容、筆數或留下產物。合成 XLSX 另以試算表工具檢查/渲染六張工作表；六份合成 PDF 以 Poppler 渲染及抽取文字，確認換頁、表頭、中文文字與可搜尋性。瀏覽器煙霧測試只使用合成資料，檢查六個 view、日期彈窗、分頁、附卷視窗、下載與零非預期第三方請求。

## 11. GitHub Pages 部署

`.github/workflows/pages.yml` 在 `main` push 或手動觸發時：

1. checkout。
2. 拒絕 Git 追蹤的 XLSX/XLS/XML/CSV/TSV、私密資料目錄及本機匯出檔。
3. 使用 Node 22 執行測試。
4. 以逐檔白名單把 `.nojekyll`、三個 HTML、`app.js`、`attachment-export.js`、`styles.css`、兩個圖片，以及固定版本的五個 vendor 程式/字型檔與四份授權檔複製到 `_site`。
5. 上傳 `_site` 並部署 Pages。

自訂網域 `phone.secbeater.com` 保留在既有 Pages/Cloudflare 設定。正式站必須驗證 HTTP、資產版本、CSP/SRI、功能及第三方 script 未執行；若 Cloudflare challenge 或 Browser Insights 仍執行或破壞頁面，應停止宣告完成並由站方停用相關功能或改為 DNS-only。

## 12. Node 匯出介面

`app.js` 輸出：`parseImportFile`、`mergeWorkspaces`、`normalizeWorkspace`、`buildAttachmentReport`、`computePhoneStats`、`computeHourBuckets`、`computeAddressHotspots`、`classifyTaiwanCounty`、`computeTaiwanCountyStats`、`computeDateRangeBounds`、`filterRecordsByDateRange`、`buildSubmissionCsv`、`collectSubmissionPhones`、`collectUniqueImeis`、`normalizePhoneText`、`tellowsUrl`（僅相容舊程式碼，畫面不使用）、`hourButtonLabel`。`attachment-export.js` 輸出六張工作表/六種 PDF 定義、`createAttachmentXlsx`、`createAttachmentPdf` 與 `safeText`。

## 13. 維護檢查表

修改前：完整閱讀本文件、確認私密檔在 repository 外、檢查工作樹。修改後：更新本文件、以合成資料測試、以無內容輸出的方式驗證真實檔、檢查 Git index/歷史/部署白名單與網路請求，並新增異動紀錄。

## 14. 異動紀錄

- 2026-07-23：依未修改的 `main`（commit `13485b4`）建立基準架構文件。
- 2026-07-23：新增中華電信地檢新版 XLSX parser、明確方向統計、日期警告及雙基地台解析；移除 Google Analytics/Tellows，加入 CSP/SRI、忽略規則、合成/私密隔離測試與 Pages 白名單部署。
- 2026-07-23：固定網頁文字資產為 LF，讓 Windows、GitHub Actions 與 Pages 使用相同位元組並維持 SRI 驗證一致。
- 2026-07-24：移除電話統計各分類前 20 名顯示上限，來電、去電與全部電話皆顯示完整排行。
- 2026-07-24：更新使用提醒的「今日重點」，公告電話統計已改為顯示全部排行。
- 2026-07-26：新增遠傳地檢新版通聯 XLSX parser，支援空白佔位欄、重複查詢區段、案件資料合併、日期警告及基地台參照。
- 2026-07-26：更新今日重點、Gemini 價格與家庭分享教學；移除線上支援格式清單，改為作者 Telegram 聯絡連結，並更新左上角 Notion 連結。
- 2026-07-26：修正同批多檔只留下最後一檔的問題，加入 workspace 合併、來源追蹤、subject/基地台去重、遠傳基地台續行附加與通聯列表 500 筆分頁。
- 2026-07-26：新增側欄附卷匯出視窗、一份六分頁 XLSX、六份可搜尋文字 PDF、固定版本本地套件/中文字型與逐檔 Pages 發布白名單。
- 2026-07-26：正式站驗證發現舊瀏覽器快取可能與新版 SRI 衝突，為靜態與延遲載入 script 加入固定發布版本查詢字串；CSP 雜湊限制維持不變。
- 2026-07-26：互換「附卷檔案匯出」與「資料匯出」的側欄位置；附卷改為主導覽動作，資料匯出改為下方工具區 view 入口。
- 2026-07-26：移除附卷視窗及下載完成訊息中的指定個資提醒文字；更新使用提醒為多檔案匯入說明，並加入更多支援類型 Notion 連結。
- 2026-07-29：熱點時間摘要新增現行 22 縣市與「未辨識」的總數、比例及多選快速篩選；台/臺統一分類，條件只影響熱點列表且不寫入 localStorage 或附卷。
- 2026-07-29：縣市篩選新增「全選」及「全部取消」草稿操作；全部取消後仍須至少重選一項才能套用。
- 2026-07-29：為 `styles.css` 加入固定發布版本查詢字串，避免正式站沿用舊版響應式樣式快取。
- 2026-08-02：新增遠傳 Order 英文原始 XLSX 與 Order XML parser，支援 raw Excel 值、UTF-8/Big5/UTF-16 XML 解碼、官方方向/類型語意、毫秒日期與起訖基地台；未知 XML 不再誤判為台哥大格式。
- 2026-08-02：資料匯入下方新增記憶體內全域日期篩選；分析畫面與附卷依包含首末整日的範圍重算，workspace JSON 與電話投單維持完整原始資料。
