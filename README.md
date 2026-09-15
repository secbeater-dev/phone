# Phone Workbench

> Maintenance requirement: read [`fullview.md`](./fullview.md) before making changes, and update it after any architecture, interface, parser, privacy, test, or deployment change.

Pure frontend CDR import and statistics tool.

This GitHub Pages build runs fully in the browser. Imported files are parsed locally and are not uploaded to a server. Large imports use a dedicated Web Worker and a temporary IndexedDB database in the user's browser. There is no server database or upload API.

## Features

- Import one multi-phone workbook containing `通聯紀錄`, `使用者資料`, and/or `網路歷程`. The derived `通聯整合歷程紀錄` sheet is excluded to prevent duplicates.
- Empty pages and legacy formats retain the original workbench interface. Legacy batches analyze every imported target and use the sidebar date dialog. Only successfully imported original multi-phone workbooks show target controls, network history and network PDF. Mixed batches use this mode when any successful source is an original multi-phone workbook.
- In multi-phone mode, search/select a target and switch between calls, network history, profile, statistics and time/hotspot analysis. Target and inclusive dates also scope attachment exports. JSON volumes preserve the source format and interface mode; older unmarked JSON uses the legacy interface.
- Stream ZIP/XML and shared strings to temporary local storage. Reads and writes are bounded by bytes; lists show up to 500 rows per page. Import/export can be cancelled without replacing the previous dataset.
- Every page visit starts empty. Explicit clear removes imported data; closing a tab attempts cleanup, and next startup removes orphan sessions. Other active tabs stay isolated. Browser notes/settings remain until separately cleared.
- Export sequential attachment detail volumes (up to 10,000 records / 16 MiB each), plus separate full-scope summary volumes. Network XLSX/PDF and reimportable JSON volumes are included. Allow multiple downloads and keep all volumes.
- Import supported `.xlsx` and `.xml` CDR files in the browser.
- Supports the Chunghwa Telecom prosecutor-office XLSX layout with whitespace-normalized headers.
- Supports the Far EasTone prosecutor-office call XLSX layout, including repeated query sections and spacer-column variants.
- Supports Far EasTone Order `QueryInfo` / `CDRInfo` XLSX and XML layouts, including raw Excel dates/identifiers and UTF-8, Big5, or UTF-16 XML decoding.
- Supports Taiwan Mobile XML flattened into XLSX, preserving station continuation rows, raw numeric IMEI and query metadata.
- Merge every successfully parsed file selected in one import batch, retain each record's source, and replace the current batch only after a new batch has at least one success.
- Review the complete call list with 500-row pagination, subject metadata, time distribution, phone statistics, and a two-pane carrier ticket CSV builder for subscriber-profile or live-location queries.
- Filter hotspot addresses by any combination of Taiwan's 22 current counties/cities plus an unrecognized-address category; 台 and 臺 are classified together, with bulk select and clear controls.
- Apply an inclusive overall date range to the call list, profile summary and IMEI list, phone statistics, time/hotspot analysis, and attachment exports; resetting restores the complete import.
- Import an independent in-memory batch for multi-number location analysis, matching distinct target phones that appear in the same Taiwan county/district within an inclusive 30-minute window; each match can expand its source location, time, phone pair, and matched base-station details while sharing the existing browser-only phone notes.
- Export local attachments for the selected target/date range, covering time distribution, hotspots, calls, network history, profile data, and count/seconds phone rankings.
- Export/import local browser settings and export parsed workspace JSON.
- `/admin.html` and missing routes show a maintenance notice.

## Local Preview

```bash
python -m http.server 8088
```

Open `http://127.0.0.1:8088/`.

Use a current 64-bit Edge or Chrome with site storage enabled. Large original multi-phone XLSX files use streaming import. Other legacy formats are limited to 16 MiB per file; legacy XLSX also requires total expanded ZIP parts within 64 MiB; independent multi-number location analysis is limited to 16 MiB per batch. Insufficient storage or unsupported browsers produce an error rather than falling back to main-thread large-file parsing.

## Development checks

```bash
npm ci
node scripts/update-assets.js
npm test
npx playwright install chromium
PHONE_BROWSER_CHANNEL=chromium npm run test:browser
```

For a release, set `version` in `scripts/release-version.json`, then run the asset updater before testing. It synchronizes the page, styles, entry scripts, Worker imports and delayed export libraries, normalizes runtime text to LF, and regenerates CSP/SRI hashes. Re-running it without changes produces identical assets. Commit the generated assets with the version setting; the development-only setting and updater are excluded from the Pages artifact.

On Windows, browser tests default to installed Edge. `PHONE_TEST_URL` points the synthetic browser suite at a deployment. The private large-file test is opt-in with `PRIVATE_MULTI_XLSX`, stays on the local test origin, emits only pass/fail and performance metrics, and saves no source data or screenshots. Never set both variables for a private-file test.

`PRIVATE_COMPAT_DIR` enables recursive local compatibility verification of XLSX/XML and XML extracted from ZIP archives into an isolated OS temporary folder outside the repository. The verifier checks original parser parity, every list page, profiles, complete rankings, hour/hotspot data and query controls, plus batch merging. Output contains only anonymous file ordinals and stage pass/fail. XSL is a display template; the application does not accept ZIP imports.

`PHONE_VISUAL_BASELINE=1` compares empty and synthetic populated pages against `b922521` on desktop/mobile, light/dark and collapsed sidebars. `PHONE_PERF=1` exercises 100,000 and 200,000 synthetic records and export cancellation. Optional private and visual tests are skipped in CI.

## Privacy

- Keep all real CDR files outside the repository.
- The deployed app contains no analytics tag or phone-number lookup link.
- GitHub Pages publishes an explicit static-file allowlist instead of the repository root.
- Attachment XLSX/PDF files are generated one bounded volume at a time in the worker using pinned same-origin libraries; they are never sent to a server.
