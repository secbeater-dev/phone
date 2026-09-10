# Multi-phone import verification

All fixtures in the repository are synthetic. Private workbook paths, contents, identifiers, and record counts are excluded. Original data is never selected at the production origin.

## Local checks

- `npm test`: 52 passed, 2 private opt-in tests skipped.
- Browser regression suite: 6 existing checks passed, plus the added BrowserContext network check passed; 2 performance/private tests are opt-in. Network auditing excludes internal browser schemes such as `edge:` and checks all HTTP(S) requests including Worker resources.
- Opt-in synthetic performance: 100,000 call rows imported in 20 s; 200,000 in 41 s. Both show bounded 500-row pages and preserve network history. Large XLSX volume download succeeded; export cancellation returned in 105 ms and the existing dataset remained queryable.
- Opt-in private large-file check: import, target switching, call/network separation, profile/statistics/hotspots, cleanup and zero unexpected data requests passed. No source values or counts were emitted.
- Desktop/mobile screenshots used only synthetic inputs and were inspected locally. Screenshots are excluded from Git and deployment.

## Performance limits

The test host has 64 GiB RAM. The private-file run applied affinity to four logical CPUs and a 512 MiB V8 old-space limit. Sampled renderer/worker working-set peak was 347 MiB; total browser working-set peak was 1,430 MiB (includes browser, GPU and service processes). Import took 177 s; indexed target switch took 46 ms. Maximum observed timer delay was 305 ms during import and 89 ms during queries.

This is not a physical 4-core / 8-GiB-machine certification. The sub-200-ms UI target was not met on every sample: synthetic runs peaked at 369/194 ms, and the private run at 305 ms. Imports completed and cancellation remained usable. Files use bounded worker/disk processing; speed and available storage still depend on the user's browser, disk, and other running applications.

## Regression coverage

Coverage includes raw aliases and physical rows, rich/shared strings and numeric identifiers/dates, ignored integrated history, per-target users, equal-time pagination, global search/sorting, invalid dates, bounded wide rows/profiles/aggregates, note-sort storage, cancellation and worker recovery, multi-tab/orphan cleanup, JSON user-only continuation volumes, XLSX/network PDF output, quota failure, rapid target changes, legacy same-name worksheets, explicit legacy subject association, expanded-ZIP legacy limits, CSP/SRI and the Pages allowlist.

First GitHub Pages release completed successfully: [Actions run](https://github.com/secbeater-dev/phone/actions/runs/34478120876). Production preflight found cached pre-deployment 404 responses for two new scripts. The follow-up release uses asset version `20260910-multi-phone-v2` to bypass those cache entries. Production-origin synthetic verification follows that release.
