# Multi-phone import verification

All fixtures in the repository are synthetic. Private workbook paths, contents, identifiers, and record counts are excluded. Original data is never selected at the production origin.

## Local checks

- `npm test`: 52 passed, 2 private opt-in tests skipped.
- Browser regression suite: 7 passed, 2 performance/private checks are opt-in. Network auditing excludes internal browser schemes such as `edge:` and checks all HTTP(S) requests including Worker resources.
- Opt-in synthetic performance: 100,000 call rows imported in 20 s; 200,000 in 41 s. Both show bounded 500-row pages and preserve network history. Large XLSX volume download succeeded; export cancellation returned in 105 ms and the existing dataset remained queryable.
- Opt-in private large-file check: import, target switching, call/network separation, profile/statistics/hotspots, cleanup and zero unexpected data requests passed. No source values or counts were emitted.
- Desktop/mobile screenshots used only synthetic inputs and were inspected locally. Screenshots are excluded from Git and deployment.

## Performance limits

The test host has 64 GiB RAM. The private-file run applied affinity to four logical CPUs and a 512 MiB V8 old-space limit. Sampled renderer/worker working-set peak was 347 MiB; total browser working-set peak was 1,430 MiB (includes browser, GPU and service processes). Import took 177 s; indexed target switch took 46 ms. Maximum observed timer delay was 305 ms during import and 89 ms during queries.

This is not a physical 4-core / 8-GiB-machine certification. The sub-200-ms UI target was not met on every sample: synthetic runs peaked at 369/194 ms, and the private run at 305 ms. Imports completed and cancellation remained usable. Files use bounded worker/disk processing; speed and available storage still depend on the user's browser, disk, and other running applications.

## Regression coverage

Coverage includes raw aliases and physical rows, rich/shared strings and numeric identifiers/dates, ignored integrated history, per-target users, equal-time pagination, global search/sorting, invalid dates, bounded wide rows/profiles/aggregates, note-sort storage, cancellation and worker recovery, multi-tab/orphan cleanup, JSON user-only continuation volumes, XLSX/network PDF output, quota failure, rapid target changes, legacy same-name worksheets, explicit legacy subject association, expanded-ZIP legacy limits, CSP/SRI and the Pages allowlist.

First GitHub Pages release completed successfully: [first Actions run](https://github.com/secbeater-dev/phone/actions/runs/34478120876). Production preflight found cached pre-deployment 404 responses for two new scripts. Asset version `20260910-multi-phone-v2` bypassed those entries; [follow-up deployment](https://github.com/secbeater-dev/phone/actions/runs/34478506647) succeeded.

At `https://phone.secbeater.com/`, all 12 checked runtime assets matched the local SHA-256 values. The full production browser regression suite passed 7 checks (2 opt-in skips): synthetic imports, target/date/kind switching, pagination, per-target profiles, XLSX/network PDF, JSON continuation volumes and reimport, cancellation, quota failure, legacy expansion guard, tab cleanup, responsive layout and network auditing.

The separate production-origin large/double-size synthetic test also passed: 100,000 call rows imported in 29 s and 200,000 in 61 s; maximum timer delays were 362/218 ms. A large XLSX detail volume downloaded successfully, cancellation returned in 118 ms, and network history remained queryable afterward. These are synthetic workload counts, not private-file counts.

The CDN injects an external beacon and inline challenge into HTML. Browser evidence showed the beacon request failed with `csp` and no HTTP response; `securitypolicyviolation` reported both scripts blocked, and there were no challenge iframes. Only same-origin GET requests reached the network, including the CDN's speculation-rules resource. The audit deliberately distinguishes CSP-blocked request attempts from successful network requests; it would fail if that beacon were allowed through.
