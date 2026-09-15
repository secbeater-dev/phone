# Release asset synchronization verification

## Baseline and changes

The original local working branch `20260827-notice-card-layout` at `34382fd` was retained. Local `main` was fast-forwarded from `591a1a3` to upstream `5a92ebb9b5ef145195e15d94bc1dbd4d7e6ac040`; repair work uses `fix/20260915-release-assets`.

Upstream added multi-phone streaming XLSX imports, browser session IndexedDB, target-scoped network/call queries, sequential XLSX/PDF/JSON volumes and Taiwan Mobile flattened XML-to-XLSX support. Legacy controls, all-target analysis and independent multi-number location analysis remain available. The synchronization brought 45 changed files relative to the old local main, including Node/Playwright tests and the expanded Pages allowlist.

The release updater previously replaced only an August version literal. An in-memory upgrade rehearsal changed the HTML release while app/import-parser/dataset-client/dataset-worker kept the September release. The repair introduces a single development-only release setting and synchronizes all runtime loaders before recalculating integrity hashes. Release: `20260915-compat-repair-v1`.

All inputs for this verification are synthetic. Private opt-in checks remain disabled. No original CDR files were changed or uploaded.

## Verification before the repair

- `npm ci`: completed; audit reported zero vulnerabilities.
- Original Node suite: 63 passed, 2 private opt-in checks skipped, zero failures.
- Chromium installed using the project's pinned Playwright dependency.

## Release verification

- The new complete-loader upgrade regression failed against the original updater, which emitted the old release rather than the fixture's requested release. After the repair both updater regressions passed, including LF/hash verification and byte-identical repeated generation.
- `npm test` after generation: 65 passed, 2 private opt-in checks skipped, zero failures (67 total).
- Windows full Chromium 151.0.7922.34 could not start: native SideBySide error 14001 reported a missing browser-version assembly. The pinned Playwright `chromium-headless-shell` at the same version started successfully and passed the focused page/Worker network test. Local browser verification uses this channel; the existing Linux CI Chromium configuration is unchanged.

### Browser compatibility and large imports

`PHONE_BROWSER_CHANNEL=chromium-headless-shell PHONE_PERF=1 npm run test:browser` passed all 13 normal browser regressions. Two optional visual tests and two private tests were skipped. The combined performance test reached its 900-second timeout after completing 100,000 synthetic calls (285 s, maximum sampled UI timer lag 186 ms); the 200,000-call/export part was not completed in that channel.

The unchanged performance test was then run alone with the installed Edge 153.0.4234.32:

```bash
PHONE_BROWSER_CHANNEL=msedge PHONE_PERF=1 node --test tests/browser/performance.test.js
```

It passed in 120 s overall. Import timings were 22 s for 100,000 synthetic calls and 47 s for 200,000, with maximum sampled UI timer lag 233 ms and 107 ms respectively. Both imports retained 500-row pages. A large XLSX volume downloaded, export cancellation returned in 121 ms, and network history remained queryable afterward. These measurements are browser/host-specific; the initial 233-ms sample exceeded the earlier aspirational 200-ms UI target. No performance-test timeout or application logic was changed to obtain this result.

The 13 normal browser cases cover desktop/mobile controls, legacy expanded-file limits, initial/legacy/multi/mixed interfaces, partial/all import failure, dates, profiles, rankings/hotspots, JSON provenance and subjects roundtrips, scoped XLSX/network PDF, import cancellation, quota errors, target switching, tab cleanup and page/Worker request auditing.

Code review and production verification are recorded after those stages complete.
