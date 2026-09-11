# Legacy interface restoration verification

Baseline interface: `b922521`. Implementation started from the latest local and remote `main`, `3d1eb8a`. Release asset version: `20260911-legacy-ui-v1`.

All committed fixtures are synthetic. Private inputs remain outside the repository; no private filenames, values, identifiers, record counts, screenshots or exports are retained in the evidence or deployment. Browser tests of private inputs use only a local origin.

## Interface and behavior

- Empty views retain the original seven views, hour chart/tiles, sidebar and empty states. The target toolbar, network entry and network PDF are hidden by default.
- Legacy formats query all imported targets and record kinds through the original table, profile, ranking and hour/hotspot layouts. Dates use the original sidebar dialog.
- Only successfully adopted sources containing `multi_phone_streaming_xlsx` enable multi-phone controls. Mixed batches, failed/cancelled imports, format switching, clear, old JSON and source-preserving JSON roundtrips are covered.
- Column resizing, notes, rank modes, hour draft/apply, county selection, address/time hotspot search, complete pagination and scoped XLSX/PDF exports remain covered by browser tests. Clearing also resets the previous batch's ticket fields.
- Source provenance survives merges and JSON volumes. IndexedDB/Worker processing, bounded pages and sequential exports remain in place.

## Local evidence

- Node suite: 63 passed, 2 private opt-in checks skipped.
- Browser suite: 13 passed, 5 opt-in visual/performance/private checks skipped. Additional targeted regressions were observed failing before the corresponding fixes and passing afterward.
- Visual suite: 60 comparisons passed against the original revision: 35 empty-page views, 20 populated legacy analysis panels and 5 expanded hotspot panels. Coverage includes desktop/mobile, light/dark and sidebar collapse. Dimensions must match; the image comparison tolerates one colour level and at most 0.005% edge pixels for Chromium raster rounding. Populated analysis panels use equal sidebar status height; actual import status wording is separately covered by behavior tests. Screenshots contain synthetic data only and are excluded from Git.
- The requested private compatibility corpus passed for anonymous inputs 1–14, including XML extracted from ZIP archives outside the repository. Existing parsers were compared with the original revision's records, stations and subject data, excluding newly added schema fields. The newly supported flattened XML XLSX was reconciled against its source call/continuation rows and has a synthetic regression fixture.
- Every private input passed browser import, all list pages with duplicate multiplicities, profile fields and six summary values, complete rankings in both modes, all hour buckets, paged hotspot counts and expanded times, search, ascending/descending duration sort and inclusive date filtering. The merged batch passed the same analysis and query controls. Request auditing found no data submission or unexpected browser errors.
- Code review findings concerning legacy data-row ranking, invalid-date metadata, intermediate hotspot time search, hotspot percentages, duplicate profile values and IMEI ordering were fixed and rechecked. No remaining concrete P1/P2 findings were reported.

## Performance and deployment

The synthetic large-file test passed with 100,000 call rows imported in 21 s and 200,000 in 47 s, under a 512-MiB V8 old-space limit. Maximum measured UI timer delays were 145 ms and 212 ms. Both showed bounded 500-row lists. A large XLSX detail volume downloaded, export cancellation returned in 86 ms, and the dataset's network history remained queryable.

The private large-file run passed import, calls/network switching, profiles/rankings/hotspots, cleanup and request auditing. With verified affinity to four logical CPUs and a 512-MiB V8 old-space limit, import took 215 s and indexed target switching took 57 ms. Sampled peak browser working set was 1,518 MiB, including browser/GPU/service processes; renderer/Worker working set peaked at 376 MiB. Maximum timer delay was 358 ms during import and 286 ms during queries.

This is not a certification on physical 4-core / 8-GiB hardware; the development host has 64 GiB RAM. The sub-200-ms UI goal was not met on every sample, but imports completed, cancellation worked, and processing stayed in bounded Worker/disk operations. Available storage, disk speed and other running applications affect actual results.

Production deployment verification is added after the release is published.
