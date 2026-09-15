# 最新版同步、相容性修復與重新上架

## Goal and constraints

Synchronize the local checkout from `34382fd` to upstream `5a92ebb`, preserve the old branch, repair release asset synchronization, verify compatibility, and publish to the existing GitHub Pages site at https://phone.secbeater.com/.

- Use only synthetic inputs; private CDR sources stay outside Git and all remote tests.
- Preserve multi-phone streaming XLSX, network history, Taiwan Mobile flattened XLSX, legacy UI, scoped queries, bounded storage and volume exports.
- Keep CSP, SRI, LF normalization and the explicit Pages allowlist.
- The user authorized merging, pushing and publishing after verification.

## Task 1: Release version and generated asset consistency

The old updater only replaces the literal August release in app/import-parser. Changing its version leaves current September references in app, import-parser, dataset-client and dataset-worker unchanged.

- Add `scripts/release-version.json` containing `{"version":"20260915-compat-repair-v1"}` as the single maintained release version.
- Update `scripts/update-assets.js` to read that setting and update every first-party loader version: HTML/CSS/entry scripts, app release constant and delayed attachment URLs, import-parser imports, dataset-client Worker URL and dataset-worker imports. Existing unversioned pinned vendor URLs used by Workers should receive the release version too so one graph is loaded consistently. Do not rewrite documentation or unrelated strings.
- Normalize runtime text to LF before hashing. Refresh delayed attachment integrity declarations, HTML entry script SRI and CSP hashes after content changes. Keep the CLI invocation `node scripts/update-assets.js` and preserve the original five entry scripts and their order.
- Update the asset assertions in `tests/app.test.js` to consume the single version setting.
- Add regression tests that execute the actual updater in isolated temporary fixtures. Prove a release upgrade updates every loader, produces matching hashes after byte changes/CRLF input, and a second run is byte-identical. Observe the upgrade regression failing before implementing the fix.
- Generate the actual release assets. Limit application code changes to loader versions unless a separately reproduced compatibility defect requires correction.
- Root agent owns README/fullview/evidence and local documents; do not edit those in this task. Report files changed, red/green evidence, tests and any concerns. Explicitly stage only task files when committing.

## Task 2: Compatibility and evidence

- Run npm ci, the Node suite and Chromium browser suite; enable PHONE_PERF for the existing 100,000/200,000 synthetic import and export cancellation test.
- Check legacy/multi/mixed imports, partial/all failure and cancellation, target/date queries, all list pages, profiles/rankings/hotspots, JSON roundtrip and XLSX/PDF exports through the existing suites.
- Update fullview and its changelog, README release instructions and a dated verification document; update the repository-external supported-formats.md and update.md.
- Perform task-scoped and final code review, resolving concrete findings before publishing.

## Task 3: Publish and verify

- Merge the tested repair branch to main, confirm remote has not advanced, push normally and wait for Pages success.
- Compare every published runtime asset with local SHA-256, check HTML version/SRI/CSP, and run the synthetic browser suite against the production URL.
- Record actual results with commit/deployment links. If this release causes a production regression, revert the repair commit and deploy the revert.
