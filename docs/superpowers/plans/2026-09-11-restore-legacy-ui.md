# Restore the original workbench interface

Approved on 2026-09-11. Visual baseline: b922521. Implementation starts from 3d1eb8a.

- Empty pages retain the original navigation, hour tiles/chart, tables, forms and dialogs.
- Legacy inputs use the original interface and all targets/all record kinds. Only successful inputs with multi_phone_streaming_xlsx provenance enable the new target/date toolbar, network navigation and network PDF. Mixed batches enable these if any successful member has this provenance.
- Failures/cancellation retain the previous dataset and presentation. Clearing restores the original empty page. Format switches update the mode only after successful adoption.
- Preserve source provenance through merged datasets and JSON volumes; unmarked older JSON uses legacy presentation.
- Keep Worker/IndexedDB, bounded reads, paging, cancellation and volume exports. Feed original DOM/components from asynchronous queries; do not rebuild a full in-memory workspace for presentation.
- Restore column resizing, notes, rank mode, date dialog, hour draft/apply, county dialog and address/time hotspot search.
- Validate all private XLSX/XML in the user-specified folder recursively, and XML inside ZIP extracted outside the repository. XSL is a presentation companion, not an import format. No ZIP upload feature.
- Private verification emits only anonymous file/stage pass/fail. No private source content, identifiers, filenames, screenshots or exports enter the repository or external services.
- Test visual parity on desktop/mobile, light/dark and collapsed sidebar, synthetic behavior and large imports. Fix every private compatibility failure, add synthetic regression cases, and rerun.
- Update fullview.md, verification evidence, release version/SRI; push GitHub only after checks; verify deployed assets and live-site flows after deployment.
