# Phone Workbench

> Maintenance requirement: read [`fullview.md`](./fullview.md) before making changes, and update it after any architecture, interface, parser, privacy, test, or deployment change.

Pure frontend CDR import and statistics tool.

This GitHub Pages build runs fully in the browser. Imported files are parsed locally and are not uploaded to a server. The map, coordinate list, coordinate completion tools, admin console, backend API, database, Docker deployment, and server-side workers are not included in this repository.

## Features

- Import supported `.xlsx` and `.xml` CDR files in the browser.
- Supports the Chunghwa Telecom prosecutor-office XLSX layout with whitespace-normalized headers.
- Supports the Far EasTone prosecutor-office call XLSX layout, including repeated query sections and spacer-column variants.
- Merge every successfully parsed file selected in one import batch, retain each record's source, and replace the current batch only after a new batch has at least one success.
- Review the complete call list with 500-row pagination, subject metadata, time distribution, phone statistics, and submission CSV output.
- Filter hotspot addresses by any combination of Taiwan's 22 current counties/cities plus an unrecognized-address category; 台 and 臺 are classified together, with bulk select and clear controls.
- Export a local attachment package: one six-sheet XLSX or six searchable-text PDFs covering time distribution, hotspots, calls, profile data, and count/seconds phone rankings.
- Export/import local browser settings and export parsed workspace JSON.
- `/admin.html` and missing routes show a maintenance notice.

## Local Preview

```bash
python -m http.server 8088
```

Open `http://127.0.0.1:8088/`.

## Privacy

- Keep all real CDR files outside the repository.
- The deployed app contains no analytics tag or phone-number lookup link.
- GitHub Pages publishes an explicit static-file allowlist instead of the repository root.
- Attachment XLSX/PDF files are created from in-memory data with pinned same-origin libraries; they are never sent to a server.
