# Phone Workbench

> Maintenance requirement: read [`fullview.md`](./fullview.md) before making changes, and update it after any architecture, interface, parser, privacy, test, or deployment change.

Pure frontend CDR import and statistics tool.

This GitHub Pages build runs fully in the browser. Imported files are parsed locally and are not uploaded to a server. The map, coordinate list, coordinate completion tools, admin console, backend API, database, Docker deployment, and server-side workers are not included in this repository.

## Features

- Import supported `.xlsx` and `.xml` CDR files in the browser.
- Supports the Chunghwa Telecom prosecutor-office XLSX layout with whitespace-normalized headers.
- Review call records, subject metadata, time distribution, phone statistics, and submission CSV output.
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
