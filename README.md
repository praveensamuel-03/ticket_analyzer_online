# CCSR Ticket Performance Analyzer

## Files
- `index.html` — dashboard structure
- `css/style.css` — layout and styling
- `js/script.js` — Excel parsing, automatic header detection, field mapping, KPIs and charts

## Run
Open `index.html` in Chrome or Edge. Chart.js and SheetJS are loaded from CDNs, so an internet connection is required.

## Data behavior
- No sample/demo data is embedded or loaded on startup.
- Upload `.xlsx`, `.xls`, or `.csv`.
- The first worksheet is scanned for a header row containing a Jira-style `Key`/`Issue key` and `Summary` or `Status` column.
- The detected headers populate mapping controls. Default mapping recognizes common Jira columns including `Key`, `Status`, `Priority`, `Brand`, `BSH Brand`, `Assignee`, and `Owner`.
- Use the mapping controls if your workbook uses different column names.
- All parsing is done locally in the browser.
