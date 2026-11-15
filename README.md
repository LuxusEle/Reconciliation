# Bank reconciliation — plain HTML, CSS and JavaScript

No Python, TypeScript, framework, package installation or build step is needed. Excel reading, matching and report generation run entirely in the browser, in a background worker. Transaction data is never uploaded.

## Use the app

1. Open the hosted page, or start a local static preview as described below.
2. Select the LankaPay **D Navi** Excel file and the bank-side **Core** Excel file.
3. Optionally select which file contains a direction column and enter its header name. Leave the header blank if unavailable.
4. Click **Generate report**.
5. Review the **Should be reversed** count and transaction table. Use the All, 911, 910, Present in D Navi, or Absent from D Navi buttons above the table, and search by account number or sequence if needed. Button counts reflect the current search. Filters affect the table only; the Excel download still includes all results.
6. Click **Download Excel report** to save all results and source transaction sheets.

The optional direction setting appears after both files are selected. Direction values are displayed as supplied; they do not change reversal eligibility because both 910 responses and issuer/acquirer 911 timeouts should be reversed. Selecting files reads them locally; it does not upload them to GitHub or another server.

## Project files

| File | Purpose |
| --- | --- |
| `index.html` | Upload form and results layout |
| `styles.css` | Page styling and responsive layout |
| `app.js` | File selection, table, search and download |
| `worker.js` | Background Excel processing |
| `reconciliation.js` | Workbook parsing, matching and Excel export |
| `vendor/` | Local Excel library and its license |
| `.nojekyll` | Serves the project as a static GitHub Pages site |

## GitHub Pages

Commit these files and folders to your repository:

- `index.html`
- `styles.css`
- `app.js`
- `worker.js`
- `reconciliation.js`
- `vendor/` (including the SheetJS license)
- `.nojekyll`

In the repository, open **Settings → Pages**, select **Deploy from a branch**, choose your branch and **/(root)**, then save. Open the URL GitHub provides after deployment. Relative asset paths support repository Pages URLs such as `https://username.github.io/repository/`.

You do not need to publish `resources/`, `tests/`, `node_modules/` or `dist/`. Do not commit bank data to a public repository. The app never fetches sample workbooks from `resources/`.

For a local preview, serve the project with any static HTTP server, such as VS Code Live Server. Opening `index.html` directly using `file://` does not support the module worker reliably.

There is no `npm run dev` or `npm run build` step in this version. The earlier Python server and TypeScript source files have been replaced by the JavaScript files above.

## Matching and export

Upload D Navi and Core `.xlsx` files from the same period. Transactions are marked **Should be reversed** when they are present in Core but absent from D Navi, or when their D Navi `Action Code` is `910` (issuer signed off) or `911` (issuer or acquirer timeout). D Navi-only 910/911 transactions are included even if they never reached Core. Approval and responses other than 910/911 are not automatically reversed when a match exists. Match keys use exact account numbers and sequences padded to six digits. If multiple D Navi records share a key, any 910 or 911 triggers reversal and all matching records are retained in the export. Missing duplicate Core rows are retained individually. Preserve account leading zeros as Excel text. Amounts and dates are shown as supplied; there is no scaling.

D Navi requires `PAN`, `System Trace Number`, and `Action Code`. Core requires `Account No`, `SeqNo`, `Amount`, and `Date`. Exactly one matching sheet per input is required, with headers in the first 30 rows. Prefilled helper rows without accounts and sequences are ignored. Files are limited to 25 MB each; large expanded workbooks may still exceed browser memory.

The Excel download contains **Should be reversed**, **Core source**, and **D Navi source**. Source sheets retain all uploaded transaction columns and source row references. D Navi fields are populated for 910/911 transactions. They are blank for Core-only transactions. Core fields are blank for D Navi-only 910/911 transactions. Each result includes a reason, response code, optional direction value, and source row references. Formulas are not executed or exported as formulas. The screen displays up to 500 rows; Excel includes the entire reversal list. The app marks transactions and does not execute reversals.

## Understanding the row counts

- The results count refers to every Core-only transaction. The on-screen table previews up to 500 of them, and search can reduce the visible count further.
- The **Should be reversed** Excel sheet contains the full result, plus its header row. The two source sheets contain all parsed transactions from their respective uploads.
- The original RIB worksheet starts with a separately selected list of failed channel transactions. This app compares every uploaded Core transaction against D Navi, so its list can be larger than the original RIB. It does not apply RIB error-code filters because those fields are not supplied by the two input files.

## Tests (optional)

With Node.js 22+, run `npm test`. No dependency installation is necessary. Tests use the local `resources/Reconcilation.xlsx` sample; keep that workbook available locally for the sample test.

## Excel library

`vendor/xlsx.mjs` is SheetJS Community Edition 0.20.3, distributed under the accompanying Apache 2.0 license. It is bundled locally, so runtime processing does not depend on a CDN. Official distribution: https://cdn.sheetjs.com/.
