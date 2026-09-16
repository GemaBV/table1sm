# Table 1 — Supplementary Materials

Read-only web view for the Excel workbook in this repository.

Keep exactly one Excel file (`.xlsx`, `.xls`, `.xlsm`, or `.xlsb`) in the repository root. The deployment automatically detects it, shows all of its worksheets, and makes the original file available for download. The workbook filename can change without requiring code changes.

## Publish with GitHub Pages

1. Push this repository to the `main` branch on GitHub.
2. In **Settings → Pages → Build and deployment**, select **GitHub Actions**.
3. The included workflow will publish the site automatically.

The original Excel workbook remains available for download from the page. The browser renders it with the SheetJS Community Edition standalone script.
