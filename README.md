# Sheet Assistant (Excel add-in)

## Before you load it
1. Host the folder on HTTPS (for example GitHub Pages, Netlify, or Azure Static Web Apps).
2. Replace every `YOUR-HOST` in `manifest.xml` with your real host.
3. Add `icon-32.png` and `icon-80.png` to the same host (required by the manifest).

## Load in Excel
- Excel on the web: Insert > Add-ins > Upload My Add-in > pick `manifest.xml`.
- Excel for Windows or Mac: Insert > My Add-ins > Upload My Add-in, or use the `yo office` dev server for local testing.

## Notes
- Do not put an AI API key in these files. Route model calls through your own backend.
- The "Run" button currently only reads the sheet. Connect a backend to make it useful.
