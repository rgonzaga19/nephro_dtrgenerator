# Nephro DTR Generator

A browser-based attendance and daily time record generator. Upload a biometric `.xlsx` punch log, review the DTR sheet, and download the completed workbook.

## Project layout

- `index.html` — GitHub Pages entry point
- `assets/css/styles.css` — interface styles
- `assets/js/app.js` — DTR processing and workbook export
- `DTR Generator.html` — redirect kept for the previous local page URL

## Run locally

From this folder, run `python -m http.server 8001` and open <http://127.0.0.1:8001/>.

## Publish with GitHub Pages

Set **Settings → Pages → Build and deployment** to **Deploy from a branch**, select `main` and `/(root)`, then save. The site uses ExcelJS from a public CDN, so workbook downloads need an internet connection.