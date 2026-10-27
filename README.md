# DTR Generator

A browser-based attendance and daily time record generator. Upload a biometric `.xlsx` punch log, review the DTR sheet, and download the completed workbook.

## Run locally

From this folder, run `python -m http.server 8001` and open <http://127.0.0.1:8001/>.

## Publish with GitHub Pages

This is a static site. The root `index.html` is the GitHub Pages entry point. In the repository settings, open **Pages**, choose **Deploy from a branch**, select the `main` branch and `/ (root)`, and save. GitHub will publish the site at the repository’s Pages URL.

The workbook export uses ExcelJS from a public CDN, so downloading `.xlsx` files requires an internet connection.