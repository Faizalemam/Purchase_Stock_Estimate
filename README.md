# Purchase Stock Estimate — GitHub Pages + Google Sheets

## What it does
- User enters only Product Code + QTY.
- Product Name, UOM and Cost come from Google Sheet `ProductMaster`.
- Line Total and Grand Total calculate automatically.
- Save Estimate writes all lines into Google Sheet `Estimates`.
- Mobile friendly and printable.

## Google Sheet
Create a Google Sheet and open Extensions → Apps Script.

1. Paste `apps-script/Code.gs`.
2. Run `setupSheets()` once.
3. In Project Settings set timezone to `Asia/Riyadh`.
4. Deploy → New deployment → Web app.
5. Execute as: Me
6. Who has access: Anyone
7. Copy the Web App URL.

### ProductMaster columns
A: Code
B: Product Name
C: UOM
D: Cost

Paste all products below row 1.

## GitHub Pages
1. Create a GitHub repository.
2. Upload `index.html`, `style.css`, `app.js`, `config.js`.
3. Open `config.js` and paste your Apps Script Web App URL.
4. GitHub repository → Settings → Pages.
5. Deploy from branch → `main` / root.
6. Open the GitHub Pages URL.

## Notes
- Product codes are matched case-insensitively in the browser.
- The server records the estimate using a generated reference like `EST-20261008-104500`.
- Costs are stored server-side in the product sheet and also written into estimate history.


## Odoo Product Master Included
This package already contains `ProductMaster.csv` generated from the uploaded Odoo product.template export.

- Products exported: 315
- Blank product codes skipped: 0
- Duplicate product codes found: 0
- Products with zero cost: 23

Import `ProductMaster.csv` into the Google Sheet tab named `ProductMaster`.
