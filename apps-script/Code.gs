const PRODUCT_SHEET = 'ProductMaster';
const ESTIMATE_SHEET = 'Estimates';
const ORDERS_SHEET = 'Orders';

const PRODUCT_CSV_URL = 'https://raw.githubusercontent.com/Faizalemam/Purchase_Stock_Estimate/main/ProductMaster.csv';

function setupProject() {
  setupSheets();
  importProductsFromGitHub();
  return 'Setup complete';
}

function importProductsFromGitHub() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh) sh = ss.insertSheet(PRODUCT_SHEET);

  const response = UrlFetchApp.fetch(PRODUCT_CSV_URL, {muteHttpExceptions: true});
  if (response.getResponseCode() !== 200) {
    throw new Error('Unable to download ProductMaster.csv from GitHub. HTTP ' + response.getResponseCode());
  }

  const rows = Utilities.parseCsv(response.getContentText('UTF-8'));
  if (!rows || rows.length < 2) throw new Error('ProductMaster.csv is empty or invalid.');

  sh.clearContents();
  sh.getRange(1, 1, rows.length, 4).setValues(rows.map(r => [
    String(r[0] || '').trim(),
    String(r[1] || '').trim(),
    String(r[2] || '').trim(),
    r[0] === 'Code' ? 'Cost' : Number(r[3] || 0)
  ]));
  sh.getRange('A1:D1').setFontWeight('bold').setBackground('#e5e7eb').setFontColor('#111827');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 300);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 110);
  if (rows.length > 1) sh.getRange(2, 4, rows.length - 1, 1).setNumberFormat('#,##0.00');

  return {ok:true, products: rows.length - 1};
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'products');
    if (action === 'products') return json_(getProducts_());
    return json_({ok:false,error:'Unknown action'});
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

function doPost(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '');
    if (action === 'saveEstimate') {
      const lines = JSON.parse(e.parameter.lines || '[]');
      const estimateDate = String(e.parameter.estimateDate || '').trim();
      return json_(saveEstimate_(lines, estimateDate));
    }
    return json_({ok:false,error:'Unknown action'});
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

function getProducts_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${PRODUCT_SHEET}`);

  const last = sh.getLastRow();
  if (last < 2) return {ok:true, products:[]};

  const values = sh.getRange(2,1,last-1,4).getValues();
  const products = values
    .filter(r => String(r[0]).trim())
    .map(r => ({
      code: String(r[0]).trim(),
      name: String(r[1] || '').trim(),
      uom: String(r[2] || '').trim(),
      cost: Number(r[3] || 0)
    }));

  return {ok:true, products};
}

function getProductMap_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${PRODUCT_SHEET}`);

  const last = sh.getLastRow();
  if (last < 2) throw new Error('Product master is empty.');

  const values = sh.getRange(2,1,last-1,4).getValues();
  const map = {};
  values.forEach((r, i) => {
    const code = String(r[0] || '').trim().toUpperCase();
    if (!code) return;
    map[code] = {
      row: i + 2,
      code: String(r[0] || '').trim(),
      name: String(r[1] || '').trim(),
      uom: String(r[2] || '').trim(),
      cost: Number(r[3] || 0)
    };
  });
  return map;
}

function saveEstimate_(lines, estimateDate) {
  if (!Array.isArray(lines) || !lines.length) throw new Error('No estimate lines received.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(estimateDate)) throw new Error('Invalid Estimate Date.');

  const ss = SpreadsheetApp.getActive();
  ensureHistorySheets_();

  const detailSh = ss.getSheetByName(ESTIMATE_SHEET);
  const orderSh = ss.getSheetByName(ORDERS_SHEET);
  const productSh = ss.getSheetByName(PRODUCT_SHEET);
  const productMap = getProductMap_();

  const tz = Session.getScriptTimeZone() || 'Asia/Riyadh';
  const now = new Date();
  const reference = 'EST-' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss');
  const rows = [];
  const masterUpdates = [];
  let grandTotal = 0;
  let validLines = 0;

  lines.forEach(x => {
    const codeKey = String(x.code || '').trim().toUpperCase();
    const qty = Number(x.qty || 0);
    if (!codeKey || qty <= 0) return;

    const p = productMap[codeKey];
    if (!p) throw new Error('Unknown Product Code: ' + codeKey);

    const enteredCost = Number(x.cost);
    const cost = Number.isFinite(enteredCost) && enteredCost >= 0 ? enteredCost : p.cost;
    const total = qty * cost;

    if (Math.abs(cost - p.cost) > 0.000001) {
      masterUpdates.push({row: p.row, cost: cost});
      p.cost = cost;
    }

    grandTotal += total;
    validLines++;

    rows.push([
      reference,
      now,
      p.code,
      p.name,
      p.uom,
      qty,
      cost,
      total,
      estimateDate,
      0
    ]);
  });

  if (!rows.length) throw new Error('No valid lines to save.');
  rows.forEach(r => r[9] = grandTotal);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    masterUpdates.forEach(u => {
      productSh.getRange(u.row, 4).setValue(u.cost).setNumberFormat('#,##0.00');
    });

    detailSh.getRange(detailSh.getLastRow()+1,1,rows.length,10).setValues(rows);
    orderSh.appendRow([
      reference,
      estimateDate,
      now,
      validLines,
      grandTotal
    ]);
  } finally {
    lock.releaseLock();
  }

  return {
    ok:true,
    reference,
    estimateDate,
    grandTotal,
    lines:validLines,
    masterCostsUpdated:masterUpdates.length
  };
}

function ensureHistorySheets_() {
  const ss = SpreadsheetApp.getActive();

  let e = ss.getSheetByName(ESTIMATE_SHEET);
  if (!e) e = ss.insertSheet(ESTIMATE_SHEET);
  e.getRange('A1:J1').setValues([[
    'Reference','Saved At','Code','Product Name','UOM','QTY','Cost','Line Total','Estimate Date','Order Grand Total'
  ]]);
  e.getRange('A1:J1').setFontWeight('bold').setBackground('#e5e7eb').setFontColor('#111827');
  e.setFrozenRows(1);
  e.getRange('G:H').setNumberFormat('#,##0.00');
  e.getRange('J:J').setNumberFormat('#,##0.00');

  let o = ss.getSheetByName(ORDERS_SHEET);
  if (!o) o = ss.insertSheet(ORDERS_SHEET);
  o.getRange('A1:E1').setValues([['Reference','Estimate Date','Saved At','Lines','Grand Total']]);
  o.getRange('A1:E1').setFontWeight('bold').setBackground('#e5e7eb').setFontColor('#111827');
  o.setFrozenRows(1);
  o.getRange('E:E').setNumberFormat('#,##0.00');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();

  let p = ss.getSheetByName(PRODUCT_SHEET);
  if (!p) p = ss.insertSheet(PRODUCT_SHEET);
  if (p.getLastRow() === 0) {
    p.getRange('A1:D1').setValues([['Code','Product Name','UOM','Cost']]);
    p.getRange('A1:D1').setFontWeight('bold').setBackground('#e5e7eb').setFontColor('#111827');
    p.setFrozenRows(1);
    p.setColumnWidths(1,4,150);
  }

  ensureHistorySheets_();
}

function json_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
