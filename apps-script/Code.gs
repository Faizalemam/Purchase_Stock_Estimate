const PRODUCT_SHEET = 'ProductMaster';
const ESTIMATE_SHEET = 'Estimates';

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
      return json_(saveEstimate_(lines));
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

function saveEstimate_(lines) {
  if (!Array.isArray(lines) || !lines.length) throw new Error('No estimate lines received.');

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ESTIMATE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ESTIMATE_SHEET);
    sh.appendRow(['Reference','Timestamp','Code','Product Name','UOM','QTY','Cost','Total Amount']);
    sh.setFrozenRows(1);
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Riyadh';
  const now = new Date();
  const reference = 'EST-' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss');
  const rows = [];
  let grandTotal = 0;

  lines.forEach(x => {
    const qty = Number(x.qty || 0);
    const cost = Number(x.cost || 0);
    const total = qty * cost;
    if (!x.code || qty <= 0) return;
    grandTotal += total;
    rows.push([
      reference,
      now,
      String(x.code),
      String(x.name || ''),
      String(x.uom || ''),
      qty,
      cost,
      total
    ]);
  });

  if (!rows.length) throw new Error('No valid lines to save.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    sh.getRange(sh.getLastRow()+1,1,rows.length,8).setValues(rows);
  } finally {
    lock.releaseLock();
  }

  return {ok:true,reference,grandTotal};
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();

  let p = ss.getSheetByName(PRODUCT_SHEET);
  if (!p) p = ss.insertSheet(PRODUCT_SHEET);
  if (p.getLastRow() === 0) {
    p.getRange('A1:D1').setValues([['Code','Product Name','UOM','Cost']]);
    p.getRange('A1:D1').setFontWeight('bold').setBackground('#0f4c81').setFontColor('#ffffff');
    p.setFrozenRows(1);
    p.setColumnWidths(1,4,150);
  }

  let e = ss.getSheetByName(ESTIMATE_SHEET);
  if (!e) e = ss.insertSheet(ESTIMATE_SHEET);
  if (e.getLastRow() === 0) {
    e.getRange('A1:H1').setValues([['Reference','Timestamp','Code','Product Name','UOM','QTY','Cost','Total Amount']]);
    e.getRange('A1:H1').setFontWeight('bold').setBackground('#0f4c81').setFontColor('#ffffff');
    e.setFrozenRows(1);
  }
}
function json_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
