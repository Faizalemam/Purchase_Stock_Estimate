const PRODUCT_SHEET = 'ProductMaster';
const ESTIMATE_SHEET = 'Estimates';
const ORDERS_SHEET = 'Orders';

const SESSION_TTL_SECONDS = 8 * 60 * 60;
const LOGIN_WINDOW_SECONDS = 60;
const MAX_LOGIN_FAILURES = 8;

const UOM_OVERRIDES = {
  '8001': 'CRT × 24 PCS',
  '8003': 'CRT × 24 PCS',
  '8002': 'CRT × 30 PCS',
  '8020': 'CRT × 28 PCS',
  '8004': 'CRT × 24 PCS'
};

function applyUomOverride_(code, uom) {
  const key = String(code || '').trim().toUpperCase();
  return UOM_OVERRIDES[key] || String(uom || '').trim();
}

function setupProject() {
  setupSheets();
  applyUomOverridesToMaster_();
  ensureSecurityConfigured_();
  return 'Setup complete';
}

function ensureSecurityConfigured_() {
  const password = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  if (!password || password.length < 10) {
    throw new Error('Security is not configured. Add Script Property APP_PASSWORD with at least 10 characters.');
  }
  return true;
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '');

    if (action === 'health') {
      return json_({ok:true, service:'Purchase Stock Estimate'});
    }

    const token = String((e && e.parameter && e.parameter.token) || '');
    requireSession_(token);

    if (action === 'products') return json_(getProducts_());
    if (action === 'orders') return json_(getOrders_(e && e.parameter ? e.parameter : {}));
    if (action === 'order') {
      const reference = String((e && e.parameter && e.parameter.reference) || '').trim();
      return json_(getOrder_(reference));
    }

    return json_({ok:false,error:'Unknown action'});
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

function doPost(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '');

    if (action === 'login') {
      const password = String(e.parameter.password || '');
      const clientId = normalizeClientId_(e.parameter.clientId || 'anonymous');
      return json_(login_(password, clientId));
    }

    if (action === 'logout') {
      const token = String(e.parameter.token || '');
      logout_(token);
      return json_({ok:true});
    }

    const token = String(e.parameter.token || '');
    requireSession_(token);

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

function login_(password, clientId) {
  ensureSecurityConfigured_();
  const cache = CacheService.getScriptCache();
  const failKey = 'login-fail:' + clientId;
  const failures = Number(cache.get(failKey) || 0);

  if (failures >= MAX_LOGIN_FAILURES) {
    throw new Error('Too many failed login attempts. Please wait one minute and try again.');
  }

  const expected = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  if (!secureEquals_(password, expected)) {
    cache.put(failKey, String(failures + 1), LOGIN_WINDOW_SECONDS);
    throw new Error('Invalid password.');
  }

  cache.remove(failKey);
  const token = Utilities.getUuid() + '-' + Utilities.getUuid();
  cache.put('session:' + token, '1', SESSION_TTL_SECONDS);
  return {ok:true, token:token, expiresIn:SESSION_TTL_SECONDS};
}

function logout_(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('session:' + token);
}

function requireSession_(token) {
  if (!token || token.length < 20) throw new Error('Unauthorized');
  const cache = CacheService.getScriptCache();
  const key = 'session:' + token;
  if (cache.get(key) !== '1') throw new Error('Session expired. Please sign in again.');
  cache.put(key, '1', SESSION_TTL_SECONDS);
  return true;
}

function secureEquals_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function normalizeClientId_(value) {
  const v = String(value || '').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 80);
  return v || 'anonymous';
}

function getProducts_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${PRODUCT_SHEET}`);
  const last = sh.getLastRow();
  if (last < 2) return {ok:true, products:[]};

  const values = sh.getRange(2, 1, last - 1, 4).getValues();
  const products = values.filter(r => String(r[0]).trim()).map(r => ({
    code:String(r[0]).trim(),
    name:String(r[1] || '').trim(),
    uom:applyUomOverride_(r[0], r[2]),
    cost:Number(r[3] || 0)
  }));
  return {ok:true, products:products};
}

function getOrders_(params) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ORDERS_SHEET);
  if (!sh || sh.getLastRow() < 2) return {ok:true, orders:[]};

  const tz = Session.getScriptTimeZone() || 'Asia/Riyadh';
  const requestedLimit = Number(params.limit || 200);
  const limit = Math.max(1, Math.min(500, Number.isFinite(requestedLimit) ? requestedLimit : 200));
  const dateFilter = String(params.date || '').trim();
  const refFilter = String(params.reference || '').trim().toUpperCase();

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  const orders = [];

  for (let i = values.length - 1; i >= 0 && orders.length < limit; i--) {
    const r = values[i];
    const reference = String(r[0] || '').trim();
    const estimateDate = normalizeSheetDate_(r[1], tz, 'yyyy-MM-dd');
    if (!reference) continue;
    if (dateFilter && estimateDate !== dateFilter) continue;
    if (refFilter && reference.toUpperCase().indexOf(refFilter) === -1) continue;

    orders.push({
      reference:reference,
      estimateDate:estimateDate,
      savedAt:normalizeSheetDate_(r[2], tz, 'yyyy-MM-dd HH:mm:ss'),
      lines:Number(r[3] || 0),
      grandTotal:Number(r[4] || 0)
    });
  }

  return {ok:true, orders:orders};
}

function getOrder_(reference) {
  reference = String(reference || '').trim();
  if (!/^EST-[A-Za-z0-9_-]{6,40}$/.test(reference)) throw new Error('Invalid reference.');

  const ss = SpreadsheetApp.getActive();
  const detailSh = ss.getSheetByName(ESTIMATE_SHEET);
  if (!detailSh || detailSh.getLastRow() < 2) throw new Error('No estimate history found.');

  const tz = Session.getScriptTimeZone() || 'Asia/Riyadh';
  const range = detailSh.getRange(2, 1, detailSh.getLastRow() - 1, 10);
  const matches = range.getValues().filter(r => String(r[0] || '').trim() === reference);
  if (!matches.length) throw new Error('Order not found.');

  const first = matches[0];
  const lines = matches.map(r => ({
    code:String(r[2] || '').trim(),
    name:String(r[3] || '').trim(),
    uom:String(r[4] || '').trim(),
    qty:Number(r[5] || 0),
    cost:Number(r[6] || 0),
    total:Number(r[7] || 0)
  }));

  return {
    ok:true,
    order:{
      reference:reference,
      savedAt:normalizeSheetDate_(first[1], tz, 'yyyy-MM-dd HH:mm:ss'),
      estimateDate:normalizeSheetDate_(first[8], tz, 'yyyy-MM-dd'),
      grandTotal:Number(first[9] || lines.reduce((s,x)=>s+x.total,0)),
      lines:lines
    }
  };
}

function normalizeSheetDate_(value, tz, pattern) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, tz, pattern);
  const text = String(value || '').trim();
  if (!text) return '';
  if (pattern === 'yyyy-MM-dd' && /^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0,10);
  return text;
}

function getProductMap_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${PRODUCT_SHEET}`);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Product master is empty.');

  const values = sh.getRange(2, 1, last - 1, 4).getValues();
  const map = {};
  values.forEach((r, i) => {
    const code = String(r[0] || '').trim().toUpperCase();
    if (!code) return;
    map[code] = {
      row:i + 2,
      code:String(r[0] || '').trim(),
      name:String(r[1] || '').trim(),
      uom:applyUomOverride_(r[0], r[2]),
      cost:Number(r[3] || 0)
    };
  });
  return map;
}

function saveEstimate_(lines, estimateDate) {
  if (!Array.isArray(lines) || !lines.length) throw new Error('No estimate lines received.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(estimateDate)) throw new Error('Invalid Estimate Date.');
  if (lines.length > 300) throw new Error('Too many estimate lines. Maximum is 300.');

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
  const masterUpdatesByRow = {};
  let grandTotal = 0;
  let validLines = 0;

  lines.forEach(x => {
    const codeKey = String(x.code || '').trim().toUpperCase();
    const qty = Number(x.qty || 0);
    if (!codeKey || !Number.isFinite(qty) || qty <= 0 || qty > 1000000) return;

    const p = productMap[codeKey];
    if (!p) throw new Error('Unknown Product Code: ' + codeKey);

    const enteredCost = Number(x.cost);
    const cost = Number.isFinite(enteredCost) && enteredCost >= 0 && enteredCost <= 10000000 ? enteredCost : p.cost;
    const total = qty * cost;
    if (!Number.isFinite(total)) throw new Error('Invalid amount for product ' + p.code);

    if (Math.abs(cost - p.cost) > 0.000001) {
      masterUpdatesByRow[p.row] = cost;
      p.cost = cost;
    }

    grandTotal += total;
    validLines++;
    rows.push([reference, now, p.code, p.name, p.uom, qty, cost, total, estimateDate, 0]);
  });

  if (!rows.length) throw new Error('No valid lines to save.');
  if (!Number.isFinite(grandTotal) || grandTotal > 1000000000) throw new Error('Invalid grand total.');
  rows.forEach(r => r[9] = grandTotal);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    Object.keys(masterUpdatesByRow).forEach(row => {
      productSh.getRange(Number(row), 4).setValue(masterUpdatesByRow[row]).setNumberFormat('#,##0.00');
    });
    applyUomOverridesToMaster_();
    detailSh.getRange(detailSh.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
    orderSh.appendRow([reference, estimateDate, now, validLines, grandTotal]);
  } finally {
    lock.releaseLock();
  }

  return {ok:true, reference:reference, estimateDate:estimateDate, grandTotal:grandTotal, lines:validLines, masterCostsUpdated:Object.keys(masterUpdatesByRow).length};
}

function applyUomOverridesToMaster_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh || sh.getLastRow() < 2) return;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  values.forEach((r, i) => {
    const code = String(r[0] || '').trim().toUpperCase();
    const override = UOM_OVERRIDES[code];
    if (override && String(r[2] || '').trim() !== override) sh.getRange(i + 2, 3).setValue(override);
  });
}

function ensureHistorySheets_() {
  const ss = SpreadsheetApp.getActive();
  let e = ss.getSheetByName(ESTIMATE_SHEET);
  if (!e) e = ss.insertSheet(ESTIMATE_SHEET);
  e.getRange('A1:J1').setValues([['Reference','Saved At','Code','Product Name','UOM','QTY','Cost','Line Total','Estimate Date','Order Grand Total']]);
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
    p.setColumnWidth(1,120);
    p.setColumnWidth(2,300);
    p.setColumnWidth(3,140);
    p.setColumnWidth(4,110);
  }
  ensureHistorySheets_();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
