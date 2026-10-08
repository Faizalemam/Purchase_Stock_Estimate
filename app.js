const CFG = window.APP_CONFIG || {};
let products = new Map();
let rowId = 0;
let authToken = sessionStorage.getItem('purchaseEstimateToken') || '';
let clientId = localStorage.getItem('purchaseEstimateClientId') || '';

if(!clientId){
  clientId = (crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`).replace(/[^a-zA-Z0-9_-]/g,'');
  localStorage.setItem('purchaseEstimateClientId', clientId);
}

const loginScreen = document.getElementById('loginScreen');
const appShell = document.getElementById('appShell');
const loginForm = document.getElementById('loginForm');
const loginPassword = document.getElementById('loginPassword');
const loginBtn = document.getElementById('loginBtn');
const loginMessage = document.getElementById('loginMessage');

const rowsEl = document.getElementById('rows');
const statusEl = document.getElementById('status');
const grandTotalEl = document.getElementById('grandTotal');
const codesList = document.getElementById('productCodes');
const savedBox = document.getElementById('savedBox');
const estimateDateEl = document.getElementById('estimateDate');

function money(v){
  return `${CFG.CURRENCY || 'SAR'} ${Number(v || 0).toFixed(2)}`;
}

function localDateISO(){
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,'0');
  const day = String(d.getDate()).padStart(2,'0');
  return `${y}-${m}-${day}`;
}

function setStatus(msg, error=false){
  statusEl.textContent = msg;
  statusEl.classList.toggle('error', error);
}

function validApi(){
  return CFG.API_URL && !CFG.API_URL.includes('PASTE_YOUR');
}

function showLogin(message=''){
  authToken = '';
  sessionStorage.removeItem('purchaseEstimateToken');
  products = new Map();
  codesList.innerHTML = '';
  appShell.classList.add('hidden');
  loginScreen.classList.remove('hidden');
  loginMessage.textContent = message;
  loginPassword.value = '';
  setTimeout(()=>loginPassword.focus(), 100);
}

function showApp(){
  loginScreen.classList.add('hidden');
  appShell.classList.remove('hidden');
}

function authError(message){
  const text = String(message || '');
  return text.includes('Unauthorized') || text.includes('Session expired');
}

async function login(password){
  if(!validApi()){
    loginMessage.textContent = 'Backend is not configured.';
    return;
  }

  loginBtn.disabled = true;
  loginMessage.textContent = 'Signing in…';

  try{
    const payload = new URLSearchParams();
    payload.set('action','login');
    payload.set('password',password);
    payload.set('clientId',clientId);

    const r = await fetch(CFG.API_URL,{method:'POST',body:payload});
    const data = await r.json();
    if(!data.ok) throw new Error(data.error || 'Login failed');

    authToken = data.token;
    sessionStorage.setItem('purchaseEstimateToken', authToken);
    loginMessage.textContent = '';
    showApp();
    clearAll();
    await loadProducts();
  }catch(e){
    showLogin(e.message || 'Login failed');
  }finally{
    loginBtn.disabled = false;
  }
}

async function logout(){
  try{
    if(authToken && validApi()){
      const payload = new URLSearchParams();
      payload.set('action','logout');
      payload.set('token',authToken);
      await fetch(CFG.API_URL,{method:'POST',body:payload});
    }
  }catch(_e){}
  showLogin('Signed out.');
}

async function loadProducts(){
  if(!validApi()){
    setStatus('Backend is not configured.', true);
    return;
  }
  if(!authToken){
    showLogin();
    return;
  }

  setStatus('Loading product master…');

  try{
    const url = `${CFG.API_URL}?action=products&token=${encodeURIComponent(authToken)}&t=${Date.now()}`;
    const r = await fetch(url,{cache:'no-store'});
    const data = await r.json();
    if(!data.ok) throw new Error(data.error || 'Unable to load products');

    products = new Map(data.products.map(p => [String(p.code).trim().toUpperCase(), p]));
    codesList.innerHTML = '';

    data.products.forEach(p=>{
      const o = document.createElement('option');
      o.value = p.code;
      o.label = p.name;
      codesList.appendChild(o);
    });

    setStatus(`${data.products.length} products loaded securely from Google Sheet. Cost can be edited manually if supplier rate is different.`);
    recalcAll();
  }catch(e){
    if(authError(e.message)) return showLogin('Session expired. Please sign in again.');
    setStatus(`Product load failed: ${e.message}`, true);
  }
}

function addRow(prefill={}){
  const id = ++rowId;
  const tr = document.createElement('tr');
  tr.dataset.id = id;
  tr.dataset.lastCode = '';
  tr.innerHTML = `
    <td class="idx"></td>
    <td><input class="code" list="productCodes" autocomplete="off" placeholder="Code" value="${prefill.code||''}"></td>
    <td><input class="name readonly" readonly></td>
    <td><input class="uom readonly" readonly></td>
    <td><input class="qty" inputmode="decimal" type="number" min="0" step="0.001" placeholder="0"></td>
    <td><input class="cost num manual-cost" inputmode="decimal" type="number" min="0" step="0.01" placeholder="0.00" title="Auto-filled from Product Master. You can edit this supplier cost manually."></td>
    <td><input class="total readonly num" readonly></td>
    <td class="remove-col"><button class="danger remove" title="Remove">×</button></td>`;
  rowsEl.appendChild(tr);

  const code = tr.querySelector('.code');
  const qty = tr.querySelector('.qty');
  const cost = tr.querySelector('.cost');

  code.addEventListener('input',()=>hydrateRow(tr));
  code.addEventListener('change',()=>hydrateRow(tr));
  qty.addEventListener('input',()=>recalcRow(tr));
  cost.addEventListener('input',()=>{
    tr.dataset.costOverridden = '1';
    recalcRow(tr);
  });

  tr.querySelector('.remove').addEventListener('click',()=>{
    tr.remove();
    renumber();
    calcGrand();
    if(!rowsEl.children.length) addRow();
  });

  renumber();
  if(prefill.code) hydrateRow(tr);
}

function hydrateRow(tr){
  const key = tr.querySelector('.code').value.trim().toUpperCase();
  const p = products.get(key);
  const codeChanged = tr.dataset.lastCode !== key;

  tr.querySelector('.code').value = key;
  tr.querySelector('.name').value = p?.name || '';
  tr.querySelector('.uom').value = p?.uom || '';

  if(codeChanged){
    tr.dataset.costOverridden = '0';
    tr.querySelector('.cost').value = p ? Number(p.cost||0).toFixed(2) : '';
    tr.dataset.lastCode = key;
  }else if(tr.dataset.costOverridden !== '1' && p){
    tr.querySelector('.cost').value = Number(p.cost||0).toFixed(2);
  }

  recalcRow(tr);
}

function recalcRow(tr){
  const qty = Number(tr.querySelector('.qty').value || 0);
  const cost = Number(tr.querySelector('.cost').value || 0);
  tr.querySelector('.total').value = (qty * cost).toFixed(2);
  calcGrand();
}

function recalcAll(){
  [...rowsEl.children].forEach(hydrateRow);
}

function calcGrand(){
  const total = [...rowsEl.children].reduce(
    (s,tr)=>s+Number(tr.querySelector('.total').value||0),0
  );
  grandTotalEl.textContent = money(total);
}

function renumber(){
  [...rowsEl.children].forEach((tr,i)=>tr.querySelector('.idx').textContent=i+1);
}

function clearAll(){
  rowsEl.innerHTML = '';
  rowId = 0;
  savedBox.classList.add('hidden');
  for(let i=0;i<8;i++) addRow();
  calcGrand();
}

function collectLines(){
  return [...rowsEl.children].map(tr=>({
    code:tr.querySelector('.code').value.trim(),
    name:tr.querySelector('.name').value,
    uom:tr.querySelector('.uom').value,
    qty:Number(tr.querySelector('.qty').value||0),
    cost:Number(tr.querySelector('.cost').value||0),
    total:Number(tr.querySelector('.total').value||0)
  })).filter(x=>x.code && x.qty>0 && x.name && x.cost>=0);
}

async function saveEstimate(){
  if(!validApi()) return setStatus('Backend is not configured.', true);
  if(!authToken) return showLogin();

  const lines = collectLines();
  const estimateDate = estimateDateEl.value;

  if(!estimateDate) return setStatus('Please select Estimate Date.', true);
  if(!lines.length) return setStatus('Enter at least one valid Product Code and QTY.', true);

  document.getElementById('saveBtn').disabled = true;
  setStatus('Saving estimate…');

  try{
    const payload = new URLSearchParams();
    payload.set('action','saveEstimate');
    payload.set('token',authToken);
    payload.set('estimateDate',estimateDate);
    payload.set('lines',JSON.stringify(lines));

    const r = await fetch(CFG.API_URL,{method:'POST',body:payload});
    const data = await r.json();
    if(!data.ok) throw new Error(data.error || 'Save failed');

    setStatus('Estimate saved successfully.');
    const savedDate = data.estimateDate || estimateDate;
    savedBox.textContent = `Saved successfully. Date: ${savedDate} | Reference: ${data.reference} | Grand Total: ${money(data.grandTotal)}`;
    savedBox.classList.remove('hidden');

    // Reload so any manually updated master costs appear immediately.
    await loadProducts();
  }catch(e){
    if(authError(e.message)) return showLogin('Session expired. Please sign in again.');
    setStatus(`Save failed: ${e.message}`, true);
  }finally{
    document.getElementById('saveBtn').disabled = false;
  }
}

loginForm.addEventListener('submit',e=>{
  e.preventDefault();
  login(loginPassword.value);
});

document.getElementById('addRowBtn').addEventListener('click',()=>addRow());
document.getElementById('clearBtn').addEventListener('click',clearAll);
document.getElementById('refreshBtn').addEventListener('click',loadProducts);
document.getElementById('logoutBtn').addEventListener('click',logout);
document.getElementById('saveBtn').addEventListener('click',saveEstimate);
document.getElementById('printBtn').addEventListener('click',()=>window.print());

estimateDateEl.value = localDateISO();

if(authToken){
  showApp();
  clearAll();
  loadProducts();
}else{
  showLogin();
}
