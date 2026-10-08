const CFG = window.APP_CONFIG || {};
let products = new Map();
let rowId = 0;
let historyMode = false;
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
const historyNotice = document.getElementById('historyNotice');
const historyNoticeText = document.getElementById('historyNoticeText');
const editActions = document.getElementById('editActions');
const saveBtn = document.getElementById('saveBtn');
const pageSubtitle = document.getElementById('pageSubtitle');
const historyModal = document.getElementById('historyModal');
const historyRows = document.getElementById('historyRows');
const historyStatus = document.getElementById('historyStatus');
const historyDate = document.getElementById('historyDate');
const historyReference = document.getElementById('historyReference');

function money(v){ return `${CFG.CURRENCY || 'SAR'} ${Number(v || 0).toFixed(2)}`; }
function localDateISO(){
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,'0');
  const day = String(d.getDate()).padStart(2,'0');
  return `${y}-${m}-${day}`;
}
function setStatus(msg, error=false){ statusEl.textContent = msg; statusEl.classList.toggle('error', error); }
function validApi(){ return CFG.API_URL && !CFG.API_URL.includes('PASTE_YOUR'); }
function authError(message){
  const text = String(message || '');
  return text.includes('Unauthorized') || text.includes('Session expired');
}
function escapeHtml(value){
  return String(value ?? '').replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));
}

function showLogin(message=''){
  authToken = '';
  sessionStorage.removeItem('purchaseEstimateToken');
  products = new Map();
  codesList.innerHTML = '';
  historyModal.classList.add('hidden');
  appShell.classList.add('hidden');
  loginScreen.classList.remove('hidden');
  loginMessage.textContent = message;
  loginPassword.value = '';
  setTimeout(()=>loginPassword.focus(),100);
}
function showApp(){ loginScreen.classList.add('hidden'); appShell.classList.remove('hidden'); }

async function login(password){
  if(!validApi()){ loginMessage.textContent='Backend is not configured.'; return; }
  loginBtn.disabled=true;
  loginMessage.textContent='Signing in…';
  try{
    const payload = new URLSearchParams();
    payload.set('action','login'); payload.set('password',password); payload.set('clientId',clientId);
    const r = await fetch(CFG.API_URL,{method:'POST',body:payload});
    const data = await r.json();
    if(!data.ok) throw new Error(data.error || 'Login failed');
    authToken = data.token;
    sessionStorage.setItem('purchaseEstimateToken',authToken);
    showApp();
    startNewEstimate();
    await loadProducts();
  }catch(e){ showLogin(e.message || 'Login failed'); }
  finally{ loginBtn.disabled=false; }
}

async function logout(){
  try{
    if(authToken && validApi()){
      const payload = new URLSearchParams();
      payload.set('action','logout'); payload.set('token',authToken);
      await fetch(CFG.API_URL,{method:'POST',body:payload});
    }
  }catch(_e){}
  showLogin('Signed out.');
}

async function apiGet(action, extra={}){
  if(!authToken) throw new Error('Unauthorized');
  const url = new URL(CFG.API_URL);
  url.searchParams.set('action',action);
  url.searchParams.set('token',authToken);
  url.searchParams.set('t',Date.now());
  Object.entries(extra).forEach(([k,v])=>{ if(v !== '' && v != null) url.searchParams.set(k,v); });
  const r = await fetch(url.toString(),{cache:'no-store'});
  const data = await r.json();
  if(!data.ok) throw new Error(data.error || 'Request failed');
  return data;
}

async function loadProducts(){
  if(!validApi()) return setStatus('Backend is not configured.',true);
  if(!authToken) return showLogin();
  setStatus('Loading product master…');
  try{
    const data = await apiGet('products');
    products = new Map(data.products.map(p=>[String(p.code).trim().toUpperCase(),p]));
    codesList.innerHTML='';
    data.products.forEach(p=>{
      const o=document.createElement('option'); o.value=p.code; o.label=p.name; codesList.appendChild(o);
    });
    setStatus(`${data.products.length} products loaded securely from Google Sheet. Cost can be edited manually if supplier rate is different.`);
    if(!historyMode) recalcAll();
  }catch(e){
    if(authError(e.message)) return showLogin('Session expired. Please sign in again.');
    setStatus(`Product load failed: ${e.message}`,true);
  }
}

function addRow(prefill={}, options={}){
  const historical = Boolean(options.historical);
  const id=++rowId;
  const tr=document.createElement('tr');
  tr.dataset.id=id;
  tr.dataset.lastCode='';
  if(historical) tr.classList.add('history-row');
  tr.innerHTML=`
    <td class="idx"></td>
    <td><input class="code" ${historical?'readonly':'list="productCodes" autocomplete="off"'} placeholder="Code" value="${escapeHtml(prefill.code||'')}"></td>
    <td><input class="name readonly" readonly value="${historical?escapeHtml(prefill.name||''):''}"></td>
    <td><input class="uom readonly" readonly value="${historical?escapeHtml(prefill.uom||''):''}"></td>
    <td><input class="qty" inputmode="decimal" type="number" min="0" step="0.001" ${historical?'readonly':''} placeholder="0" value="${historical?escapeHtml(prefill.qty??''):''}"></td>
    <td><input class="cost num manual-cost" inputmode="decimal" type="number" min="0" step="0.01" ${historical?'readonly':''} placeholder="0.00" value="${historical?Number(prefill.cost||0).toFixed(2):''}"></td>
    <td><input class="total readonly num" readonly value="${historical?Number(prefill.total||0).toFixed(2):''}"></td>
    <td class="remove-col"><button class="danger remove" title="Remove" ${historical?'disabled':''}>×</button></td>`;
  rowsEl.appendChild(tr);

  if(!historical){
    const code=tr.querySelector('.code');
    const qty=tr.querySelector('.qty');
    const cost=tr.querySelector('.cost');
    code.addEventListener('input',()=>hydrateRow(tr));
    code.addEventListener('change',()=>hydrateRow(tr));
    qty.addEventListener('input',()=>recalcRow(tr));
    cost.addEventListener('input',()=>{ tr.dataset.costOverridden='1'; recalcRow(tr); });
    tr.querySelector('.remove').addEventListener('click',()=>{
      tr.remove(); renumber(); calcGrand(); if(!rowsEl.children.length) addRow();
    });
    if(prefill.code) hydrateRow(tr);
  }
  renumber();
}

function hydrateRow(tr){
  const key=tr.querySelector('.code').value.trim().toUpperCase();
  const p=products.get(key);
  const codeChanged=tr.dataset.lastCode!==key;
  tr.querySelector('.code').value=key;
  tr.querySelector('.name').value=p?.name||'';
  tr.querySelector('.uom').value=p?.uom||'';
  if(codeChanged){
    tr.dataset.costOverridden='0';
    tr.querySelector('.cost').value=p?Number(p.cost||0).toFixed(2):'';
    tr.dataset.lastCode=key;
  }else if(tr.dataset.costOverridden!=='1' && p){
    tr.querySelector('.cost').value=Number(p.cost||0).toFixed(2);
  }
  recalcRow(tr);
}
function recalcRow(tr){
  const qty=Number(tr.querySelector('.qty').value||0);
  const cost=Number(tr.querySelector('.cost').value||0);
  tr.querySelector('.total').value=(qty*cost).toFixed(2);
  calcGrand();
}
function recalcAll(){ [...rowsEl.children].forEach(tr=>{ if(!tr.classList.contains('history-row')) hydrateRow(tr); }); }
function calcGrand(){
  const total=[...rowsEl.children].reduce((s,tr)=>s+Number(tr.querySelector('.total').value||0),0);
  grandTotalEl.textContent=money(total);
}
function renumber(){ [...rowsEl.children].forEach((tr,i)=>tr.querySelector('.idx').textContent=i+1); }

function startNewEstimate(){
  historyMode=false;
  rowsEl.innerHTML=''; rowId=0;
  historyNotice.classList.add('hidden');
  editActions.classList.remove('hidden');
  saveBtn.classList.remove('hidden'); saveBtn.disabled=false;
  estimateDateEl.disabled=false;
  estimateDateEl.value=localDateISO();
  pageSubtitle.textContent='Enter Product Code and Quantity only.';
  savedBox.classList.add('hidden');
  for(let i=0;i<8;i++) addRow();
  calcGrand();
  setStatus('Ready.');
}
function clearAll(){ startNewEstimate(); }

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

function preparePrint(){
  let printedIndex=0;
  [...rowsEl.children].forEach(tr=>{
    const valid=Boolean(tr.querySelector('.code').value.trim() && tr.querySelector('.name').value.trim() && Number(tr.querySelector('.qty').value||0)>0);
    tr.classList.toggle('print-hide',!valid);
    if(valid) tr.querySelector('.idx').textContent=++printedIndex;
  });
}
function cleanupPrint(){ [...rowsEl.children].forEach(tr=>tr.classList.remove('print-hide')); renumber(); }
function printEstimate(){
  if(!collectLines().length){ setStatus('No valid estimate available to print.',true); return; }
  window.print();
}

async function saveEstimate(){
  if(historyMode) return setStatus('Historical estimates are read-only. Click New Estimate to create a new one.',true);
  if(!validApi()) return setStatus('Backend is not configured.',true);
  if(!authToken) return showLogin();
  const lines=collectLines();
  const estimateDate=estimateDateEl.value;
  if(!estimateDate) return setStatus('Please select Estimate Date.',true);
  if(!lines.length) return setStatus('Enter at least one valid Product Code and QTY.',true);

  saveBtn.disabled=true; setStatus('Saving estimate…');
  try{
    const payload=new URLSearchParams();
    payload.set('action','saveEstimate'); payload.set('token',authToken); payload.set('estimateDate',estimateDate); payload.set('lines',JSON.stringify(lines));
    const r=await fetch(CFG.API_URL,{method:'POST',body:payload});
    const data=await r.json();
    if(!data.ok) throw new Error(data.error||'Save failed');
    setStatus('Estimate saved successfully.');
    const savedDate=data.estimateDate||estimateDate;
    savedBox.textContent=`Saved successfully. Date: ${savedDate} | Reference: ${data.reference} | Grand Total: ${money(data.grandTotal)}`;
    savedBox.classList.remove('hidden');
    await loadProducts();
  }catch(e){
    if(authError(e.message)) return showLogin('Session expired. Please sign in again.');
    setStatus(`Save failed: ${e.message}`,true);
  }finally{ saveBtn.disabled=false; }
}

function openHistory(){
  historyModal.classList.remove('hidden');
  loadOrderHistory();
}
function closeHistory(){ historyModal.classList.add('hidden'); }
async function loadOrderHistory(){
  historyStatus.textContent='Loading order history…'; historyStatus.classList.remove('error'); historyRows.innerHTML='';
  try{
    const data=await apiGet('orders',{date:historyDate.value,reference:historyReference.value.trim(),limit:'300'});
    if(!data.orders.length){
      historyStatus.textContent='No saved estimates found for this filter.';
      return;
    }
    historyStatus.textContent=`${data.orders.length} saved estimate(s) found.`;
    data.orders.forEach(order=>{
      const tr=document.createElement('tr');
      tr.innerHTML=`<td>${escapeHtml(order.estimateDate)}</td><td class="history-ref">${escapeHtml(order.reference)}</td><td>${escapeHtml(order.savedAt)}</td><td>${Number(order.lines||0)}</td><td>${escapeHtml(money(order.grandTotal))}</td><td><button class="primary history-open">Open</button></td>`;
      tr.querySelector('.history-open').addEventListener('click',()=>loadHistoricalOrder(order.reference));
      historyRows.appendChild(tr);
    });
  }catch(e){
    if(authError(e.message)) return showLogin('Session expired. Please sign in again.');
    historyStatus.textContent=`History load failed: ${e.message}`; historyStatus.classList.add('error');
  }
}

async function loadHistoricalOrder(reference){
  historyStatus.textContent='Opening estimate…'; historyStatus.classList.remove('error');
  try{
    const data=await apiGet('order',{reference});
    const order=data.order;
    historyMode=true;
    rowsEl.innerHTML=''; rowId=0;
    order.lines.forEach(line=>addRow(line,{historical:true}));
    estimateDateEl.value=order.estimateDate||'';
    estimateDateEl.disabled=true;
    grandTotalEl.textContent=money(order.grandTotal);
    historyNoticeText.textContent=`Reference: ${order.reference} | Date: ${order.estimateDate} | Saved: ${order.savedAt}`;
    historyNotice.classList.remove('hidden');
    editActions.classList.add('hidden');
    saveBtn.classList.add('hidden');
    pageSubtitle.textContent='Historical saved estimate — read only.';
    savedBox.textContent=`Historical record. Reference: ${order.reference} | Grand Total: ${money(order.grandTotal)}`;
    savedBox.classList.remove('hidden');
    setStatus('Historical estimate loaded. You can print or save it as PDF.');
    closeHistory();
    window.scrollTo({top:0,behavior:'smooth'});
  }catch(e){
    if(authError(e.message)) return showLogin('Session expired. Please sign in again.');
    historyStatus.textContent=`Unable to open estimate: ${e.message}`; historyStatus.classList.add('error');
  }
}

loginForm.addEventListener('submit',e=>{ e.preventDefault(); login(loginPassword.value); });
document.getElementById('addRowBtn').addEventListener('click',()=>addRow());
document.getElementById('clearBtn').addEventListener('click',startNewEstimate);
document.getElementById('refreshBtn').addEventListener('click',loadProducts);
document.getElementById('logoutBtn').addEventListener('click',logout);
document.getElementById('saveBtn').addEventListener('click',saveEstimate);
document.getElementById('printBtn').addEventListener('click',printEstimate);
document.getElementById('historyBtn').addEventListener('click',openHistory);
document.getElementById('closeHistoryBtn').addEventListener('click',closeHistory);
document.getElementById('newEstimateBtn').addEventListener('click',startNewEstimate);
document.getElementById('historySearchBtn').addEventListener('click',loadOrderHistory);
document.getElementById('historyResetBtn').addEventListener('click',()=>{ historyDate.value=''; historyReference.value=''; loadOrderHistory(); });
historyModal.addEventListener('click',e=>{ if(e.target===historyModal) closeHistory(); });
document.addEventListener('keydown',e=>{ if(e.key==='Escape' && !historyModal.classList.contains('hidden')) closeHistory(); });
window.addEventListener('beforeprint',preparePrint);
window.addEventListener('afterprint',cleanupPrint);

estimateDateEl.value=localDateISO();
if(authToken){ showApp(); startNewEstimate(); loadProducts(); } else { showLogin(); }
