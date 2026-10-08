const CFG = window.APP_CONFIG || {};
let products = new Map();
let rowId = 0;

const rowsEl = document.getElementById("rows");
const statusEl = document.getElementById("status");
const grandTotalEl = document.getElementById("grandTotal");
const codesList = document.getElementById("productCodes");
const savedBox = document.getElementById("savedBox");

function money(v){
  return `${CFG.CURRENCY || "SAR"} ${Number(v || 0).toFixed(2)}`;
}
function setStatus(msg, error=false){
  statusEl.textContent = msg;
  statusEl.classList.toggle("error", error);
}
function validApi(){
  return CFG.API_URL && !CFG.API_URL.includes("PASTE_YOUR");
}
async function loadProducts(){
  if(!validApi()){
    setStatus("Setup required: paste Apps Script Web App URL in config.js", true);
    return;
  }
  setStatus("Loading product master…");
  try{
    const r = await fetch(`${CFG.API_URL}?action=products&t=${Date.now()}`);
    const data = await r.json();
    if(!data.ok) throw new Error(data.error || "Unable to load products");
    products = new Map(data.products.map(p => [String(p.code).trim().toUpperCase(), p]));
    codesList.innerHTML = "";
    data.products.forEach(p=>{
      const o=document.createElement("option");
      o.value=p.code;
      o.label=p.name;
      codesList.appendChild(o);
    });
    setStatus(`${data.products.length} products loaded from Google Sheet.`);
    recalcAll();
  }catch(e){
    setStatus(`Product load failed: ${e.message}`, true);
  }
}
function addRow(prefill={}){
  const id=++rowId;
  const tr=document.createElement("tr");
  tr.dataset.id=id;
  tr.innerHTML=`
    <td class="idx"></td>
    <td><input class="code" list="productCodes" autocomplete="off" placeholder="Code" value="${prefill.code||""}"></td>
    <td><input class="name readonly" readonly></td>
    <td><input class="uom readonly" readonly></td>
    <td><input class="qty" inputmode="decimal" type="number" min="0" step="0.001" placeholder="0"></td>
    <td><input class="cost readonly num" readonly></td>
    <td><input class="total readonly num" readonly></td>
    <td class="remove-col"><button class="danger remove" title="Remove">×</button></td>`;
  rowsEl.appendChild(tr);

  const code=tr.querySelector(".code"), qty=tr.querySelector(".qty");
  code.addEventListener("input",()=>hydrateRow(tr));
  code.addEventListener("change",()=>hydrateRow(tr));
  qty.addEventListener("input",()=>recalcRow(tr));
  tr.querySelector(".remove").addEventListener("click",()=>{
    tr.remove(); renumber(); calcGrand();
    if(!rowsEl.children.length) addRow();
  });
  renumber();
  if(prefill.code) hydrateRow(tr);
}
function hydrateRow(tr){
  const key=tr.querySelector(".code").value.trim().toUpperCase();
  const p=products.get(key);
  tr.querySelector(".code").value=key;
  tr.querySelector(".name").value=p?.name || "";
  tr.querySelector(".uom").value=p?.uom || "";
  tr.querySelector(".cost").value=p ? Number(p.cost||0).toFixed(2) : "";
  recalcRow(tr);
}
function recalcRow(tr){
  const qty=Number(tr.querySelector(".qty").value || 0);
  const cost=Number(tr.querySelector(".cost").value || 0);
  tr.querySelector(".total").value=(qty*cost).toFixed(2);
  calcGrand();
}
function recalcAll(){
  [...rowsEl.children].forEach(hydrateRow);
}
function calcGrand(){
  const total=[...rowsEl.children].reduce((s,tr)=>s+Number(tr.querySelector(".total").value||0),0);
  grandTotalEl.textContent=money(total);
}
function renumber(){
  [...rowsEl.children].forEach((tr,i)=>tr.querySelector(".idx").textContent=i+1);
}
function clearAll(){
  rowsEl.innerHTML=""; rowId=0; savedBox.classList.add("hidden");
  for(let i=0;i<8;i++) addRow();
  calcGrand();
}
function collectLines(){
  return [...rowsEl.children].map(tr=>({
    code:tr.querySelector(".code").value.trim(),
    name:tr.querySelector(".name").value,
    uom:tr.querySelector(".uom").value,
    qty:Number(tr.querySelector(".qty").value||0),
    cost:Number(tr.querySelector(".cost").value||0),
    total:Number(tr.querySelector(".total").value||0)
  })).filter(x=>x.code && x.qty>0 && x.name);
}
async function saveEstimate(){
  if(!validApi()) return setStatus("Please configure Apps Script URL first.", true);
  const lines=collectLines();
  if(!lines.length) return setStatus("Enter at least one valid Product Code and QTY.", true);

  document.getElementById("saveBtn").disabled=true;
  setStatus("Saving estimate…");
  try{
    const payload = new URLSearchParams();
    payload.set("action","saveEstimate");
    payload.set("lines",JSON.stringify(lines));
    const r=await fetch(CFG.API_URL,{method:"POST",body:payload});
    const data=await r.json();
    if(!data.ok) throw new Error(data.error || "Save failed");
    setStatus("Estimate saved successfully.");
    savedBox.textContent=`Saved successfully. Reference: ${data.reference} | Grand Total: ${money(data.grandTotal)}`;
    savedBox.classList.remove("hidden");
  }catch(e){
    setStatus(`Save failed: ${e.message}`, true);
  }finally{
    document.getElementById("saveBtn").disabled=false;
  }
}

document.getElementById("addRowBtn").addEventListener("click",()=>addRow());
document.getElementById("clearBtn").addEventListener("click",clearAll);
document.getElementById("refreshBtn").addEventListener("click",loadProducts);
document.getElementById("saveBtn").addEventListener("click",saveEstimate);
document.getElementById("printBtn").addEventListener("click",()=>window.print());

clearAll();
loadProducts();
