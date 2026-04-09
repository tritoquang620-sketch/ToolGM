const screens = [...document.querySelectorAll('.screen')];
const backBtn = document.getElementById('backBtn');
const homeBtn = document.getElementById('homeBtn');
const singleDialog = document.getElementById('singleDialog');
const pairDialog = document.getElementById('pairDialog');
const singleForm = document.getElementById('singleForm');
const pairForm = document.getElementById('pairForm');
let screenStack = ['screen-home'];
let packingCache = {single:[], pair:[]};

function showScreen(id){
  screens.forEach(s=>s.classList.toggle('active', s.id===id));
}
function pushScreen(id){
  if (screenStack.at(-1)!==id) screenStack.push(id);
  showScreen(id);
}
function goBack(){
  if (screenStack.length>1) screenStack.pop();
  showScreen(screenStack.at(-1));
}

document.querySelectorAll('[data-open]').forEach(btn=>{
  btn.onclick = ()=>pushScreen(btn.dataset.open);
});
backBtn.onclick = goBack;
homeBtn.onclick = ()=>{screenStack=['screen-home']; showScreen('screen-home');};

document.querySelectorAll('.subtab').forEach(btn=>{
  btn.onclick = ()=>{
    document.querySelectorAll('.subtab').forEach(x=>x.classList.remove('active'));
    document.querySelectorAll('.tab-pane').forEach(x=>x.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab+'Tab').classList.add('active');
  };
});

function qs(form){return new FormData(form);}
function esc(v){return String(v ?? '').replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));}

async function api(url, options={}){
  const res = await fetch(url, options);
  if (!res.ok) {
    let msg = 'Lỗi xử lý';
    try{ msg = await res.text(); }catch{}
    throw new Error(msg);
  }
  return res.json();
}

function renderSingleTable(){
  const wrap = document.getElementById('singleTableWrap');
  const rows = packingCache.single || [];
  wrap.innerHTML = `<div class="table-wrap"><table>
    <thead><tr><th>STT</th><th>Item</th><th>Rev</th><th>Packing</th><th>Note</th><th>Action</th></tr></thead>
    <tbody>${rows.map((r,i)=>`<tr>
      <td>${i+1}</td><td class="left">${esc(r.item)}</td><td>${esc(r.rev)}</td><td>${esc(r.packing)}</td><td class="left">${esc(r.note||'')}</td>
      <td><div class="action-group"><button class="small-btn" onclick='editSingle(${JSON.stringify(r).replace(/'/g,"&#39;")})'>Sửa</button><button class="small-btn" onclick='removeSingle(${r.id})'>Xóa</button></div></td>
    </tr>`).join('')}</tbody></table></div>`;
}

function renderPairTable(){
  const wrap = document.getElementById('pairTableWrap');
  const rows = packingCache.pair || [];
  wrap.innerHTML = `<div class="table-wrap"><table>
    <thead><tr><th>STT</th><th>Item 1</th><th>Rev 1</th><th>Item 2</th><th>Rev 2</th><th>Packing</th><th>Note</th><th>Action</th></tr></thead>
    <tbody>${rows.map((r,i)=>`<tr>
      <td>${i+1}</td><td class="left">${esc(r.item1)}</td><td>${esc(r.rev1)}</td><td class="left">${esc(r.item2)}</td><td>${esc(r.rev2)}</td><td>${esc(r.packing)}</td><td class="left">${esc(r.note||'')}</td>
      <td><div class="action-group"><button class="small-btn" onclick='editPair(${JSON.stringify(r).replace(/'/g,"&#39;")})'>Sửa</button><button class="small-btn" onclick='removePair(${r.id})'>Xóa</button></div></td>
    </tr>`).join('')}</tbody></table></div>`;
}

async function loadPacking(){
  packingCache = await api('/api/packing');
  renderSingleTable();
  renderPairTable();
}

window.editSingle = (row) => {
  singleForm.id.value = row.id;
  singleForm.item.value = row.item;
  singleForm.rev.value = row.rev;
  singleForm.packing.value = row.packing;
  singleForm.note.value = row.note || '';
  singleDialog.showModal();
};
window.editPair = (row) => {
  pairForm.id.value = row.id;
  pairForm.item1.value = row.item1;
  pairForm.rev1.value = row.rev1;
  pairForm.item2.value = row.item2;
  pairForm.rev2.value = row.rev2;
  pairForm.packing.value = row.packing;
  pairForm.note.value = row.note || '';
  pairDialog.showModal();
};
window.removeSingle = async (id) => {
  if (!confirm('Xóa dòng mã đơn này?')) return;
  await fetch(`/api/packing/single/${id}`, {method:'DELETE'});
  await loadPacking();
};
window.removePair = async (id) => {
  if (!confirm('Xóa dòng mã đôi này?')) return;
  await fetch(`/api/packing/pair/${id}`, {method:'DELETE'});
  await loadPacking();
};

document.getElementById('addSingleBtn').onclick = ()=>{singleForm.reset(); singleForm.id.value=''; singleDialog.showModal();};
document.getElementById('addPairBtn').onclick = ()=>{pairForm.reset(); pairForm.id.value=''; pairDialog.showModal();};

singleForm.addEventListener('submit', async (e)=>{
  e.preventDefault();
  const id = singleForm.id.value;
  await fetch(id ? `/api/packing/single/${id}` : '/api/packing/single', {method: id ? 'PUT' : 'POST', body: qs(singleForm)});
  singleDialog.close();
  await loadPacking();
});
pairForm.addEventListener('submit', async (e)=>{
  e.preventDefault();
  const id = pairForm.id.value;
  await fetch(id ? `/api/packing/pair/${id}` : '/api/packing/pair', {method: id ? 'PUT' : 'POST', body: qs(pairForm)});
  pairDialog.close();
  await loadPacking();
});

document.getElementById('importSingleInput').addEventListener('change', async (e)=>{
  if (!e.target.files[0]) return;
  const fd = new FormData(); fd.append('file', e.target.files[0]);
  await fetch('/api/packing/import-single', {method:'POST', body:fd});
  e.target.value='';
  await loadPacking();
});
document.getElementById('importPairInput').addEventListener('change', async (e)=>{
  if (!e.target.files[0]) return;
  const fd = new FormData(); fd.append('file', e.target.files[0]);
  await fetch('/api/packing/import-pair', {method:'POST', body:fd});
  e.target.value='';
  await loadPacking();
});

function renderImgResult(data){
  const box = document.getElementById('imgResult');
  box.innerHTML = '';
  ['CPT','OP','GP'].forEach(group=>{
    const count = data.counts?.[group] || 0;
    const href = data.downloads?.[group];
    if (!count) return;
    const el = document.createElement('div');
    el.className = 'result-card';
    el.innerHTML = `<div><strong>${group}</strong>: ${count} đơn</div>${href ? `<div><a href="${href}">Tải file IMG</a></div>`:''}`;
    box.appendChild(el);
  });
}
function renderExcelResult(data){
  const box = document.getElementById('excelResult');
  box.innerHTML = `<div class="result-card"><div>Đã xử lý ${data.records.length} file</div><div><a href="${data.download}">Tải file Excel</a></div></div>`;
}

document.getElementById('runImgBtn').onclick = async ()=>{
  const files = document.getElementById('imgFiles').files;
  if (!files.length) return alert('Chọn PDF trước');
  const fd = new FormData(); [...files].forEach(f=>fd.append('files',f));
  document.getElementById('imgResult').innerHTML = '<div class="result-card">Đang xử lý...</div>';
  try{ renderImgResult(await api('/api/process/img', {method:'POST', body:fd})); }
  catch(err){ document.getElementById('imgResult').innerHTML = `<div class="result-card">${esc(err.message)}</div>`; }
};

document.getElementById('runExcelBtn').onclick = async ()=>{
  const files = document.getElementById('excelFiles').files;
  if (!files.length) return alert('Chọn PDF trước');
  const fd = new FormData(); [...files].forEach(f=>fd.append('files',f));
  document.getElementById('excelResult').innerHTML = '<div class="result-card">Đang xử lý...</div>';
  try{ renderExcelResult(await api('/api/process/excel', {method:'POST', body:fd})); }
  catch(err){ document.getElementById('excelResult').innerHTML = `<div class="result-card">${esc(err.message)}</div>`; }
};

if ('serviceWorker' in navigator) {
  window.addEventListener('load', ()=>navigator.serviceWorker.register('/service-worker.js').catch(()=>{}));
}
loadPacking();
