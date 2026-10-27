import {recoll,opsFromSession,opsFromBuildingRows,CONTRACT_COL,OPE_COL,STATUS_COL} from './recoll.js';

const $=id=>document.getElementById(id);
const state={sf:null,sfName:'',et:[],result:null};
const MAX_SHOWN=300;

const readBook=file=>file.arrayBuffer().then(b=>XLSX.read(b,{type:'array',cellNF:true}));
const cellValue=c=>!c?null:c.t==='n'&&c.z&&XLSX.SSF.is_date(c.z)?XLSX.SSF.format('yyyy-mm-dd',c.v):c.t==='e'?null:c.v;
function sheetAoa(ws){
  if(!ws?.['!ref']) return [];
  const r=XLSX.utils.decode_range(ws['!ref']), out=[];
  for(let i=r.s.r;i<=r.e.r;i++){
    const row=[]; for(let j=r.s.c;j<=r.e.c;j++) row.push(cellValue(ws[XLSX.utils.encode_cell({r:i,c:j})]));
    if(row.some(v=>v!=null&&v!=='')) out.push(row);
  }
  return out;
}

async function loadSf(file){
  const wb=await readBook(file);
  const names=['OPERATIONS',...wb.SheetNames.filter(n=>n!=='OPERATIONS')].filter(n=>wb.Sheets[n]);
  for(const n of names){
    const aoa=sheetAoa(wb.Sheets[n]);
    const h=aoa.findIndex((r,i)=>i<5&&r.includes(CONTRACT_COL)&&r.includes(OPE_COL));
    if(h>=0){ state.sf={header:aoa[h].map(v=>v??''),rows:aoa.slice(h+1)}; state.sfName=`${file.name} — feuille ${n}`; return; }
  }
  throw new Error(`Aucune feuille de « ${file.name} » ne contient les colonnes « ${CONTRACT_COL} » et « ${OPE_COL} ».`);
}
async function loadEt(file){
  if(/\.json$/i.test(file.name)) return {name:file.name,ops:opsFromSession(JSON.parse(await file.text()),file.name)};
  const wb=await readBook(file), ws=wb.Sheets['Données par bâtiment'];
  if(!ws) throw new Error(`« ${file.name} » n’a pas de feuille « Données par bâtiment » (export ExtracTerre attendu).`);
  const aoa=sheetAoa(ws), h=aoa[0];
  return {name:file.name,ops:opsFromBuildingRows(aoa.slice(1).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]]))),file.name)};
}

function toast(msg,kind=''){ const t=document.createElement('div'); t.className='toast '+kind; t.textContent=msg; $('toasts').append(t); setTimeout(()=>t.remove(),6000); }
const esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

function run(){
  renderSources();
  if(!state.sf||!state.et.length){ state.result=null; render(); return; }
  try{ state.result=recoll(state.sf,state.et.flatMap(e=>e.ops),{keepSalesforce:$('keepSf').checked}); }
  catch(e){ state.result=null; toast(e.message,'error'); }
  render();
}

function renderSources(){
  $('sfInfo').textContent=state.sf?`${state.sfName} · ${state.sf.rows.length} lignes`:'Aucun fichier';
  $('etInfo').innerHTML=state.et.length?state.et.map((e,i)=>`<span class="chip">${esc(e.name)} · ${e.ops.length} opérations <button data-rm="${i}" title="Retirer">×</button></span>`).join(''):'Aucun fichier';
}

const FILTERS={
  all:()=>true,
  ok:s=>s.startsWith('Recollé'),
  check:s=>/Conflit|ExtracTerre seul|à vérifier|remplace/.test(s),
  sf:s=>s==='Salesforce seul'
};
function render(){
  const r=state.result;
  $('results').hidden=!r;
  if(!r) return;
  const s=r.stats;
  $('kpis').innerHTML=[
    [s.lignes,'lignes (1 par bâtiment)'],[s.recolles,'opérations recollées'],[s.salesforceSeul,'opérations Salesforce seules'],
    [s.etSeul,'opérations ExtracTerre seules',s.etSeul],[s.conflits,'conflits contrat / OPE',s.conflits],
    [s.aVerifier,'bâtiments à vérifier',s.aVerifier],[s.ecarts,'écarts de valeur',s.ecarts]
  ].map(([n,l,alert])=>`<div class="kpi${alert?' alert':''}"><b>${n}</b><span>${l}</span></div>`).join('');
  const si=r.header.indexOf(STATUS_COL), q=$('search').value.trim().toLowerCase(), f=FILTERS[$('filter').value];
  const rows=r.rows.filter(row=>f(row[si]||'')&&(!q||row.some(v=>v!=null&&String(v).toLowerCase().includes(q))));
  const shown=rows.slice(0,MAX_SHOWN);
  const order=[si,...r.header.map((_,i)=>i).filter(i=>i!==si)];
  $('table').innerHTML=`<thead><tr>${order.map(i=>`<th>${esc(r.header[i])}</th>`).join('')}</tr></thead><tbody>${shown.map(row=>`<tr class="${/Conflit|ExtracTerre seul/.test(row[si])?'bad':/à vérifier|remplace/.test(row[si])?'warn':''}">${order.map(i=>`<td>${esc(row[i])}</td>`).join('')}</tr>`).join('')}</tbody>`;
  $('shown').textContent=`${rows.length} ligne(s)${rows.length>MAX_SHOWN?` — ${MAX_SHOWN} premières affichées ; l’export contient tout`:''}`;
  $('ecartsCount').textContent=r.ecarts.length;
  $('ecarts').innerHTML=r.ecarts.length?`<thead><tr><th>Contrat</th><th>OPE</th><th>Bâtiment</th><th>Colonne</th><th>Tous contrats</th><th>ExtracTerre</th><th>Retenu</th><th>Source</th></tr></thead><tbody>${r.ecarts.map(e=>`<tr><td>${esc(e.contrat)}</td><td>${esc(e.ope)}</td><td>${esc(e.batiment)}</td><td>${esc(e.colonne)}</td><td>${esc(e.salesforce)}</td><td>${esc(e.extracterre)}</td><td>${esc(e.retenu)}</td><td>${esc(e.source)}</td></tr>`).join('')}</tbody>`:'';
}

const ecartsAoa=r=>[['Contrat','OPE','Bâtiment','Colonne','Tous contrats','ExtracTerre','Retenu','Source'],...r.ecarts.map(e=>[e.contrat,e.ope,e.batiment,e.colonne,e.salesforce,e.extracterre,e.retenu,e.source])];
function exportXlsx(){
  const r=state.result, wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([r.header,...r.rows]),'OPERATIONS');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(ecartsAoa(r)),'Écarts');
  XLSX.writeFile(wb,`TotalRecoll_operations_${new Date().toISOString().slice(0,10)}.xlsx`);
}
async function copyForSheets(){
  const r=state.result, ci=r.header.indexOf(CONTRACT_COL);
  const cell=(v,j)=>{ if(v==null) return ''; const s=String(v).replace(/[\t\r\n]+/g,' '); return j===ci&&/^0\d+$/.test(s)?`="${s}"`:s; }; // ="0000…" garde les zéros au collage dans Google Sheets
  const tsv=[r.header,...r.rows].map(row=>r.header.map((_,j)=>cell(row[j],j)).join('\t')).join('\n');
  await navigator.clipboard.writeText(tsv);
  toast(`${r.rows.length} lignes copiées : dans Google Sheets, clique la cellule A1 puis colle (Cmd+V).`,'success');
}

$('sfFile').addEventListener('change',async e=>{ const f=e.target.files[0]; if(!f) return; try{ await loadSf(f); run(); }catch(err){ toast(err.message,'error'); } e.target.value=''; });
$('etFile').addEventListener('change',async e=>{
  for(const f of e.target.files){ try{ state.et.push(await loadEt(f)); }catch(err){ toast(err.message,'error'); } }
  e.target.value=''; run();
});
$('etInfo').addEventListener('click',e=>{ const i=e.target.dataset.rm; if(i!=null){ state.et.splice(+i,1); run(); } });
$('keepSf').addEventListener('change',run);
$('filter').addEventListener('change',render);
$('search').addEventListener('input',render);
$('xlsxBtn').addEventListener('click',exportXlsx);
$('copyBtn').addEventListener('click',()=>copyForSheets().catch(e=>toast('Copie impossible : '+e.message,'error')));
