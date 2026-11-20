// Page autonome « Quartiers NPNRU » : même calcul que l'onglet ANRU de TotalRecoll (recoll.js partagé), présentée pour un client.
// Tous contrats + session(s) ExtracTerre déposés dans la page, lus localement ; rien n'est envoyé ni enregistré.
import {recoll,opsFromSession,opsFromBuildingRows,relKey,CONTRACT_COL,OPE_COL,BUILDING_COL,empty,anruIndex,ANRU_COL,ANRU_QP_COL} from '../recoll.js?v=26';

const $=id=>document.getElementById(id);
const esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const nf=n=>n.toLocaleString('fr-FR');
const plural=(n,one,many)=>`${nf(n)} ${n>1?many:one}`;
const state={sf:null,sfName:'',et:[],result:null,anru:null,mentionsOff:new Set(),map:null,layer:null};
// Mêmes choix d'écarts que dans TotalRecoll (même site) : les coordonnées retenues sont donc identiques.
try{ state.decisions=JSON.parse(localStorage.getItem('totalrecoll.decisions')||'{}'); }catch{ state.decisions={}; }

/* ---------- Lecture des fichiers (reprise de TotalRecoll) ---------- */
const readBook=file=>file.arrayBuffer().then(b=>XLSX.read(b,{type:'array',cellNF:true}));
const cellValue=c=>!c?null:c.t==='n'&&c.z&&XLSX.SSF.is_date(c.z)?XLSX.SSF.format('yyyy-mm-dd',c.v):c.t==='e'?null:c.v;
function sheetAoa(ws){
  if(!ws?.['!ref']) return [];
  const r=XLSX.utils.decode_range(ws['!ref']), out=[];
  for(let i=r.s.r;i<=r.e.r;i++){ const row=[]; for(let j=r.s.c;j<=r.e.c;j++) row.push(cellValue(ws[XLSX.utils.encode_cell({r:i,c:j})])); if(row.some(v=>v!=null&&v!=='')) out.push(row); }
  return out;
}
async function readSessionHead(file){
  const MARK=/"projects":\s*\[\{?\s*$/, reader=file.stream().pipeThrough(new TextDecoderStream()).getReader(); const parts=[]; let tail='';
  while(true){ const {value,done}=await reader.read(); if(done) break;
    const chunk=tail+value, cut=chunk.lastIndexOf('\n'); const lines=cut>=0?chunk.slice(0,cut):''; tail=cut>=0?chunk.slice(cut+1):chunk;
    if(lines){ const ls=lines.split('\n'); const k=ls.findIndex(l=>MARK.test(l));
      if(k>=0){ parts.push([...ls.slice(0,k),ls[k].replace(/\{\s*$/,'')].join('\n')); try{ reader.cancel(); }catch{} return JSON.parse(parts.join('\n').replace(/^﻿?[^{]*/,'')+']}}'); }
      parts.push(lines); } }
  if(tail) parts.push(tail);
  try{ return JSON.parse(parts.join('\n').replace(/^﻿?[^{]*/,'')); }catch(e){ throw new Error(`Session ExtracTerre illisible (${e.message}).`); }
}
async function readSf(file){
  const wb=await readBook(file);
  for(const n of ['OPERATIONS',...wb.SheetNames.filter(n=>n!=='OPERATIONS')].filter(n=>wb.Sheets[n])){
    const aoa=sheetAoa(wb.Sheets[n]); const h=aoa.findIndex((r,i)=>i<30&&r.includes(CONTRACT_COL)&&r.includes(OPE_COL));
    if(h>=0){ const head=aoa[h].map(v=>String(v??'').trim()), keep=head.map((v,j)=>v||!h?j:-1).filter(j=>j>=0), ci=head.indexOf(CONTRACT_COL);
      const footer=r=>{ const c=String(r[ci]??'').trim(); if(/^\d+(?:\.0+)?$/.test(c)) return false; return r.some(v=>/^(?:total|somme|nombre|copyright\b.*|informations confidentielles\b.*)$/i.test(String(v??'').trim())); };
      const rows=aoa.slice(h+1).filter(r=>!footer(r)).map(r=>keep.map(j=>r[j]??null)).filter(r=>r.some(v=>v!=null&&v!==''));
      return {header:keep.map(j=>head[j]),rows}; }
  }
  return null;
}
async function readEt(file){
  if(/\.json$/i.test(file.name)){ const head=await readSessionHead(file); if(head.kind==='observatoire') throw new Error(`« ${file.name} » ne contient que des recherches, pas de bâtiments.`); return {name:file.name,ops:opsFromSession(head,file.name)}; }
  const wb=await readBook(file), ws=wb.Sheets['Données par bâtiment'];
  if(!ws) return null;
  const aoa=sheetAoa(ws), h=aoa[0], wr=wb.Sheets['Fiabilité par bâtiment'], ra=wr?sheetAoa(wr):null, rh=ra?.[0]||[];
  const relRows=ra?ra.slice(1).map(r=>Object.fromEntries(rh.map((k,i)=>[k,r[i]]).filter(([,v])=>relKey(v)))):[];
  return {name:file.name,ops:opsFromBuildingRows(aoa.slice(1).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]]))),file.name,relRows)};
}

/* ---------- Interface ---------- */
function toast(text,kind=''){ const t=document.createElement('div'); t.className=`toast ${kind}`; t.textContent=text; $('toasts').append(t); setTimeout(()=>t.remove(),kind==='error'?7000:3500); }
const busy=(on,text='')=>{ $('busy').hidden=!on; $('busyText').textContent=text; };
const frame=()=>new Promise(r=>requestAnimationFrame(()=>setTimeout(r,0)));

// Un fichier déposé n'importe où : .json = ExtracTerre ; .xlsx = Tous contrats s'il en a les colonnes, sinon export ExtracTerre.
async function addFiles(files,kind=''){
  for(const f of files){
    busy(true,`Lecture de « ${f.name} »…`); await frame();
    try{
      if(kind!=='et'&&!/\.json$/i.test(f.name)){ const sf=await readSf(f); if(sf){ state.sf=sf; state.sfName=f.name; continue; } if(kind==='sf') throw new Error(`« ${f.name} » n’a pas les colonnes « ${CONTRACT_COL} » et « ${OPE_COL} ».`); }
      const et=await readEt(f); if(!et) throw new Error(`« ${f.name} » n’est ni un export Tous contrats ni un export ExtracTerre.`);
      state.et=[...state.et.filter(x=>x.name!==et.name),et];
    }catch(e){ toast(e.message,'error'); }
  }
  busy(false); run();
}
function renderLoad(){
  $('sfState').textContent=state.sf?`${state.sfName} : ${plural(state.sf.rows.length,'ligne','lignes')}`:'Choisir ou déposer le fichier';
  $('etState').textContent=state.et.length?state.et.map(e=>`${e.name} : ${plural(e.ops.length,'opération','opérations')}`).join(', '):'Choisir ou déposer les fichiers';
  document.querySelector('.drop[data-kind=sf]').classList.toggle('ok',!!state.sf); document.querySelector('.drop[data-kind=et]').classList.toggle('ok',state.et.length>0);
  $('loadMsg').textContent=!state.anru?'Chargement des quartiers NPNRU…':state.sf&&!state.et.length?'Ajoutez maintenant la session ExtracTerre.':!state.sf&&state.et.length?'Ajoutez maintenant l’export Tous contrats.':'';
}
function run(){
  renderLoad();
  if(!(state.sf&&state.et.length&&state.anru)) return;
  try{ state.result=recoll(state.sf,state.et.flatMap(e=>e.ops),{keepSalesforce:false,decisions:state.decisions,anru:state.anru}); }
  catch(e){ toast(e.message,'error'); return; }
  $('load').hidden=true; $('summary').hidden=false; $('reload').hidden=false;
  render();
}

/* ---------- ANRU (même règle que l'onglet de TotalRecoll) ---------- */
const MENTIONS_COL='Opération: Mentions', NO_MENTION='(mention non renseignée)';
const mentionsOf=v=>{ const l=String(v??'').split(/<br\s*\/?>|&lt;br&gt;|\n/i).map(x=>x.trim()).filter(Boolean); return l.length?l:[NO_MENTION]; };
const col=n=>state.result.header.indexOf(n);
function anruRows(){
  const r=state.result; if(!r) return []; const a=col(ANRU_COL), la=col('Latitude'), lo=col('Longitude'), mc=col(MENTIONS_COL);
  if(a<0) return [];
  return r.rows.map((row,i)=>({i,lat:Number(row[la]),lon:Number(row[lo]),mentions:mentionsOf(mc<0?'':row[mc])})).filter(p=>r.rows[p.i][a]==='Oui');
}
const anruShown=(all=anruRows())=>all.filter(p=>p.mentions.some(m=>!state.mentionsOff.has(m)));
function rowLabel(i){
  const r=state.result, row=r.rows[i], h=n=>row[r.header.indexOf(n)];
  return {title:h("Nom de l'opération")||h(OPE_COL)||h(CONTRACT_COL)||'Opération sans identifiant',sub:[h(CONTRACT_COL),h(OPE_COL),h(BUILDING_COL),h('Nom de la société: Nom de la société')].filter(v=>!empty(v)).join(', ')};
}
// « Nom (QP…, intérêt national) » éventuellement suivi de « à 10 m au plus du périmètre » : nom du quartier et catégorie.
const quartierOf=s=>{ const t=String(s||''), m=t.match(/^(.*?)\s*\(([^,()]*),\s*intérêt (national|régional)\)/); return m?{key:`${m[1]}|${m[2]}`,nom:m[1],qp:m[2],cat:m[3]}:{key:t,nom:t||'Quartier non précisé',qp:'',cat:/national/.test(t)?'national':'régional'}; };

function render(){
  const r=state.result, all=anruRows(), qi=col(ANRU_QP_COL), ai=col(ANRU_COL);
  const counts=new Map(); all.forEach(p=>p.mentions.forEach(m=>counts.set(m,(counts.get(m)||0)+1)));
  const list=[...counts].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0],'fr'));
  $('mentions').innerHTML=list.length?list.map(([m,n])=>`<label class="mention"><input type="checkbox" data-mention="${esc(m)}" ${state.mentionsOff.has(m)?'':'checked'}><span>${esc(m)}</span><b>${nf(n)}</b></label>`).join(''):'<p class="empty">Aucun bâtiment en quartier NPNRU dans ces données.</p>';
  const pts=anruShown(all);
  const ops=new Set(pts.map(p=>r.meta[p.i].group||p.i)).size;
  const qs=new Map(); for(const p of pts){ const q=quartierOf(r.rows[p.i][qi]); const g=qs.get(q.key)||{...q,pts:[],ops:new Set()}; g.pts.push(p); g.ops.add(r.meta[p.i].group||p.i); qs.set(q.key,g); }
  const nat=[...qs.values()].filter(q=>q.cat==='national').length, reg=qs.size-nat;
  const unsure=r.rows.filter(row=>row[ai]===''&&row[qi]).length, edge=all.filter(p=>/du périmètre$/.test(r.rows[p.i][qi]||'')).length;
  $('headline').innerHTML=pts.length?`<span class="fig">${nf(pts.length)}</span> ${pts.length>1?'bâtiments certifiés':'bâtiment certifié'} dans <span class="fig">${nf(qs.size)}</span> ${qs.size>1?'quartiers':'quartier'} du NPNRU`:'Aucun bâtiment certifié en quartier NPNRU pour les mentions choisies';
  $('figures').innerHTML=`<div><dt>Opérations</dt><dd>${nf(ops)}</dd></div><div><dt>Quartiers d’intérêt national</dt><dd>${nf(nat)}</dd></div><div><dt>Quartiers d’intérêt régional</dt><dd>${nf(reg)}</dd></div>`;
  $('note').textContent=[pts.length<all.length?`${plural(all.length,'bâtiment','bâtiments')} en quartier NPNRU au total, ${nf(pts.length)} avec les mentions choisies.`:'',edge?`${plural(edge,'bâtiment est compté','bâtiments sont comptés')} à 10 m au plus d’un périmètre.`:'',unsure?`${plural(unsure,'ligne non géolocalisée','lignes non géolocalisées')} dans une commune ayant un quartier NPNRU, à vérifier.`:''].filter(Boolean).join(' ');
  const qList=[...qs.values()].sort((a,b)=>b.pts.length-a.pts.length||a.nom.localeCompare(b.nom,'fr'));
  $('quartiers').innerHTML=qList.map((q,k)=>`<li><button type="button" class="q ${q.cat==='national'?'nat':'reg'}" data-q="${k}"><span class="q-name">${esc(q.nom)}</span><span class="q-meta">${esc([q.qp,`intérêt ${q.cat}`].filter(Boolean).join(', '))}</span><span class="q-count">${plural(q.pts.length,'bâtiment','bâtiments')}, ${plural(q.ops.size,'opération','opérations')}</span></button></li>`).join('');
  $('quartiers').querySelectorAll('[data-q]').forEach(b=>b.onclick=()=>{ const q=qList[+b.dataset.q]; if(!state.map) return; const bounds=L.latLngBounds(q.pts.map(p=>[p.lat,p.lon])); state.map.flyToBounds(bounds.pad(.6),{maxZoom:16,duration:.9}); });
  drawPoints(pts,qi);
}

/* ---------- Carte ---------- */
const STYLE={national:{color:'#C1440E',weight:2,fillColor:'#C1440E',fillOpacity:.18},régional:{color:'#2563A8',weight:2,dashArray:'5 4',fillColor:'#2563A8',fillOpacity:.1}};
function initMap(){
  if(typeof L==='undefined'){ toast('Fond de carte indisponible (pas de connexion).','error'); return; }
  state.map=L.map('map',{zoomControl:false,preferCanvas:true,minZoom:5,maxZoom:19}).setView([46.6,2.4],6);
  L.control.zoom({position:'topright'}).addTo(state.map);
  L.tileLayer('https://data.geopf.fr/wmts?SERVICE=WMTS&REQUEST=GetTile&VERSION=1.0.0&LAYER=GEOGRAPHICALGRIDSYSTEMS.PLANIGNV2&STYLE=normal&TILEMATRIXSET=PM&FORMAT=image/png&TILEMATRIX={z}&TILEROW={y}&TILECOL={x}',{maxZoom:19,attribution:'© IGN Géoplateforme, quartiers NPNRU : ANCT / ANRU'}).addTo(state.map);
}
function drawQuartiers(){
  if(!state.map||!state.anru) return;
  L.featureGroup(state.anru.quartiers.map(q=>L.polygon(q.rings,STYLE[q.cat]||STYLE.national).bindPopup(`<strong>${esc(q.nom)}</strong><br>${esc(q.qp)}, intérêt ${esc(q.cat)}<br><span class="pop-muted">${esc(q.communes.map(c=>c[1]).join(', '))}</span>`))).addTo(state.map);
}
function drawPoints(pts,qi){
  if(!state.map) return; const r=state.result;
  state.layer?.remove();
  state.layer=L.featureGroup(pts.map(p=>{ const x=rowLabel(p.i), nat=/intérêt national/.test(r.rows[p.i][qi]||'');
    return L.circleMarker([p.lat,p.lon],{radius:7,color:'#ffffff',weight:2,fillColor:nat?'#C1440E':'#2563A8',fillOpacity:1})
      .bindPopup(`<strong>${esc(x.title)}</strong><br>${esc(x.sub)}<br>${esc(r.rows[p.i][qi])}<br><span class="pop-muted">${esc(p.mentions.join(', '))}</span>`); })).addTo(state.map);
  setTimeout(()=>state.map.invalidateSize(),0);
  // Cadrage sur la métropole quand elle a des bâtiments ; l'outre-mer est accessible par ses raccourcis.
  const metro=pts.filter(p=>territoryOf(p)==='Métropole'), frame=metro.length?metro:pts;
  if(frame.length) state.map.fitBounds(L.latLngBounds(frame.map(p=>[p.lat,p.lon])).pad(.2),{maxZoom:15});
  const by=new Map(); for(const p of pts){ const t=territoryOf(p); by.set(t,[...(by.get(t)||[]),p]); }
  const terr=[...by].sort((a,b)=>(a[0]==='Métropole'?-1:b[0]==='Métropole'?1:a[0].localeCompare(b[0],'fr')));
  $('territories').hidden=terr.length<2;
  $('territories').innerHTML=terr.map(([t,xs],k)=>`<button type="button" data-t="${k}">${esc(t)} <b>${nf(xs.length)}</b></button>`).join('');
  $('territories').querySelectorAll('[data-t]').forEach(b=>b.onclick=()=>{ const xs=terr[+b.dataset.t][1]; state.map.flyToBounds(L.latLngBounds(xs.map(p=>[p.lat,p.lon])).pad(.3),{maxZoom:14,duration:.9}); });
}
// Territoire d'un bâtiment (boîtes englobantes larges).
const TERRITORIES=[['Guadeloupe',15.8,16.6,-61.9,-61.0],['Martinique',14.3,14.95,-61.3,-60.75],['Guyane',2,6,-55,-51],['La Réunion',-21.5,-20.8,55.1,55.9],['Mayotte',-13.1,-12.6,44.9,45.4]];
const territoryOf=p=>TERRITORIES.find(([,a,b,c,d])=>p.lat>=a&&p.lat<=b&&p.lon>=c&&p.lon<=d)?.[0]||'Métropole';

/* ---------- Export CSV des bâtiments affichés (toutes colonnes, « ; » et BOM pour Excel) ---------- */
function exportCsv(){
  const r=state.result, pts=anruShown(); if(!pts.length){ toast('Aucun bâtiment à exporter : cochez au moins une mention.','error'); return; }
  const cell=v=>{ const s=v==null?'':typeof v==='number'?String(v).replace('.',','):String(v); return /[;"\n\r]/.test(s)?`"${s.replace(/"/g,'""')}"`:s; };
  const lines=[r.header,...pts.map(p=>r.header.map((_,j)=>r.rows[p.i][j]))].map(l=>l.map(cell).join(';'));
  const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob(['﻿'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'}));
  a.download=`NPNRU_batiments_${new Date().toISOString().slice(0,10)}.csv`; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}

/* ---------- Câblage ---------- */
$('sfFile').addEventListener('change',e=>{ addFiles([...e.target.files],'sf'); e.target.value=''; });
$('etFile').addEventListener('change',e=>{ addFiles([...e.target.files],'et'); e.target.value=''; });
document.querySelectorAll('.drop').forEach(z=>{ z.addEventListener('dragover',e=>{ e.preventDefault(); z.classList.add('over'); }); z.addEventListener('dragleave',()=>z.classList.remove('over'));
  z.addEventListener('drop',e=>{ e.preventDefault(); e.stopPropagation(); z.classList.remove('over'); addFiles([...e.dataTransfer.files],z.dataset.kind); }); });
document.addEventListener('dragover',e=>e.preventDefault());
document.addEventListener('drop',e=>{ e.preventDefault(); if(e.dataTransfer?.files?.length) addFiles([...e.dataTransfer.files]); });
$('mentions').addEventListener('change',e=>{ const m=e.target.dataset.mention; if(m==null) return; e.target.checked?state.mentionsOff.delete(m):state.mentionsOff.add(m); render(); });
$('mentionsAll').onclick=()=>{ state.mentionsOff.clear(); render(); };
$('mentionsNone').onclick=()=>{ anruRows().forEach(p=>p.mentions.forEach(m=>state.mentionsOff.add(m))); render(); };
$('csv').onclick=exportCsv;
$('reload').onclick=()=>{ $('load').hidden=false; $('summary').hidden=true; $('reload').hidden=true; renderLoad(); };
$('full').onclick=()=>{ if(document.fullscreenElement) document.exitFullscreen(); else document.documentElement.requestFullscreen?.(); };
document.addEventListener('fullscreenchange',()=>{ $('full').textContent=document.fullscreenElement?'Quitter le plein écran':'Plein écran'; setTimeout(()=>state.map?.invalidateSize(),150); });

initMap(); renderLoad();
fetch('../data/npnru.json').then(r=>r.ok?r.json():Promise.reject(new Error(`HTTP ${r.status}`)))
  .then(d=>{ state.anru=anruIndex(d); drawQuartiers(); run(); })
  .catch(e=>{ toast(`Quartiers NPNRU non chargés (${e.message}).`,'error'); $('loadMsg').textContent='Les quartiers NPNRU n’ont pas pu être chargés : rechargez la page.'; });
