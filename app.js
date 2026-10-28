import {recoll,opsFromSession,opsFromBuildingRows,CONTRACT_COL,OPE_COL,STATUS_COL,BUILDING_COL,KINDS,TECH_LABELS,targetColumn,empty} from './recoll.js?v=3';

const $=id=>document.getElementById(id);
const esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmt=v=>v==null?'':typeof v==='number'?v.toLocaleString('fr-FR',{maximumFractionDigits:6}):String(v);
const PAGE=100;
const KEY_COLS=[CONTRACT_COL,OPE_COL,"Nom de l'opération",'Nom de la société: Nom de la société','Référentiel',BUILDING_COL,'Total logements','Nombre de logements','Ville','Latitude','Longitude','Cep projet','Bbio projet','DH projet','IC construction'];
const state={filter:'recolle',sf:null,sfName:'',sfSheet:'',et:[],result:null,decisions:{},page:0,tab:'table',map:null,mapLayer:null};
try{ state.decisions=JSON.parse(localStorage.getItem('totalrecoll.decisions')||'{}'); }catch{}
const saveDecisions=()=>{ try{ localStorage.setItem('totalrecoll.decisions',JSON.stringify(state.decisions)); }catch{} };

/* ---------- Lecture des fichiers ---------- */
const readBook=file=>file.arrayBuffer().then(b=>XLSX.read(b,{type:'array',cellNF:true}));
// Dates Excel → AAAA-MM-JJ (lisible par Google Sheets et l'Observatoire, sans décalage de fuseau).
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
  for(const n of ['OPERATIONS',...wb.SheetNames.filter(n=>n!=='OPERATIONS')].filter(n=>wb.Sheets[n])){
    const aoa=sheetAoa(wb.Sheets[n]);
    const h=aoa.findIndex((r,i)=>i<5&&r.includes(CONTRACT_COL)&&r.includes(OPE_COL));
    if(h>=0){ state.sf={header:aoa[h].map(v=>v??''),rows:aoa.slice(h+1)}; state.sfName=file.name; state.sfSheet=n; return; }
  }
  throw new Error(`Aucune feuille de « ${file.name} » ne contient les colonnes « ${CONTRACT_COL} » et « ${OPE_COL} ».`);
}
async function loadEt(file){
  if(/\.json$/i.test(file.name)) return {name:file.name,kind:'Session JSON',ops:opsFromSession(JSON.parse(await file.text()),file.name)};
  const wb=await readBook(file), ws=wb.Sheets['Données par bâtiment'];
  if(!ws) throw new Error(`« ${file.name} » n’a pas de feuille « Données par bâtiment » (export ExtracTerre attendu).`);
  const aoa=sheetAoa(ws), h=aoa[0];
  return {name:file.name,kind:'Export Excel',ops:opsFromBuildingRows(aoa.slice(1).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]]))),file.name)};
}
const busy=(on,text='')=>{ $('busy').hidden=!on; $('busyText').textContent=text; };
const nextFrame=()=>new Promise(r=>requestAnimationFrame(()=>setTimeout(r,0)));
async function handleSf(file){
  if(!file) return; busy(true,`Lecture de « ${file.name} »… (quelques secondes pour un gros fichier)`); await nextFrame();
  try{ await loadSf(file); toast(`Tous contrats : ${state.sf.rows.length} lignes lues.`,'success'); run(); }
  catch(e){ toast(e.message,'error'); } finally{ busy(false); }
}
async function handleEt(files){
  for(const f of files){ busy(true,`Lecture de « ${f.name} »…`); await nextFrame();
    try{ const s=await loadEt(f); state.et=[...state.et.filter(x=>x.name!==s.name),s]; toast(`${f.name} : ${s.ops.length} opérations.`,'success'); }
    catch(e){ toast(e.message,'error'); } }
  busy(false); run();
}

/* ---------- Recollement ---------- */
function run(){
  renderSources();
  state.result=null;
  if(state.sf&&state.et.length){
    try{ state.result=recoll(state.sf,state.et.flatMap(e=>e.ops),{keepSalesforce:document.querySelector('input[name=default]:checked').value==='sf',decisions:state.decisions}); }
    catch(e){ toast(e.message,'error'); }
  }
  const ok=!!state.result;
  if(!ok){ $('matchLine').innerHTML=''; $('matchBar').innerHTML=''; $('matchLegend').innerHTML=''; }
  $('step-control').hidden=!ok; $('step-export').hidden=!ok; $('nextHint').hidden=ok;
  renderSteps();
  if(ok) renderControl();
}
function renderSteps(){
  const ok=!!state.result, loaded=!!state.sf&&state.et.length>0;
  document.querySelectorAll('.step').forEach(a=>{ const n=+a.dataset.step;
    a.classList.toggle('done',n===1?loaded:ok&&n===2);
    a.classList.toggle('active',n===1?!loaded:n===2?ok:false);
    a.setAttribute('aria-disabled',String(n>1&&!ok)); });
}
function renderSources(){
  const sf=state.sf;
  $('dropSf').classList.toggle('loaded',!!sf);
  $('sfCards').innerHTML=sf?`<div class="card"><div class="grow"><div class="name">${esc(state.sfName)}</div><div class="meta">${sf.rows.length.toLocaleString('fr-FR')} lignes, ${sf.header.length} colonnes, feuille ${esc(state.sfSheet)}</div></div><button class="x" data-rm-sf title="Retirer" aria-label="Retirer ${esc(state.sfName)}">×</button></div>`:'';
  $('dropEt').classList.toggle('loaded',state.et.length>0);
  $('etCards').innerHTML=state.et.map((e,i)=>`<div class="card"><div class="grow"><div class="name">${esc(e.name)}</div><div class="meta">${e.ops.length} opérations, ${e.ops.reduce((a,o)=>a+o.buildings.length,0)} bâtiments (${e.kind.toLowerCase()})</div></div><button class="x" data-rm-et="${i}" title="Retirer" aria-label="Retirer ${esc(e.name)}">×</button></div>`).join('');
}

/* ---------- Étape 2 : contrôle ---------- */
const kindOf=i=>state.result.meta[i];
const FILTERS={all:()=>true,recolle:m=>m.kind==='recolle'||m.kind==='vide',check:m=>m.check,sf:m=>m.kind==='sf',et:m=>m.kind==='et',conflit:m=>m.kind==='conflit'};
// Une opération ExtracTerre = un trait : recollée (traverse les deux registres), à vérifier, absente de Tous contrats, en conflit.
function etOperations(){
  const ops=new Map();
  state.result.meta.forEach((m,i)=>{ if(m.kind==='sf') return; const k=m.kind==='recolle'||m.kind==='vide'?'g:'+m.group+'|'+(m.source||''):'s:'+m.source;
    const o=ops.get(k); if(!o) ops.set(k,{kind:m.kind==='vide'?'recolle':m.kind,check:!!m.check,row:i}); else if(m.check) o.check=true; });
  const order={recolle:0,check:1,et:2,conflit:3}, tone=o=>o.kind==='recolle'&&o.check?'check':o.kind;
  return [...ops.values()].map(o=>({...o,tone:tone(o)})).sort((a,b)=>order[a.tone]-order[b.tone]||a.row-b.row);
}
function renderControl(){
  const r=state.result, s=r.stats, n=v=>v.toLocaleString('fr-FR');
  const ops=etOperations(), count=t=>ops.filter(o=>o.tone===t).length;
  $('matchLine').innerHTML=`<b>${n(s.recolles)}</b><span>sur ${n(s.etOperations)} opérations ExtracTerre recollées</span><small>${n(s.viaDeux)} par contrat et OPE, ${n(s.viaContrat)} par contrat seul, ${n(s.viaOpe)} par OPE seule</small>`;
  $('matchBar').innerHTML=ops.map((o,k)=>{ const x=rowLabel(o.row); return `<span class="t t-${o.tone}" style="--k:${k}" data-open="${o.row}" title="${esc(x.title)}${x.sub?'\n'+esc(x.sub):''}"></span>`; }).join('');
  $('matchLegend').innerHTML=[['recolle','Recollées'],['check','À vérifier'],['et','Absentes de Tous contrats'],['conflit','En conflit']].filter(([t])=>count(t)||t==='recolle').map(([t,l])=>`<span><i class="t-${t}"></i>${l} ${n(count(t))}</span>`).join('');
  const geo=geoRows().length, issues=r.meta.filter(m=>m.check).length;
  const chip=(filter,label,v,tone='')=>`<button class="chip ${tone}" data-kpi="${filter}" aria-pressed="${state.filter===filter}" type="button">${label}<b>${n(v)}</b></button>`;
  $('kpis').innerHTML=chip('recolle','Recollées',r.meta.filter(m=>m.kind==='recolle'||m.kind==='vide').length,'et')+chip('check','À contrôler',issues,issues?'check':'')
    +chip('sf','Tous contrats seul',r.meta.filter(m=>m.kind==='sf').length,'sf')+chip('et','ExtracTerre seul',r.meta.filter(m=>m.kind==='et').length)
    +(s.conflits?chip('conflit','En conflit',r.meta.filter(m=>m.kind==='conflit').length,'bad'):'')+chip('all','Toutes les lignes',s.lignes);
  $('nIssues').textContent=issues; $('nEcarts').textContent=s.ecarts; $('nMap').textContent=geo;
  $('exportLine').innerHTML=`<b>${n(s.lignes)} lignes × ${r.header.length} colonnes</b> prêtes pour la feuille OPERATIONS de l’Observatoire`;
  renderTable(); renderIssues(); renderEcarts(); renderColumns(); if(state.tab==='map') renderMap();
}
function visibleColumns(){
  const r=state.result, si=r.header.indexOf(STATUS_COL), preset=$('preset').value;
  let cols=r.header.map((_,i)=>i).filter(i=>i!==si);
  if(preset==='key') cols=KEY_COLS.map(h=>r.header.indexOf(h)).filter(i=>i>=0);
  if(preset==='et'){ const used=new Set(r.meta.flatMap(m=>m.et)); const id=[CONTRACT_COL,OPE_COL,"Nom de l'opération",BUILDING_COL].map(h=>r.header.indexOf(h)); cols=[...new Set([...id,...cols.filter(i=>used.has(i))])].filter(i=>i>=0); }
  return [si,...cols];
}
function filteredRows(){
  const r=state.result, q=$('search').value.trim().toLowerCase(), f=FILTERS[state.filter];
  const out=[]; r.rows.forEach((row,i)=>{ if(f(r.meta[i])&&(!q||row.some(v=>v!=null&&String(v).toLowerCase().includes(q)))) out.push(i); });
  return out;
}
const statusHtml=(i)=>{ const m=kindOf(i), t=state.result.rows[i][state.result.header.indexOf(STATUS_COL)]; return `<span class="status s-${m.kind} ${m.check?'s-check':''}" title="${esc(t)}">${esc(m.check&&m.kind==='recolle'?'À vérifier':KINDS[m.kind])}</span>`; };
function renderTable(){
  const r=state.result, idx=filteredRows(), cols=visibleColumns(), si=cols[0];
  const pages=Math.max(1,Math.ceil(idx.length/PAGE)); state.page=Math.min(state.page,pages-1);
  const slice=idx.slice(state.page*PAGE,(state.page+1)*PAGE);
  $('table').innerHTML=`<thead><tr>${cols.map(c=>`<th scope="col">${esc(r.header[c])}</th>`).join('')}</tr></thead><tbody>${slice.map(i=>{ const et=new Set(r.meta[i].et);
    return `<tr tabindex="0" data-row="${i}">${cols.map(c=>c===si?`<td>${statusHtml(i)}</td>`:`<td class="${et.has(c)?'et':''}" title="${esc(fmt(r.rows[i][c]))}">${esc(fmt(r.rows[i][c]))}</td>`).join('')}</tr>`; }).join('')}</tbody>`;
  $('shown').textContent=idx.length?`${idx.length.toLocaleString('fr-FR')} ligne(s) — ${state.page*PAGE+1} à ${Math.min(idx.length,(state.page+1)*PAGE)}`:'Aucune ligne ne correspond.';
  $('pager').innerHTML=pages>1?`<button data-page="-1" ${state.page?'':'disabled'} aria-label="Page précédente">‹</button><span>${state.page+1} / ${pages}</span><button data-page="1" ${state.page<pages-1?'':'disabled'} aria-label="Page suivante">›</button>`:'';
}
function rowLabel(i){
  const r=state.result, row=r.rows[i], h=n=>row[r.header.indexOf(n)];
  return {title:h("Nom de l'opération")||h(OPE_COL)||h(CONTRACT_COL)||`Sans identifiant — ${String(r.meta[i].source||'').split(' › ').pop()}`,sub:[h(CONTRACT_COL),h(OPE_COL),h(BUILDING_COL),h('Nom de la société: Nom de la société')].filter(v=>!empty(v)).join(', ')};
}
const ISSUE_TEXT={
  conflit:['Conflits contrat / OPE','Le numéro de contrat et le numéro d’opération désignent deux opérations différentes de Tous contrats. Rien n’a été fusionné : corrigez l’identifiant faux dans ExtracTerre ou dans Salesforce, puis rechargez.','bad'],
  et:['Absentes de Tous contrats','Opérations ExtracTerre dont ni le contrat ni l’OPE ne figurent dans Tous contrats (contrat récent, numéro mal saisi, ou projet sans identifiant). Elles sont ajoutées en fin de tableau.',''],
  check:['Bâtiments à vérifier','Tous contrats a déjà plusieurs lignes pour ce contrat, mais pas autant que de bâtiments ExtracTerre : la première ligne sert de base et ses valeurs techniques ne sont pas reprises. Vérifiez le découpage.','']
};
function renderIssues(){
  const r=state.result, groups={conflit:[],et:[],check:[]};
  r.meta.forEach((m,i)=>{ if(m.kind==='conflit') groups.conflit.push(i); else if(m.kind==='et') groups.et.push(i); else if(m.check) groups.check.push(i); });
  const html=Object.entries(groups).filter(([,l])=>l.length).map(([k,l])=>{ const [t,d,tone]=ISSUE_TEXT[k];
    return `<div class="issue-group"><h3>${t} (${l.length})</h3><p>${d}</p>${l.slice(0,200).map(i=>{ const x=rowLabel(i); const st=r.rows[i][r.header.indexOf(STATUS_COL)];
      return `<div class="issue ${tone}"><div class="grow"><b>${esc(x.title)}</b><div>${esc(x.sub)}</div><div>${esc(st)}</div></div><button class="btn light small" data-open="${i}" type="button">Détail</button></div>`; }).join('')}${l.length>200?`<p class="muted">… ${l.length-200} autres : filtrez le tableau sur « À contrôler ».</p>`:''}</div>`; }).join('');
  $('issues').innerHTML=html||'<div class="empty">Rien à contrôler : tout est recollé sans conflit.</div>';
}
function renderEcarts(){
  const r=state.result, sel=$('ecartCol'), cur=sel.value;
  const colsUsed=[...new Set(r.ecarts.map(e=>e.colonne))].sort();
  sel.innerHTML='<option value="">Toutes les colonnes</option>'+colsUsed.map(c=>`<option ${c===cur?'selected':''}>${esc(c)}</option>`).join('');
  const list=r.ecarts.filter(e=>!sel.value||e.colonne===sel.value);
  $('ecarts').innerHTML=list.length?`<thead><tr><th>Contrat</th><th>OPE</th><th>Bâtiment</th><th>Colonne</th><th>Tous contrats</th><th>ExtracTerre</th><th>Valeur retenue</th><th>Source</th></tr></thead><tbody>${list.map(e=>`<tr><td>${esc(e.contrat)}</td><td>${esc(e.ope)}</td><td>${esc(e.batiment)}</td><td>${esc(e.colonne)}</td><td class="num v-sf">${esc(fmt(e.salesforce))}</td><td class="num v-et">${esc(fmt(e.extracterre))}</td>
    <td><span class="seg" role="group" aria-label="Valeur retenue"><button type="button" data-dec="${esc(e.key)}" data-v="sf" aria-pressed="${e.retenu==='Salesforce'}">Tous contrats</button><button type="button" data-dec="${esc(e.key)}" data-v="et" aria-pressed="${e.retenu==='ExtracTerre'}">ExtracTerre</button></span>${e.decided?'':' <small class="muted">défaut</small>'}</td><td title="${esc(e.source)}">${esc(e.source)}</td></tr>`).join('')}</tbody>`:'<tbody><tr><td class="empty">Aucun écart : ExtracTerre ne contredit aucune valeur déjà présente.</td></tr></tbody>';
}
function renderColumns(){
  const r=state.result, filled={};
  r.meta.forEach(m=>m.et.forEach(c=>{ filled[c]=(filled[c]||0)+1; }));
  const seen=new Set(), rows=[];
  for(const l of TECH_LABELS){ const h=targetColumn(l), c=r.header.indexOf(h); if(c<0||seen.has(l)) continue; seen.add(l);
    rows.push([l,h,c>=r.sfWidth?'ajoutée':h===l?'même nom':'correspondance',filled[c]||0]); }
  rows.sort((a,b)=>b[3]-a[3]);
  $('columns').innerHTML=`<thead><tr><th>Libellé ExtracTerre</th><th>Colonne du tableau</th><th>Type</th><th>Cellules remplies</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x[0])}</td><td>${esc(x[1])}</td><td>${x[2]}</td><td class="num">${x[3]}</td></tr>`).join('')}</tbody>`;
}

/* ---------- Carte ---------- */
function geoRows(){
  const r=state.result, la=r.header.indexOf('Latitude'), lo=r.header.indexOf('Longitude'); if(la<0||lo<0) return [];
  return r.rows.map((row,i)=>({i,lat:Number(row[la]),lon:Number(row[lo])})).filter(p=>Number.isFinite(p.lat)&&Number.isFinite(p.lon)&&p.lat!==0);
}
function renderMap(){
  const pts=geoRows(), r=state.result, gi=r.header.indexOf('Géolocalisation');
  const failed=gi<0?0:r.rows.filter(row=>/^Non géolocalisée/.test(row[gi]||'')).length;
  $('mapInfo').textContent=pts.length?`${pts.length} bâtiment(s) placés au point adresse IGN.${failed?` ${failed} adresse(s) non géolocalisée(s) dans ExtracTerre (motif dans la colonne « Géolocalisation »).`:''}`:'Aucun bâtiment géolocalisé : les coordonnées viennent d’ExtracTerre v2.9.0 et plus (colonnes Latitude / Longitude).';
  if(typeof L==='undefined'){ $('mapInfo').textContent+=' (fond de carte indisponible hors connexion)'; return; }
  if(!state.map){
    state.map=L.map('map',{scrollWheelZoom:true}).setView([46.6,2.4],6);
    L.tileLayer('https://data.geopf.fr/wmts?SERVICE=WMTS&REQUEST=GetTile&VERSION=1.0.0&LAYER=GEOGRAPHICALGRIDSYSTEMS.PLANIGNV2&STYLE=normal&TILEMATRIXSET=PM&FORMAT=image/png&TILEMATRIX={z}&TILEROW={y}&TILECOL={x}',{maxZoom:19,attribution:'© IGN — Géoplateforme'}).addTo(state.map);
  }
  setTimeout(()=>state.map.invalidateSize(),0);
  state.mapLayer?.remove();
  state.mapLayer=L.featureGroup(pts.map(p=>{ const x=rowLabel(p.i);
    return L.circleMarker([p.lat,p.lon],{radius:7,color:'#fff',weight:2,fillColor:'#17865A',fillOpacity:.9}).bindPopup(`<b>${esc(x.title)}</b><br>${esc(x.sub)}<br><button class="btn small" data-open="${p.i}" type="button">Détail</button>`); })).addTo(state.map);
  if(pts.length) state.map.fitBounds(state.mapLayer.getBounds().pad(.15),{maxZoom:16});
}

/* ---------- Détail d'une ligne ---------- */
function openRow(i){
  const r=state.result, row=r.rows[i], m=r.meta[i], x=rowLabel(i), et=new Set(m.et), st=row[r.header.indexOf(STATUS_COL)];
  $('drawerKind').textContent=m.check&&m.kind==='recolle'?'À vérifier':KINDS[m.kind];
  $('drawerTitle').textContent=x.title; $('drawerSub').textContent=x.sub;
  const dl=list=>`<dl class="kv">${list.map(c=>`<dt>${esc(r.header[c])}</dt><dd class="${et.has(c)?'et':''}">${esc(fmt(row[c]))}</dd>`).join('')}</dl>`;
  const etCols=[...et], sfCols=r.header.map((_,c)=>c).filter(c=>!et.has(c)&&c!==r.header.indexOf(STATUS_COL)&&!empty(row[c]));
  const ec=r.ecarts.filter(e=>e.rowIndex===i);
  const la=Number(row[r.header.indexOf('Latitude')]), lo=Number(row[r.header.indexOf('Longitude')]);
  $('drawerBody').innerHTML=`<p class="note ${m.kind==='conflit'?'bad':''}" ${m.check||m.kind==='conflit'?'':'hidden'}>${esc(st)}</p>
    ${Number.isFinite(la)&&Number.isFinite(lo)&&la?`<p><a class="btn light small" target="_blank" rel="noopener" href="https://www.geoportail.gouv.fr/carte?c=${lo},${la}&z=19&l0=ORTHOIMAGERY.ORTHOPHOTOS::GEOPORTAIL:OGC:WMTS(1)&permalink=yes">Voir sur Géoportail ↗</a></p>`:''}
    ${ec.length?`<h4 class="h-check">Écarts (${ec.length})</h4>${ec.map(e=>`<div class="issue"><div class="grow"><b>${esc(e.colonne)}</b><div><span class="v-sf">Tous contrats : ${esc(fmt(e.salesforce))}</span>, <span class="v-et">ExtracTerre : ${esc(fmt(e.extracterre))}</span></div></div><span class="seg"><button type="button" data-dec="${esc(e.key)}" data-v="sf" aria-pressed="${e.retenu==='Salesforce'}">Tous contrats</button><button type="button" data-dec="${esc(e.key)}" data-v="et" aria-pressed="${e.retenu==='ExtracTerre'}">ExtracTerre</button></span></div>`).join('')}`:''}
    <h4 class="h-et">Données ExtracTerre (${etCols.length})</h4>${etCols.length?dl(etCols):'<p class="muted">Aucune.</p>'}
    <h4 class="h-sf">Données Tous contrats (${sfCols.length})</h4>${dl(sfCols)}
    ${m.source?`<p class="muted">Source ExtracTerre : ${esc(m.source)}</p>`:''}`;
  $('drawer').hidden=false; $('drawer').dataset.row=i; $('drawerClose').focus();
}
const closeDrawer=()=>{ $('drawer').hidden=true; };

/* ---------- Export ---------- */
function exportXlsx(){
  const r=state.result, wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([r.header,...r.rows]),'OPERATIONS');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Contrat','OPE','Bâtiment','Colonne','Tous contrats','ExtracTerre','Retenu','Choix','Source'],...r.ecarts.map(e=>[e.contrat,e.ope,e.batiment,e.colonne,e.salesforce,e.extracterre,e.retenu,e.decided?'manuel':'par défaut',e.source])]),'Écarts');
  const si=r.header.indexOf(STATUS_COL);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Nature','Contrat','OPE','Opération','Bâtiment','Détail'],...r.meta.map((m,i)=>[m,i]).filter(([m])=>m.check||m.kind==='conflit'||m.kind==='et').map(([m,i])=>{ const row=r.rows[i], h=n=>row[r.header.indexOf(n)]; return [KINDS[m.kind],h(CONTRACT_COL),h(OPE_COL),h("Nom de l'opération"),h(BUILDING_COL),row[si]]; })]),'À contrôler');
  XLSX.writeFile(wb,`TotalRecoll_OPERATIONS_${new Date().toISOString().slice(0,10)}.xlsx`);
  toast('Fichier téléchargé.','success');
}
async function copyForSheets(){
  const r=state.result, ci=r.header.indexOf(CONTRACT_COL);
  // ="0000…" garde les zéros du numéro de contrat au collage dans Google Sheets.
  const cell=(v,j)=>{ if(v==null) return ''; const s=String(v).replace(/[\t\r\n]+/g,' '); return j===ci&&/^0\d+$/.test(s)?`="${s}"`:s; };
  await navigator.clipboard.writeText([r.header,...r.rows].map(row=>r.header.map((_,j)=>cell(row[j],j)).join('\t')).join('\n'));
  toast(`${r.rows.length.toLocaleString('fr-FR')} lignes copiées : collez-les en A1 de la feuille OPERATIONS.`,'success');
}

/* ---------- Divers ---------- */
function toast(msg,kind=''){ const t=document.createElement('div'); t.className='toast '+kind; t.textContent=msg; $('toasts').append(t); setTimeout(()=>t.remove(),6000); }
function setTab(tab){
  state.tab=tab;
  document.querySelectorAll('.tab').forEach(b=>b.setAttribute('aria-selected',String(b.dataset.tab===tab)));
  document.querySelectorAll('.panel').forEach(p=>{ p.hidden=p.dataset.panel!==tab; });
  if(tab==='map') renderMap();
}
function decide(key,v){ state.decisions[key]=v; saveDecisions(); const open=$('drawer').hidden?null:+$('drawer').dataset.row; run(); if(open!=null) openRow(open); }
function setupDrop(zone,input,handler){
  // Le clic relancé sur le champ caché remonterait jusqu'ici : on l'ignore pour ne pas rouvrir le sélecteur.
  input.addEventListener('click',e=>e.stopPropagation());
  zone.addEventListener('click',e=>{ if(e.target!==input&&!e.target.closest('button')) input.click(); });
  zone.addEventListener('keydown',e=>{ if((e.key==='Enter'||e.key===' ')&&e.target===zone){ e.preventDefault(); input.click(); } });
  ['dragenter','dragover'].forEach(t=>zone.addEventListener(t,e=>{ e.preventDefault(); zone.classList.add('drag'); }));
  ['dragleave','drop'].forEach(t=>zone.addEventListener(t,()=>zone.classList.remove('drag')));
  zone.addEventListener('drop',e=>{ e.preventDefault(); handler([...e.dataTransfer.files]); });
  // Le champ n'est vidé qu'après lecture : certains navigateurs invalident le fichier dès qu'on le vide.
  input.addEventListener('change',async()=>{ await handler([...input.files]); input.value=''; });
}
setupDrop($('dropSf'),$('sfFile'),files=>handleSf(files[0]));
setupDrop($('dropEt'),$('etFile'),files=>handleEt(files));
document.addEventListener('click',e=>{
  const t=e.target.closest('button,[data-row],[data-open]'); if(!t) return;
  if(t.matches('[data-rm-sf]')){ state.sf=null; run(); }
  else if(t.matches('[data-rm-et]')){ state.et.splice(+t.dataset.rmEt,1); run(); }
  else if(t.matches('[data-kpi]')){ state.filter=t.dataset.kpi; state.page=0; setTab('table'); renderControl(); }
  else if(t.matches('[data-goto]')) setTab(t.dataset.goto);
  else if(t.matches('.tab')) setTab(t.dataset.tab);
  else if(t.matches('[data-page]')){ state.page+=+t.dataset.page; renderTable(); }
  else if(t.matches('[data-dec]')) decide(t.dataset.dec,t.dataset.v);
  else if(t.matches('[data-open]')) openRow(+t.dataset.open);
  else if(t.matches('[data-row]')) openRow(+t.dataset.row);
});
document.addEventListener('keydown',e=>{
  if(e.key==='Escape'&&!$('drawer').hidden) closeDrawer();
  if(e.key==='Enter'&&e.target.matches?.('tr[data-row]')) openRow(+e.target.dataset.row);
});
$('drawerClose').addEventListener('click',closeDrawer);
document.querySelectorAll('input[name=default]').forEach(i=>i.addEventListener('change',run));
$('preset').addEventListener('change',renderTable);
let searchTimer; $('search').addEventListener('input',()=>{ clearTimeout(searchTimer); searchTimer=setTimeout(()=>{ state.page=0; renderTable(); },180); });
$('ecartCol').addEventListener('change',renderEcarts);
const bulk=v=>{ state.result.ecarts.filter(e=>!$('ecartCol').value||e.colonne===$('ecartCol').value).forEach(e=>{ state.decisions[e.key]=v; }); saveDecisions(); run(); };
$('allEt').addEventListener('click',()=>bulk('et'));
$('allSf').addEventListener('click',()=>bulk('sf'));
$('resetDec').addEventListener('click',()=>{ state.decisions={}; saveDecisions(); run(); });
$('xlsxBtn').addEventListener('click',exportXlsx);
$('copyBtn').addEventListener('click',()=>copyForSheets().catch(e=>toast('Copie impossible : '+e.message,'error')));
$('helpBtn').addEventListener('click',()=>$('help').showModal());
renderSteps();
