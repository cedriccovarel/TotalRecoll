import {REL_ORDER,REL_LABEL,relKey,recoll,opsFromSession,opsFromBuildingRows,CONTRACT_COL,OPE_COL,STATUS_COL,BUILDING_COL,KINDS,TECH_LABELS,targetColumn,empty,anruIndex,ANRU_COL,ANRU_QP_COL,confortOf,oddValue,COLUMN_KEY,sortRowIndexes,ZONES,zoneOfDep,depOfRow,confortStats,SYSTEMS,DH_ZT,DH_ZNT,publishTable,researchFromSession,mergeResearch,researchRows,researchStats,NO_REF} from './recoll.js?v=22';
import {verificationItems,verifiedSearch,verificationProgress,highlightRanges} from './recoll.js?v=22';

const $=id=>document.getElementById(id);
const esc=v=>String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmt=v=>v==null?'':typeof v==='number'?v.toLocaleString('fr-FR',{maximumFractionDigits:6}):String(v);
const PAGE=100;
const KEY_COLS=[CONTRACT_COL,OPE_COL,"Nom de l'opération",'Nom de la société: Nom de la société','Référentiel',BUILDING_COL,'Total logements','Nombre de logements','Ville','Latitude','Longitude','Cep projet','Bbio projet','DH projet','IC construction','ANRU'];
const state={filter:'recolle',sf:null,sfName:'',sfSheet:'',et:[],result:null,decisions:{},page:0,sort:{col:null,dir:''},tab:'table',map:null,mapLayer:null,anru:null,anruMap:null,anruLayer:null,mentionsOff:new Set(),zonesGeo:null,zoneMap:null,zone:'',rs:{grid:'',ref:'',owner:'',q:''}};
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
// Session ExtracTerre : seule la tête (« operations ») sert ici. Lue en flux jusqu'à « "projects": [ » : le texte des documents qui
// suit n'est pas lu (une session de 550 Mo dépasse la taille maximale d'une chaîne de Chrome → « Unexpected end of JSON input »).
async function readSessionHead(file){
  const MARK=/"projects":\s*\[\{?\s*$/, reader=file.stream().pipeThrough(new TextDecoderStream()).getReader(); const parts=[]; let tail='';
  while(true){ const {value,done}=await reader.read(); if(done) break;
    const chunk=tail+value, cut=chunk.lastIndexOf('\n'); const lines=cut>=0?chunk.slice(0,cut):''; tail=cut>=0?chunk.slice(cut+1):chunk;
    if(lines){ const ls=lines.split('\n'); const k=ls.findIndex(l=>MARK.test(l));
      if(k>=0){ parts.push([...ls.slice(0,k),ls[k].replace(/\{\s*$/,'')].join('\n')); try{ reader.cancel(); }catch{} return JSON.parse(parts.join('\n').replace(/^\uFEFF?[^{]*/,'')+']}}'); }
      parts.push(lines); } }
  if(tail) parts.push(tail);
  try{ return JSON.parse(parts.join('\n').replace(/^\uFEFF?[^{]*/,'')); }catch(e){ throw new Error(`Session JSON illisible (${e.message}).`); }
}
async function loadSf(file){
  const wb=await readBook(file);
  for(const n of ['OPERATIONS',...wb.SheetNames.filter(n=>n!=='OPERATIONS')].filter(n=>wb.Sheets[n])){
    const aoa=sheetAoa(wb.Sheets[n]);
    // Export Salesforce brut : titre, date et filtres au-dessus des en-têtes (ligne 14 environ), colonnes sans en-tête
    // (A, C ; retirées seulement dans ce cas, la mise en page OPERATIONS reste intacte) et lignes Total / Nombre / mentions légales en bas : en-têtes cherchés sur 30 lignes, le reste écarté.
    const h=aoa.findIndex((r,i)=>i<30&&r.includes(CONTRACT_COL)&&r.includes(OPE_COL));
    if(h>=0){ const head=aoa[h].map(v=>String(v??'').trim()), keep=head.map((v,j)=>v||!h?j:-1).filter(j=>j>=0);
      const ci=head.indexOf(CONTRACT_COL);
      const footer=r=>{ const c=String(r[ci]??'').trim(); if(/^\d+(?:\.0+)?$/.test(c)) return false;
        return r.some(v=>/^(?:total|somme|nombre|copyright\b.*|informations confidentielles\b.*)$/i.test(String(v??'').trim())); };
      const rows=aoa.slice(h+1).filter(r=>!footer(r)).map(r=>keep.map(j=>r[j]??null)).filter(r=>r.some(v=>v!=null&&v!==''));
      state.sf={header:keep.map(j=>head[j]),rows}; state.sfName=file.name; state.sfSheet=n; return; }
  }
  throw new Error(`Aucune feuille de « ${file.name} » ne contient les colonnes « ${CONTRACT_COL} » et « ${OPE_COL} ».`);
}
async function loadEt(file){
  // Session : bâtiments (recollement) + résultats de l'onglet Recherches ; export « observatoire » : recherches seules.
  if(/\.json$/i.test(file.name)){ const head=await readSessionHead(file), research=researchFromSession(head,file.name);
    return {name:file.name,kind:head.kind==='observatoire'?'Recherches':'Session JSON',ops:head.kind==='observatoire'?[]:opsFromSession(head,file.name),research}; }
  const wb=await readBook(file), ws=wb.Sheets['Données par bâtiment'];
  if(!ws) throw new Error(`« ${file.name} » n’a pas de feuille « Données par bâtiment » (export ExtracTerre attendu).`);
  const aoa=sheetAoa(ws), h=aoa[0];
  // ExtracTerre ≥ 2.34 : feuille « Fiabilité par bâtiment », même grille (ligne à ligne) avec le niveau de chaque valeur.
  const wr=wb.Sheets['Fiabilité par bâtiment'], ra=wr?sheetAoa(wr):null, rh=ra?.[0]||[];
  const relRows=ra?ra.slice(1).map(r=>Object.fromEntries(rh.map((k,i)=>[k,r[i]]).filter(([,v])=>relKey(v)))):[];
  return {name:file.name,kind:'Export Excel',ops:opsFromBuildingRows(aoa.slice(1).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]]))),file.name,relRows)};
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
    try{ state.result=recoll(state.sf,state.et.flatMap(e=>e.ops),{keepSalesforce:document.querySelector('input[name=default]:checked').value==='sf',decisions:state.decisions,anru:state.anru}); }
    catch(e){ toast(e.message,'error'); }
  }
  const ok=!!state.result;
  if(!ok){ $('matchLine').innerHTML=''; $('matchBar').innerHTML=''; $('matchLegend').innerHTML=''; }
  $('step-control').hidden=!ok; $('step-export').hidden=!ok; $('nextHint').hidden=ok;
  renderSteps();
  if(ok){ renderControl(); const sc=document.querySelector('.stitch'); sc.classList.remove('scan'); void sc.offsetWidth; sc.classList.add('scan'); }
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
  $('etCards').innerHTML=state.et.map((e,i)=>`<div class="card"><div class="grow"><div class="name">${esc(e.name)}</div><div class="meta">${e.ops.length} opérations, ${e.ops.reduce((a,o)=>a+o.buildings.length,0)} bâtiments${e.research?.ops.length?`, ${e.research.ops.length} avec recherches`:''} (${e.kind.toLowerCase()})</div></div><button class="x" data-rm-et="${i}" title="Retirer" aria-label="Retirer ${esc(e.name)}">×</button></div>`).join('');
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
  $('nIssues').textContent=issues; $('nEcarts').textContent=s.ecarts; $('nMap').textContent=geo; $('nAnru').textContent=anruRows().length||''; $('nConfort').textContent=confortRows().length||''; $('nResearch').textContent=researchAll().length||'';
  $('exportLine').innerHTML=`<b>${n(s.lignes)} lignes × ${r.header.length} colonnes</b> prêtes pour la feuille OPERATIONS de l’Observatoire`;
  renderTable(); renderIssues(); renderEcarts(); renderColumns(); if(state.tab==='map') renderMap(); if(state.tab==='anru') renderAnru(); if(state.tab==='confort') renderConfort(); if(state.tab==='research') renderResearch();
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
  return sortRowIndexes(out,r.rows,state.sort.col,state.sort.dir);
}
const statusHtml=(i)=>{ const m=kindOf(i), t=state.result.rows[i][state.result.header.indexOf(STATUS_COL)]; return `<span class="status s-${m.kind} ${m.check?'s-check':''}" title="${esc(t)}">${esc(m.check&&m.kind==='recolle'?'À vérifier':KINDS[m.kind])}</span>`; };
function renderTable(){
  const r=state.result, idx=filteredRows(), cols=visibleColumns(), si=cols[0];
  const pages=Math.max(1,Math.ceil(idx.length/PAGE)); state.page=Math.min(state.page,pages-1);
  const slice=idx.slice(state.page*PAGE,(state.page+1)*PAGE);
  const arrows=c=>{ const on=d=>state.sort.col===c&&state.sort.dir===d; return `<span class="sort" role="group" aria-label="Trier ${esc(r.header[c])}"><button type="button" data-sort="${c}" data-dir="asc" aria-pressed="${on('asc')}" title="Ordre croissant">▲</button><button type="button" data-sort="${c}" data-dir="desc" aria-pressed="${on('desc')}" title="Ordre décroissant">▼</button></span>`; };
  const sortAttr=c=>state.sort.col===c?` aria-sort="${state.sort.dir==='asc'?'ascending':'descending'}"`:'';
  $('table').innerHTML=`<thead><tr>${cols.map(c=>`<th scope="col"${sortAttr(c)}><span class="th-in">${esc(r.header[c])}${arrows(c)}</span></th>`).join('')}</tr></thead><tbody>${slice.map(i=>{ const et=new Set(r.meta[i].et);
    return `<tr tabindex="0" data-row="${i}">${cols.map(c=>{ if(c===si) return `<td>${statusHtml(i)}</td>`; const v=r.rows[i][c], odd=oddValue(COLUMN_KEY[r.header[c]],v);
      const rk=et.has(c)?(r.meta[i].rel?.[c]||'inconnue'):'', ec=ecartTip(i,r.header[c]);
      return `<td class="${et.has(c)?'et':''}${odd?' odd':''}${rk?` rel-${rk}`:''}${ec?' sfdiff':''}" title="${esc([odd?`${fmt(v)} — ${odd}`:fmt(v),ec,rk?`Fiabilité ExtracTerre : ${REL_LABEL[rk]||'inconnue (export antérieur à la v2.34)'}`:''].filter(Boolean).join('\n'))}">${esc(fmt(v))}</td>`; }).join('')}</tr>`; }).join('')}</tbody>`;
  $('shown').textContent=idx.length?`${idx.length.toLocaleString('fr-FR')} ligne(s) — ${state.page*PAGE+1} à ${Math.min(idx.length,(state.page+1)*PAGE)}`:'Aucune ligne ne correspond.';
  $('pager').innerHTML=pages>1?`<button data-page="-1" ${state.page?'':'disabled'} aria-label="Page précédente">‹</button><span>${state.page+1} / ${pages}</span><button data-page="1" ${state.page<pages-1?'':'disabled'} aria-label="Page suivante">›</button>`:'';
}
// Valeur différente entre Tous contrats (Salesforce) et ExtracTerre (document le plus récent) : toujours rappelée au survol.
function ecartTip(i,colName){
  const r=state.result; if(!r) return ''; if(!r.ecartIndex){ r.ecartIndex=new Map(); for(const e of r.ecarts) r.ecartIndex.set(e.rowIndex+'|'+e.colonne,e); }
  const e=r.ecartIndex.get(i+'|'+colName); if(!e) return '';
  return `Salesforce (Tous contrats) : ${fmt(e.salesforce)}\nExtracTerre (document le plus récent) : ${fmt(e.extracterre)}\nRetenu : ${e.retenu}`;
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
  $('ecarts').innerHTML=list.length?`<thead><tr><th>Contrat</th><th>OPE</th><th>Bâtiment</th><th>Colonne</th><th>Tous contrats</th><th>ExtracTerre</th><th>Valeur retenue</th><th>Source</th></tr></thead><tbody>${list.map(e=>`<tr><td>${esc(e.contrat)}</td><td>${esc(e.ope)}</td><td>${esc(e.batiment)}</td><td>${esc(e.colonne)}</td><td class="num v-sf${oddValue(COLUMN_KEY[e.colonne],e.salesforce)?' odd':''}" title="${esc(oddValue(COLUMN_KEY[e.colonne],e.salesforce))}">${esc(fmt(e.salesforce))}</td><td class="num v-et${oddValue(COLUMN_KEY[e.colonne],e.extracterre)?' odd':''}" title="${esc(oddValue(COLUMN_KEY[e.colonne],e.extracterre))}">${esc(fmt(e.extracterre))}</td>
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
    return L.circleMarker([p.lat,p.lon],{radius:7,color:'#1A0B08',weight:2,fillColor:'#4FE0D2',fillOpacity:.9}).bindPopup(`<b>${esc(x.title)}</b><br>${esc(x.sub)}<br><button class="btn small" data-open="${p.i}" type="button">Détail</button>`); })).addTo(state.map);
  if(pts.length) state.map.fitBounds(state.mapLayer.getBounds().pad(.15),{maxZoom:16});
}

/* ---------- ANRU : bâtiments en quartier NPNRU, filtrés par mention ---------- */
const MENTIONS_COL='Opération: Mentions', NO_MENTION='(mention non renseignée)';
// « BEE+<br>Option TFPB<br> » : une mention par ligne dans Salesforce.
const mentionsOf=v=>{ const l=String(v??'').split(/<br\s*\/?>|&lt;br&gt;|\n/i).map(x=>x.trim()).filter(Boolean); return l.length?l:[NO_MENTION]; };
function anruRows(){
  const r=state.result; if(!r) return []; const a=r.header.indexOf(ANRU_COL), la=r.header.indexOf('Latitude'), lo=r.header.indexOf('Longitude'), mc=r.header.indexOf(MENTIONS_COL);
  if(a<0) return [];
  return r.rows.map((row,i)=>({i,lat:Number(row[la]),lon:Number(row[lo]),mentions:mentionsOf(mc<0?'':row[mc])})).filter(p=>r.rows[p.i][a]==='Oui');
}
const anruShown=(all=anruRows())=>all.filter(p=>p.mentions.some(m=>!state.mentionsOff.has(m)));
// Export des bâtiments affichés sur la carte ANRU (mentions cochées), toutes colonnes ; « ; » et BOM pour Excel.
function exportAnruCsv(){
  const r=state.result, pts=anruShown(); if(!pts.length){ toast('Aucune opération à exporter : cochez au moins une mention.','error'); return; }
  const mc=r.header.indexOf(MENTIONS_COL), val=(row,j)=>j===mc&&row[j]!=null?mentionsOf(row[j]).filter(m=>m!==NO_MENTION).join(' / '):row[j];
  downloadCsv('ANRU',r.header,pts.map(p=>r.header.map((_,j)=>val(r.rows[p.i],j))));
  toast(`${pts.length} bâtiment(s) exporté(s), ${new Set(pts.map(p=>r.meta[p.i].group||p.i)).size} opération(s).`,'success');
}
// « ; » et BOM pour Excel ; nombres à virgule.
function downloadCsv(name,header,rows){
  const cell=v=>{ const s=v==null?'':typeof v==='number'?String(v).replace('.',','):String(v); return /[;"\n\r]/.test(s)?`"${s.replace(/"/g,'""')}"`:s; };
  const csv='\ufeff'+[header,...rows].map(row=>row.map(cell).join(';')).join('\r\n');
  const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'})); a.download=`TotalRecoll_${name}_${new Date().toISOString().slice(0,10)}.csv`; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}

/* ---------- Recherches ExtracTerre, croisées avec le référentiel de Tous contrats ---------- */
let rsCache=null;
// 18/11 — vérification humaine des « Oui » : décisions mémorisées (navigateur + fichier), appliquées aux statistiques et au CSV.
// 18/11/2026 — nouvelles règles de comptage (Non = indice contraire, Sans objet = nul) : clé v2, les anciennes réponses sont
// effacées au premier chargement (demande de l'utilisateur : recommencer la vérification avec ces règles).
const VERIF_LS='totalrecoll.verifications.v2';
try{ localStorage.removeItem('totalrecoll.verifications'); }catch{}
try{ state.verif=JSON.parse(localStorage.getItem(VERIF_LS)||'{}'); }catch{ state.verif={}; }
let verifRev=0; const saveVerif=()=>{ verifRev++; try{ localStorage.setItem(VERIF_LS,JSON.stringify(state.verif)); }catch{} };
function researchData(){
  const parts=state.et.map(e=>e.research).filter(Boolean), key=parts.map(p=>p.ops.length).join(',')+'|'+state.et.map(e=>e.name).join(',')+'|'+(state.sf?.rows.length||0)+state.sfName;
  if(rsCache?.key!==key){ const m=mergeResearch(parts); rsCache={key,grids:m.grids,raw:researchRows(m,state.sf),rev:-1}; }
  if(rsCache.rev!==verifRev){ rsCache.rows=rsCache.raw.map(o=>({...o,searches:Object.fromEntries(Object.entries(o.searches).map(([gid,sr])=>{ const g=rsCache.grids.find(x=>String(x.id)===gid); return [gid,g?verifiedSearch(o,g,state.verif):sr]; }))})); rsCache.rev=verifRev; }
  return rsCache;
}
const researchAll=()=>researchData().rows;
function researchSelected(raw=false){
  const d=researchData(), rows=raw?d.raw:d.rows, f=state.rs, q=f.q.toLowerCase();
  return rows.filter(o=>(!f.ref||o.refs.includes(f.ref))&&(!f.owner||o.owner===f.owner)&&(!q||`${o.contract} ${o.ope} ${o.name} ${o.owner} ${o.refs.join(' ')}`.toLowerCase().includes(q)));
}
const pct1=(a,b)=>b?Math.round(a/b*1000)/10:0;
function renderResearch(){
  const {grids,rows}=researchData(), box=$('research'), f=state.rs;
  if(!grids.length||!rows.length){ ['rsGrid','rsRef','rsOwner'].forEach(id=>$(id).innerHTML=''); box.innerHTML='<p class="muted">Aucun résultat de recherche dans les fichiers ExtracTerre chargés. Déposez les fichiers de session (« Exporter la session ») : ils contiennent les grilles et les résultats de l’onglet Recherches.</p>'; return; }
  if(!grids.some(g=>g.id===f.grid)) f.grid=grids[0].id; const g=grids.find(x=>x.id===f.grid);
  const opt=(v,l,sel)=>`<option value="${esc(v)}"${v===sel?' selected':''}>${esc(l)}</option>`, cnt=(list,k)=>list.filter(k).length;
  const refs=[...new Set(rows.flatMap(o=>o.refs))].sort((a,b)=>(a===NO_REF)-(b===NO_REF)||a.localeCompare(b,'fr'));
  if(f.ref&&!refs.includes(f.ref)) f.ref='';
  const owners=[...new Set(rows.map(o=>o.owner))].sort((a,b)=>a.localeCompare(b,'fr')); if(f.owner&&!owners.includes(f.owner)) f.owner='';
  $('rsGrid').innerHTML=grids.map(x=>opt(x.id,`Grille : ${x.name||x.id}`,f.grid)).join('');
  $('rsRef').innerHTML=opt('',`Tous les référentiels (${refs.length})`,f.ref)+refs.map(r=>opt(r,`${r} (${cnt(rows,o=>o.refs.includes(r)&&o.searches[g.id])})`,f.ref)).join('');
  $('rsOwner').innerHTML=opt('',`Tous les maîtres d’ouvrage (${owners.length})`,f.owner)+owners.map(o=>opt(o,o,f.owner)).join('');
  const st=researchStats(researchSelected(),g), N=st.rows.length, n=v=>v.toLocaleString('fr-FR');
  { const pr=verificationProgress(verificationItems(researchSelected(true),g,state.verif)); $('nVerif').textContent=pr.todo?`(${pr.todo.toLocaleString('fr-FR')})`:''; }
  const notSf=st.rows.filter(o=>!o.inSf).length;
  box.innerHTML=`<div class="rs-kpis"><div class="rs-kpi"><b>${n(N)}</b><span>opération(s) analysée(s) avec « ${esc(g.name||g.id)} »${notSf?` · ${notSf} hors Tous contrats`:''}</span></div>
    <div class="rs-kpi"><b>${n(new Set(st.rows.map(o=>o.owner)).size)}</b><span>maître(s) d’ouvrage</span></div>
    ${(()=>{ const pr=verificationProgress(verificationItems(researchSelected(true),g,state.verif)); return pr.total?`<div class="rs-kpi"><b>${n(pr.ok+pr.ko+pr.pa+pr.so)} / ${n(pr.total)}</b><span>« Oui » vérifiés (${n(pr.ok)} acceptés · ${n(pr.pa)} partiels · ${n(pr.ko)} rejetés · ${n(pr.so)} sans objet)</span></div>`:''; })()}
    ${st.themes.map(x=>`<div class="rs-kpi"><b>${pct1(x.ok,x.applicable)} %</b><span>${esc(x.t.name)} : indices suffisants (${x.ok} / ${x.applicable})</span></div>`).join('')}</div>
  ${N?st.themes.map(x=>`<section class="rs-theme"><h3>${esc(x.t.name)}</h3><p class="muted">${x.applicable} opération(s) concernée(s) · minimum ${Number(x.t.min)||0} critère(s)</p>
    <div class="rs-legend"><span><i style="background:var(--et)"></i>Indice trouvé</span><span><i style="background:var(--rust)"></i>Indice contraire / hors seuil</span></div>
    ${x.crit.map(c=>`<div class="rs-crit"><div>${esc(c.c.label)}${c.min!=null?`<small>valeurs lues : min ${fmt(c.min)} · médiane ${fmt(c.med)} · max ${fmt(c.max)}${c.c.unit?' '+esc(c.c.unit):''}</small>`:''}</div>
      <div class="rs-track" role="img" aria-label="${pct1(c.ok,c.pool)} % trouvés, ${pct1(c.ko,c.pool)} % contraires"><span class="ok" style="width:${pct1(c.ok,c.pool)}%"></span><span class="ko" style="width:${pct1(c.ko,c.pool)}%"></span></div>
      <div class="rs-pct">${pct1(c.ok,c.pool)} %<small>${c.ok} / ${c.pool}${c.ko?` · ${c.ko} contraire(s)`:''}</small></div></div>`).join('')}</section>`).join('')
    +`<div class="scroll"><table class="grid"><thead><tr><th scope="col">Contrat</th><th scope="col">OPE</th><th scope="col">Opération</th><th scope="col">Maître d’ouvrage</th><th scope="col">Référentiel</th>${(g.themes||[]).map(t=>`<th scope="col">${esc(t.name)}</th>`).join('')}<th scope="col">Export</th></tr></thead><tbody>
      ${st.rows.map(o=>{ const s=o.searches[g.id]; return `<tr><td>${esc(o.contract||'—')}</td><td>${esc(o.ope||'—')}</td><td title="${esc(o.name)}">${esc(o.name)}</td><td title="${esc(o.owner)}">${esc(o.owner)}</td><td title="${esc(o.refs.join(' / '))}">${esc(o.refs.join(' / '))}</td>${(g.themes||[]).map(t=>{ const th=s.themes?.[t.name]; return th?.applicable===false?'<td class="rs-na">–</td>':th?.ok?`<td class="rs-ok">✓ ${th.found??''}</td>`:`<td class="rs-ko">✗ ${th?.found??0}</td>`; }).join('')}<td>${esc(String(o.at).slice(0,10))}</td></tr>`; }).join('')}</tbody></table></div>`
    :'<p class="muted">Aucune opération de cette sélection n’a été analysée avec cette grille.</p>'}`;
}
function exportResearchCsv(){
  const {grids}=researchData(), g=grids.find(x=>x.id===state.rs.grid); if(!g){ toast('Aucun résultat de recherche à exporter.','error'); return; }
  const rows=researchSelected().filter(o=>o.searches[g.id]);
  const head=['Contrat','OPE','Opération','Maître d’ouvrage','Référentiel',...(g.themes||[]).map(t=>`${t.name} — suffisant`),...(g.criteria||[]).map(c=>`${c.theme} — ${c.label}`),'Date de l’export ExtracTerre'];
  downloadCsv(`Recherches_${String(g.name||g.id).replace(/[^\w-]+/g,'_')}`,head,rows.map(o=>{ const s=o.searches[g.id];
    return [o.contract,o.ope,o.name,o.owner,o.refs.join(' / '),...(g.themes||[]).map(t=>{ const th=s.themes?.[t.name]; return th?.applicable===false?'Non applicable':th?.ok?'Oui':'Non'; }),...(g.criteria||[]).map(c=>s.criteria?.[c.id]?.value??''),String(o.at).slice(0,10)]; }));
  toast(`${rows.length} opération(s) exportée(s).`,'success');
}

/* ---------- Vérification des « Oui » : tableau ou cartes à glisser (accepter / rejeter) ---------- */
const vf={view:'table',crit:'',status:'todo',limit:150,idx:0,history:[]};
function verifData(){ const {grids}=researchData(), g=grids.find(x=>x.id===state.rs.grid); if(!g) return {g:null,all:[],list:[]};
  const all=verificationItems(researchSelected(true),g,state.verif);
  const list=all.filter(i=>(!vf.crit||i.crit.id===vf.crit)&&(vf.status==='all'||(vf.status==='todo'?!i.decision:i.decision===vf.status)));
  return {g,all,list}; }
// Texte de la preuve avec les passages qui ont déclenché le critère (mots-clés) et les mots de contexte surlignés.
const markHtml=(text,crit)=>{ const t=String(text||''), rs=highlightRanges(t,crit); let out='',k=0; for(const r of rs){ out+=esc(t.slice(k,r.s))+`<mark class="hl-${r.kind}">${esc(t.slice(r.s,r.e))}</mark>`; k=r.e; } return out+esc(t.slice(k)); };
const evHtml=(e,crit)=>`<div class="ev"><div class="ev-src">${esc(e.doc||'Document non précisé')}${e.page!==''&&e.page!=null?` · p. ${esc(String(e.page))}`:''}</div><div class="ev-line">${markHtml(e.line,crit)}</div>${e.ctx&&e.ctx.trim()!==String(e.line||'').trim()?`<div class="ev-ctx">${markHtml(e.ctx,crit)}</div>`:''}</div>`;
const reqHtml=i=>`<div class="req-label">${esc(i.crit.label)}</div><div class="req-meta">${esc(i.crit.theme||'')} · réponse : <b>${i.status==='partiel'?'Partiel':'Oui'}</b>${i.manual?' (saisie manuelle dans ExtracTerre)':''}${i.value&&!/^Oui$/i.test(String(i.value))?` · ${esc(String(i.value))}`:''}</div>${i.crit.description?`<div class="req-desc">${esc(i.crit.description)}</div>`:''}`;
const opHtml=o=>`<b>${esc(o.contract||o.ope||'—')}</b> ${esc(o.name)}<small>${esc(o.owner)} · ${esc(o.refs.join(' / '))}</small>`;
function verifProgress(all){ const p=verificationProgress(all), pc=p.total?Math.round((p.ok+p.ko+p.pa+p.so)/p.total*100):0;
  $('verifProgress').innerHTML=`<div class="vp-bar" role="progressbar" aria-label="Avancement de la vérification" aria-valuenow="${pc}" aria-valuemin="0" aria-valuemax="100"><span class="ok" style="width:${p.total?p.ok/p.total*100:0}%"></span><span class="pa" style="width:${p.total?p.pa/p.total*100:0}%"></span><span class="ko" style="width:${p.total?p.ko/p.total*100:0}%"></span><span class="so" style="width:${p.total?p.so/p.total*100:0}%"></span></div><small>${(p.ok+p.ko+p.pa+p.so).toLocaleString('fr-FR')} / ${p.total.toLocaleString('fr-FR')} vérifiés · ✓ ${p.ok.toLocaleString('fr-FR')} · ◐ ${p.pa.toLocaleString('fr-FR')} · ✗ ${p.ko.toLocaleString('fr-FR')} · ⊘ ${p.so.toLocaleString('fr-FR')}</small>`; }
function setDecision(key,d){ const prev=state.verif[key]?.d||''; if(d===prev||!d) delete state.verif[key]; else state.verif[key]={d,at:new Date().toISOString()}; saveVerif(); return prev; }
function renderVerif(){
  const {g,all,list}=verifData(), body=$('verifBody');
  if(!g){ body.innerHTML='<p class="muted">Aucune grille de recherche chargée.</p>'; return; }
  $('verifSub').textContent=`${g.name||g.id} · ${researchSelected(true).filter(o=>o.searches[g.id]).length.toLocaleString('fr-FR')} opération(s) de la sélection de l’onglet Recherches`;
  const crits=(g.criteria||[]).filter(c=>all.some(i=>i.crit.id===c.id)), cnt=id=>all.filter(i=>i.crit.id===id&&!i.decision).length;
  if(vf.crit&&!crits.some(c=>c.id===vf.crit)) vf.crit='';
  $('vfCrit').innerHTML=`<option value="">Tous les critères (${all.filter(i=>!i.decision).length} à vérifier)</option>`+crits.map(c=>`<option value="${esc(c.id)}"${c.id===vf.crit?' selected':''}>${esc(c.theme)} — ${esc(c.label)} (${cnt(c.id)})</option>`).join('');
  $('vfStatus').value=vf.status; document.querySelectorAll('#verifDlg [data-view]').forEach(b=>b.setAttribute('aria-checked',String(b.dataset.view===vf.view)));
  verifProgress(all);
  if(!list.length){ body.innerHTML=`<p class="verif-empty">${vf.status==='todo'?'Tout est vérifié pour cette sélection ✓':'Aucune réponse pour ce filtre.'}</p>`; return; }
  if(vf.view==='cards') return renderVerifCard(list);
  const rows=list.slice(0,vf.limit);
  body.innerHTML=`<div class="scroll"><table class="grid verif-table"><thead><tr><th scope="col">Opération</th><th scope="col">Exigence</th><th scope="col">Vérification (preuve et source)</th><th scope="col">Décision</th></tr></thead><tbody>
    ${rows.map(i=>`<tr data-key="${esc(i.key)}" class="${i.decision?'d-'+i.decision:''}"><td class="v-op">${opHtml(i.op)}</td><td class="v-req">${reqHtml(i)}</td><td class="v-ev">${i.evidence.map(e=>evHtml(e,i.crit)).join('')||'<span class="muted">Aucune preuve exportée</span>'}</td>
      <td class="v-dec">${i.inherited?`<small class="v-inh">✓ validé d’office (même exigence : ${esc(i.inherited)})</small>`:''}<button type="button" class="vbtn ok${i.decision==='ok'?' on':''}" data-d="ok" aria-pressed="${i.decision==='ok'}">✓ Accepter</button><button type="button" class="vbtn pa${i.decision==='pa'?' on':''}" data-d="pa" aria-pressed="${i.decision==='pa'}">◐ Partiel</button><button type="button" class="vbtn ko${i.decision==='ko'?' on':''}" data-d="ko" aria-pressed="${i.decision==='ko'}">✗ Rejeter</button><button type="button" class="vbtn so${i.decision==='so'?' on':''}" data-d="so" aria-pressed="${i.decision==='so'}" title="La phrase n’a pas de lien avec l’exigence">⊘ Sans objet</button></td></tr>`).join('')}</tbody></table></div>
    ${list.length>rows.length?`<p class="verif-more"><button class="btn quiet small" type="button" id="vfMore">Afficher ${Math.min(150,list.length-rows.length)} de plus (${(list.length-rows.length).toLocaleString('fr-FR')} restantes)</button></p>`:''}`;
}
function renderVerifCard(list){
  if(vf.idx>=list.length) vf.idx=Math.max(0,list.length-1); const i=list[vf.idx], e=i.evidence[0], more=i.evidence.slice(1);
  const q=String(i.crit.label||'').replace(/\s*\(condition\)\s*$/i,'').replace(/[.\s]+$/,'');
  $('verifBody').innerHTML=`<div class="vcard-wrap"><p class="vcard-count">${vf.idx+1} / ${list.length.toLocaleString('fr-FR')} à vérifier</p>
    <article class="vcard ${i.decision?'d-'+i.decision:''}" id="vcard" data-key="${esc(i.key)}" tabindex="0" aria-label="Carte à glisser : à droite pour Oui, à gauche pour Non">
      <div class="vstamp ok" aria-hidden="true">OUI</div><div class="vstamp ko" aria-hidden="true">NON</div><div class="vstamp pa" aria-hidden="true">PARTIEL</div><div class="vstamp so" aria-hidden="true">SANS OBJET</div>
      <p class="vc-op">${esc(i.op.contract||i.op.ope||'—')} · ${esc(i.op.name)}</p>
      <h3 class="vc-q">${esc(q)} ?</h3>
      ${i.status==='partiel'?'<p class="vc-partial">Réponse ExtracTerre : Partiel</p>':''}
      ${e?`<blockquote class="vc-quote">${markHtml(e.line,i.crit)}</blockquote>
      <p class="vc-src">${esc(e.doc||'Document non précisé')}${e.page!==''&&e.page!=null?` · page ${esc(String(e.page))}`:''}</p>`:'<p class="vc-quote vc-none">Aucune preuve exportée</p>'}
      ${(e&&e.ctx&&e.ctx.trim()!==String(e.line||'').trim())||more.length||i.crit.description?`<details class="vc-more"><summary>Contexte${more.length?` · ${more.length} autre(s) preuve(s)`:''}${i.crit.description?' · précision de l’exigence':''}</summary>
        ${i.crit.description?`<p class="vc-desc">${esc(i.crit.description)}</p>`:''}${e&&e.ctx?`<p class="vc-ctx">${markHtml(e.ctx,i.crit)}</p>`:''}${more.map(x=>evHtml(x,i.crit)).join('')}</details>`:''}
    </article>
    <div class="vcard-actions"><button type="button" class="vbig ko" data-card="ko">✗ Non</button><button type="button" class="vbig pa" data-card="pa">◐ Partiel</button><button type="button" class="vbig ok" data-card="ok">✓ Oui</button></div>
    <p class="vcard-so"><button type="button" class="vbig so" data-card="so" title="La phrase n’a pas de lien avec l’exigence">⊘ Sans objet</button></p>
    <p class="vcard-undo"><button type="button" class="btn quiet small" data-card="undo"${vf.history.length?'':' disabled'}>↶ Annuler la dernière décision</button></p>
    <p class="muted small vcard-help">Glisser à droite = Oui · à gauche = Non · vers le haut = Partiel · vers le bas = Sans objet · clavier → ← ↑ ↓ · Retour arrière = annuler</p></div>`;
  const card=$('vcard'); let x0=null,y0=0,dx=0,dy=0; card.focus({preventScroll:true});
  card.addEventListener('pointerdown',e=>{ if(e.target.closest('a,button,summary,details')) return; x0=e.clientX; y0=e.clientY; dx=0; dy=0; card.setPointerCapture(e.pointerId); card.classList.add('drag'); });
  card.addEventListener('pointermove',e=>{ if(x0===null) return; dx=e.clientX-x0; dy=e.clientY-y0; const vert=Math.abs(dy)>40&&Math.abs(dy)>Math.abs(dx); card.style.transform=vert?`translateY(${dy}px)`:`translateX(${dx}px) rotate(${dx/22}deg)`; card.dataset.swipe=vert?(dy<0?'pa':'so'):dx>40?'ok':dx<-40?'ko':''; });
  const end=()=>{ if(x0===null) return; x0=null; card.classList.remove('drag'); if(dy<-110&&Math.abs(dy)>Math.abs(dx)) cardDecide('pa'); else if(dy>110&&Math.abs(dy)>Math.abs(dx)) cardDecide('so'); else if(Math.abs(dx)>110) cardDecide(dx>0?'ok':'ko'); else { card.style.transform=''; card.dataset.swipe=''; } };
  card.addEventListener('pointerup',end); card.addEventListener('pointercancel',end);
}
function cardDecide(d){ const {list}=verifData(), i=list[vf.idx]; if(!i) return; const prev=state.verif[i.key]?.d||''; if(prev!==d) setDecision(i.key,d); vf.history.push({key:i.key,prev});
  const card=$('vcard'); if(card){ card.style.transition='transform .22s ease-out,opacity .22s'; card.style.transform=d==='pa'?'translateY(-500px)':d==='so'?'translateY(500px)':`translateX(${d==='ok'?600:-600}px) rotate(${d==='ok'?18:-18}deg)`; card.style.opacity='0'; }
  setTimeout(()=>{ if(vf.status!=='todo') vf.idx++; renderVerif(); },200); }
function cardUndo(){ const h=vf.history.pop(); if(!h) return; if(h.prev) state.verif[h.key]={d:h.prev,at:new Date().toISOString()}; else delete state.verif[h.key]; saveVerif();
  const {list}=verifData(), k=list.findIndex(i=>i.key===h.key); if(k>=0) vf.idx=k; renderVerif(); }
function openVerif(){ vf.idx=0; vf.history=[]; vf.limit=150; renderVerif(); $('verifDlg').showModal(); }
$('rsVerif').addEventListener('click',openVerif);
$('verifClose').addEventListener('click',()=>$('verifDlg').close());
$('verifDlg').addEventListener('close',()=>renderResearch());
$('vfCrit').addEventListener('change',e=>{ vf.crit=e.target.value; vf.idx=0; vf.limit=150; renderVerif(); });
$('vfStatus').addEventListener('change',e=>{ vf.status=e.target.value; vf.idx=0; vf.limit=150; renderVerif(); });
document.querySelectorAll('#verifDlg [data-view]').forEach(b=>b.addEventListener('click',()=>{ vf.view=b.dataset.view; vf.idx=0; renderVerif(); }));
$('verifBody').addEventListener('click',e=>{
  if(e.target.closest('#vfMore')){ vf.limit+=150; renderVerif(); return; }
  const c=e.target.closest('[data-card]'); if(c){ if(c.dataset.card==='undo') cardUndo(); else cardDecide(c.dataset.card); return; }
  const b=e.target.closest('.vbtn'); if(!b) return; const tr=b.closest('tr'), key=tr.dataset.key; setDecision(key,b.dataset.d); const d=state.verif[key]?.d||'';
  tr.className=d?'d-'+d:''; tr.querySelectorAll('.vbtn').forEach(x=>{ const on=x.dataset.d===d; x.classList.toggle('on',on); x.setAttribute('aria-pressed',String(on)); }); verifProgress(verifData().all); });
$('verifDlg').addEventListener('keydown',e=>{ if(vf.view!=='cards'||e.target.closest('select,input')) return;
  if(e.key==='ArrowRight'){ e.preventDefault(); cardDecide('ok'); } else if(e.key==='ArrowLeft'){ e.preventDefault(); cardDecide('ko'); } else if(e.key==='ArrowUp'){ e.preventDefault(); cardDecide('pa'); } else if(e.key==='ArrowDown'){ e.preventDefault(); cardDecide('so'); } else if(e.key==='Backspace'){ e.preventDefault(); cardUndo(); } });
$('vfReset').addEventListener('click',()=>{ const n=Object.keys(state.verif).length; if(!n){ toast('Aucune décision à effacer.','info'); return; }
  if($('vfReset').dataset.armed!=='1'){ $('vfReset').dataset.armed='1'; $('vfReset').textContent=`Confirmer : effacer ${n} décision(s)`; setTimeout(()=>{ $('vfReset').dataset.armed=''; $('vfReset').textContent='Tout effacer'; },5000); return; }
  state.verif={}; saveVerif(); $('vfReset').dataset.armed=''; $('vfReset').textContent='Tout effacer'; vf.history=[]; vf.idx=0; renderVerif(); toast(`${n} décision(s) effacée(s).`,'success'); });
$('vfExport').addEventListener('click',()=>{ const blob=new Blob([JSON.stringify({format:'totalrecoll-verifications',version:1,exportedAt:new Date().toISOString(),decisions:state.verif},null,1)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=`TotalRecoll_verifications_${new Date().toISOString().slice(0,10)}.json`; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),2000);
  toast(`${Object.keys(state.verif).length} décision(s) exportée(s).`,'success'); });
$('vfImport').addEventListener('change',async e=>{ const f=e.target.files?.[0]; if(!f) return;
  try{ const j=JSON.parse(await f.text()); if(j?.format!=='totalrecoll-verifications'||typeof j.decisions!=='object') throw new Error('Ce fichier n’est pas un export de vérifications TotalRecoll.');
    let n=0; for(const [k,v] of Object.entries(j.decisions)){ if(!v||!['ok','ko','pa','so'].includes(v.d)) continue; const cur=state.verif[k]; if(!cur||String(v.at||'')>=String(cur.at||'')){ state.verif[k]=v; n++; } }
    saveVerif(); renderVerif(); toast(`${n} décision(s) reprise(s).`,'success'); }catch(err){ toast(err.message||String(err),'error'); }
  e.target.value=''; });

/* ---------- Confort d'été : valeurs normalisées par bâtiment ---------- */
const CONFORT_HEAD=['Contrat','OPE',"Nom de l'opération",'Bâtiment','DH zones traversantes','DH zones non traversantes','Système constructif','Brasseur d’air','Bâtiment climatisé','À vérifier'];
function confortRows(){
  const r=state.result; if(!r) return [];
  return r.rows.map((row,i)=>{ const get=h=>{ const j=r.header.indexOf(h); return j<0?undefined:row[j]; }; const c=confortOf(get);
    return {i,cells:[get(CONTRACT_COL),get(OPE_COL),get("Nom de l'opération"),get(BUILDING_COL),c.dhZt,c.dhZnt,c.system,c.fan,c.clim,c.note]}; })
    .filter(x=>x.cells.slice(4).some(v=>!empty(v)));
}
/* ---------- Confort d'été : carte des zones climatiques ---------- */
const ZONE_COLORS={H1a:'#1F5FA8',H1b:'#3B88D0',H1c:'#78B8E8',H2a:'#3FA586',H2b:'#B9A33F',H2c:'#E0912F',H2d:'#E5672A',H3:'#C1440E'};
// Point dans un département (anneaux [lat,lon], pair-impair) : dernier recours quand ni code postal ni nom d'opération ne donnent le département.
const inRings=(la,lo,polys)=>polys.some(rings=>rings.reduce((inside,ring)=>{ let c=false; for(let i=0,j=ring.length-1;i<ring.length;j=i++){ const [yi,xi]=ring[i],[yj,xj]=ring[j]; if((yi>la)!==(yj>la)&&lo<(xj-xi)*(la-yi)/(yj-yi)+xi) c=!c; } return c?!inside:inside; },false));
function zoneItems(){
  const r=state.result; if(!r) return [];
  const col=h=>r.header.indexOf(h), lat=col('Latitude'), lon=col('Longitude');
  return r.rows.map((row,i)=>{ const get=h=>{ const j=col(h); return j<0?undefined:row[j]; }, c=confortOf(get);
    let dep=depOfRow(get), how=dep?'':'non trouvé';
    if(!dep&&state.zonesGeo&&lat>=0){ const la=Number(row[lat]), lo=Number(row[lon]); if(la&&lo){ const d=state.zonesGeo.departements.find(d=>inRings(la,lo,d.rings)); if(d){ dep=d.code; how='géolocalisation'; } } }
    const num=v=>empty(v)?null:Number(v);
    return {i,op:r.meta[i].group||get(CONTRACT_COL)||'ligne'+i,dep,zone:zoneOfDep(dep),how,dhZt:num(c.dhZt),dhZnt:num(c.dhZnt),system:c.system,fan:c.fan,clim:c.clim}; })
    .filter(x=>x.dhZt!=null||x.dhZnt!=null||x.system||x.fan||x.clim);
}
const fr=v=>v==null?'—':v.toLocaleString('fr-FR',{maximumFractionDigits:1});
function renderZoneCard(st){
  const z=state.zone, s=z?st.zones[z]:st.all;
  const line=(l,x)=>`<dt>${l}</dt><dd>${fr(x.mean)}</dd><dd>${fr(x.median)} <small>(${x.n})</small></dd>`;
  const pct=(l,x,base,why)=>`<dt>${l}</dt><dd>${x.pct==null?'—':fr(x.pct)+' %'}</dd><dd><small title="${why}">${x.yes} / ${base} op.</small></dd>`;
  $('zoneCard').innerHTML=`<h3>${z?'Zone '+esc(z):'Toutes zones'}</h3><div class="sub">${s.ops} opération(s), ${s.buildings} bâtiment(s)${z?'':` — cliquer sur une zone de la carte`}</div>
    <dl><dt class="hd">DH (°C.h)</dt><dd class="hd">Moyenne</dd><dd class="hd">Médiane (n)</dd>${line('Zones traversantes',s.zt)}${line('Zones non traversantes',s.znt)}</dl>
    <dl><dt class="hd">Opérations équipées</dt><dd class="hd">Part</dd><dd class="hd">Oui / opérations</dd>${pct('Brasseurs d’air',s.fan,s.fan.total,'Sur toutes les opérations de la zone : sans mention de brasseur = Non.')}${pct('Climatisation active',s.clim,s.clim.total,'Sur toutes les opérations de la zone : sans mention de climatisation = Non.')}</dl>
    <div class="sub">Parts calculées sur les ${s.ops} opération(s) de la zone : une opération sans mention compte comme « Non ».</div>
    ${z?'<button class="btn quiet small" id="zoneAll" type="button">Voir toutes les zones</button>':''}`;
  $('zoneSystems').innerHTML=`<caption>DH selon le système constructif — ${z?'zone '+esc(z):'toutes zones'}</caption><thead><tr><th scope="col">Système constructif</th><th scope="col">Bâtiments</th><th scope="col">DH ZT moyenne</th><th scope="col">DH ZT médiane (n)</th><th scope="col">DH ZNT moyenne</th><th scope="col">DH ZNT médiane (n)</th></tr></thead><tbody>${s.systems.map(x=>`<tr><th scope="row">${esc(x.system)}</th><td>${x.buildings}</td><td>${fr(x.zt.mean)}</td><td>${fr(x.zt.median)} (${x.zt.n})</td><td>${fr(x.znt.mean)}</td><td>${fr(x.znt.median)} (${x.znt.n})</td></tr>`).join('')}</tbody>`;
}
function renderZones(){
  const items=zoneItems(), st=confortStats(items), geo=items.filter(x=>x.how==='géolocalisation').length;
  $('zoneInfo').innerHTML=`${items.length} bâtiment(s) avec une donnée de confort d’été, ${items.length-st.unzoned} placés dans une zone climatique${geo?` (dont ${geo} par la géolocalisation)`:''}${st.unzoned?`, ${st.unzoned} sans département connu (hors carte)`:''}.<span class="zone-legend">${ZONES.map(z=>`<span><i style="background:${ZONE_COLORS[z]}"></i>${z} <b>${st.zones[z].ops}</b></span>`).join('')}</span>`;
  renderZoneCard(st);
  if(typeof L==='undefined'||!state.zonesGeo){ if(!state.zonesGeo) $('zoneInfo').append(' (zones climatiques en cours de chargement)'); return; }
  if(!state.zoneMap){
    state.zoneMap=L.map('zoneMap',{scrollWheelZoom:true,preferCanvas:true,zoomSnap:.25}).setView([46.6,2.4],5.5);
    L.tileLayer('https://data.geopf.fr/wmts?SERVICE=WMTS&REQUEST=GetTile&VERSION=1.0.0&LAYER=GEOGRAPHICALGRIDSYSTEMS.PLANIGNV2&STYLE=normal&TILEMATRIXSET=PM&FORMAT=image/png&TILEMATRIX={z}&TILEROW={y}&TILECOL={x}',{maxZoom:12,opacity:.35,attribution:'© IGN — Géoplateforme, Admin Express'}).addTo(state.zoneMap);
    state.zoneLayers={};
    for(const d of state.zonesGeo.departements){ const p=L.polygon(d.rings,{color:'#170A07',weight:.6,fillColor:ZONE_COLORS[d.zone],fillOpacity:.6});
      p.on('click',()=>{ state.zone=state.zone===d.zone?'':d.zone; renderZones(); });
      p.bindTooltip(`${esc(d.nom)} (${esc(d.code)}) — zone ${esc(d.zone)}`,{sticky:true}); p.addTo(state.zoneMap); (state.zoneLayers[d.zone]??=[]).push(p); }
  }
  setTimeout(()=>state.zoneMap.invalidateSize(),0);
  for(const [z,ps] of Object.entries(state.zoneLayers)) ps.forEach(p=>p.setStyle({fillOpacity:!state.zone||state.zone===z?.6:.15,weight:state.zone===z?1.2:.6}));
  state.zoneMarks?.remove();
  state.zoneMarks=L.featureGroup(ZONES.filter(z=>st.zones[z].ops).map(z=>{ const b=L.featureGroup(state.zoneLayers[z]).getBounds().getCenter(), s=st.zones[z];
    return L.marker(b,{icon:L.divIcon({className:'',html:''}),interactive:false}).bindTooltip(`${z} · ${s.ops} op.<br>DH ZT ${fr(s.zt.median)} · ZNT ${fr(s.znt.median)}`,{permanent:true,direction:'center',className:'zone-label'}); })).addTo(state.zoneMap);
}

function renderConfort(){
  renderZones();
  const r=state.result, rows=confortRows(), n=k=>rows.filter(x=>!empty(x.cells[k])).length;
  $('confortInfo').textContent=rows.length?`${rows.length} bâtiment(s) avec au moins une valeur sur ${r.rows.length} ligne(s) : DH ZT ${n(4)}, DH ZNT ${n(5)}, système constructif ${n(6)}, brasseur d’air ${n(7)}, climatisation ${n(8)}, à vérifier ${n(9)}.`:'Aucune de ces valeurs dans les données recollées.';
  const dhCol={4:DH_ZT,5:DH_ZNT};
  $('confort').innerHTML=rows.length?`<thead><tr>${CONFORT_HEAD.map(h=>`<th scope="col">${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(x=>`<tr tabindex="0" data-row="${x.i}">${x.cells.map((v,k)=>{ const ec=dhCol[k]?ecartTip(x.i,dhCol[k]):''; return `<td${ec?' class="sfdiff"':''} title="${esc([fmt(v),ec].filter(Boolean).join('\n'))}">${esc(fmt(v))}</td>`; }).join('')}</tr>`).join('')}</tbody>`:'';
}
function exportConfortCsv(){
  const rows=confortRows(); if(!rows.length){ toast('Aucune valeur de confort d’été à exporter.','error'); return; }
  downloadCsv('Confort_ete',CONFORT_HEAD,rows.map(x=>x.cells)); toast(`${rows.length} bâtiment(s) exporté(s).`,'success');
}

function renderAnru(){
  const r=state.result, all=anruRows(), qi=r.header.indexOf(ANRU_QP_COL), ai=r.header.indexOf(ANRU_COL);
  if(!state.anru){ $('anruInfo').textContent='Quartiers ANRU en cours de chargement…'; return; }
  const counts=new Map(); all.forEach(p=>p.mentions.forEach(m=>counts.set(m,(counts.get(m)||0)+1)));
  const list=[...counts].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0],'fr'));
  $('mentions').innerHTML=list.length?list.map(([m,n])=>`<label><input type="checkbox" data-mention="${esc(m)}" ${state.mentionsOff.has(m)?'':'checked'}><span>${esc(m)}</span><b>${n}</b></label>`).join(''):'<p class="muted">Aucun bâtiment en quartier ANRU.</p>';
  const pts=anruShown(all);
  const ops=new Set(pts.map(p=>r.meta[p.i].group||p.i)).size, unsure=r.rows.filter(row=>row[ai]===''&&row[qi]).length, edge=all.filter(p=>/du périmètre$/.test(r.rows[p.i][qi]||'')).length;
  $('anruInfo').innerHTML=`${pts.length} bâtiment(s) en quartier ANRU affiché(s), ${ops} opération(s)${pts.length<all.length?` sur ${all.length} bâtiment(s)`:''}${edge?`, dont ${edge} à 10 m au plus d’un périmètre`:''}.${unsure?` ${unsure} ligne(s) non géolocalisée(s) dans une commune qui a un quartier NPNRU : colonne ANRU laissée vide, à vérifier.`:''}<span class="anru-legend"><span><i style="border:2px solid #C1440E;background:rgba(193,68,14,.2)"></i>Quartier d’intérêt national</span><span><i style="border:2px dashed #2563A8"></i>Quartier d’intérêt régional</span><span><i style="background:#C1440E;border-radius:50%"></i><i style="background:#2563A8;border-radius:50%"></i>Bâtiment (couleur de son quartier)</span></span>`;
  if(typeof L==='undefined'){ $('anruInfo').append(' (fond de carte indisponible hors connexion)'); return; }
  if(!state.anruMap){
    state.anruMap=L.map('anruMap',{scrollWheelZoom:true,preferCanvas:true}).setView([46.6,2.4],6);
    L.tileLayer('https://data.geopf.fr/wmts?SERVICE=WMTS&REQUEST=GetTile&VERSION=1.0.0&LAYER=GEOGRAPHICALGRIDSYSTEMS.PLANIGNV2&STYLE=normal&TILEMATRIXSET=PM&FORMAT=image/png&TILEMATRIX={z}&TILEROW={y}&TILECOL={x}',{maxZoom:19,attribution:'© IGN — Géoplateforme ; quartiers NPNRU : ANCT / ANRU'}).addTo(state.anruMap);
    L.featureGroup(state.anru.quartiers.map(q=>L.polygon(q.rings,q.cat==='national'?{color:'#C1440E',weight:2,fillOpacity:.2}:{color:'#2563A8',weight:2,dashArray:'5 4',fillOpacity:.12}).bindPopup(`<b>${esc(q.nom)}</b><br>${esc(q.qp)}, intérêt ${esc(q.cat)}<br>${esc(q.communes.map(c=>c[1]).join(', '))}`))).addTo(state.anruMap);
  }
  setTimeout(()=>state.anruMap.invalidateSize(),0);
  state.anruLayer?.remove();
  state.anruLayer=L.featureGroup(pts.map(p=>{ const x=rowLabel(p.i);
    const nat=/intérêt national/.test(r.rows[p.i][qi]||'');
    return L.circleMarker([p.lat,p.lon],{radius:7,color:'#1A0B08',weight:2,fillColor:nat?'#C1440E':'#2563A8',fillOpacity:.95}).bindPopup(`<b>${esc(x.title)}</b><br>${esc(x.sub)}<br>${esc(r.rows[p.i][qi])}<br>${esc(p.mentions.join(', '))}<br><button class="btn small" data-open="${p.i}" type="button">Détail</button>`); })).addTo(state.anruMap);
  if(pts.length) state.anruMap.fitBounds(state.anruLayer.getBounds().pad(.2),{maxZoom:15});
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

// Observatoire (onglet qui a ouvert TotalRecoll, même origine) : seules les valeurs écrites par ExtracTerre
// partent, d'onglet à onglet ; l'Observatoire ne remplit que les cellules vides de son Sheet. Rien n'est stocké.
function observatoireWindow(){ try{ const w=window.opener; return w&&!w.closed&&w.location.origin===location.origin?w:null; }catch{ return null; } }
// v2.34 — seuil de fiabilité des valeurs envoyées (choix mémorisé ; message v1 gardé : l'Observatoire ignore « rel ») : par défaut « Probable » (lecture mesurée juste ≥ 85 %
// face à Salesforce) ; « Toutes » envoie aussi les valeurs à vérifier et celles des exports antérieurs (fiabilité inconnue).
function minRel(){ try{ return localStorage.getItem('totalrecoll.minRel')??'probable'; }catch{ return 'probable'; } }
function observatoirePayload(){
  const r=state.result, ci=r.header.indexOf(CONTRACT_COL), oi=r.header.indexOf(OPE_COL), bi=r.header.indexOf(BUILDING_COL);
  const buildings=[]; let skipped=0;
  r.meta.forEach((m,i)=>{
    if(m.kind!=='recolle'&&m.kind!=='et'||!m.et?.length) return;
    const row=r.rows[i], cells={};
    const rel={}, min=REL_ORDER[minRel()]||0;
    m.et.forEach(j=>{ if(j===ci||j===oi||j===bi||empty(row[j])) return; const k=m.rel?.[j]||''; if((REL_ORDER[k]||0)<min){ skipped++; return; } const v=row[j]; cells[r.header[j]]=typeof v==='number'||typeof v==='boolean'?v:String(v); if(k) rel[r.header[j]]=k; });
    if(Object.keys(cells).length) buildings.push({contract:row[ci]??'',ope:row[oi]??'',building:bi>=0?String(row[bi]??''):'',cells,rel});
  });
  return {type:'totalrecoll:operations',v:1,sentAt:new Date().toISOString(),minReliability:minRel(),skipped,buildings};
}
function sendToObservatoire(){
  const w=observatoireWindow();
  if(!w){ toast('Ouvre TotalRecoll depuis l’Observatoire (fenêtre Données › Complément TotalRecoll) pour pouvoir l’envoyer.','error'); return; }
  const p=observatoirePayload();
  if(!p.buildings.length){ toast(p.skipped?`Aucune valeur au niveau de fiabilité choisi : ${p.skipped.toLocaleString('fr-FR')} valeur(s) écartée(s). Choisissez « toutes » pour les envoyer quand même.`:'Aucune donnée ExtracTerre recollée à envoyer.','error'); return; }
  w.postMessage(p,location.origin);
  toast(`${p.buildings.length.toLocaleString('fr-FR')} bâtiments envoyés à l’Observatoire${p.skipped?` · ${p.skipped.toLocaleString('fr-FR')} valeur(s) écartée(s) (fiabilité insuffisante)`:''}…`);
}
window.addEventListener('message',e=>{
  if(e.origin!==location.origin||e.source!==observatoireWindow()||e.data?.type!=='newosb:recoll:ack') return;
  const s=e.data.stats||{}, n=v=>Number(v||0).toLocaleString('fr-FR');
  if(!e.data.ok){ toast('Observatoire : '+(e.data.error||'complément refusé'),'error'); return; }
  const more=[s.addedRows?`${n(s.addedRows)} bâtiments ajoutés aux opérations`:'',s.correctedCells?`${n(s.correctedCells)} valeurs illisibles du Sheet corrigées`:''].filter(Boolean).join(' · ');
  toast(`Observatoire complété : ${n(s.filledCells)} cellules remplies sur ${n(s.rowsCompleted)} lignes · ${n(s.ecarts)} écarts (Sheet conservé) · ${n(s.unmatchedOperations)} opérations absentes du Sheet${more?' · '+more:''}.`,'success');
});

/* ---------- Publication dans le Google Sheet de l'Observatoire (script OPERATIONS ≥ 06.47) ---------- */
// Le tableau recollé remplace l'onglet OPERATIONS (le script le sauvegarde d'abord). Mêmes règles de fiabilité que
// « Envoyer à l'Observatoire » : une valeur ExtracTerre sous le niveau choisi n'est pas publiée (la valeur Salesforce
// d'origine est remise si ExtracTerre l'avait remplacée). La clé reste en mémoire de l'onglet, jamais enregistrée.
const OBS_SOURCE_KEY='prestaterre-v29-connected-source', PUB_CHUNK=300;
let pubBridge=null, pubPopup=false, pubStage='';
function openPublish(){
  const r=state.result; if(!r){ toast('Charge d’abord Tous contrats.','error'); return; }
  let url=''; try{ url=JSON.parse(localStorage.getItem(OBS_SOURCE_KEY)||'null')?.url||''; }catch{}
  if(!$('pubUrl').value) $('pubUrl').value=url;
  const t=publishTable(r,minRel()), ops=new Set(r.rows.map(x=>x[r.header.indexOf(CONTRACT_COL)]||x[r.header.indexOf(OPE_COL)])).size;
  $('pubSummary').textContent=`${t.rows.length.toLocaleString('fr-FR')} lignes (${ops.toLocaleString('fr-FR')} opérations) × ${t.header.length} colonnes. Valeurs ExtracTerre publiées : ${$('relMin').selectedOptions[0]?.textContent||'toutes'}${t.skipped?` (${t.skipped.toLocaleString('fr-FR')} valeur(s) écartée(s))`:''}.`;
  setPub(''); $('pubGo').disabled=false; $('pubDlg').showModal(); ($('pubUrl').value?$('pubKey'):$('pubUrl')).focus();
}
function setPub(msg,kind=''){ const s=$('pubStatus'); s.textContent=msg; s.className=kind; }
async function runPublish(e){
  e.preventDefault();
  const url=$('pubUrl').value.trim(), key=$('pubKey').value.trim();
  if(!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/.test(url)){ setPub('Adresse du script invalide : elle se termine par /exec.','err'); return; }
  if(key.length<16){ setPub('Clé d’accès trop courte.','err'); return; }
  if(!window.NEWOSB_BRIDGE){ setPub('Pont Apps Script non chargé : recharge la page.','err'); return; }
  const t=publishTable(state.result,minRel());
  if(!confirm(`Remplacer l’onglet OPERATIONS par ${t.rows.length.toLocaleString('fr-FR')} lignes ? Son contenu actuel est sauvegardé dans OPERATIONS_SAUVEGARDE.`)) return;
  $('pubGo').disabled=true;
  pubBridge=pubBridge||window.NEWOSB_BRIDGE.create('totalrecoll-publication'); pubBridge.setKey(key);
  try{
    // Fenêtre Google ouverte pendant le clic si l'iframe a déjà échoué (autorisation Google à valider).
    if(pubPopup) pubBridge.preparePopup(url);
    pubStage='begin'; setPub('Connexion au script…');
    const b=await pubBridge.request(url,{mode:'publishbegin',headers:t.header,total:t.rows.length},120000);
    for(let o=0;o<t.rows.length;o+=PUB_CHUNK){
      setPub(`Envoi des lignes ${(o+1).toLocaleString('fr-FR')} à ${Math.min(o+PUB_CHUNK,t.rows.length).toLocaleString('fr-FR')} sur ${t.rows.length.toLocaleString('fr-FR')}…`);
      await pubBridge.request(url,{mode:'publishchunk',id:b.id,offset:o,rows:t.rows.slice(o,o+PUB_CHUNK)},180000);
    }
    pubStage='commit'; setPub('Remplacement de l’onglet OPERATIONS (sauvegarde de l’ancien contenu)…');
    const c=await pubBridge.request(url,{mode:'publishcommit',id:b.id},360000);
    setPub(`Publié : ${Number(c.rows).toLocaleString('fr-FR')} lignes × ${c.columns} colonnes. Ancien contenu (${Number(c.backupRows).toLocaleString('fr-FR')} lignes) dans ${c.backupSheet}. Recharge l’Observatoire pour voir les nouvelles données.`,'ok');
    toast('Observatoire publié.','success'); pubPopup=false;
  }catch(err){
    const auth=err&&err.authError;
    if(!auth&&!pubPopup&&/ne répond pas|indisponible/i.test(String(err?.message||''))){ pubPopup=true; setPub('Le script ne répond pas dans la page : clique à nouveau sur « Publier », une fenêtre Google s’ouvrira (valide l’autorisation si Google la demande).','err'); }
    else setPub((auth?'Accès refusé : ':'Publication interrompue : ')+(err?.message||err)+(auth?'':pubStage==='commit'?' — vérifie l’onglet OPERATIONS : son ancien contenu est dans OPERATIONS_SAUVEGARDE.':' — OPERATIONS n’a pas été modifié.'),'err');
    if(auth){ pubBridge.clearKey(); $('pubKey').value=''; }
  }finally{ $('pubGo').disabled=false; try{ pubBridge.closePopupSoon(800); }catch{} }
}
$('pubBtn').addEventListener('click',openPublish);
$('pubForm').addEventListener('submit',runPublish);
$('pubCancel').addEventListener('click',()=>$('pubDlg').close());

/* ---------- Divers ---------- */
function toast(msg,kind=''){ const t=document.createElement('div'); t.className='toast '+kind; t.textContent=msg; $('toasts').append(t); setTimeout(()=>t.remove(),6000); }
function setTab(tab){
  state.tab=tab;
  document.querySelectorAll('.tab').forEach(b=>b.setAttribute('aria-selected',String(b.dataset.tab===tab)));
  document.querySelectorAll('.panel').forEach(p=>{ p.hidden=p.dataset.panel!==tab; });
  if(tab==='map') renderMap();
  if(tab==='anru') renderAnru();
  if(tab==='confort') renderConfort();
  if(tab==='research') renderResearch();
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
  if(t.matches('[data-sort]')){ e.stopPropagation(); const col=+t.dataset.sort, dir=t.dataset.dir, same=state.sort.col===col&&state.sort.dir===dir; state.sort=same?{col:null,dir:''}:{col,dir}; state.page=0; renderTable(); return; }
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
$('obsBtn').addEventListener('click',sendToObservatoire);
{ const sel=$('relMin'); if(sel){ sel.value=minRel(); sel.addEventListener('change',()=>{ try{ localStorage.setItem('totalrecoll.minRel',sel.value); }catch{} }); } }
$('copyBtn').addEventListener('click',()=>copyForSheets().catch(e=>toast('Copie impossible : '+e.message,'error')));
$('helpBtn').addEventListener('click',()=>$('help').showModal());
$('mentions').addEventListener('change',e=>{ const m=e.target.dataset.mention; if(m==null) return; e.target.checked?state.mentionsOff.delete(m):state.mentionsOff.add(m); renderAnru(); });
$('anruCsv').addEventListener('click',exportAnruCsv);
$('confortCsv').addEventListener('click',exportConfortCsv);
$('rsCsv').addEventListener('click',exportResearchCsv);
['rsGrid','rsRef','rsOwner'].forEach(id=>$(id).addEventListener('change',e=>{ state.rs[{rsGrid:'grid',rsRef:'ref',rsOwner:'owner'}[id]]=e.target.value; renderResearch(); }));
let rsTimer; $('rsQ').addEventListener('input',e=>{ clearTimeout(rsTimer); rsTimer=setTimeout(()=>{ state.rs.q=e.target.value.trim(); renderResearch(); },180); });
$('zoneCard').addEventListener('click',e=>{ if(e.target.id==='zoneAll'){ state.zone=''; renderZones(); } });
$('mentionsAll').addEventListener('click',()=>{ state.mentionsOff.clear(); renderAnru(); });
$('mentionsNone').addEventListener('click',()=>{ anruRows().forEach(p=>p.mentions.forEach(m=>state.mentionsOff.add(m))); renderAnru(); });
fetch('data/zones-climatiques.json').then(r=>r.ok?r.json():Promise.reject(new Error(`HTTP ${r.status}`))).then(d=>{ state.zonesGeo=d; if(state.tab==='confort') renderConfort(); })
  .catch(e=>toast(`Zones climatiques non chargées (${e.message}) : carte du confort d’été indisponible.`,'error'));
fetch('data/npnru.json').then(r=>r.ok?r.json():Promise.reject(new Error(`HTTP ${r.status}`))).then(d=>{ state.anru=anruIndex(d); if(state.sf&&state.et.length) run(); })
  .catch(e=>toast(`Quartiers ANRU non chargés (${e.message}) : la colonne ANRU ne sera pas remplie.`,'error'));
renderSteps();
