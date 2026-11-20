// Recollement Salesforce (tableau « Tous contrats », 1 ligne = 1 bâtiment) × ExtracTerre (données techniques par bâtiment).
import {ET_FIELDS} from './fields.js?v=4';

export const CONTRACT_COL='Numéro du contrat', OPE_COL='Opération: Code interne';
export const STATUS_COL='Recollement', BUILDING_COL='Bâtiment';

// Libellé ExtracTerre → colonne du tableau OPERATIONS quand les noms diffèrent.
export const ALIAS={
  'Année de construction':'Année','Structure':'Mode constructif',
  'Planchers hauts':'Planchers hauts structure','Planchers hauts isolant':'Planchers hauts type isolant',
  'Planchers bas isolant':'Planchers bas type isolant',
  'Menuiseries matériau':'Menuiseries extérieures matériau','Menuiseries vitrage':'Menuiseries extérieures vitrage',
  'Menuiseries occultations':'Menuiseries extérieures occultations',
  'ECS':'ECS après travaux','Refroidissement':'Refroidissement après travaux','Ventilation':'Ventilation après travaux',
  'DH':'DH projet','DH Max':'DH max','Tic':'TIC projet','Tic ref':'TIC ref',
  'Bbio':'Bbio projet','Bbio Max':'Bbio max','Gain Bbio':'Gain BBio',
  'Cep':'Cep projet','Cep Ref':'Cep max','Gain Cep':'Gain CEP',
  'Cepnr':'Cep,nr projet','Cepnr Max':'Cep,nr max','Gain Cepnr':'GAIN CEPnr',
  'Ubat avant travaux':'Ubat initial','Ubat après travaux':'Ubat projet','Cep avant travaux':'Cep initial',
  'IC énergie Max':'IC énergie max','IC construction Max':'IC construction max','IC construction Max 2028':'IC construction seuil 2028',
  'DPE Énergie avant travaux':'DPE énergie avant travaux','DPE Énergie après travaux final':'DPE énergie après travaux final',
  'Tags':'TAG'
};
export const TECH_LABELS=[...ET_FIELDS.map(f=>f[1]),'Tags'];
const TECH_SET=new Set(TECH_LABELS);
export const targetColumn=l=>ALIAS[l]||l;

// Nature de chaque ligne produite (filtres et couleurs de l'interface).
export const KINDS={recolle:'Recollée',vide:'Recollée sans bâtiment ExtracTerre',sf:'Salesforce seul',et:'ExtracTerre seul',conflit:'Conflit contrat / OPE'};

export function normContract(v){
  const m=String(v??'').match(/^\s*0*(\d{1,8})(?:\.0+)?\s*$/);
  return m&&m[1]!=='0'?m[1].padStart(8,'0'):'';
}
export function normOpe(v){
  const m=String(v??'').match(/OPE[\s_-]*(\d{1,6})\b/i);
  return m?'OPE-'+m[1].padStart(6,'0'):'';
}
export const empty=v=>v==null||String(v).trim()===''||v==='#N/A';
// Nom d'opération ou de dossier Box « 0000xxxx - 26 - VILLE - … » : le n° de contrat est en tête (zéros de tête obligatoires).
export function leadContract(v){
  const m=String(v??'').match(/^\s*(0{2,}\d{4,6})(?=\s*(?:[-–—_]|\s\d{4,8}\b|$))/);
  return m?normContract(m[1]):'';
}
const contractOf=v=>normContract(v)||leadContract(v);
function same(a,b){
  const na=Number(String(a).replace(',','.')), nb=Number(String(b).replace(',','.'));
  if(Number.isFinite(na)&&Number.isFinite(nb)) return Math.abs(na-nb)<=1e-6*Math.max(1,Math.abs(na));
  return String(a).trim().toLowerCase()===String(b).trim().toLowerCase();
}

// Une opération ExtracTerre : {contract, code, source, buildings:[{libellé:valeur}]}. Les identifiants sont cherchés partout.
// Fiabilité ExtracTerre ≥ 2.34 (banc de vérité Salesforce) : certifiee > sure > probable > averifier ; '' = inconnue
// (export plus ancien). Lue dans la session (« reliability ») ou la feuille « Fiabilité par bâtiment » de l'export Excel.
export const REL_ORDER={certifiee:4,sure:3,probable:2,averifier:1};
export const REL_LABEL={certifiee:'Certifiée',sure:'Sûre',probable:'Probable',averifier:'À vérifier'};
export function relKey(v){ const s=String(v??'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().replace(/[^a-z]/g,''); return REL_ORDER[s]?s:''; }
// docAlt[i] : { libellé: {value,file,page,chosen} } — valeurs lues dans les documents par ExtracTerre quand il a retenu Salesforce
// (chosen=false), ou quand l'utilisateur y a choisi les documents (chosen=true ; la valeur du bâtiment est déjà celle des documents).
export function etOperation(buildings,ids=[],source='',rel=[],docAlt=[]){
  const all=[...ids,...buildings.flatMap(b=>[b['Contrat: Numéro du contrat'],b['Code interne'],b['Opération'],b['Projet'],b['Nom opération']])];
  const contract=all.map(contractOf).find(Boolean)||'';
  const code=all.map(normOpe).find(Boolean)||'';
  return {contract,code,source,buildings:buildings.map(b=>withDh(Object.fromEntries(Object.entries(b).filter(([k,v])=>TECH_SET.has(k)&&!empty(v))))),
    rel:buildings.map((b,i)=>Object.fromEntries(Object.entries(rel[i]||{}).map(([k,v])=>[k,relKey(v)]).filter(([,v])=>v))),
    docAlt:buildings.map((b,i)=>docAlt[i]||{})};
}
// ExtracTerre ≥ 2.11 : DH ZT / DH ZNT. « DH projet » reçoit le plus défavorable des deux ; un DH sans type de zone
// (exports plus anciens) va en DH ZNT, sauf bâtiment uniquement traversant (DH ZT) — même règle qu'ExtracTerre.
export const DH_ZT='DH Zone traversante (DH ZT)', DH_ZNT='DH Zone non traversante (DH ZNT)';
const num=v=>{ const n=Number(String(v??'').replace(/\s/g,'').replace(',','.')); return String(v??'').trim()!==''&&Number.isFinite(n)?n:null; };
const yes=v=>!empty(v)&&!/^(?:non|aucun|0)$/i.test(String(v).trim());
export function withDh(b){
  const o={...b};
  if(num(o[DH_ZT])===null&&num(o[DH_ZNT])===null&&num(o.DH)!==null) o[yes(o['Logement traversant'])&&!yes(o['Logement non traversant'])?DH_ZT:DH_ZNT]=o.DH;
  const z=[o[DH_ZT],o[DH_ZNT]].map(num).filter(x=>x!==null); if(z.length) o.DH=Math.max(...z);
  return o;
}
// Clé d'une décision d'écart (stable d'un recollement à l'autre).
export const ecartKey=e=>`${e.contrat||e.ope}|${e.batiment}|${e.colonne}`;

// sf = {header:[], rows:[[]]} ; etOps = [etOperation] (la dernière source chargée l'emporte sur une même opération).
// keepSalesforce : choix par défaut en cas d'écart ; decisions[ecartKey] = 'sf' | 'et' le remplace écart par écart.
export function recoll(sf,etOps,{keepSalesforce=false,decisions={},anru=null}={}){
  const header=[...sf.header];
  const ci=header.indexOf(CONTRACT_COL), oi=header.indexOf(OPE_COL);
  if(ci<0||oi<0) throw new Error(`Colonnes « ${CONTRACT_COL} » et « ${OPE_COL} » introuvables dans le tableau Tous contrats.`);
  // Groupes Salesforce : lignes partageant le même contrat (sinon la même OPE).
  const groups=new Map(), byContract=new Map(), byOpe=new Map();
  sf.rows.forEach((r,i)=>{
    const c=normContract(r[ci]), o=normOpe(r[oi]);
    const key=c?'c'+c:o?'o'+o:'row'+i;
    if(!groups.has(key)) groups.set(key,{key,rows:[],contracts:new Set(),opes:new Set()});
    const g=groups.get(key); g.rows.push(i); if(c) g.contracts.add(c); if(o) g.opes.add(o);
    if(c) byContract.set(c,g); if(o&&!byOpe.has(o)) byOpe.set(o,g);
  });
  // Rattachement de chaque opération ExtracTerre.
  const matched=new Map(), orphans=[];
  const unique=new Map(); etOps.forEach((op,i)=>unique.set(op.contract||op.code||'n'+i,op));
  for(const op of unique.values()){
    const gc=op.contract&&byContract.get(op.contract), go=op.code&&byOpe.get(op.code);
    const g=gc||go;
    if(gc&&go&&gc!==go){ orphans.push({op,kind:'conflit',status:`Conflit : contrat ${op.contract} et ${op.code} désignent deux opérations Salesforce différentes`}); continue; }
    if(g&&op.contract&&g.contracts.size&&!g.contracts.has(op.contract)){ orphans.push({op,kind:'conflit',status:`Conflit : ${op.code} est rattachée au contrat ${[...g.contracts][0]} dans Salesforce, pas ${op.contract}`}); continue; }
    if(g&&op.code&&g.opes.size&&!g.opes.has(op.code)){ orphans.push({op,kind:'conflit',status:`Conflit : le contrat ${op.contract} porte ${[...g.opes][0]} dans Salesforce, pas ${op.code}`}); continue; }
    if(!g){ orphans.push({op,kind:'et',status:op.contract||op.code?'ExtracTerre seul (absent de Tous contrats)':'ExtracTerre seul (ni contrat ni OPE trouvés)'}); continue; }
    let status=gc&&go?'Recollé (contrat + OPE)':gc?'Recollé (contrat)':'Recollé (OPE)', replaced=false;
    if(matched.has(g)){ status+=' — remplace une autre source ExtracTerre'; replaced=true; }
    matched.set(g,{op,status,replaced,via:gc&&go?'both':gc?'contract':'ope'});
  }
  // Colonnes : celles du tableau, puis les champs ExtracTerre sans équivalent qui ont une valeur.
  const used=new Set(etOps.flatMap(op=>op.buildings.flatMap(Object.keys)));
  // Colonnes ajoutées sous leur nom OPERATIONS (« Cep projet », « Cep max »…) : c'est sous ce nom que les valeurs sont écrites.
  const extra=[...new Set(TECH_LABELS.filter(l=>used.has(l)&&!header.includes(targetColumn(l))&&l!==BUILDING_COL).map(targetColumn))];
  const sfWidth=header.length;
  header.push(BUILDING_COL,...extra,...(anru?[ANRU_COL,ANRU_QP_COL]:[]),STATUS_COL);
  const col=Object.fromEntries(header.map((h,i)=>[h,i]));
  const techCols=[...new Set(TECH_LABELS.map(targetColumn).filter(h=>h in col))];
  const rows=[], meta=[], ecarts=[];
  const stats={recolles:0,salesforceSeul:0,etSeul:0,conflits:0,lignes:0,ecarts:0,aVerifier:0,etOperations:unique.size,etBatiments:0,viaContrat:0,viaOpe:0,viaDeux:0};
  const width=header.length;
  // Écrit les valeurs d'un bâtiment ExtracTerre dans la ligne ; renvoie les colonnes remplies par ExtracTerre.
  const overlay=(row,b,op,rowIndex,rel={},relOut=null,alt={})=>{
    const filled=[];
    for(const [label,v] of Object.entries(b)){
      const h=targetColumn(label), j=col[h]; if(j==null) continue;
      if(!empty(row[j])&&!same(row[j],v)&&h!==BUILDING_COL){
        const e={contrat:normContract(row[ci])||op.contract,ope:normOpe(row[oi])||op.code,batiment:b[BUILDING_COL]||'',colonne:h,salesforce:row[j],extracterre:v,source:op.source,rowIndex};
        // Valeur des documents choisie dans ExtracTerre à la place de Salesforce : ExtracTerre retenu par défaut.
        e.key=ecartKey(e); const choice=decisions[e.key]||(alt[label]?.chosen?'et':keepSalesforce?'sf':'et'); if(alt[label]?.chosen){ e.choixExtracTerre=true; e.source=`${e.source||''} — valeur des documents choisie dans ExtracTerre (Salesforce : ${alt[label].sf??''})`; }
        e.retenu=choice==='sf'?'Salesforce':'ExtracTerre'; e.decided=!!decisions[e.key]; ecarts.push(e);
        if(choice==='sf') continue;
      }
      row[j]=v; filled.push(j); if(relOut) relOut[j]=rel[label]||(label==='DH'?rel[DH_ZNT]||rel[DH_ZT]:'')||'';
    }
    // ExtracTerre a retenu Salesforce alors que ses documents donnent une autre valeur : écart « Salesforce ↔ documents »,
    // Salesforce retenu par défaut (priorité inchangée), bascule possible vers la valeur des documents.
    for(const [label,a] of Object.entries(alt)){
      if(a.chosen) continue; const h=targetColumn(label), j=col[h]; if(j==null||h===BUILDING_COL) continue;
      const cur=row[j]; if(empty(cur)||empty(a.value)||same(cur,a.value)||ecarts.some(x=>x.rowIndex===rowIndex&&x.colonne===h)) continue;
      const e={contrat:normContract(row[ci])||op.contract,ope:normOpe(row[oi])||op.code,batiment:b[BUILDING_COL]||'',colonne:h,salesforce:cur,extracterre:a.value,docs:true,
        source:`Documents lus par ExtracTerre : ${a.file||'document'}${a.page?' p. '+a.page:''} (ExtracTerre avait retenu Salesforce)`,rowIndex};
      e.key=ecartKey(e)+'|docs'; const choice=decisions[e.key]||'sf';
      e.retenu=choice==='sf'?'Salesforce':'ExtracTerre'; e.decided=!!decisions[e.key]; ecarts.push(e);
      if(choice==='et'){ row[j]=a.value; if(!filled.includes(j)) filled.push(j); if(relOut) relOut[j]=rel[label]||''; }
    }
    return filled;
  };
  for(const g of groups.values()){
    const m=matched.get(g);
    if(!m||!m.op.buildings.length){
      g.rows.forEach(i=>{ const r=sf.rows[i].slice(0,sfWidth); r.length=width; r[col[STATUS_COL]]=m?'Recollé — aucun bâtiment ExtracTerre':'Salesforce seul'; rows.push(r); meta.push({kind:m?'vide':'sf',check:false,et:[],group:g.key}); });
      if(m) stats.recolles++; else stats.salesforceSeul++; continue;
    }
    stats.recolles++; stats[m.via==='both'?'viaDeux':m.via==='contract'?'viaContrat':'viaOpe']++;
    const k=g.rows.length, n=m.op.buildings.length;
    m.op.buildings.forEach((b,i)=>{
      // Lignes Salesforce déjà découpées par bâtiment : appariement dans l'ordre ; sinon la 1re ligne sert de base.
      const base=sf.rows[g.rows[k===n?i:0]].slice(0,sfWidth); base.length=width;
      const mismatch=k>1&&k!==n;
      if(mismatch) techCols.forEach(h=>{ base[col[h]]=undefined; });
      const rel={}; const et=overlay(base,b,m.op,rows.length,m.op.rel?.[i]||{},rel,m.op.docAlt?.[i]||{});
      let st=m.status; if(mismatch){ st+=` — à vérifier : ${k} lignes Salesforce pour ${n} bâtiments`; stats.aVerifier++; }
      base[col[STATUS_COL]]=st; rows.push(base); meta.push({kind:'recolle',check:mismatch||m.replaced,et,rel,group:g.key,source:m.op.source,op:m.op,bi:i});
    });
    stats.etBatiments+=n;
  }
  for(const {op,kind,status} of orphans){
    kind==='conflit'?stats.conflits++:stats.etSeul++;
    (op.buildings.length?op.buildings:[{}]).forEach((b,bi)=>{
      const r=new Array(width); r[ci]=op.contract; r[oi]=op.code; const rel={}; const et=overlay(r,b,op,rows.length,op.rel?.[bi]||{},rel,op.docAlt?.[bi]||{});
      r[col[STATUS_COL]]=`${status} — ${op.source}`; rows.push(r); meta.push({kind,check:true,et,rel,group:'',source:op.source,op,bi});
    });
    stats.etBatiments+=op.buildings.length;
  }
  // Format 0000xxxx et OPE-XXXXXX partout.
  rows.forEach(r=>{ if(!empty(r[ci])) r[ci]=normContract(r[ci])||r[ci]; if(!empty(r[oi])) r[oi]=normOpe(r[oi])||r[oi]; });
  if(anru){ const h=n=>col[n]; rows.forEach(r=>{ const x=anruOf(anru,{lat:r[h('Latitude')],lon:r[h('Longitude')],cp:r[h('Code postal')],ville:r[h('Ville')]}); r[col[ANRU_COL]]=x.anru; r[col[ANRU_QP_COL]]=x.quartier; }); }
  stats.lignes=rows.length; stats.ecarts=ecarts.length;
  return {header,rows,meta,ecarts,stats,sfWidth,extra};
}

// ExtracTerre ≥ 3.2.0 : valeurs retenues portant un écart Salesforce ↔ documents (documentValue quand Salesforce est retenu,
// sfChoice 'doc' quand l'utilisateur a choisi les documents).
function docAltOf(o,building,nRows){ const out={};
  for(const f of o.finals||[]){ if(f.building!==building&&!(f.building==='Bâtiment unique'&&nRows===1)) continue; const label=KEY2LABEL[f.field]||f.field;
    if(f.sfChoice==='doc') out[label]={value:f.value,chosen:true,sf:f.sfValue};
    else if(f.documentValue!==undefined&&f.documentValue!==null&&f.documentValue!=='') out[label]={value:f.documentValue,file:f.documentFile||'',page:f.documentPage||'',chosen:false}; }
  return out; }
// Session ExtracTerre (.json) → opérations ; les clés internes sont traduites en libellés.
const KEY2LABEL=Object.fromEntries(ET_FIELDS);
export function opsFromSession(json,source){
  if(json?.format!=='extracterre-session'||!Array.isArray(json.operations)) throw new Error('Ce JSON n’est pas une session ExtracTerre.');
  return json.operations.map(o=>({id:o.id||'',...etOperation(
    (o.rows||[]).map(r=>({...Object.fromEntries(Object.entries(r).map(([k,v])=>[KEY2LABEL[k]||k,v])),'Contrat: Numéro du contrat':r.contract_number,'Code interne':r.internal_code})),
    [o.identity?.contract,o.identity?.code,o.title],`${source} › ${o.title||o.identity?.name||'sans titre'}`,
    (o.rows||[]).map(r=>Object.fromEntries(Object.entries(o.reliability?.[r.building]||{}).map(([k,v])=>[KEY2LABEL[k]||k,v]))),
    (o.rows||[]).map(r=>docAltOf(o,r.building,(o.rows||[]).length)))}));
}
// Export Excel ExtracTerre (feuille « Données par bâtiment », lignes déjà en libellés) → opérations groupées par contrat/OPE.
export function opsFromBuildingRows(rows,source,relRows=[]){
  const groups=new Map();
  rows.forEach((r,i)=>{
    const c=normContract(r['Contrat: Numéro du contrat'])||contractOf(r['Projet'])||contractOf(r['Opération'])||leadContract(r['Nom opération']);
    const o=normOpe(r['Code interne']);
    const key=c||o||'row'+i;
    (groups.get(key)||groups.set(key,[]).get(key)).push(Object.assign(r,{__i:i}));
  });
  return [...groups.values()].map(b=>etOperation(b,[],`${source} › ${b[0]['Nom opération']||b[0]['Projet']||'sans titre'}`,b.map(r=>relRows[r.__i]||{})));
}

// ANRU (NPNRU) : 480 quartiers (intérêt national, arrêté du 29/04/2015 ; intérêt régional, arrêté du 15/01/2019), périmètres QP 2015.
// Bâtiment géolocalisé : Oui s'il est dans un périmètre ou à 10 m au plus de son bord, sinon Non. Sans coordonnées : Non si sa commune n'a aucun quartier
// NPNRU ; vide (à vérifier) si elle en a un. shortcut: périmètre strict, sans la bande de 300 m autour du quartier.
export const ANRU_COL='ANRU', ANRU_QP_COL='Quartier ANRU';
const normCity=v=>String(v??'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().replace(/\bcedex\b.*$/,'').replace(/[-'’]/g,' ')
  .replace(/\bste?\b/g,m=>m==='st'?'saint':'sainte').replace(/\s+\d+\s*(?:er|e|eme)?(?:\s+arrondissement)?\s*$/,'').replace(/\s+/g,' ').trim();
const deptOf=(cp,insee)=>{ const s=String(insee||cp||'').trim(); if(/^97/.test(s)) return s.slice(0,3); if(/^(?:2A|2B|20)/i.test(s)) return '20'; return s.slice(0,2); };
export function anruIndex(data){
  const qs=(data?.quartiers||[]).map(q=>{ let s=90,w=180,n=-90,e=-180; for(const r of q.rings) for(const [la,lo] of r){ if(la<s)s=la; if(la>n)n=la; if(lo<w)w=lo; if(lo>e)e=lo; } return {...q,bbox:[s,w,n,e]}; });
  const cities=new Set(qs.flatMap(q=>q.communes.map(([insee,lib])=>deptOf('',insee)+'|'+normCity(lib))));
  return {quartiers:qs,cities};
}
// Règle pair-impair sur tous les anneaux : gère les trous et les quartiers en plusieurs parties.
const inRings=(rings,la,lo)=>{ let inside=false; for(const r of rings) for(let i=0,j=r.length-1;i<r.length;j=i++){ const [ai,oi]=r[i],[aj,oj]=r[j]; if((ai>la)!==(aj>la)&&lo<(oj-oi)*(la-ai)/(aj-ai)+oi) inside=!inside; } return inside; };
export function anruAt(index,lat,lon){
  return index.quartiers.find(q=>lat>=q.bbox[0]&&lat<=q.bbox[2]&&lon>=q.bbox[1]&&lon<=q.bbox[3]&&inRings(q.rings,lat,lon))||null;
}
// Distance (m) du point au bord le plus proche d'un quartier, pour les quartiers à moins de ~60 m (sinon null).
export const ANRU_EDGE_M=10;
function nearEdge(index,la,lo){
  const c=Math.cos(la*Math.PI/180), X=x=>x*c*111320, Y=y=>y*110540, px=X(lo), py=Y(la), pad=.0006; let best=null;
  for(const q of index.quartiers){ if(la<q.bbox[0]-pad||la>q.bbox[2]+pad||lo<q.bbox[1]-pad/c||lo>q.bbox[3]+pad/c) continue;
    for(const r of q.rings) for(let i=1;i<r.length;i++){ const ax=X(r[i-1][1]),ay=Y(r[i-1][0]),dx=X(r[i][1])-ax,dy=Y(r[i][0])-ay,t=Math.max(0,Math.min(1,((px-ax)*dx+(py-ay)*dy)/((dx*dx+dy*dy)||1))),d=Math.hypot(px-ax-t*dx,py-ay-t*dy); if(!best||d<best.d) best={d,q}; } }
  return best;
}
export function anruOf(index,{lat,lon,cp,ville}){
  const la=num(lat), lo=num(lon);
  if(la!==null&&lo!==null&&la!==0){ const q=anruAt(index,la,lo); if(q) return {anru:'Oui',quartier:`${q.nom} (${q.qp}, intérêt ${q.cat})`,q};
    // Point adresse à 10 m au plus du périmètre (rue en limite de quartier) : intégré, distance indiquée.
    const e=nearEdge(index,la,lo); return e&&e.d<=ANRU_EDGE_M?{anru:'Oui',quartier:`${e.q.nom} (${e.q.qp}, intérêt ${e.q.cat}), à ${Math.round(e.d)} m du périmètre`,q:e.q}:{anru:'Non',quartier:''}; }
  const c=normCity(ville); if(!c||!String(cp??'').trim()) return {anru:'',quartier:''};
  return index.cities.has(deptOf(cp)+'|'+c)?{anru:'',quartier:'Commune avec quartier NPNRU : bâtiment non géolocalisé'}:{anru:'Non',quartier:''};
}

// Confort d'été : DH ZT / DH ZNT, système constructif, brasseur d'air, climatisation, normalisés par bâtiment.
// get(colonne) lit la ligne recollée. Une valeur absente ou ambiguë reste vide, avec la raison dans « note ».
export const SYSTEMS={iti:'Béton ou brique + isolation intérieure ou isolation répartie',ite:'Béton ou brique + isolation extérieure',bois:'Ossature bois'};
const nt=v=>String(v??'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().trim();
const RE_BOIS=/ossature\s*bois|\bmob\b|\bcob\b|\bclt\b|bois\s*massif|colombage|pans?\s*de\s*bois/;
const RE_MASS=/beton|brique|parpaing|\bblocs?\b|pierre|meuliere|moellon|maconn|porotherm|monomur|banche|prefabriqu|\bvoiles?\b|terre\s*cuite|agglo/;
const RE_REPART=/monomur|beton\s*cellulaire|isolation\s*repartie|thermo\s*brique|brique\s*isolante/;
const yesNo=v=>{ const s=nt(v); return /^oui\b/.test(s)?true:/^non\b/.test(s)?false:null; };
export function confortOf(get){
  // Structure des murs, sinon mode constructif (« Poutrelles-hourdis », un plancher, ne décide pas si l'autre colonne le fait).
  const cands=[get('Parois verticales structure'),get('Mode constructif')].filter(v=>!empty(v)), raw=cands.find(v=>RE_BOIS.test(nt(v))||RE_MASS.test(nt(v)))??cands[0];
  const notes=[], dh=l=>num(get(l)), struct=nt(raw);
  // ITE / ITI : la colonne Oui / Non prime ; sinon un isolant ITE / ITI renseigné, ou la mention dans le type d'isolant.
  const ins=nt(get('Parois verticales type d’isolant'));
  const side=(col,isoCol,re)=>yesNo(get(col))??(!empty(get(isoCol))||re.test(ins)?true:null);
  const ite=side('ITE','Parois verticales ITE type d’isolant',/\bite\b|exterieur|isofacade/), iti=side('ITI','Parois verticales ITI type d’isolant',/\biti\b|interieur|doublage|doublissimo|complexe isolant/);
  let system='';
  if(RE_BOIS.test(struct)) system=SYSTEMS.bois;
  else if(RE_MASS.test(struct)){
    if(ite&&iti) notes.push('ITE et ITI : à trancher');
    else if(ite) system=SYSTEMS.ite;
    else if(iti||RE_REPART.test(struct)) system=SYSTEMS.iti;
    else notes.push('isolation des murs non précisée');
  } else if(struct) notes.push(`structure hors catégories : ${raw}`);
  else if(ite||iti) notes.push('structure non renseignée');
  const fanN=num(get('Nombre de brasseurs d’air')), fanYN=yesNo(get('Brasseur d’air'));
  const fan=fanYN!==null?fanYN:fanN!==null?fanN>0:null;
  const cool=nt(get('Refroidissement après travaux'));
  const clim=yesNo(get('Climatisation active'))??(/^(?:aucun|sans|pas de|neant|non)\b/.test(cool)?false:/climatis|split|groupe froid|reversible|\bvrv\b|\bdrv\b/.test(cool)?true:null);
  const on=b=>b===null?'':b?'Oui':'Non';
  return {dhZt:dh(DH_ZT),dhZnt:dh(DH_ZNT),system,fan:on(fan),clim:on(clim),note:notes.join(' ; ')};
}

// Valeurs plausibles mais étranges, surlignées en jaune (jamais modifiées) — mêmes règles qu'ExtracTerre (js/sanity.js oddValue) :
// Tic < 24 °C ; DH < 10 ou > 1 250 °C.h ; Cep / Cep,nr après travaux > 300 kWhep/m².an, avant travaux > 800 ; fioul après travaux.
export const ODD_RANGES={tic:[24,null],dh:[10,1250],dh_zt:[10,1250],dh_znt:[10,1250],cep:[null,300],cepnr:[null,300],cep_after_final:[null,300],cep_before:[null,800]};
const ODD_AFTER_VECTORS=new Set(['heating_vector_after','heating_mode_after','ecs_vector_after']);
export function oddValue(key,v){
  if(empty(v)) return '';
  if(ODD_AFTER_VECTORS.has(key)) return /\b(?:fioul|fuel|mazout)\b/i.test(String(v))?'Fioul après travaux : inhabituel':'';
  const r=ODD_RANGES[key]; if(!r) return '';
  const n=typeof v==='number'?v:Number(String(v).replace(/[\s\u00a0]/g,'').replace(',','.')); if(!Number.isFinite(n)) return '';
  if(r[0]!==null&&n<r[0]) return `Valeur très faible (< ${r[0]}) : à vérifier`;
  if(r[1]!==null&&n>r[1]) return `Valeur très élevée (> ${r[1].toLocaleString('fr-FR')}) : à vérifier`;
  return '';
}
// Colonne du tableau → clé ExtracTerre (pour appliquer oddValue).
export const COLUMN_KEY=Object.fromEntries(ET_FIELDS.map(([k,l])=>[targetColumn(l),k]));
// Tri d'une colonne : nombres (y compris « 12,5 ») par valeur, textes sans casse ni accents ; cellules vides toujours en fin.
const sortNum=v=>typeof v==='number'?v:/^-?\d[\d\s\u00a0]*(?:[.,]\d+)?$/.test(String(v).trim())?Number(String(v).replace(/[\s\u00a0]/g,'').replace(',','.')):null;
export function sortRowIndexes(idx,rows,col,dir){
  if(col==null||!dir) return idx; const s=dir==='desc'?-1:1;
  return [...idx].sort((a,b)=>{ const x=rows[a][col], y=rows[b][col], ex=empty(x), ey=empty(y); if(ex||ey) return ex-ey;
    const nx=sortNum(x), ny=sortNum(y); if(nx!==null&&ny!==null) return s*(nx-ny)||a-b;
    return s*String(x).localeCompare(String(y),'fr',{numeric:true,sensitivity:'base'})||a-b; });
}

// Carte confort d'été : zone climatique RE2020 par département (arrêté du 4 août 2021, mêmes zones que la RT2012).
export const ZONES=['H1a','H1b','H1c','H2a','H2b','H2c','H2d','H3'];
const ZONE_DEPS={H1a:'02 14 27 59 60 62 75 76 77 78 80 91 92 93 94 95',H1b:'08 10 45 51 52 54 55 57 58 61 67 68 70 88 89 90',H1c:'01 03 05 15 19 21 23 25 38 39 42 43 63 69 71 73 74 87',H2a:'22 29 35 50 56',H2b:'16 17 18 28 36 37 41 44 49 53 72 79 85 86',H2c:'09 12 24 31 32 33 40 46 47 64 65 81 82',H2d:'04 07 26 48 84',H3:'06 11 13 2A 2B 20 30 34 66 83'};
const DEP_ZONE=Object.fromEntries(Object.entries(ZONE_DEPS).flatMap(([z,l])=>l.split(' ').map(d=>[d,z])));
export const zoneOfDep=d=>DEP_ZONE[String(d??'').toUpperCase()]||'';
// Département d'une ligne : code postal, sinon code INSEE, sinon « 0000xxxx - 62 - … » en tête du nom d'opération. null si inconnu.
export function depOfRow(get){
  for(const v of [get('Code postal'),get('Code INSEE')]){ const s=String(v??'').trim(); if(/^(?:2[AB]|\d{2})\d{3}$/i.test(s)) return /^20/.test(s)?'20':s.slice(0,2).toUpperCase(); }
  for(const v of [get('Département'),get('Departement')]){ const m=String(v??'').match(/^\s*(2[AB]|\d{2})\b/i); if(m) return m[1].toUpperCase(); }
  const m=String(get("Nom de l'opération")??'').match(/^\s*\d{8}\s*-\s*(2[AB]|\d{2})\s*-/i); return m?m[1].toUpperCase():null;
}
const stat=xs=>{ if(!xs.length) return {n:0,mean:null,median:null}; const s=[...xs].sort((a,b)=>a-b), h=s.length>>1;
  return {n:s.length,mean:Math.round(s.reduce((a,b)=>a+b,0)/s.length*10)/10,median:Math.round((s.length%2?s[h]:(s[h-1]+s[h])/2)*10)/10}; };
// items : [{op, zone, dhZt, dhZnt, system, fan, clim}] (un par bâtiment). DH : statistiques sur les bâtiments.
// Brasseur / climatisation : part de toutes les opérations de la zone (une opération compte « Oui » si un de ses bâtiments l'est ;
// « Non » lu ou absence de mention comptent comme « Non »).
export function confortStats(items){
  const share=(list,k)=>{ const ops=new Map(); for(const x of list){ if(!x[k]) continue; ops.set(x.op,ops.get(x.op)==='Oui'||x[k]==='Oui'?'Oui':'Non'); }
    const yes=[...ops.values()].filter(v=>v==='Oui').length, total=new Set(list.map(x=>x.op)).size;
    return {yes,known:ops.size,total,pct:total?Math.round(yes/total*1000)/10:null}; };
  const dh=list=>({zt:stat(list.map(x=>x.dhZt).filter(v=>v!=null)),znt:stat(list.map(x=>x.dhZnt).filter(v=>v!=null))});
  const one=list=>({ops:new Set(list.map(x=>x.op)).size,buildings:list.length,...dh(list),fan:share(list,'fan'),clim:share(list,'clim'),
    systems:Object.values(SYSTEMS).map(s=>({system:s,...dh(list.filter(x=>x.system===s)),buildings:list.filter(x=>x.system===s).length}))});
  return {all:one(items),zones:Object.fromEntries(ZONES.map(z=>[z,one(items.filter(x=>x.zone===z))])),unzoned:items.filter(x=>!x.zone).length};
}

// Tableau publié dans l'onglet OPERATIONS de l'Observatoire (voir app.js « Publication »).
export function publishTable(r,minLevel){
  const min=REL_ORDER[minLevel]||0, keep=new Set([r.header.indexOf(CONTRACT_COL),r.header.indexOf(OPE_COL),r.header.indexOf(BUILDING_COL)]);
  const restore=new Map(r.ecarts.filter(e=>e.retenu==='ExtracTerre').map(e=>[`${e.rowIndex}|${e.colonne}`,e.salesforce]));
  let skipped=0;
  const rows=r.rows.map((row,i)=>{ const out=Array.from({length:r.header.length},(_,j)=>{ const v=row[j]; return v===undefined||v===null?'':typeof v==='number'?v:String(v); });
    const m=r.meta[i]; (m?.et||[]).forEach(j=>{ if(keep.has(j)||(REL_ORDER[m.rel?.[j]||'']||0)>=min) return; const k=`${i}|${r.header[j]}`; out[j]=restore.has(k)?String(restore.get(k)??''):''; skipped++; });
    return out; });
  return {header:r.header.map(h=>String(h??'')),rows,skipped};
}

// ---------- Recherches ExtracTerre (onglet Recherches d'ExtracTerre, anciennement « Observatoire des recherches ») ----------
// Lues dans la tête d'une session (operations[].searches + grids) ou dans l'export « ExtracTerre_observatoire_….json ».
// Le référentiel, le maître d'ouvrage et le nom viennent de Tous contrats (Salesforce fait foi), par contrat puis par OPE.
export function researchFromSession(json,source=''){
  if(json?.format!=='extracterre-session'||!Array.isArray(json.operations)) throw new Error('Ce JSON n’est pas une session ExtracTerre.');
  const grids=(Array.isArray(json.grids)?json.grids:[]).filter(g=>g&&g.id&&Array.isArray(g.criteria));
  const ops=json.operations.filter(o=>o&&o.searches&&Object.keys(o.searches).length).map(o=>({
    contract:contractOf(o.identity?.contract)||leadContract(o.title)||leadContract(o.identity?.name),
    ope:normOpe(o.identity?.code)||normOpe(o.title),
    name:String(o.identity?.name||o.title||''),owner:String(o.identity?.owner||''),
    searches:o.searches,at:String(json.exportedAt||''),source}));
  return {grids,ops};
}
const researchKey=o=>o.contract||o.ope||o.name.trim().toLowerCase();
// Plusieurs fichiers : une opération par contrat (sinon OPE, sinon nom) ; pour chaque grille, l'export le plus récent l'emporte.
export function mergeResearch(parts){
  const grids=new Map(), ops=new Map();
  for(const p of parts){ for(const g of p.grids||[]) grids.set(String(g.id),g);
    for(const o of p.ops||[]){ const k=researchKey(o); if(!k) continue; const prev=ops.get(k);
      if(!prev){ ops.set(k,{...o,searches:{...o.searches}}); continue; }
      for(const [gid,s] of Object.entries(o.searches)) if(!prev.searches[gid]||o.at>=prev.at) prev.searches[gid]=s;
      if(o.at>=prev.at) Object.assign(prev,{name:o.name||prev.name,owner:o.owner||prev.owner,at:o.at}); } }
  return {grids:[...grids.values()],ops:[...ops.values()]};
}
export const NO_REF='Non trouvé dans Tous contrats';
// Tous contrats → par contrat et par OPE : référentiel(s) (tous les bâtiments), maître d'ouvrage, nom de l'opération.
export function sfIndex(sf){
  const byC=new Map(), byO=new Map(); if(!sf) return {byC,byO};
  const h=sf.header, col=(...names)=>{ for(const n of names){ const j=h.indexOf(n); if(j>=0) return j; } return -1; };
  const cC=h.indexOf(CONTRACT_COL), cO=h.indexOf(OPE_COL), cR=h.findIndex(x=>/^r[ée]f[ée]rentiels?$/i.test(String(x).trim()));
  const cM=col('Nom de la société: Nom de la société','Maître d’ouvrage',"Maître d'ouvrage"), cN=col("Nom de l'opération",'Nom de l’opération');
  for(const r of sf.rows){ const c=cC>=0?normContract(r[cC]):'', o=cO>=0?normOpe(r[cO]):''; if(!c&&!o) continue;
    const e=(c&&byC.get(c))||(o&&byO.get(o))||{refs:new Set(),owner:'',name:''};
    const ref=cR>=0?String(r[cR]??'').trim():''; if(ref&&!empty(ref)) ref.split(/\s*;\s*/).filter(Boolean).forEach(x=>e.refs.add(x));
    if(!e.owner&&cM>=0&&!empty(r[cM])) e.owner=String(r[cM]).trim(); if(!e.name&&cN>=0&&!empty(r[cN])) e.name=String(r[cN]).trim();
    if(c) byC.set(c,e); if(o) byO.set(o,e); }
  return {byC,byO};
}
export function researchRows(merged,sf){
  const ix=sfIndex(sf);
  return merged.ops.map(o=>{ const e=(o.contract&&ix.byC.get(o.contract))||(o.ope&&ix.byO.get(o.ope))||null;
    return {...o,inSf:!!e,refs:e&&e.refs.size?[...e.refs].sort((a,b)=>a.localeCompare(b,'fr')):[e?'Référentiel non renseigné':NO_REF],
      owner:e?.owner||o.owner||'Non renseigné',name:e?.name||o.name}; });
}
// Statistiques d'une grille sur une sélection d'opérations (mêmes règles que l'observatoire d'ExtracTerre).
export function researchStats(rows,grid){
  const withG=rows.filter(o=>o.searches?.[grid.id]), num=v=>(Array.isArray(v)?v:[]).map(Number).filter(Number.isFinite);
  const themes=(grid.themes||[]).map(t=>{ const pool=withG.filter(o=>o.searches[grid.id].themes?.[t.name]?.applicable!==false);
    const crit=(grid.criteria||[]).filter(c=>c.theme===t.name).map(c=>{ const rs=pool.map(o=>o.searches[grid.id].criteria?.[c.id]).filter(Boolean);
      const vals=rs.flatMap(r=>num(r.values)).sort((a,b)=>a-b);
      return {c,pool:pool.length,ok:rs.filter(r=>r.found).length,ko:rs.filter(r=>!r.found&&r.contrary).length,
        min:vals.length?vals[0]:null,med:vals.length?vals[Math.floor(vals.length/2)]:null,max:vals.length?vals[vals.length-1]:null}; });
    return {t,applicable:pool.length,ok:pool.filter(o=>o.searches[grid.id].themes?.[t.name]?.ok).length,crit}; });
  return {rows:withG,themes};
}

/* ---------- Vérification humaine des « Oui » des recherches (18/11/2026) ----------
   Chaque réponse Oui / Partiel d'un critère est montrée avec son exigence et ses preuves (document, page, ligne, contexte) ;
   l'utilisateur l'accepte ou la rejette. Une réponse rejetée ne compte plus comme trouvée (statistiques, CSV, thèmes recalculés). */
export const verifKey=(o,gid,cid)=>`${researchKey(o)}|${gid}|${cid}`;
const verifStatus=x=>!x?'':x.partial?'partiel':x.found?'oui':'';
// 19/11/2026 — règles v4 d'ExtracTerre (3.0.5) : pour ces critères la phrase de preuve a changé ; une décision prise sur
// l'ancienne phrase n'est plus appliquée (elle reste enregistrée et revient si l'on recharge l'ancienne session). Une décision
// enregistrée avec sa phrase (line) ne vaut que pour cette phrase.
export const RULES_CHANGED={'securite-logements':{4:['etage_escaliers','hall_sas','parking_pieton']}};
const evLine=e=>String(e?.line||'').replace(/\s+/g,' ').trim();
export function decisionApplies(dec,grid,critId,line){
  if(!dec?.d) return false;
  if(dec.line!=null) return evLine({line:dec.line})===evLine({line});
  const ch=RULES_CHANGED[grid?.id]||{}; return !Object.entries(ch).some(([v,ids])=>(Number(grid?.rulesVersion)||0)>=Number(v)&&ids.includes(critId));
}
export function verificationItems(rows,grid,decisions={}){
  const out=[], valid={};
  for(const o of rows){ const s=o.searches?.[grid.id]; if(!s) continue;
    for(const c of grid.criteria||[]){ const x=s.criteria?.[c.id], st=verifStatus(x); if(!st) continue;
      const ev=(x.evidence||[]).filter(e=>e&&e.ok!==false&&!e.negative);
      const key=verifKey(o,grid.id,c.id), evidence=ev.length?ev:(x.evidence||[]), own=decisions[key], ok=decisionApplies(own,grid,c.id,evidence[0]?.line);
      if(ok) valid[key]=own;
      out.push({key,op:o,crit:c,status:st,manual:x.manual||'',value:x.value??'',evidence,decision:ok?own.d:'',...(own?.d&&!ok?{stale:true}:{})}); } }
  // 18/11/2026 — un Oui donné une fois dans une opération vaut pour les cartes en attente de cette même opération qui portent
  // sur le même critère (opération chargée deux fois) ou montrent la même phrase (« 44.2 » des fenêtres et du hall) : elles
  // sont validées d'office (decision 'ok', inherited) et ne sont plus redemandées. Rien n'est enregistré pour elles.
  const opId=o=>o.contract||o.ope||String(o.name||'').trim().toLowerCase(), line=i=>hlNorm(String(i.evidence[0]?.line||'')).replace(/[^a-z0-9.,/+*%-]+/g,' ').trim();
  // 18/11/2026 — même phrase, même exigence, toutes opérations : la décision donnée une fois (Oui, Partiel, Non, Sans objet) vaut
  // partout où cette phrase est la preuve de cette exigence (CCTP types répétés). La décision propre à une carte l'emporte ; si la
  // même phrase a reçu plusieurs décisions, la plus récente est reprise.
  { const bySentence=new Map();
    for(const i of out){ const own=valid[i.key]; if(!own?.d) continue; const k=`${i.crit.id}|${line(i)}`; if(!line(i)) continue;
      const prev=bySentence.get(k); if(!prev||String(own.at||'')>=String(prev.at||'')) bySentence.set(k,{d:own.d,at:own.at,from:i.op}); }
    for(const i of out){ if(i.decision) continue; const src=bySentence.get(`${i.crit.id}|${line(i)}`); if(src){ i.decision=src.d; i.inherited=`même phrase, ${src.from.contract||src.from.ope||src.from.name}`; i.sameSentence=true; } } }
  const byOp=new Map(); for(const i of out){ const k=opId(i.op); if(!byOp.has(k)) byOp.set(k,[]); byOp.get(k).push(i); }
  for(const list of byOp.values()){ const yes=list.filter(i=>i.decision==='ok'); if(!yes.length) continue;
    for(const i of list){ if(i.decision) continue; const src=yes.find(y=>y.crit.id===i.crit.id||(line(y)&&line(y)===line(i)));
      if(src){ i.decision='ok'; i.inherited=src.crit.label||src.crit.id; } } }
  return out;
}
/* 18/11/2026 — Contrôle des décisions d'ExtracTerre : une carte par exigence et par opération dont ExtracTerre a décidé Oui,
   Partiel ou Non (avec ses preuves). La décision manuelle remplace celle d'ExtracTerre pour cette opération seulement (aucune
   reprise vers d'autres cartes) et l'emporte sur la vérification des « Oui ». Enregistrée à part (totalrecoll.controles). */
export const etDecision=x=>!x?'':x.partial?'pa':x.found?'ok':x.contrary?'ko':'';
export const CTRL_LABEL={ok:'Oui',pa:'Partiel',ko:'Non',so:'Sans objet'};
export function controlItems(rows,grid,controls={}){
  const out=[];
  for(const o of rows){ const s=o.searches?.[grid.id]; if(!s) continue;
    for(const c of grid.criteria||[]){ const x=s.criteria?.[c.id], et=etDecision(x); if(!et) continue;
      const all=(x.evidence||[]).filter(Boolean), pos=all.filter(e=>e.ok!==false&&!e.negative), neg=all.filter(e=>e.ok===false||e.negative);
      const ev=et==='ko'?(neg.length?neg:all):(pos.length?pos:all), key=verifKey(o,grid.id,c.id);
      out.push({key,op:o,crit:c,et,status:et==='pa'?'partiel':et==='ok'?'oui':'non',manual:x.manual||'',value:x.value??'',evidence:ev,decision:controls[key]?.d||''}); } }
  return out;
}
// Décision de contrôle appliquée à un critère (absolue : ce que l'utilisateur a décidé, quel que soit ExtracTerre).
function controlled(x,d){
  if(d==='ok') return {...x,found:true,partial:false,contrary:false,controlled:true,value:'Oui (contrôle)'};
  if(d==='pa') return {...x,found:false,partial:true,contrary:false,controlled:true,value:'Partiel (contrôle)'};
  if(d==='ko') return {...x,found:false,partial:false,contrary:true,controlled:true,value:'Non conforme (contrôle)'};
  return {...x,found:false,partial:false,contrary:false,controlled:true,offTopic:true,value:''};
}
// Résultats d'une grille pour une opération, avec les décisions appliquées (rejet = non trouvé ; thèmes recalculés).
// controls (contrôle des décisions d'ExtracTerre) l'emporte sur decisions (vérification des « Oui »).
export function verifiedSearch(o,grid,decisions={},controls={}){
  const s=o.searches?.[grid.id]; if(!s) return s;
  let changed=false; const criteria={};
  for(const [cid,x] of Object.entries(s.criteria||{})){ const k=verifKey(o,grid.id,cid), cd=controls[k]?.d;
    if(cd&&etDecision(x)){ const y=controlled(x,cd); criteria[cid]=y; if(etDecision(y)!==etDecision(x)) changed=true; continue; }
    const d=decisions[k]?.d;
    // « Non » : la phrase est bien liée à l'exigence mais ne la remplit pas → indice contraire (non conforme).
    if(d==='ko'&&(x.found||x.partial)){ criteria[cid]={...x,found:false,partial:false,contrary:true,rejected:true,value:'Non conforme (vérification)'}; changed=true; }
        // « Sans objet » : compté comme nul (ni trouvé, ni contraire), comme si la phrase n'avait pas été lue.
    else if(d==='so'&&(x.found||x.partial)){ criteria[cid]={...x,found:false,partial:false,contrary:false,rejected:true,offTopic:true,value:''}; changed=true; }
    // « Partiel » (vérifié) : comme dans ExtracTerre, ne compte pas comme trouvé pour le thème, mais n'est pas un Non.
    else if(d==='pa'&&(x.found||x.partial)){ criteria[cid]={...x,found:false,partial:true,verified:true,value:'Partiel (vérifié)'}; changed=true; }
    else criteria[cid]=d==='ok'?{...x,verified:true}:x; }
  if(!changed) return {...s,criteria};
  const themes={};
  for(const t of grid.themes||[]){ const own=(grid.criteria||[]).filter(c=>c.theme===t.name&&c.id!==t.applicableIf), req=own.filter(c=>(t.required||[]).includes(c.id)), opt=own.filter(c=>!(t.required||[]).includes(c.id));
    const applicable=!t.applicableIf||!!criteria[t.applicableIf]?.found, found=opt.filter(c=>criteria[c.id]?.found).length, requiredOk=req.every(c=>criteria[c.id]?.found);
    themes[t.name]={...(s.themes?.[t.name]||{}),applicable,found,min:t.min,required:req.length,requiredOk,ok:applicable?(requiredOk&&found>=(Number(t.min)||0)):null}; }
  return {...s,criteria,themes};
}
// Décisions effectives (propres + reprises) pour les statistiques.
export function effectiveDecisions(items){ const m={}; for(const i of items) if(i.decision) m[i.key]={d:i.decision,inherited:!!i.inherited}; return m; }
export function verificationProgress(items){ const c=d=>items.filter(i=>i.decision===d).length, ok=c('ok'), ko=c('ko'), pa=c('pa'), so=c('so'); return {total:items.length,ok,ko,pa,so,todo:items.length-ok-ko-pa-so}; }

/* Surlignage d'une preuve : passages qui ont déclenché le critère (mots-clés de la grille, kind 'k') et mots de contexte
   (kind 'c'). Même normalisation qu'ExtracTerre (sans accents ni casse, apostrophes = espace), caractère par caractère pour
   garder les positions du texte d'origine. */
const hlNorm=s=>[...String(s||'')].map(ch=>{ const b=ch.normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase(); return /[’'`]/.test(ch)?' ':(b.length===1?b:ch.toLowerCase().slice(0,1)||' '); }).join('');
const hlEsc=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
function hlRegex(term){ const t=String(term||'').trim(); if(!t) return null;
  try{ if(/^re:/i.test(t)) return new RegExp(t.slice(3),'gi');
    const body=hlNorm(t).trim().split(/[\s-]+/).map(hlEsc).join('[\\s-]+'); return new RegExp(`(?<![a-z0-9])${body}(?![a-z0-9])`,'gi'); }catch{ return null; } }
export function highlightRanges(text,crit){
  const n=hlNorm(text), out=[];
  for(const [kind,terms] of [['k',[...(crit?.keywords||[]),...(crit?.allOf||[])]],['c',crit?.context||[]]]) for(const t of terms){ const rx=hlRegex(t); if(!rx) continue; let m, guard=0;
    while((m=rx.exec(n))&&guard++<20){ if(!m[0]){ rx.lastIndex++; continue; } out.push({s:m.index,e:m.index+m[0].length,kind}); } }
  out.sort((a,b)=>a.s-b.s||(a.kind==='k'?-1:1)); const merged=[];
  for(const r of out){ const last=merged[merged.length-1]; if(last&&r.s<last.e){ if(r.kind==='k'&&last.kind==='c'){ last.kind='k'; } last.e=Math.max(last.e,r.e); } else merged.push({...r}); }
  return merged;
}

// ---------- Adresses saisies dans TotalRecoll (lignes non géolocalisées) ----------
// Une correction = {address, postcode, city, lat, lon, insee, precision, label} pour un bâtiment d'une opération d'un fichier
// ExtracTerre. Clé : « fichier|opération|bâtiment ». Écrite dans la session comme saisie manuelle (ExtracTerre ne remplace jamais
// une coordonnée saisie), et dans la tête de la session (lue par TotalRecoll).
export const GEO_FIX_FIELDS={address:'Adresse',postcode:'Code postal',city:'Ville',lat:'Latitude',lon:'Longitude',insee:'Code INSEE',precision:'Géolocalisation'};
const GEO_FIX_KEYS={address:'address',postcode:'postcode',city:'city',lat:'latitude',lon:'longitude',insee:'insee_code',precision:'geocode_precision'};
export const GEO_FIX_SOURCE='TotalRecoll — adresse saisie à la main';
export const geoFixKey=(file,op,building)=>`${file}|${op?.id||op?.contract||op?.code||''}|${building||''}`;
export function geoFixValues(fix){ return Object.fromEntries(Object.entries(GEO_FIX_KEYS).filter(([k])=>fix?.[k]!==undefined&&fix[k]!==null&&fix[k]!=='').map(([k,f])=>[f,fix[k]])); }
// Opération ExtracTerre avec les corrections de ce fichier appliquées (copie ; l'original reste intact).
export function withGeoFixes(op,file,fixes){
  let hit=false; const buildings=op.buildings.map(b=>{ const f=fixes[geoFixKey(file,op,b[BUILDING_COL])]; if(!f) return b; hit=true;
    return {...b,...Object.fromEntries(Object.entries(GEO_FIX_FIELDS).filter(([k])=>f[k]!==undefined&&f[k]!==null&&f[k]!=='').map(([k,l])=>[l,f[k]]))}; });
  return hit?{...op,buildings}:op;
}
// Réponse IGN (feature GeoJSON) → correction ; null pour un résultat trop vague (commune seule).
const GEO_TYPES={housenumber:'Adresse exacte (numéro)',street:'Voie (sans numéro)',locality:'Lieu-dit'};
export function geoFixFromFeature(f){
  const p=f?.properties||{}, [lon,lat]=f?.geometry?.coordinates||[]; if(!Number.isFinite(lat)||!Number.isFinite(lon)||!GEO_TYPES[p.type]) return null;
  return {address:p.name||'',postcode:p.postcode||'',city:p.city||'',lat:Math.round(lat*1e6)/1e6,lon:Math.round(lon*1e6)/1e6,insee:p.citycode||'',precision:`${GEO_TYPES[p.type]}, saisie TotalRecoll`,label:p.label||''};
}
// Session : tête (operations[].rows) et projet du workspace (saisies manuelles des bâtiments). byBuilding : {bâtiment: correction}.
export function patchSessionOperation(o,byBuilding){
  for(const r of o?.rows||[]){ const f=byBuilding[r.building]; if(f) Object.assign(r,geoFixValues(f)); }
  return o;
}
export function patchSessionProject(p,byBuilding,at=new Date().toISOString()){
  p.manualValues={...(p.manualValues||{})}; p.manualSources={...(p.manualSources||{})};
  for(const [b,f] of Object.entries(byBuilding)) for(const [field,v] of Object.entries(geoFixValues(f))){ const k=`${b}|${field}`;
    p.manualValues[k]=v; p.manualSources[k]={docId:'',fileName:GEO_FIX_SOURCE,page:'',excerpt:f.label||'',method:'manual:totalrecoll-geocode',provenanceNote:`Adresse saisie dans TotalRecoll le ${at.slice(0,10)}, géolocalisée par l’IGN`}; }
  return p;
}
