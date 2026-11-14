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
export function etOperation(buildings,ids=[],source='',rel=[]){
  const all=[...ids,...buildings.flatMap(b=>[b['Contrat: Numéro du contrat'],b['Code interne'],b['Opération'],b['Projet'],b['Nom opération']])];
  const contract=all.map(contractOf).find(Boolean)||'';
  const code=all.map(normOpe).find(Boolean)||'';
  return {contract,code,source,buildings:buildings.map(b=>withDh(Object.fromEntries(Object.entries(b).filter(([k,v])=>TECH_SET.has(k)&&!empty(v))))),
    rel:buildings.map((b,i)=>Object.fromEntries(Object.entries(rel[i]||{}).map(([k,v])=>[k,relKey(v)]).filter(([,v])=>v)))};
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
  const overlay=(row,b,op,rowIndex,rel={},relOut=null)=>{
    const filled=[];
    for(const [label,v] of Object.entries(b)){
      const h=targetColumn(label), j=col[h]; if(j==null) continue;
      if(!empty(row[j])&&!same(row[j],v)&&h!==BUILDING_COL){
        const e={contrat:normContract(row[ci])||op.contract,ope:normOpe(row[oi])||op.code,batiment:b[BUILDING_COL]||'',colonne:h,salesforce:row[j],extracterre:v,source:op.source,rowIndex};
        e.key=ecartKey(e); const choice=decisions[e.key]||(keepSalesforce?'sf':'et');
        e.retenu=choice==='sf'?'Salesforce':'ExtracTerre'; e.decided=!!decisions[e.key]; ecarts.push(e);
        if(choice==='sf') continue;
      }
      row[j]=v; filled.push(j); if(relOut) relOut[j]=rel[label]||(label==='DH'?rel[DH_ZNT]||rel[DH_ZT]:'')||'';
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
      const rel={}; const et=overlay(base,b,m.op,rows.length,m.op.rel?.[i]||{},rel);
      let st=m.status; if(mismatch){ st+=` — à vérifier : ${k} lignes Salesforce pour ${n} bâtiments`; stats.aVerifier++; }
      base[col[STATUS_COL]]=st; rows.push(base); meta.push({kind:'recolle',check:mismatch||m.replaced,et,rel,group:g.key,source:m.op.source});
    });
    stats.etBatiments+=n;
  }
  for(const {op,kind,status} of orphans){
    kind==='conflit'?stats.conflits++:stats.etSeul++;
    (op.buildings.length?op.buildings:[{}]).forEach((b,bi)=>{
      const r=new Array(width); r[ci]=op.contract; r[oi]=op.code; const rel={}; const et=overlay(r,b,op,rows.length,op.rel?.[bi]||{},rel);
      r[col[STATUS_COL]]=`${status} — ${op.source}`; rows.push(r); meta.push({kind,check:true,et,rel,group:'',source:op.source});
    });
    stats.etBatiments+=op.buildings.length;
  }
  // Format 0000xxxx et OPE-XXXXXX partout.
  rows.forEach(r=>{ if(!empty(r[ci])) r[ci]=normContract(r[ci])||r[ci]; if(!empty(r[oi])) r[oi]=normOpe(r[oi])||r[oi]; });
  if(anru){ const h=n=>col[n]; rows.forEach(r=>{ const x=anruOf(anru,{lat:r[h('Latitude')],lon:r[h('Longitude')],cp:r[h('Code postal')],ville:r[h('Ville')]}); r[col[ANRU_COL]]=x.anru; r[col[ANRU_QP_COL]]=x.quartier; }); }
  stats.lignes=rows.length; stats.ecarts=ecarts.length;
  return {header,rows,meta,ecarts,stats,sfWidth,extra};
}

// Session ExtracTerre (.json) → opérations ; les clés internes sont traduites en libellés.
const KEY2LABEL=Object.fromEntries(ET_FIELDS);
export function opsFromSession(json,source){
  if(json?.format!=='extracterre-session'||!Array.isArray(json.operations)) throw new Error('Ce JSON n’est pas une session ExtracTerre.');
  return json.operations.map(o=>etOperation(
    (o.rows||[]).map(r=>({...Object.fromEntries(Object.entries(r).map(([k,v])=>[KEY2LABEL[k]||k,v])),'Contrat: Numéro du contrat':r.contract_number,'Code interne':r.internal_code})),
    [o.identity?.contract,o.identity?.code,o.title],`${source} › ${o.title||o.identity?.name||'sans titre'}`,
    (o.rows||[]).map(r=>Object.fromEntries(Object.entries(o.reliability?.[r.building]||{}).map(([k,v])=>[KEY2LABEL[k]||k,v])))));
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
// Brasseur / climatisation : part des opérations (une opération compte « Oui » si un de ses bâtiments l'est ; inconnue si aucun n'est renseigné).
export function confortStats(items){
  const share=(list,k)=>{ const ops=new Map(); for(const x of list){ if(!x[k]) continue; ops.set(x.op,ops.get(x.op)==='Oui'||x[k]==='Oui'?'Oui':'Non'); }
    const yes=[...ops.values()].filter(v=>v==='Oui').length; return {yes,known:ops.size,pct:ops.size?Math.round(yes/ops.size*1000)/10:null}; };
  const dh=list=>({zt:stat(list.map(x=>x.dhZt).filter(v=>v!=null)),znt:stat(list.map(x=>x.dhZnt).filter(v=>v!=null))});
  const one=list=>({ops:new Set(list.map(x=>x.op)).size,buildings:list.length,...dh(list),fan:share(list,'fan'),clim:share(list,'clim'),
    systems:Object.values(SYSTEMS).map(s=>({system:s,...dh(list.filter(x=>x.system===s)),buildings:list.filter(x=>x.system===s).length}))});
  return {all:one(items),zones:Object.fromEntries(ZONES.map(z=>[z,one(items.filter(x=>x.zone===z))])),unzoned:items.filter(x=>!x.zone).length};
}
