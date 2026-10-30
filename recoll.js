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
function same(a,b){
  const na=Number(String(a).replace(',','.')), nb=Number(String(b).replace(',','.'));
  if(Number.isFinite(na)&&Number.isFinite(nb)) return Math.abs(na-nb)<=1e-6*Math.max(1,Math.abs(na));
  return String(a).trim().toLowerCase()===String(b).trim().toLowerCase();
}

// Une opération ExtracTerre : {contract, code, source, buildings:[{libellé:valeur}]}. Les identifiants sont cherchés partout.
export function etOperation(buildings,ids=[],source=''){
  const all=[...ids,...buildings.flatMap(b=>[b['Contrat: Numéro du contrat'],b['Code interne'],b['Opération'],b['Projet']])];
  const contract=all.map(normContract).find(Boolean)||'';
  const code=all.map(normOpe).find(Boolean)||'';
  return {contract,code,source,buildings:buildings.map(b=>withDh(Object.fromEntries(Object.entries(b).filter(([k,v])=>TECH_SET.has(k)&&!empty(v)))))};
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
  const extra=TECH_LABELS.filter(l=>used.has(l)&&!header.includes(targetColumn(l))&&l!==BUILDING_COL);
  const sfWidth=header.length;
  header.push(BUILDING_COL,...extra,...(anru?[ANRU_COL,ANRU_QP_COL]:[]),STATUS_COL);
  const col=Object.fromEntries(header.map((h,i)=>[h,i]));
  const techCols=[...new Set(TECH_LABELS.map(targetColumn).filter(h=>h in col))];
  const rows=[], meta=[], ecarts=[];
  const stats={recolles:0,salesforceSeul:0,etSeul:0,conflits:0,lignes:0,ecarts:0,aVerifier:0,etOperations:unique.size,etBatiments:0,viaContrat:0,viaOpe:0,viaDeux:0};
  const width=header.length;
  // Écrit les valeurs d'un bâtiment ExtracTerre dans la ligne ; renvoie les colonnes remplies par ExtracTerre.
  const overlay=(row,b,op,rowIndex)=>{
    const filled=[];
    for(const [label,v] of Object.entries(b)){
      const h=targetColumn(label), j=col[h]; if(j==null) continue;
      if(!empty(row[j])&&!same(row[j],v)&&h!==BUILDING_COL){
        const e={contrat:normContract(row[ci])||op.contract,ope:normOpe(row[oi])||op.code,batiment:b[BUILDING_COL]||'',colonne:h,salesforce:row[j],extracterre:v,source:op.source,rowIndex};
        e.key=ecartKey(e); const choice=decisions[e.key]||(keepSalesforce?'sf':'et');
        e.retenu=choice==='sf'?'Salesforce':'ExtracTerre'; e.decided=!!decisions[e.key]; ecarts.push(e);
        if(choice==='sf') continue;
      }
      row[j]=v; filled.push(j);
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
      const et=overlay(base,b,m.op,rows.length);
      let st=m.status; if(mismatch){ st+=` — à vérifier : ${k} lignes Salesforce pour ${n} bâtiments`; stats.aVerifier++; }
      base[col[STATUS_COL]]=st; rows.push(base); meta.push({kind:'recolle',check:mismatch||m.replaced,et,group:g.key,source:m.op.source});
    });
    stats.etBatiments+=n;
  }
  for(const {op,kind,status} of orphans){
    kind==='conflit'?stats.conflits++:stats.etSeul++;
    (op.buildings.length?op.buildings:[{}]).forEach(b=>{
      const r=new Array(width); r[ci]=op.contract; r[oi]=op.code; const et=overlay(r,b,op,rows.length);
      r[col[STATUS_COL]]=`${status} — ${op.source}`; rows.push(r); meta.push({kind,check:true,et,group:'',source:op.source});
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
    [o.identity?.contract,o.identity?.code,o.title],`${source} › ${o.title||o.identity?.name||'sans titre'}`));
}
// Export Excel ExtracTerre (feuille « Données par bâtiment », lignes déjà en libellés) → opérations groupées par contrat/OPE.
export function opsFromBuildingRows(rows,source){
  const groups=new Map();
  rows.forEach((r,i)=>{
    const c=normContract(r['Contrat: Numéro du contrat'])||normContract(r['Projet'])||normContract(r['Opération']);
    const o=normOpe(r['Code interne']);
    const key=c||o||'row'+i;
    (groups.get(key)||groups.set(key,[]).get(key)).push(r);
  });
  return [...groups.values()].map(b=>etOperation(b,[],`${source} › ${b[0]['Nom opération']||b[0]['Projet']||'sans titre'}`));
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
