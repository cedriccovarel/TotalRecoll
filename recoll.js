// Recollement Salesforce (tableau « Tous contrats », 1 ligne = 1 bâtiment) × ExtracTerre (données techniques par bâtiment).
import {ET_FIELDS} from './fields.js';

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
const TECH_LABELS=[...ET_FIELDS.map(f=>f[1]),'Tags'];
const TECH_SET=new Set(TECH_LABELS);

export function normContract(v){
  const m=String(v??'').match(/^\s*0*(\d{1,8})(?:\.0+)?\s*$/);
  return m&&m[1]!=='0'?m[1].padStart(8,'0'):'';
}
export function normOpe(v){
  const m=String(v??'').match(/OPE[\s_-]*(\d{1,6})\b/i);
  return m?'OPE-'+m[1].padStart(6,'0'):'';
}
const empty=v=>v==null||String(v).trim()===''||v==='#N/A';
function same(a,b){
  const na=Number(String(a).replace(',','.')), nb=Number(String(b).replace(',','.'));
  if(Number.isFinite(na)&&Number.isFinite(nb)) return Math.abs(na-nb)<=1e-6*Math.max(1,Math.abs(na));
  if(a instanceof Date||b instanceof Date) return +new Date(a)===+new Date(b);
  return String(a).trim().toLowerCase()===String(b).trim().toLowerCase();
}

// Une opération ExtracTerre : {contract, code, source, buildings:[{libellé:valeur}]}. Les identifiants sont cherchés partout.
export function etOperation(buildings,ids=[],source=''){
  const all=[...ids,...buildings.flatMap(b=>[b['Contrat: Numéro du contrat'],b['Code interne'],b['Opération'],b['Projet']])];
  const contract=all.map(normContract).find(Boolean)||'';
  const code=all.map(normOpe).find(Boolean)||'';
  return {contract,code,source,buildings:buildings.map(b=>Object.fromEntries(Object.entries(b).filter(([k,v])=>TECH_SET.has(k)&&!empty(v))))};
}

// sf = {header:[], rows:[[]]} ; etOps = [etOperation] (la dernière source chargée l'emporte sur une même opération).
// keepSalesforce : en cas d'écart, la valeur déjà présente dans Tous contrats est conservée (sinon ExtracTerre l'emporte).
export function recoll(sf,etOps,{keepSalesforce=false}={}){
  const header=[...sf.header];
  const ci=header.indexOf(CONTRACT_COL), oi=header.indexOf(OPE_COL);
  if(ci<0||oi<0) throw new Error(`Colonnes « ${CONTRACT_COL} » et « ${OPE_COL} » introuvables dans le tableau Tous contrats.`);
  // Groupes Salesforce : lignes consécutives ou non partageant le même contrat (sinon la même OPE).
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
    let g=gc||go, status;
    if(gc&&go&&gc!==go){ orphans.push({op,status:`Conflit : contrat ${op.contract} et ${op.code} pointent vers deux opérations Salesforce`}); continue; }
    if(g&&op.contract&&g.contracts.size&&!g.contracts.has(op.contract)){ orphans.push({op,status:`Conflit : ${op.code} est rattachée au contrat ${[...g.contracts][0]} dans Salesforce, pas ${op.contract}`}); continue; }
    if(g&&op.code&&g.opes.size&&!g.opes.has(op.code)){ orphans.push({op,status:`Conflit : contrat ${op.contract} porte ${[...g.opes][0]} dans Salesforce, pas ${op.code}`}); continue; }
    if(!g){ orphans.push({op,status:op.contract||op.code?'ExtracTerre seul (absent de Tous contrats)':'ExtracTerre seul (ni contrat ni OPE trouvés)'}); continue; }
    status=gc&&go?'Recollé (contrat + OPE)':gc?'Recollé (contrat)':'Recollé (OPE)';
    if(matched.has(g)) status+=' — remplace une autre source ExtracTerre';
    matched.set(g,{op,status});
  }
  // Colonnes : celles du tableau, puis les champs ExtracTerre sans équivalent qui ont une valeur.
  const target=l=>ALIAS[l]||l;
  const used=new Set(etOps.flatMap(op=>op.buildings.flatMap(Object.keys)));
  const extra=TECH_LABELS.filter(l=>used.has(l)&&!header.includes(target(l))&&l!==BUILDING_COL);
  header.push(BUILDING_COL,...extra,STATUS_COL);
  const col=Object.fromEntries(header.map((h,i)=>[h,i]));
  const techCols=new Set(TECH_LABELS.map(target).filter(h=>h in col));
  const rows=[], ecarts=[], stats={recolles:0,salesforceSeul:0,etSeul:0,conflits:0,lignes:0,ecarts:0,aVerifier:0};
  const width=header.length;
  const overlay=(row,b,op)=>{
    for(const [label,v] of Object.entries(b)){
      const h=target(label); const j=col[h]; if(j==null) continue;
      if(!empty(row[j])&&!same(row[j],v)&&h!==BUILDING_COL){
        ecarts.push({contrat:op.contract,ope:op.code,batiment:b[BUILDING_COL]||'',colonne:h,salesforce:row[j],extracterre:v,retenu:keepSalesforce?'Salesforce':'ExtracTerre',source:op.source});
        if(keepSalesforce) continue;
      }
      row[j]=v;
    }
  };
  for(const g of groups.values()){
    const m=matched.get(g);
    if(!m||!m.op.buildings.length){
      g.rows.forEach(i=>{const r=sf.rows[i].slice(0,sf.header.length); r.length=width; r[col[STATUS_COL]]=m?'Recollé — aucun bâtiment ExtracTerre':'Salesforce seul'; rows.push(r);});
      stats.salesforceSeul+=m?0:1; continue;
    }
    stats.recolles++;
    const k=g.rows.length, n=m.op.buildings.length;
    m.op.buildings.forEach((b,i)=>{
      // Lignes Salesforce déjà découpées par bâtiment : appariement dans l'ordre ; sinon la 1re ligne sert de base.
      const base=sf.rows[g.rows[k===n?i:0]].slice(0,sf.header.length); base.length=width;
      if(k>1&&k!==n) techCols.forEach(j=>{base[col[j]]=undefined;});
      overlay(base,b,m.op);
      let st=m.status; if(k>1&&k!==n){st+=` — à vérifier : ${k} lignes Salesforce pour ${n} bâtiments`; stats.aVerifier++;}
      base[col[STATUS_COL]]=st; rows.push(base);
    });
  }
  for(const {op,status} of orphans){
    status.startsWith('Conflit')?stats.conflits++:stats.etSeul++;
    (op.buildings.length?op.buildings:[{}]).forEach(b=>{
      const r=new Array(width); r[ci]=op.contract; r[oi]=op.code; overlay(r,b,op); r[col[STATUS_COL]]=`${status} — ${op.source}`; rows.push(r);
    });
  }
  // Format 0000xxxx et OPE-XXXXXX partout.
  rows.forEach(r=>{ if(!empty(r[ci])) r[ci]=normContract(r[ci])||r[ci]; if(!empty(r[oi])) r[oi]=normOpe(r[oi])||r[oi]; });
  stats.lignes=rows.length; stats.ecarts=ecarts.length;
  return {header,rows,ecarts,stats};
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
