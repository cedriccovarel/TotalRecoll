// Données fictives uniquement.
import assert from 'node:assert/strict';
import {recoll,opsFromSession,opsFromBuildingRows,normContract,normOpe,ecartKey,withDh,DH_ZT,DH_ZNT,targetColumn,leadContract} from './recoll.js';

assert.equal(normContract(258),'00000258');
assert.equal(normContract('00004398'),'00004398');
assert.equal(normContract('OPE-000123'),'');
assert.equal(normOpe('ope 123'),'OPE-000123');
// N° de contrat en tête d'un nom d'opération / dossier Box (exports ExtracTerre sans colonne contrat).
assert.equal(leadContract('00004701 - 26 - VILLE FICTIVE - Résidence'),'00004701');
assert.equal(leadContract('00004371  00004371 - 27 - VILLE - Fictif'),'00004371');
assert.equal(leadContract('0000123'),'00000123');
assert.equal(leadContract('2373 - PARIS - FICTIF'),'');
assert.equal(leadContract('1271L'),'');
assert.equal(leadContract('Résidence 00004701'),'');
{ const ops=opsFromBuildingRows([{'Projet':'00004701 - 26 - VILLE FICTIVE','Bâtiment':'A','Cep':60},{'Projet':'00004701 - 26 - VILLE FICTIVE','Bâtiment':'B','Cep':61}],'fictif.xlsx');
  assert.equal(ops.length,1); assert.equal(ops[0].contract,'00004701'); assert.equal(ops[0].buildings.length,2); }
console.log('Contrat lu dans le nom OK');

const header=['Numéro du contrat','Opération: Code interne','Nom de l\'opération','Total logements','Cep projet','Mode constructif'];
const sf={header,rows:[
  [1001,'OPE-000001','Les Tilleuls',40,null,null],      // 1 ligne, 2 bâtiments ExtracTerre
  [1002,'OPE-000002','Le Parc',20,90,'Béton'],          // 2 lignes = 2 bâtiments (appariement)
  [1002,'OPE-000002','Le Parc',20,95,'Béton'],
  [1003,'OPE-000003','Rien',10,null,null],              // Salesforce seul
  [1004,'OPE-000004','Conflit',5,null,null],
  [1005,'OPE-000005','Trois lignes',30,80,null],         // 3 lignes Salesforce / 2 bâtiments
  [1005,'OPE-000005','Trois lignes',30,81,null],
  [1005,'OPE-000005','Trois lignes',30,82,null]
]};
const session={format:'extracterre-session',operations:[
  {identity:{contract:'00001001',code:''},rows:[{building:'A',cep:70,structure:'Bois',housing_count:25},{building:'B',cep:72,housing_count:15}]},
  {identity:{contract:'',code:'OPE-000002'},rows:[{building:'Bât 1',cep:90},{building:'Bât 2',cep:99,structure:'Béton'}]},
  {identity:{contract:'00001004',code:'OPE-000001'},rows:[{building:'X',cep:1}]},
  {identity:{contract:'00009999',code:''},rows:[{building:'Seul',cep:50}]},
  {identity:{contract:'00001005',code:'OPE-000005'},rows:[{building:'B1',cep:60},{building:'B2',cep:61}]}
]};
const ops=opsFromSession(session,'session.json');
const out=recoll(sf,ops);
const H=out.header, col=n=>H.indexOf(n), st=col('Recollement');
const rowsOf=c=>out.rows.filter(r=>r[0]===c);

// 1001 : une ligne Salesforce dupliquée par bâtiment, colonnes administratives recopiées.
const t=rowsOf('00001001');
assert.equal(t.length,2);
assert.deepEqual(t.map(r=>r[col('Bâtiment')]),['A','B']);
assert.deepEqual(t.map(r=>r[col('Cep projet')]),[70,72]);
assert.equal(t[1][col('Total logements')],40);
assert.equal(t[0][col('Nombre de logements')],25);
assert.equal(t[0][st],'Recollé (contrat)');
assert.equal(t[0][1],'OPE-000001');

// 1002 : recollé par OPE seule, appariement ligne à ligne, écart signalé sur Cep (95 → 99).
const p=rowsOf('00001002');
assert.deepEqual(p.map(r=>r[col('Cep projet')]),[90,99]);
assert.equal(p[0][st],'Recollé (OPE)');
assert.equal(out.ecarts.length,1);
assert.equal(out.ecarts[0].salesforce,95);

// 1003 conservé tel quel ; 1004 + OPE-000001 → conflit ; 9999 → ExtracTerre seul.
assert.equal(rowsOf('00001003')[0][st],'Salesforce seul');
assert.match(rowsOf('00001004').find(r=>/Conflit/.test(r[st]))[st],/Conflit/);
assert.match(rowsOf('00009999')[0][st],/^ExtracTerre seul \(absent de Tous contrats\) — session.json › /);
assert.equal(rowsOf('00001004').length,2); // la ligne Salesforce + la ligne en conflit

// 1005 : 3 lignes / 2 bâtiments → base = 1re ligne, valeurs techniques Salesforce vidées, à vérifier.
const v=rowsOf('00001005');
assert.deepEqual(v.map(r=>r[col('Cep projet')]),[60,61]);
assert.match(v[0][st],/à vérifier : 3 lignes Salesforce pour 2 bâtiments/);
assert.equal(out.stats.aVerifier,2);

// Option : la valeur Salesforce est conservée en cas d'écart.
const keep=recoll(sf,ops,{keepSalesforce:true});
assert.equal(keep.rows.filter(r=>r[0]==='00001002')[1][col('Cep projet')],95);
assert.equal(keep.ecarts[0].retenu,'Salesforce');

// Export Excel ExtracTerre (libellés) : regroupement par contrat, alias de colonnes.
const xl=opsFromBuildingRows([
  {'Code interne':'OPE-000003','Contrat: Numéro du contrat':'00001003','Bâtiment':'Unique','Cep':88,'Structure':'Bois','Maître d\'ouvrage: Nom de la société':'Ignoré'},
],'multi.xlsx');
const out2=recoll(sf,[...ops,...xl]);
const r3=out2.rows.find(r=>r[0]==='00001003');
assert.equal(r3[col('Cep projet')],88);
assert.equal(r3[col('Mode constructif')],'Bois');
assert.equal(out2.header.includes('Maître d\'ouvrage: Nom de la société'),false);

// Nature des lignes, cellules issues d'ExtracTerre, décision écart par écart.
const k=i=>out.meta[out.rows.findIndex(r=>r[0]===i)].kind;
assert.deepEqual(['00001001','00001003','00009999'].map(k),['recolle','sf','et']);
assert.ok(out.meta[out.rows.findIndex(r=>r[0]==='00001001')].et.includes(col('Cep projet')));
assert.equal(out.stats.viaContrat,1); assert.equal(out.stats.viaOpe,1); assert.equal(out.stats.conflits,1);
const dec=recoll(sf,ops,{decisions:{[out.ecarts[0].key]:'sf'}});
assert.equal(dec.rows.filter(r=>r[0]==='00001002')[1][col('Cep projet')],95);
assert.equal(dec.ecarts[0].decided,true);
assert.equal(ecartKey(out.ecarts[0]),'00001002|Bât 2|Cep projet');
assert.throws(()=>recoll({header:['x'],rows:[]},[]),/introuvables/);
assert.throws(()=>opsFromSession({},'x'),/session ExtracTerre/);
// DH ZT / DH ZNT (ExtracTerre ≥ 2.11) : DH projet = le plus défavorable ; DH sans zone → ZNT (ZT si uniquement traversant).
assert.deepEqual(withDh({[DH_ZT]:500,[DH_ZNT]:'712,5'}),{[DH_ZT]:500,[DH_ZNT]:'712,5',DH:712.5});
assert.deepEqual(withDh({DH:640}),{DH:640,[DH_ZNT]:640});
assert.deepEqual(withDh({DH:640,'Logement traversant':'Oui'}),{DH:640,'Logement traversant':'Oui',[DH_ZT]:640});
assert.equal(targetColumn('DH'),'DH projet');
const sdh=opsFromSession({format:'extracterre-session',operations:[{title:'00001001',identity:{contract:'00001001'},rows:[{building:'Bât 1',dh_zt:450,dh_znt:820,wall_ite:'Oui',cooling_active:'Non'}]}]},'s');
assert.deepEqual([sdh[0].buildings[0].DH,sdh[0].buildings[0].ITE,sdh[0].buildings[0]['Climatisation active']],[820,'Oui','Non']);
console.log('TotalRecoll — vérifications OK');

// ANRU : quartier FICTIF carré avec un trou ; commune avec / sans quartier pour les bâtiments non géolocalisés.
{
  const {anruIndex,anruOf,ANRU_COL,ANRU_QP_COL}=await import('./recoll.js');
  const sq=(a,b,c,d)=>[[a,c],[a,d],[b,d],[b,c],[a,c]];
  const idx=anruIndex({quartiers:[{qp:'QP099001',nom:'Les Fictifs',cat:'national',communes:[['99001','Saint-Fictif']],rings:[sq(45,45.01,5,5.01),sq(45.004,45.006,5.004,5.006)]}]});
  assert.equal(anruOf(idx,{lat:45.002,lon:5.002}).anru,'Oui');
  assert.equal(anruOf(idx,{lat:'45,002',lon:'5,002'}).quartier,'Les Fictifs (QP099001, intérêt national)');
  assert.equal(anruOf(idx,{lat:45.005,lon:5.005}).anru,'Non','dans le trou');
  assert.equal(anruOf(idx,{lat:45.02,lon:5.002}).anru,'Non');
  assert.equal(anruOf(idx,{lat:45.0101,lon:5.002}).anru,'Non','à 11 m du bord');
  assert.equal(anruOf(idx,{lat:45.01008,lon:5.002}).quartier,'Les Fictifs (QP099001, intérêt national), à 9 m du périmètre','à 9 m du bord : intégré');
  assert.equal(anruOf(idx,{cp:'99100',ville:'ST FICTIF CEDEX 2'}).anru,'','commune avec quartier, sans coordonnées : à vérifier');
  assert.equal(anruOf(idx,{cp:'99100',ville:'Autreville'}).anru,'Non');
  assert.equal(anruOf(idx,{cp:'',ville:''}).anru,'');
  const r=recoll({header:['Numéro du contrat','Opération: Code interne','Code postal','Ville'],rows:[[1,'OPE-000001','99100','Saint-Fictif']]},
    opsFromSession({format:'extracterre-session',operations:[{identity:{contract:'00000001'},rows:[{building:'A',latitude:45.002,longitude:5.002}]}]},'s.json'),{anru:idx});
  assert.deepEqual([r.rows[0][r.header.indexOf(ANRU_COL)],r.rows[0][r.header.indexOf(ANRU_QP_COL)],r.header.at(-1)],['Oui','Les Fictifs (QP099001, intérêt national)','Recollement']);
}
console.log('ANRU OK');

// Confort d'été (données fictives).
{
  const {confortOf,SYSTEMS}=await import('./recoll.js');
  const c=o=>confortOf(k=>o[k]);
  assert.equal(c({'Parois verticales structure':'Ossature bois'}).system,SYSTEMS.bois);
  assert.equal(c({'Mode constructif':'CLT'}).system,SYSTEMS.bois);
  assert.equal(c({'Parois verticales structure':'Brique terre cuite','ITE':'Oui'}).system,SYSTEMS.ite);
  assert.equal(c({'Parois verticales structure':'Béton armé','Parois verticales ITI type d’isolant':'PSE'}).system,SYSTEMS.iti);
  assert.equal(c({'Parois verticales structure':'Monomur 37 cm'}).system,SYSTEMS.iti);
  assert.equal(c({'Mode constructif':'Bloc béton / parpaing','Parois verticales type d’isolant':'Laine de roche en ITE'}).system,SYSTEMS.ite);
  assert.deepEqual([c({'Parois verticales structure':'Béton','ITE':'Oui','ITI':'Oui'}).system,c({'Parois verticales structure':'Béton','ITE':'Oui','ITI':'Oui'}).note],['','ITE et ITI : à trancher']);
  assert.equal(c({'Parois verticales structure':'Béton armé','ITE':'Non','Parois verticales type d’isolant':'Laine en ITE'}).note,'isolation des murs non précisée');
  assert.equal(c({'Mode constructif':'Structure métallique'}).note,'structure hors catégories : Structure métallique');
  const x=c({[DH_ZT]:'512,4',[DH_ZNT]:700,'Nombre de brasseurs d’air':0,'Refroidissement après travaux':'Split / multisplit'});
  assert.deepEqual([x.dhZt,x.dhZnt,x.fan,x.clim],[512.4,700,'Non','Oui']);
  assert.deepEqual([c({'Brasseur d’air':'Oui','Climatisation active':'Non','Refroidissement après travaux':'Split'}).fan,c({'Climatisation active':'Non','Refroidissement après travaux':'Split'}).clim,c({'Refroidissement après travaux':'Aucun'}).clim,c({}).fan],['Oui','Non','Non','']);
}
console.log('Confort d’été OK');
{
  const {confortOf,SYSTEMS}=await import('./recoll.js');
  assert.equal(confortOf(k=>({'Parois verticales structure':'Poutrelles-hourdis','Mode constructif':'Brique','Parois verticales type d’isolant':'Complexe isolant Doublissimo'})[k]).system,SYSTEMS.iti);
}
console.log('Confort d’été (2) OK');

// Valeurs étranges et tri par colonne (données fictives).
{
  const {oddValue,COLUMN_KEY,sortRowIndexes}=await import('./recoll.js');
  assert.ok(oddValue('tic',21)); assert.equal(oddValue('tic','27,4'),''); assert.ok(oddValue('dh_znt',1400)); assert.ok(oddValue('dh',2)); assert.equal(oddValue('dh_zt',600),'');
  assert.ok(oddValue('cep','450,5')); assert.equal(oddValue('cep_before',650),''); assert.ok(oddValue('heating_vector_after','Chaudière fioul')); assert.equal(oddValue('heating_vector_before','Fioul'),'');
  assert.equal(oddValue('tic',''),''); assert.equal(COLUMN_KEY[targetColumn('Tic')],'tic');
  const rows=[['12,5','b'],[3,'A'],['','c'],[100,'é']];
  assert.deepEqual(sortRowIndexes([0,1,2,3],rows,0,'asc'),[1,0,3,2],'nombres croissants, vide en fin');
  assert.deepEqual(sortRowIndexes([0,1,2,3],rows,0,'desc'),[3,0,1,2],'nombres décroissants, vide en fin');
  assert.deepEqual(sortRowIndexes([0,1,2,3],rows,1,'asc'),[1,0,2,3],'textes sans casse');
  assert.deepEqual(sortRowIndexes([0,1,2,3],rows,null,''),[0,1,2,3],'sans tri : ordre d’origine');
}
console.log('Valeurs étranges et tri OK');

// Export Salesforce brut (aucune colonne technique) : les colonnes ExtracTerre sont ajoutées sous leur nom OPERATIONS
// (« Cep projet », « Cep max »…) et reçoivent les valeurs (avant : colonnes « Cep », « Cep Ref » ajoutées mais vides).
{ const sfRaw={header:['Numéro du contrat','Opération: Code interne','Statut'],rows:[['00004794','OPE-000111','En cours']]};
  const s={format:'extracterre-session',operations:[{title:'Fictive',identity:{contract:'00004794'},rows:[{contract_number:'00004794',building:'Bât A',cep:60,cep_max:80,cepnr:55,cepnr_max:70,bbio:50,bbio_max:70,tic:26,tic_ref:30,cep_cooling:4}]}]};
  const r=recoll(sfRaw,opsFromSession(s,'s.json')), get=h=>r.rows[0][r.header.indexOf(h)];
  for(const [h,v] of [['Cep projet',60],['Cep max',80],['Cep,nr projet',55],['Cep,nr max',70],['Bbio projet',50],['Bbio max',70],['TIC projet',26],['TIC ref',30],['Cep refroidissement',4]]) assert.equal(get(h),v,h);
  assert.ok(!r.header.includes('Cep')&&!r.header.includes('Cep Ref'),'pas de colonne au nom ExtracTerre');
  assert.ok(r.meta[0].et.includes(r.header.indexOf('Cep max')),'Cep max transmis à l’Observatoire');
  console.log('Export Salesforce brut : colonnes techniques OK'); }

// ExtracTerre ≥ 2.34 : fiabilité de chaque valeur transmise (session « reliability », feuille « Fiabilité par bâtiment »).
{ const {relKey}=await import('./recoll.js');
  assert.equal(relKey('Sûre'),'sure'); assert.equal(relKey('À vérifier'),'averifier'); assert.equal(relKey('certifiee'),'certifiee'); assert.equal(relKey('x'),'');
  const sfRaw={header:['Numéro du contrat','Opération: Code interne','Statut'],rows:[['00004794','OPE-000111','En cours']]};
  const s={format:'extracterre-session',operations:[{title:'Fictive',identity:{contract:'00004794'},rows:[{contract_number:'00004794',building:'Bât A',cep:60,cep_max:80}],reliability:{'Bât A':{cep:'averifier',cep_max:'sure'}}}]};
  const r=recoll(sfRaw,opsFromSession(s,'s.json')), j=h=>r.header.indexOf(h);
  assert.equal(r.meta[0].rel[j('Cep projet')],'averifier'); assert.equal(r.meta[0].rel[j('Cep max')],'sure');
  // Export Excel : feuille de fiabilité alignée ligne à ligne
  const rows=[{'Contrat: Numéro du contrat':'00004794','Bâtiment':'Bât A','Cep':60,'Cep Ref':80}];
  const r2=recoll(sfRaw,opsFromBuildingRows(rows,'x.xlsx',[{'Cep':'Probable','Cep Ref':'Certifiée'}]));
  const j2=h=>r2.header.indexOf(h); assert.equal(r2.meta[0].rel[j2('Cep projet')],'probable'); assert.equal(r2.meta[0].rel[j2('Cep max')],'certifiee');
  console.log('Fiabilité ExtracTerre transmise OK'); }

{ // Carte confort d'été : zones climatiques et statistiques (données fictives)
  const {zoneOfDep,depOfRow,confortStats,SYSTEMS,ZONES}=await import('./recoll.js');
  const fs=await import('node:fs'); const geo=JSON.parse(fs.readFileSync(new URL('./data/zones-climatiques.json',import.meta.url),'utf8'));
  assert.equal(geo.departements.length,96); assert.ok(geo.departements.every(d=>ZONES.includes(d.zone)&&zoneOfDep(d.code)===d.zone));
  assert.equal(zoneOfDep('62'),'H1a'); assert.equal(zoneOfDep('33'),'H2c'); assert.equal(zoneOfDep('2B'),'H3'); assert.equal(zoneOfDep('971'),'');
  assert.equal(depOfRow(k=>({'Code postal':'20090'})[k]),'20'); assert.equal(zoneOfDep('20'),'H3');
  assert.equal(depOfRow(k=>({'Code postal':'','Nom de l\'opération':'00009999 - 69 - VILLE FICTIVE - TEST'})[k]),'69');
  assert.equal(depOfRow(k=>({'Nom de l\'opération':'Opération sans département'})[k]),null);
  const s=confortStats([
    {op:'A',zone:'H1a',dhZt:400,dhZnt:600,system:SYSTEMS.ite,fan:'Oui',clim:'Non'},
    {op:'A',zone:'H1a',dhZt:500,dhZnt:null,system:SYSTEMS.ite,fan:'Non',clim:''},
    {op:'B',zone:'H1a',dhZt:300,dhZnt:null,system:SYSTEMS.bois,fan:'Non',clim:'Oui'},
    {op:'C',zone:'',dhZt:900,dhZnt:null,system:'',fan:'',clim:''}]);
  const h=s.zones.H1a;
  assert.deepEqual([h.ops,h.buildings,h.zt.mean,h.zt.median,h.zt.n,h.znt.n],[2,3,400,400,3,1]);
  assert.deepEqual(h.fan,{yes:1,known:2,total:2,pct:50}); assert.deepEqual(h.clim,{yes:1,known:2,total:2,pct:50});
  { const t=confortStats([{op:'A',zone:'H3',fan:'Oui',clim:''},{op:'B',zone:'H3',fan:'',clim:'Non'},{op:'C',zone:'H3',dhZt:500,fan:'',clim:''}]).zones.H3;
    assert.deepEqual(t.fan,{yes:1,known:1,total:3,pct:33.3}); assert.deepEqual(t.clim,{yes:0,known:1,total:3,pct:0});
    // Climatisation aussi sur toutes les opérations : 1 « Oui » sur 4 (un « Non » lu, deux sans mention).
    const u=confortStats([{op:'A',zone:'H3',clim:'Oui'},{op:'B',zone:'H3',clim:''},{op:'C',zone:'H3',dhZt:400,clim:''},{op:'D',zone:'H3',clim:'Non'}]).zones.H3;
    assert.deepEqual(u.clim,{yes:1,known:2,total:4,pct:25}); }
  assert.equal(h.systems.find(x=>x.system===SYSTEMS.ite).zt.mean,450); assert.equal(s.unzoned,1); assert.equal(s.all.zt.median,450);
}
console.log('Carte confort d’été (zones climatiques) OK');

// Publication dans l'onglet OPERATIONS de l'Observatoire : fiabilité appliquée, valeur Salesforce remise si ExtracTerre l'avait remplacée.
{
  const {publishTable,CONTRACT_COL,OPE_COL,BUILDING_COL}=await import('./recoll.js');
  const header=[CONTRACT_COL,OPE_COL,'Ville','Bbio projet','Cep projet',BUILDING_COL,'Recollement'];
  const r={header,rows:[['00001234','OPE-000001','Lille',41,undefined,'Bât A','Recollé']],
    meta:[{kind:'recolle',et:[2,3,4,5],rel:{2:'sure',3:'averifier',4:'probable'}}],
    ecarts:[{rowIndex:0,colonne:'Bbio projet',salesforce:'39',retenu:'ExtracTerre'}]};
  const t=publishTable(r,'probable');
  assert.deepEqual(t.rows[0],['00001234','OPE-000001','Lille','39','','Bât A','Recollé']); // Bbio « à vérifier » → valeur Salesforce ; Cep vide → ''
  assert.equal(t.skipped,1);
  assert.deepEqual(publishTable(r,'').rows[0].slice(2,5),['Lille',41,'']);
  console.log('Publication OPERATIONS OK');
}
{ // Recherches ExtracTerre × référentiel Tous contrats (données fictives)
  const {researchFromSession,mergeResearch,researchRows,researchStats,NO_REF}=await import('./recoll.js');
  const grid={id:'g',name:'Grille',themes:[{name:'Thème',min:1}],criteria:[{id:'c1',theme:'Thème',label:'Critère 1'},{id:'c2',theme:'Thème',label:'Critère 2'}]};
  const s=(ok,f1,v)=>({themes:{'Thème':{ok,found:f1?1:0}},criteria:{c1:{found:f1,values:v},c2:{found:false,contrary:!f1}}});
  const a=researchFromSession({format:'extracterre-session',exportedAt:'2026-11-01',grids:[grid],operations:[
    {title:'00001111 - 59 - A',identity:{contract:'',owner:'MOA doc'},searches:{g:s(true,true,[10])}},
    {title:'B',identity:{contract:'2222',code:'OPE-000002'},searches:{g:s(false,false,[])}},
    {title:'C',identity:{contract:'3333'},searches:{}},
    {title:'D',identity:{contract:'9999'},searches:{g:s(true,true,[30])}}]},'p1');
  assert.equal(a.ops.length,3); assert.equal(a.ops[0].contract,'00001111');
  const b=researchFromSession({format:'extracterre-session',kind:'observatoire',exportedAt:'2026-11-10',grids:[grid],operations:[{title:'B',identity:{contract:'00002222'},searches:{g:s(true,true,[20])}}]},'p2');
  const m=mergeResearch([a,b]); assert.equal(m.ops.length,3); assert.equal(m.ops.find(o=>o.contract==='00002222').searches.g.themes['Thème'].ok,true); // plus récent
  const sf={header:['Numéro du contrat','Opération: Code interne','Référentiel','Nom de la société: Nom de la société',"Nom de l'opération"],rows:[
    ['00001111','OPE-000001','NF Habitat','MOA SF','Op A'],['00001111','OPE-000001','NF Habitat HQE','MOA SF','Op A'],['00002222','OPE-000002','','MOA 2','Op B']]};
  const rows=researchRows(m,sf);
  const A=rows.find(o=>o.contract==='00001111'); assert.deepEqual(A.refs,['NF Habitat','NF Habitat HQE']); assert.equal(A.owner,'MOA SF'); assert.equal(A.name,'Op A');
  assert.deepEqual(rows.find(o=>o.contract==='00002222').refs,['Référentiel non renseigné']);
  const D=rows.find(o=>o.contract==='00009999'); assert.deepEqual(D.refs,[NO_REF]); assert.equal(D.inSf,false);
  const st=researchStats(rows.filter(o=>o.refs.includes('NF Habitat')),grid);
  assert.equal(st.rows.length,1); assert.equal(st.themes[0].ok,1); assert.equal(st.themes[0].crit[0].ok,1); assert.equal(st.themes[0].crit[0].med,10);
  const all=researchStats(rows,grid); assert.equal(all.rows.length,3); assert.equal(all.themes[0].crit[0].med,20);
  console.log('Recherches × référentiel OK');
}

// Vérification des « Oui » des recherches (données fictives).
{ const {verificationItems,verifiedSearch,verificationProgress,verifKey}=await import('./recoll.js');
  const grid={id:'g',themes:[{name:'Parkings',min:1,required:['p_portail'],applicableIf:'p_present'}],criteria:[{id:'p_present',theme:'Parkings',label:'Parking'},{id:'p_portail',theme:'Parkings',label:'Portail'},{id:'p_badge',theme:'Parkings',label:'Badge'}]};
  const o={contract:'00009999',ope:'',name:'Fictif',searches:{g:{criteria:{p_present:{found:true,evidence:[{doc:'CCTP.pdf',page:3,line:'parking en sous-sol',ok:true}]},p_portail:{found:true,evidence:[{doc:'CCTP.pdf',page:4,line:'portail motorise parking',ok:true}]},p_badge:{found:true,evidence:[{doc:'C.pdf',page:1,line:'sans badge',ok:false},{doc:'C.pdf',page:2,line:'lecteur vigik porte parking',ok:true}]}},themes:{Parkings:{applicable:true,found:1,ok:true}}}}};
  let it=verificationItems([o],grid,{}); assert.equal(it.length,3); assert.equal(it[2].evidence.length,1,'seules les preuves qui justifient le Oui');
  const dec={[verifKey(o,'g','p_present')]:{d:'ko'}};
  const v=verifiedSearch(o,grid,dec); assert.equal(v.criteria.p_present.found,false); assert.equal(v.themes.Parkings.applicable,false,'parking rejeté → thème non applicable');
  const v2=verifiedSearch(o,grid,{[verifKey(o,'g','p_portail')]:{d:'ko'}}); assert.equal(v2.themes.Parkings.ok,false,'critère obligatoire rejeté → thème insuffisant');
  assert.equal(verifiedSearch(o,grid,{[verifKey(o,'g','p_badge')]:{d:'ok'}}).criteria.p_badge.verified,true);
  it=verificationItems([o],grid,dec); assert.deepEqual(verificationProgress(it),{total:3,ok:0,ko:1,pa:0,so:0,todo:2});
  { const v3=verifiedSearch(o,grid,{[verifKey(o,'g','p_portail')]:{d:'pa'}}); assert.equal(v3.criteria.p_portail.partial,true); assert.equal(v3.criteria.p_portail.found,false); assert.equal(v3.themes.Parkings.ok,false,'partiel ne compte pas comme trouvé');
    assert.deepEqual(verificationProgress(verificationItems([o],grid,{[verifKey(o,'g','p_badge')]:{d:'pa'}})),{total:3,ok:0,ko:0,pa:1,so:0,todo:2}); }
  { const g2={id:'g',themes:[],criteria:[{id:'fen_442',theme:'F',label:'Vitrage 44.2 des fenêtres'},{id:'hall_vitrage',theme:'H',label:'Vitrage du hall'},{id:'fen_volets',theme:'F',label:'Volets'}]};
    const L='Vitrage feuilleté 44.2 sur les ensembles menuisés'; const op2={contract:'00009998',ope:'',name:'Fictif 2',searches:{g:{criteria:{fen_442:{found:true,evidence:[{doc:'a',page:1,line:L,ok:true}]},hall_vitrage:{found:true,evidence:[{doc:'b',page:2,line:L,ok:true}]},fen_volets:{found:true,evidence:[{doc:'c',page:3,line:'Volets anti-relevage',ok:true}]}}}}};
    const dup={...op2,name:'Fictif 2 bis'}; // même contrat, chargée deux fois
    let it2=verificationItems([op2,dup],g2,{[verifKey(op2,'g','fen_442')]:{d:'ok'}});
    assert.equal(it2.filter(i=>!i.decision).length,2,'seules les cartes Volets restent (même phrase / même critère validés d’office)');
    assert.equal(it2.find(i=>i.crit.id==='hall_vitrage').inherited,'Vitrage 44.2 des fenêtres');
    it2=verificationItems([op2],g2,{[verifKey(op2,'g','fen_442')]:{d:'pa'}}); assert.equal(it2.filter(i=>!i.decision).length,2,'un Partiel ne se propage pas'); }
  { const vs=verifiedSearch(o,grid,{[verifKey(o,'g','p_badge')]:{d:'so'}}); assert.equal(vs.criteria.p_badge.found,false); assert.equal(vs.criteria.p_badge.offTopic,true);
    assert.equal(verificationProgress(verificationItems([o],grid,{[verifKey(o,'g','p_badge')]:{d:'so'}})).so,1); }
  console.log('Vérification des Oui OK'); }
{ const {highlightRanges}=await import('./recoll.js');
  const crit={keywords:['re:portail[^\\n]{0,40}(?:parking|sous[- ]sol|garage)'],context:['telecommande','badge']};
  const t='Portail coulissant motorisé du parking avec ouverture par télécommande.';
  const r=highlightRanges(t,crit); assert.deepEqual(r.map(x=>[t.slice(x.s,x.e),x.kind]),[['Portail coulissant motorisé du parking','k'],['télécommande','c']]);
  assert.deepEqual(highlightRanges('rien à voir',crit),[]);
  { const t2='Verrou automatique anti-relevage du tablier.'; assert.deepEqual(highlightRanges(t2,{keywords:['anti relevage'],context:[]}).map(x=>t2.slice(x.s,x.e)),['anti-relevage']); }
  console.log('Surlignage des preuves OK'); }
