// Données fictives uniquement.
import assert from 'node:assert/strict';
import {recoll,opsFromSession,opsFromBuildingRows,normContract,normOpe,ecartKey,withDh,DH_ZT,DH_ZNT,targetColumn} from './recoll.js';

assert.equal(normContract(258),'00000258');
assert.equal(normContract('00004398'),'00004398');
assert.equal(normContract('OPE-000123'),'');
assert.equal(normOpe('ope 123'),'OPE-000123');

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
