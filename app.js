/* Neurosurg Clinic – archivio pazienti e visite (IndexedDB, solo su questo dispositivo), esame neurologico, report, esportazione Excel e backup */
(function(){
'use strict';
var NX=window.NX;
var DB={pat:[],vis:[]},P=null,V=null,S={},NEWV={},HQ='';
var esc=function(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})};
var cap=function(s){s=String(s);return s.charAt(0).toUpperCase()+s.slice(1)};
var lc=function(s){s=String(s);return s.charAt(0).toLowerCase()+s.slice(1)};
var lcw=function(s){s=String(s);return /^[A-ZÀ-Ý][a-zà-ÿ]/.test(s)?lc(s):s};   // non tocca sigle come ACDF, GCS
var trimDot=function(s){return String(s).replace(/[.;\s]+$/,'')};
var $=function(id){return document.getElementById(id)};
var EMPTY='Nessun dato inserito: compila almeno una sezione.';
function td(){var d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')}
function dateIt(v){if(!v)return '';var p=String(v).split('-');return p.length===3?p[2]+'/'+p[1]+'/'+p[0]:v}
function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,8)}
function ageAt(nasc,at){if(!nasc)return NaN;var b=nasc.split('-').map(Number),a=(at||td()).split('-').map(Number);
  var y=a[0]-b[0];if(a[1]<b[1]||(a[1]===b[1]&&a[2]<b[2]))y--;return y}
function pname(p){return ((p.cognome||'')+' '+(p.nome||'')).trim()||'(senza nome)'}

/* ---------- archivio (IndexedDB) ---------- */
var idb=null;
function openDB(){return new Promise(function(res){
  try{var r=indexedDB.open('neurosurg-exam',1);
    r.onupgradeneeded=function(){var d=r.result;['pat','vis'].forEach(function(s){if(!d.objectStoreNames.contains(s))d.createObjectStore(s,{keyPath:'id'})})};
    r.onsuccess=function(){idb=r.result;res()};r.onerror=r.onblocked=function(){res()};
  }catch(e){res()}})}
function getAll(s){return new Promise(function(res){if(!idb)return res([]);
  try{var q=idb.transaction(s).objectStore(s).getAll();q.onsuccess=function(){res(q.result||[])};q.onerror=function(){res([])}}catch(e){res([])}})}
function put(s,o){if(!idb)return;try{var tx=idb.transaction(s,'readwrite');tx.objectStore(s).put(o);tx.oncomplete=savedMark;tx.onerror=storeErr}catch(e){storeErr()}}
function savedMark(){var el=$('saved'),d=new Date();if(!el)return;el.className='saved';
  el.textContent='✓ Salvato sull\'iPad alle '+String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0')}
function savingMark(){var el=$('saved');if(el){el.className='saved pending';el.textContent='Salvataggio…'}}
function del(s,id){if(!idb)return;try{idb.transaction(s,'readwrite').objectStore(s).delete(id)}catch(e){storeErr()}}
function banner(t){var b=$('banner');b.hidden=!t;b.textContent=t||''}
function storeErr(){banner('Errore di salvataggio: esporta subito un backup.')}
var tv=0,tp=0;
function saveV(){if(!V)return;V.updated=Date.now();savingMark();clearTimeout(tv);tv=setTimeout(function(){tv=0;put('vis',V)},300)}
function saveP(){if(!P)return;P.updated=Date.now();savingMark();clearTimeout(tp);tp=setTimeout(function(){tp=0;put('pat',P)},300)}
function flush(){if(tv){clearTimeout(tv);tv=0;if(V)put('vis',V)}if(tp){clearTimeout(tp);tp=0;if(P)put('pat',P)}}
function meta(){try{return JSON.parse(localStorage.getItem('nx-meta')||'{}')}catch(e){return {}}}
function setMeta(o){try{localStorage.setItem('nx-meta',JSON.stringify(Object.assign(meta(),o)))}catch(e){}}
function patOf(id){return DB.pat.filter(function(p){return p.id===id})[0]}
function visitsOf(pid){return DB.vis.filter(function(v){return v.pid===pid}).sort(function(a,b){
  return (b.data['d.date']||'').localeCompare(a.data['d.date']||'')||b.created-a.created})}

/* ---------- campi definiti in fields.js ---------- */
var CAT={},FIELD={},GROUPS=[];
NX.CAT.forEach(function(c){CAT[c.id]=c;if(GROUPS.indexOf(c.g)<0)GROUPS.push(c.g);
  c.f.forEach(function(f){if(f[2]!=='h'){f[0]='x.'+c.id+'.'+f[0];FIELD[f[0]]=f}})});
[NX.ANAM,NX.QUAD,NX.TER,NX.DEC].forEach(function(l){l.forEach(function(f){FIELD[f[0]]=f})});
function vt(id,i){var t=NX.VT.filter(function(x){return x[0]===id})[0];return t?t[i]:''}
function fl(f){return f[4]||''}
function fvis(f){var x=fl(f),t=S['d.tipo'];
  if(x.indexOf('P')>-1&&t==='pre')return false;if(x.indexOf('F')>-1&&t!=='fu')return false;if(x.indexOf('D')>-1&&t!=='dim')return false;return true}
function calcVal(f,s,p){try{var r=f[3](s,{age:ageAt(p&&p.nascita,s['d.date'])});
  return r===undefined||r===null||r===''||(typeof r==='number'&&!isFinite(r))?'':r}catch(e){return ''}}

/* ---------- componenti ---------- */
function seg(k,opts,multi){
  return '<div class="seg" role="group" data-k="'+k+'"'+(multi?' data-multi="1"':'')+'>'+opts.map(function(o){
    var a=typeof o==='string'?[o,o,'']:o, v=a[0], on=multi?(S[k]||[]).indexOf(v)>-1:S[k]===v;
    return '<button type="button" data-v="'+esc(v)+'" data-t="'+(a[2]||'')+'" aria-pressed="'+on+'">'+esc(a[1])+'</button>';
  }).join('')+'</div>';
}
function sel(k,opts,al,ph){
  return '<select data-k="'+k+'" aria-label="'+esc(al||k)+'"><option value="">'+(ph||'–')+'</option>'+opts.map(function(o){
    var a=typeof o==='string'?[o,o]:o;
    return '<option value="'+esc(a[0])+'"'+(S[k]===a[0]?' selected':'')+'>'+esc(a[1])+'</option>';
  }).join('')+'</select>';
}
function inp(k,ph,al,type,extra){
  return '<input type="'+(type||'text')+'" data-k="'+k+'" placeholder="'+esc(ph||'')+'" aria-label="'+esc(al||ph||k)+'" value="'+esc(S[k]||'')+'" '+(extra||'')+'>';
}
function area(k,ph,al){return '<textarea data-k="'+k+'" placeholder="'+esc(ph||'')+'" aria-label="'+esc(al||k)+'">'+esc(S[k]||'')+'</textarea>'}
function field(label,html){return '<div class="field"><span>'+label+'</span>'+html+'</div>'}
function two(label,sub,base,mk,extra){
  return '<div class="row2"><div class="lbl"><b>'+label+'</b>'+(sub?'<small>'+sub+'</small>':'')+'</div>'+
    ['dx','sx'].map(function(s){return '<div class="cell"><span class="side">'+(s==='dx'?'Dx':'Sx')+'</span>'+mk(base+'.'+s,label+(s==='dx'?' destra':' sinistra'),s)+'</div>'}).join('')+(extra||'')+'</div>';
}
function head(id,title,desc,norm){
  return '<div class="ph"><div><h2>'+title+'</h2><p class="desc">'+desc+'</p></div>'+(norm?'<button class="btn" type="button" data-norm="'+id+'">Tutto nella norma</button>':'')+'</div>'+
    (GUIDE[id]?'<details class="guide"><summary>Come si esegue</summary><dl>'+GUIDE[id].map(function(g){return '<dt>'+g[0]+'</dt><dd>'+g[1]+'</dd>'}).join('')+'</dl></details>':'');
}
function catSel(attr,cur){
  return '<select '+attr+' aria-label="Patologia"><option value="">Scegli la patologia…</option>'+GROUPS.map(function(g){
    return '<optgroup label="'+esc(g)+'">'+NX.CAT.filter(function(c){return c.g===g}).map(function(c){
      return '<option value="'+c.id+'"'+(cur===c.id?' selected':'')+'>'+esc(c.l)+'</option>'}).join('')+'</optgroup>'}).join('')+'</select>';
}
function ctl(f){var k=f[0],l=f[1],t=f[2],o=f[3];
  if(t==='s')return seg(k,o);if(t==='m')return seg(k,o,true);if(t==='l')return sel(k,o,l);
  if(t==='n')return inp(k,'',l,'number','inputmode="decimal" step="any"');if(t==='d')return inp(k,'',l,'date');
  if(t==='a')return area(k,o||'',l);
  if(t==='c'){var r=calcVal(f,S,P);return '<output class="calc" data-calc="'+k+'">'+esc(r===''?'–':r)+'</output>'}
  return inp(k,o||'',l)}
function renderFields(list){
  var vis=list.filter(function(f){return f[2]==='h'||fvis(f)});
  vis=vis.filter(function(f,i){return f[2]!=='h'||(vis[i+1]&&vis[i+1][2]!=='h')});
  var h='',buf=[];
  function out(){if(buf.length){h+='<div class="grid">'+buf.join('')+'</div>';buf=[]}}
  vis.forEach(function(f){
    if(f[2]==='h'){out();h+='<h3>'+esc(f[1])+'</h3>';return}
    var big=f[2]==='a'||fl(f).indexOf('W')>-1||((f[2]==='s'||f[2]==='m')&&f[3].length>3);
    var html=field(esc(f[1]),ctl(f));
    if(big){out();h+=html}else buf.push(html)});
  out();return '<div class="fields">'+h+'</div>'}

/* guide "Come si esegue" – riviste sui testi di riferimento (Neurological Examination Made Easy, Oxford Handbook of Neurology, Problem Based Neurosurgery, Plum & Posner) */
var GUIDE={
 cos:[
  ['GCS – Occhi (E)','4 spontanea · 3 alla voce · 2 al dolore · 1 nessuna.'],
  ['GCS – Verbale (V)','5 orientato · 4 confuso · 3 parole inappropriate · 2 suoni incomprensibili · 1 nessuna.'],
  ['GCS – Motoria (M)','6 esegue ordini · 5 localizza il dolore · 4 retrae · 3 flessione anomala (decorticazione) · 2 estensione (decerebrazione) · 1 nessuna. Stimolo centrale (trapezio, sopraorbitario) per la localizzazione; periferico (letto ungueale) per la flessione.'],
  ['GCS – Come annotarlo','Riporta sempre E, V e M separati: il totale da solo è ambiguo. Coma = non apre gli occhi, non parla, non esegue ordini (in genere GCS ≤8). Componente non valutabile: C (occhi chiusi per edema), T (intubato), D (afasia); non segnare 1.'],
  ['Paziente in coma: tronco encefalico','Pupille, riflesso corneale, oculocefalico (occhi di bambola; non eseguirlo se possibile instabilità cervicale), tosse e respiro. Nel paziente intubato il FOUR score aggiunge riflessi del tronco e respiro.'],
  ['Orientamento','Chiedi data e giorno (tempo), dove si trova (spazio), nome ed età (persona).'],
  ['Linguaggio','Ascolta l\'eloquio spontaneo; fai denominare 2 oggetti, ripetere una frase, eseguire un ordine semplice e uno a più passaggi.'],
  ['Neglect','Tocca contemporaneamente i due lati (estinzione); fai bisecare una linea o descrivere la stanza.']],
 nc:[
  ['I – Olfatto','Odore noto (caffè) su ciascuna narice a occhi chiusi. Raramente eseguito di routine.'],
  ['II – Visus e campo visivo','Visus per occhio (lettura, conta dita). Campo visivo per confronto: di fronte al paziente, un occhio coperto, il paziente fissa il tuo occhio; dita mosse nei 4 quadranti, poi un target rosso portato dalla periferia verso il centro.'],
  ['Campo visivo nelle lesioni sellari','Confronta il target rosso ai due lati del meridiano verticale: desaturazione o perdita nei campi temporali di entrambi gli occhi = compressione chiasmatica.'],
  ['Pupille','Diametro in mm e forma; luce diretta e consensuale per ciascun occhio. Anisocoria con midriasi areattiva = urgenza (III NC).'],
  ['III–IV–VI – Motilità oculare','Segui il dito disegnando una H, testa ferma. Chiedi se vede doppio; osserva ptosi e nistagmo nelle posizioni estreme.'],
  ['V – Trigemino','Tocca fronte, guancia e mento ai due lati (V1–V2–V3). Fai serrare i denti palpando i masseteri. Riflesso corneale: tocca la cornea (non la sclera) di lato con un filo di cotone; atteso ammiccamento bilaterale.'],
  ['VII – Facciale','Alza le sopracciglia, chiudi forte gli occhi, mostra i denti, gonfia le guance. Centrale = risparmia la fronte; periferico = coinvolge tutta l\'emifaccia.'],
  ['VIII – Acustico','Sfrega le dita vicino a ciascun orecchio; se alterato, Weber e Rinne con diapason.'],
  ['IX–X','Fai dire "A": velo simmetrico, ugola centrale (devia verso il lato sano). Valuta voce e deglutizione.'],
  ['XI – Accessorio','Alzare le spalle e ruotare il capo contro resistenza.'],
  ['XII – Ipoglosso','Lingua fuori: devia verso il lato leso; osserva trofismo e fascicolazioni.']],
 mot:[
  ['Tono','Paziente rilassato: mobilizza passivamente gomito, polso e ginocchio, prima lentamente poi rapidamente. Spasticità = resistenza che aumenta con la velocità; rigidità = resistenza costante.'],
  ['Mingazzini I (arti superiori)','Braccia tese in avanti, palmi in alto, occhi chiusi, per 20–30 s; puoi dare un colpetto verso il basso. Pronazione o slivellamento di un braccio = deficit piramidale lieve.'],
  ['Mingazzini II (arti inferiori)','Supino, anche e ginocchia flesse a 90°, gambe sollevate, per 20–30 s. Positivo se una gamba cade o slivella.'],
  ['Scala MRC','0 nessuna contrazione · 1 contrazione visibile senza movimento · 2 movimento a gravità eliminata · 3 vince la gravità ma non la resistenza · 4 vince una resistenza moderata (spesso distinto in 4−, 4, 4+) · 5 forza normale.'],
  ['Tecnica','Stabilizza il segmento prossimale, chiedi il movimento e opponi resistenza; confronta subito con il lato opposto. Se il dolore limita lo sforzo, segna "su base antalgica".'],
  ['Mano mielopatica','Finger escape: dita tese e addotte, il mignolo (poi l\'anulare) si abduce e flette. Grip and release: aprire e chiudere il pugno più rapidamente possibile per 10 s; meno di 20 cicli suggerisce mielopatia.']],
 rif:[
  ['Tecnica generale','Paziente rilassato, segmento in posizione intermedia. Colpo secco con il martelletto lasciato oscillare dal polso; confronta subito con il controlaterale.'],
  ['Bicipitale (C5–C6)','Gomito flesso a 90°, appoggia il pollice sul tendine del bicipite e percuoti il pollice.'],
  ['Stiloradiale (C5–C6)','Avambraccio semiflesso in posizione neutra; percuoti il radio 3–5 cm sopra il polso. Se manca la flessione del gomito e compare solo la flessione delle dita = riflesso invertito (lesione a livello C5–C6).'],
  ['Tricipitale (C7)','Braccio sostenuto, gomito flesso a 90°; percuoti il tendine subito sopra l\'olecrano.'],
  ['Flessori delle dita (C8)','Il paziente aggancia le dita semiflesse alle tue; percuoti le tue dita.'],
  ['Rotuleo (L3–L4)','Seduto con gambe penzoloni, o supino con ginocchio sostenuto in lieve flessione; percuoti il tendine sotto la rotula.'],
  ['Achilleo (S1–S2)','Porta il piede in lieve dorsiflessione passiva e percuoti il tendine d\'Achille. Se dubbio: paziente in ginocchio su una sedia con i piedi fuori dal bordo.'],
  ['Manovra di Jendrassik','Se i ROT sono ipoevocabili. Arti inferiori: il paziente aggancia le mani e tira mentre percuoti. Arti superiori: fai serrare i denti o stringere il pugno controlaterale.'],
  ['Addominali e cremasterico','Addominali superiori (T8–T10) e inferiori (T10–T12): stimolo strisciato verso l\'ombelico. Cremasterico (L1–L2): faccia mediale della coscia.'],
  ['Babinski','Con punta smussa striscia il margine laterale della pianta dal tallone verso le dita, curvando sotto le teste metatarsali. Estensione dell\'alluce, con o senza apertura a ventaglio delle dita = positivo. Alternativa: Chaddock (margine laterale del dorso del piede). Attenzione: se l\'estensore dell\'alluce è paralizzato l\'alluce non può estendersi.'],
  ['Hoffmann','Mano rilassata, polso in lieve estensione: tieni il dito medio e fai uno scatto sull\'unghia. Flessione di pollice e indice = positivo; significativo se asimmetrico o con altri segni piramidali.'],
  ['Clono','Dorsiflessione rapida del piede, mantenuta: più di 3 battiti = patologico. Clono rotuleo: spingi bruscamente la rotula verso il basso a gamba estesa; sempre patologico.']],
 sens:[
  ['Tattile e dolorifica','Cotone e punta smussa (spillo monouso); confronta destra/sinistra e prossimale/distale, a occhi chiusi.'],
  ['Dermatomeri di riferimento','C6 pollice · C7 dito medio · C8 mignolo · T4 capezzolo · T10 ombelico · L4 malleolo mediale · L5 dorso del piede e alluce · S1 margine laterale del piede · S2–S4 sella.'],
  ['Sensibilità sellare','Perineo e regione perianale. Da testare sempre se sospetti cauda equina.'],
  ['Propriocezione e vibratoria','Muovi l\'alluce su/giù a occhi chiusi; diapason 128 Hz sul malleolo.'],
  ['Segno di Lhermitte','Flessione del capo: scossa elettrica lungo il dorso o gli arti. Mielopatia cervicale, demielinizzazione.'],
  ['Coordinazione','Indice-naso, calcagno-ginocchio, movimenti alternati rapidi delle mani.'],
  ['Romberg','Piedi uniti, braccia lungo i fianchi: occhi aperti, poi chiusi per 30 s; stai pronto a sostenerlo. Se cade già a occhi aperti il test non è eseguibile. Positivo = deficit propriocettivo o vestibolare; non è positivo nella sindrome cerebellare.'],
  ['Marcia','Osserva una camminata di 10 m andata e ritorno con giro. Poi punte (S1), talloni (L4–L5) e tandem.'],
  ['Marce tipiche','Parkinsoniana: a piccoli passi, festinante, tronco flesso, pendolarità ridotta, girata "en bloc". Magnetica (idrocefalo normoteso): piedi che sembrano incollati, base allargata, difficoltà ad avviarsi e nelle rotazioni, mentre a letto il movimento di pedalare è conservato. Paraparetica spastica (a forbice): rigida, punte che strisciano (mielopatia, lesione parasagittale). Atassica cerebellare: base allargata, devia verso il lato della lesione. Atassica sensitiva: tallonante, peggiora a occhi chiusi. Falciante: arto spastico che ruota in fuori. Steppante: piede cadente (L5, SPE).'],
  ['Idrocefalo normoteso: test cronometrati','Test del cammino di 10 m (tempo e numero di passi) e Timed Up and Go, prima e dopo tap test (sottrazione di circa 30–50 mL di liquor). Annota i valori nella scheda della patologia.']],
 seg:[
  ['Rigidità nucale','Flessione passiva del capo verso lo sterno: resistenza o dolore = positivo. Non eseguire se possibile instabilità cervicale (trauma, artrite reumatoide).'],
  ['Kernig e Brudzinski','Kernig: anca a 90°, estendi il ginocchio → dolore o resistenza; se unilaterale pensa a una radicolopatia. Brudzinski: flettendo il capo si flettono le anche.'],
  ['Lasègue (L4–S1)','Supino, solleva l\'arto esteso dal tallone. Positivo se compare dolore radicolare sotto il ginocchio tra 30° e 60°, accentuato dalla dorsiflessione del piede: annota l\'angolo. Sotto 30° è aspecifico; oltre 60° pensa a sacroiliaca o ischiocrurali.'],
  ['Lasègue crociato','Solleva l\'arto sano: positivo se compare dolore nell\'arto affetto. Suggerisce ernia voluminosa o frammento espulso.'],
  ['Wassermann (L2–L4)','Prono, ginocchio flesso al massimo, estendi l\'anca: dolore alla faccia anteriore della coscia. Diagnosi differenziale: ascesso o ematoma dello psoas.'],
  ['Sfinteri','Chiedi di ritenzione, incontinenza e urgenza. Se sospetti cauda equina: sensibilità sellare, tono e contrazione volontaria all\'esplorazione rettale, riflesso anale (S4–S5), residuo post-minzionale.'],
  ['Cauda equina: allarme','Dolore o deficit bilaterali, anestesia a sella, ritenzione o incontinenza, ipotono anale, riflesso anale assente: urgenza chirurgica.']]
};

/* ---------- dati di struttura dell'esame ---------- */
var N=[['Normale','Norma','ok'],['Alterato','Alterato','alt']];
var NC=[
 ['1','NC I','Olfatto','Olfatto'],['2','NC II','Visus, campo visivo, fundus','Visus/campo visivo'],
 ['p','NC II–III','Pupille e riflesso fotomotore','Pupille'],['oc','NC III–IV–VI','Motilità oculare, diplopia, ptosi, nistagmo','Motilità oculare'],
 ['5','NC V','Sensibilità facciale, masseteri, corneale','Trigemino'],['7','NC VII','Mimica facciale','Facciale'],
 ['8','NC VIII','Udito, vertigine','Acustico-vestibolare'],['9','NC IX–X','Velo, ugola, deglutizione, voce','Glossofaringeo/vago'],
 ['11','NC XI','Sternocleidomastoideo, trapezio','Accessorio'],['12','NC XII','Lingua','Ipoglosso']
];
var M=[
 ['delt','Deltoide','C5','ul','deltoide'],['bic','Bicipite','C5–C6','ul','bicipite'],['est','Estensori del polso','C6','ul','estensori del polso'],
 ['tric','Tricipite','C7','ul','tricipite'],['fl','Flessori delle dita','C8','ul','flessori delle dita'],['int','Interossei','T1','ul','interossei'],
 ['ilio','Ileopsoas (flessione coscia)','L2','ll','flessione della coscia'],['quad','Quadricipite (estensione gamba)','L3','ll','estensione della gamba'],
 ['ta','Tibiale anteriore (dorsiflessione piede)','L4','ll','dorsiflessione del piede'],['ehl','Estensore lungo alluce','L5','ll','dorsiflessione dell\'alluce'],
 ['gast','Gastrocnemio (flessione plantare)','S1','ll','flessione plantare']
];
var ROT=[['bic','Bicipitale','C5–C6','ul'],['sr','Stiloradiale','C6','ul'],['tric','Tricipitale','C7','ul'],['rot','Rotuleo','L3–L4','ll'],['ach','Achilleo','S1','ll']];
/* modelli: schede dell'esame nascoste e righe mostrate */
var TPL={
 compl:{hide:[],ul:true,ll:true,hof:true,m1:true,ten:true},
 cran:{hide:[],ul:true,ll:true,hof:true,m1:true,ten:false},
 cerv:{hide:['cos','nc'],ul:true,ll:true,hof:true,m1:true,ten:false},
 lomb:{hide:['cos','nc'],ul:false,ll:true,hof:false,m1:false,ten:true}
};
function T(){return TPL[S['d.tpl']]||TPL.compl}
function visM(){return M.filter(function(m){return T()[m[3]]})}
function visR(){return ROT.filter(function(r){return T()[r[3]]})}
function showTab(id){if(id==='dec')return !!S['d.tipo']&&S['d.tipo']!=='pre';return T().hide.indexOf(id)<0}
var GR=[['0','0'],['1','1+'],['2','2+'],['3','3+'],['4','4+']];
var GRTERM={'0':'assenti','1':'ipoevocabili','3':'vivaci','4':'vivaci con clono'};
var SCALE=[['asia','ASIA (A–E)'],['nihss','NIHSS'],['kps','Karnofsky KPS'],['mrs','mRS (0–6)']];
var SINT=['Parestesie','Disestesia','Allodinia','Dolore neuropatico'];
var GAIT=['Antalgica','Camptocormica','Parkinsoniana (festinante)','Magnetica','A piccoli passi','Girata en bloc','Falciante','Steppante','Atassica','A base allargata','Spastica'];
var FIXED=['d.tipo','d.tpl'];

var PANELS=[
 {id:'vis',label:'Visita'},{id:'anam',label:'Anamnesi',re:/^a\./},{id:'quad',label:'Quadro clinico',re:/^q\./},{id:'ter',label:'Terapia',re:/^t\./},
 {id:'cat',label:'Patologia',re:/^x\./},{id:'dec',label:'Decorso',re:/^o\./},
 {id:'cos',label:'Coscienza',re:/^c\./},{id:'nc',label:'Nervi cranici',re:/^nc\./},
 {id:'mot',label:'Motorio',re:/^(m|f)\./},{id:'rif',label:'Riflessi',re:/^r\./},{id:'sens',label:'Sensibilità e marcia',re:/^s\./},
 {id:'seg',label:'Segni e scale',re:/^g\./},{id:'con',label:'Conclusioni',re:/^z\./},{id:'rep',label:'Report'},{id:'glo',label:'Glossario'}
];

/* ---------- pannelli della visita ---------- */
var VIEW={
vis:function(){var a=ageAt(P.nascita,S['d.date']);
 return head('vis','Visita',(isFinite(a)?'Età alla data della visita: '+a+' anni. ':'')+'Tipo di visita, patologia e sezioni dell\'esame da mostrare.')+
 (V.copied?'<div class="hint">Anamnesi, terapia e dati della patologia sono stati copiati dalla visita del '+dateIt(V.copied)+': verificali.</div>':'')+
 '<div class="stack">'+field('Tipo di visita',seg('d.tipo',NX.VT.map(function(t){return [t[0],t[2],'']})))+
 '<div class="grid">'+field('Data della visita',inp('d.date','','Data','date'))+field('Patologia',catSel('data-k="d.cat"',S['d.cat']))+'</div>'+
 field('Diagnosi / intervento (testo libero)',inp('d.dx','es. ernia discale L4–L5 sinistra','Diagnosi'))+
 field('Lato dominante',seg('d.dom',['Destro','Sinistro']))+
 field('Esame neurologico: sezioni da mostrare',seg('d.tpl',[['lomb','Lombare'],['cerv','Cervicale'],['cran','Cranico'],['compl','Completo']]))+
 '<p class="desc">Lombare: arti inferiori, Lasègue, sella e sfinteri. Cervicale: quattro arti, Hoffmann, mielopatia. Cranico: coscienza, GCS, nervi cranici. Viene proposto in base alla patologia.</p></div>';},

anam:function(){return head('anam','Anamnesi','Patologica remota, abitudini e dati antropometrici. Viene copiata nelle visite successive.')+renderFields(NX.ANAM)},
quad:function(){return head('quad','Quadro clinico','Motivo della visita, sintomi e loro andamento.')+renderFields(NX.QUAD)},
ter:function(){return head('ter','Terapia','Terapia in corso al momento della visita e modifiche impostate.')+renderFields(NX.TER)},
cat:function(){var c=CAT[S['d.cat']];
 if(!c)return head('cat','Patologia','Scegli la patologia nella scheda Visita.');
 return head('cat',esc(c.l),'Dati per la ricerca: compila ciò che conosci. Le voci sull\'intervento e sull\'esito compaiono nelle visite dopo l\'intervento.')+renderFields(c.f)},
dec:function(){return head('dec','Decorso','Intervento, complicanze ed esito. Le voci cambiano con il tipo di visita.')+renderFields(NX.DEC)},

cos:function(){return head('cos','Coscienza e funzioni superiori','GCS, vigilanza, orientamento, linguaggio.','cos')+
 '<div class="cols">'+field('GCS Occhi',sel('c.e',[['4','4'],['3','3'],['2','2'],['1','1']],'GCS occhi'))+
 field('GCS Verbale',sel('c.v',[['5','5'],['4','4'],['3','3'],['2','2'],['1','1'],['T','T (intubato)']],'GCS verbale'))+
 field('GCS Motoria',sel('c.m',[['6','6'],['5','5'],['4','4'],['3','3'],['2','2'],['1','1']],'GCS motoria'))+'</div>'+
 '<div class="stack" style="margin-top:14px">'+
 field('Vigilanza',seg('c.vig',[['Vigile','Vigile','ok'],['Soporoso','Soporoso','alt'],['Stuporoso','Stuporoso','alt'],['Coma','Coma','alt']]))+
 field('Collaborazione',seg('c.coop',[['Collaborante','Collaborante','ok'],['Poco collaborante','Poco','alt'],['Non collaborante','Non collaborante','alt']]))+
 field('Orientamento',seg('c.ori',[['Orientato nei tre parametri','Orientato','ok'],['Disorientato nel tempo','Tempo','alt'],['Disorientato nel tempo e nello spazio','Tempo e spazio','alt'],['Disorientato nei tre parametri','Tutti e tre','alt']]))+
 field('Linguaggio',seg('c.ling',[['Fluente','Fluente','ok'],['Afasia motoria','Afasia motoria','alt'],['Afasia sensitiva','Afasia sensitiva','alt'],['Afasia globale','Afasia globale','alt'],['Afasia anomica','Anomica','alt']]))+
 '<div class="grid">'+field('Disartria',seg('c.dis',[['Assente','Assente','ok'],['Presente','Presente','alt']]))+
 field('Neglect',seg('c.neg',[['Assente','Assente','ok'],['Presente','Presente','alt']]))+'</div>'+
 field('Note (memoria, attenzione, aprassia, comportamento)',inp('c.nota','','Note funzioni superiori'))+'</div>';},

nc:function(){return head('nc','Nervi cranici','Per ogni lato: Norma o Alterato. Lascia vuoto ciò che non hai valutato.','nc')+
 NC.map(function(r){return two(r[1],r[2],'nc.'+r[0],function(k){return seg(k,N)},
   '<div class="wide">'+inp('nc.'+r[0]+'.n','Dettaglio (es. paresi facciale periferica, anisocoria 4/2 mm)','Dettaglio '+r[1])+'</div>');}).join('');},

mot:function(){return head('mot','Esame motorio','Tono, trofismo, prove di Mingazzini e forza MRC per segmento.','mot')+
 '<div class="stack">'+
 field('Tono',seg('m.tono',[['Normale','Normale','ok'],['Ipotono','Ipotono','alt'],['Spasticità lieve','Spasticità lieve','alt'],['Spasticità moderata/severa','Spasticità mod./severa','alt'],['Rigidità','Rigidità','alt']]))+
 '<div class="grid">'+field('Trofismo',seg('m.trof',[['Conservato','Conservato','ok'],['Ipotrofia','Ipotrofia','alt']]))+
 '</div><div class="grid">'+
 (T().m1?field('Mingazzini I (arti superiori)',seg('m.m1',[['Mantenuto','Mantenuto','ok'],['slivellamento a destra','Slivella dx','alt'],['slivellamento a sinistra','Slivella sx','alt']])):'')+
 field('Mingazzini II (arti inferiori)',seg('m.m2',[['Mantenuto','Mantenuto','ok'],['slivellamento a destra','Slivella dx','alt'],['slivellamento a sinistra','Slivella sx','alt']]))+'</div>'+
 field('Deficit di forza',seg('m.antalg',[['antalgica','Su base antalgica']],true))+'</div>'+
 '<h3>Forza (scala MRC 0–5)</h3><div id="syn" class="hint" hidden></div>'+
 visM().map(function(m){return two(m[1],m[2],'f.'+m[0],function(k,al){return sel(k,['0','1','2','3','4','5'],al)});}).join('');},

rif:function(){return head('rif','Riflessi e segni piramidali','Gradazione ROT: 0 assente · 1+ ipoevocabile · 2+ normale · 3+ vivace · 4+ con clono.','rif')+
 '<h3>Riflessi osteotendinei</h3>'+
 visR().map(function(r){return two(r[1],r[2],'r.'+r[0],function(k,al){return sel(k,GR,al)});}).join('')+
 '<h3>Segni piramidali</h3>'+
 two('Babinski','','r.bab',function(k){return seg(k,[['Flessorio','Flessorio','ok'],['Estensorio','Estensorio','alt'],['Indifferente','Indifferente','']])})+
 (T().hof?two('Hoffmann','','r.hof',function(k){return seg(k,[['Negativo','Neg','ok'],['Positivo','Pos','alt']])}):'')+
 two('Clono','achilleo / rotuleo','r.clo',function(k){return seg(k,[['Assente','Assente','ok'],['Presente','Presente','alt']])});},

sens:function(){return head('sens','Sensibilità, coordinazione e marcia','Sintomi sensitivi, sensibilità sellare, coordinazione, deambulazione.','sens')+
 '<div class="stack">'+
 field('Sensibilità tattile/dolorifica',seg('s.sens',[['Conservata','Conservata','ok'],['Ipoestesia','Ipoestesia','alt'],['Anestesia','Anestesia','alt'],['Iperestesia','Iperestesia','alt']]))+
 field('Sintomi sensitivi associati',seg('s.sint',SINT,true))+
 '<div class="grid">'+field('Distribuzione',inp('s.dist','es. territorio L5 sx, distale > prossimale','Distribuzione'))+field('Livello sensitivo',inp('s.liv','es. T10','Livello sensitivo'))+'</div>'+
 '<div class="grid">'+field('Sensibilità sellare',seg('s.sella',[['Conservata','Conservata','ok'],['Ipoestesia','Ipoestesia','alt'],['Anestesia','Anestesia','alt']]))+
 field('Propriocezione / vibratoria',seg('s.prop',[['Conservata','Conservata','ok'],['Ridotta','Ridotta','alt']]))+'</div>'+
 '<div class="grid">'+field('Prove di coordinazione',seg('s.coord',[['Corrette','Corrette','ok'],['Dismetria','Dismetria','alt']]))+
 field('Romberg',seg('s.romb',[['Negativo','Negativo','ok'],['Positivo','Positivo','alt']]))+'</div>'+
 '<h3>Marcia</h3>'+
 field('Deambulazione',seg('s.marcia',[['Autonoma','Autonoma','ok'],['Con ausilio','Con ausilio','alt'],['Non possibile','Non possibile','alt']]))+
 field('Punte e talloni',seg('s.punte',[['Possibile','Possibile','ok'],['Non possibile','Non possibile','alt']]))+
 field('Caratteristiche',seg('s.tipo',GAIT,true))+'</div>';},

seg:function(){return head('seg','Segni meningei, tensione radicolare, sfinteri e scale','Lasègue e Wassermann: indica anche i gradi se positivo.','seg')+
 '<div class="stack">'+field('Segni meningei (rigidità nucale, Kernig, Brudzinski)',seg('g.men',[['Assenti','Assenti','ok'],['Presenti','Presenti','alt']]))+'</div>'+
 (T().ten?'<h3>Segni di tensione radicolare</h3>'+
 [['Lasègue','las'],['Lasègue crociato','lcr'],['Wassermann','was']].map(function(r){
   return two(r[0],'','g.'+r[1],function(k){return seg(k,[['Negativo','Neg','ok'],['Positivo','Pos','alt']])+'<div class="deg">'+inp(k+'g','gradi (es. 10-20)','Gradi '+r[0])+'</div>'});}).join(''):'')+
 '<h3>Sfinteri e rachide</h3><div class="stack">'+
 field('Sfinteri',seg('g.sf',[['riferiti controllati','Controllati','ok'],['con urgenza minzionale','Urgenza','alt'],['con ritenzione urinaria','Ritenzione','alt'],['con incontinenza','Incontinenza','alt']]))+
 field('Tono sfinterico anale',seg('g.ton',[['Normale','Normale','ok'],['Ridotto','Ridotto','alt']]))+
 '<div class="grid">'+field('Rachide (dolore, contrattura, motilità)',inp('g.rach','','Rachide'))+field('Ferita / drenaggi (post-op)',inp('g.fer','','Ferita'))+'</div></div>'+
 '<h3>Scale generali</h3><p class="desc">Le scale specifiche (mJOA, ODI, Hunt-Hess, iNPHGS…) sono nella scheda della patologia.</p><div class="cols" style="margin-top:8px">'+SCALE.map(function(s){return field(s[1],inp('g.sc.'+s[0],'','Scala '+s[1]))}).join('')+'</div>';},

con:function(){return head('con','Conclusioni','Confronto con l\'esame precedente, imaging, indicazioni.')+
 '<div class="stack">'+
 '<div class="grid">'+field('Rispetto all\'esame precedente',seg('z.conf',['Invariato','Migliorato','Peggiorato']))+field('Data esame precedente',inp('z.confd','','Data esame precedente','date'))+'</div>'+
 field('Dettaglio del confronto',inp('z.confn','es. peggioramento della destrezza manuale sx','Dettaglio confronto'))+
 field('Correlazione con l\'imaging',area('z.img','','Imaging'))+
 field('Indicazioni',area('z.ind','es. rivalutazione ogni 4 ore, TC di controllo','Indicazioni'))+
 field('Firma',inp('z.firma','','Firma'))+'</div>';},

rep:function(){return head('rep','Report','Generato dai dati inseriti: solo le voci compilate. Puoi modificare il testo prima di copiarlo.')+
 '<div class="rep"><div class="acts"><button class="btn fill" type="button" id="regen">Aggiorna dai dati</button></div>'+
 '<p class="note" id="repnote" hidden>Testo modificato a mano: viene salvato così. "Aggiorna dai dati" lo rigenera e perde le modifiche.</p>'+
 '<h3>Testo per cartella</h3><textarea id="rep1" style="min-height:300px" aria-label="Testo per cartella"></textarea>'+
 '<div class="acts"><button class="btn" type="button" data-copy="rep1">Copia testo</button></div>'+
 '<h3>Referto strutturato</h3><textarea id="rep2" style="min-height:340px" aria-label="Referto strutturato"></textarea>'+
 '<div class="acts"><button class="btn" type="button" data-copy="rep2">Copia referto</button></div>'+
 '<p class="note">Rileggi sempre il testo prima di inserirlo in cartella: il report riporta solo ciò che hai inserito. Non contiene nome e cognome del paziente.</p></div>';},

glo:function(){return head('glo','Glossario','-paresi = deficit parziale (MRC 1–4/5) · -plegia = deficit completo (0/5).')+
 '<div class="field" style="margin-bottom:8px"><input type="text" id="gsearch" placeholder="Cerca un termine (es. emiparesi, disestesia, Hoffmann)" aria-label="Cerca nel glossario"></div><div id="gout" class="g-wrap"></div>';}
};

var G=[
['Deficit motorio','Monoparesi / monoplegia','Un solo arto (parziale / completo). Es. monoparesi crurale sinistra: radicolopatia, lesione parasagittale'],
['Deficit motorio','Emiparesi / emiplegia','Metà del corpo: arto superiore e inferiore dello stesso lato, spesso con la faccia. Lesione emisferica controlaterale'],
['Deficit motorio','Paraparesi / paraplegia','Entrambi gli arti inferiori. Lesione midollare toracica o lombare, cauda equina'],
['Deficit motorio','Tetraparesi / tetraplegia','Tutti e quattro gli arti. Lesione midollare cervicale'],
['Deficit motorio','Diparesi brachiale','Entrambi gli arti superiori (es. sindrome centromidollare)'],
['Deficit motorio','Emiparesi crociata','Deficit dei nervi cranici da un lato, degli arti dal lato opposto: lesione del tronco encefalico'],
['Deficit motorio','Brachiale / crurale','Prevalenza all\'arto superiore / inferiore'],
['Deficit motorio','Faciobrachiale','Faccia e arto superiore: lesione corticale o capsulare'],
['Sensibilità','Ipoestesia / anestesia','Sensibilità ridotta / assente'],
['Sensibilità','Iperestesia','Sensibilità aumentata'],
['Sensibilità','Parestesie','Formicolii spontanei, non dolorosi'],
['Sensibilità','Disestesia','Sensazione anomala e spiacevole, anche da stimolo normale'],
['Sensibilità','Allodinia','Dolore da uno stimolo normalmente non doloroso'],
['Sensibilità','Iperalgesia','Dolore eccessivo da uno stimolo doloroso'],
['Sensibilità','Anestesia a sella','Perdita di sensibilità perineale e perianale: segno di cauda equina'],
['Sensibilità','Emianestesia','Perdita della sensibilità di metà corpo'],
['Sensibilità','Livello sensitivo','Dermatomero sotto cui la sensibilità è alterata (lesione midollare)'],
['Tono e riflessi','Spasticità','Ipertono velocità-dipendente, a "lama di coltello": lesione del 1° motoneurone'],
['Tono e riflessi','Rigidità','Ipertono costante, a "tubo di piombo" o a ruota dentata: sindromi extrapiramidali'],
['Tono e riflessi','Ipotono / flaccidità','Tono ridotto: lesione del 2° motoneurone, fase acuta'],
['Tono e riflessi','Iperreflessia','ROT vivaci, con eventuale clono'],
['Tono e riflessi','Areflessia / iporeflessia','ROT assenti / ridotti'],
['Tono e riflessi','Segno di Babinski','Estensione dell\'alluce alla stimolazione plantare: lesione piramidale'],
['Tono e riflessi','Segno di Hoffmann','Flessione di pollice e indice dopo flick del 3° dito: lesione piramidale cervicale'],
['Tono e riflessi','Clono','Contrazioni ritmiche da stiramento rapido (piede, rotula)'],
['Tono e riflessi','1° motoneurone (piramidale)','Forza ridotta, tono aumentato, ROT vivaci, trofismo conservato, Babinski presente'],
['Tono e riflessi','2° motoneurone (periferico)','Forza ridotta, tono ridotto, ROT ridotti/assenti, ipotrofia e fascicolazioni, Babinski assente'],
['Segni','Lasègue','Dolore radicolare alla flessione dell\'arto inferiore esteso sul bacino (L5–S1). Si annota l\'angolo'],
['Segni','Lasègue crociato','Dolore all\'arto affetto sollevando quello controlaterale: ernia voluminosa o mediale'],
['Segni','Wassermann','Dolore alla faccia anteriore della coscia all\'estensione dell\'anca (radici L2–L4)'],
['Segni','Mingazzini','Mantenimento dell\'arto in posizione antigravitaria: I (superiori), II (inferiori da supino)'],
['Segni','Romberg','Instabilità a piedi uniti e occhi chiusi: deficit propriocettivo'],
['Segni','Segni meningei','Rigidità nucale, Kernig, Brudzinski: irritazione meningea'],
['Altri termini','Afasia (motoria / sensitiva / globale)','Disturbo del linguaggio: produzione / comprensione / entrambe'],
['Altri termini','Disartria','Difficoltà di articolazione della parola, linguaggio conservato'],
['Altri termini','Atassia','Incoordinazione (cerebellare o sensitiva)'],
['Altri termini','Dismetria','Errore di misura nel raggiungere un bersaglio'],
['Altri termini','Adiadococinesia','Incapacità di eseguire movimenti alternati rapidi'],
['Altri termini','Emianopsia','Perdita di metà del campo visivo'],
['Altri termini','Anisocoria / midriasi / miosi','Pupille diseguali / dilatate / ristrette'],
['Altri termini','Diplopia','Visione doppia: paralisi dei nervi oculomotori'],
['Altri termini','Ptosi','Caduta della palpebra'],
['Altri termini','Nistagmo','Oscillazione ritmica involontaria dei globi oculari'],
['Altri termini','Neglect (eminattenzione)','Mancata attenzione a un lato dello spazio o del corpo'],
['Altri termini','Aprassia','Incapacità di eseguire movimenti finalizzati senza deficit di forza'],
['Altri termini','Claudicatio neurogena','Dolore e debolezza alla marcia, alleviati dalla flessione del tronco (stenosi lombare)'],
['Marcia','Marcia falciante','Arto inferiore spastico che ruota in fuori: emiparesi'],
['Marcia','Marcia steppante','Piede cadente sollevato in alto: deficit L5 o nervo sciatico popliteo esterno'],
['Marcia','Marcia paraparetica spastica (a forbice)','Arti inferiori rigidi, passi corti, punte che strisciano, ginocchia che si incrociano: mielopatia, lesione parasagittale bilaterale'],
['Marcia','Marcia atassica','A base allargata, instabile: cerebellare o sensitiva'],
['Marcia','Marcia parkinsoniana','A piccoli passi, festinante (accelera come inseguendo il baricentro), tronco flesso in avanti, pendolarità delle braccia ridotta, girata "en bloc". Può associarsi a camptocormia'],
['Marcia','Marcia magnetica','Piedi che sembrano incollati al suolo, passi brevi, base allargata, difficoltà ad avviarsi e a girarsi: idrocefalo normoteso'],
['Marcia','Triade di Hakim (idrocefalo normoteso)','Disturbo della marcia, incontinenza urinaria, deterioramento cognitivo'],
['Marcia','Camptocormia','Flessione anteriore del tronco in stazione eretta e alla marcia, che si riduce da supino. Parkinson, distonie, patologia lombare'],
['Altri termini','Radicolopatia / mielopatia / plessopatia','Sofferenza di radice / midollo / plesso'],
['Altri termini','Cauda equina','Compressione delle radici sacrali: anestesia a sella, disturbi sfinterici, deficit bilaterali. Urgenza chirurgica']
];

/* ---------- normali ---------- */
function setAll(o){Object.keys(o).forEach(function(k){S[k]=o[k]})}
var NORM={
 cos:function(){setAll({'c.e':'4','c.v':'5','c.m':'6','c.vig':'Vigile','c.coop':'Collaborante','c.ori':'Orientato nei tre parametri','c.ling':'Fluente','c.dis':'Assente','c.neg':'Assente'})},
 nc:function(){NC.forEach(function(r){S['nc.'+r[0]+'.dx']='Normale';S['nc.'+r[0]+'.sx']='Normale'})},
 mot:function(){setAll({'m.tono':'Normale','m.trof':'Conservato','m.m2':'Mantenuto'});if(T().m1)S['m.m1']='Mantenuto';
   visM().forEach(function(m){S['f.'+m[0]+'.dx']='5';S['f.'+m[0]+'.sx']='5'})},
 rif:function(){visR().forEach(function(r){S['r.'+r[0]+'.dx']='2';S['r.'+r[0]+'.sx']='2'});
   ['dx','sx'].forEach(function(s){S['r.bab.'+s]='Flessorio';if(T().hof)S['r.hof.'+s]='Negativo';S['r.clo.'+s]='Assente'})},
 sens:function(){setAll({'s.sens':'Conservata','s.sella':'Conservata','s.prop':'Conservata','s.coord':'Corrette','s.romb':'Negativo','s.marcia':'Autonoma','s.punte':'Possibile'})},
 seg:function(){setAll({'g.men':'Assenti','g.sf':'riferiti controllati','g.ton':'Normale'});
   if(T().ten)['las','lcr','was'].forEach(function(r){S['g.'+r+'.dx']='Negativo';S['g.'+r+'.sx']='Negativo'})}
};

/* ---------- report ---------- */
function numIt(x){return typeof x==='number'?String(x).replace('.',','):x}
function fval(f,s){var v=s[f[0]];
  if(f[2]==='c')return numIt(calcVal(f,s,P));
  if(v===undefined||v===''||(Array.isArray(v)&&!v.length))return '';
  if(Array.isArray(v))return v.map(lcw).join(', ');
  if(f[2]==='d')return dateIt(v);
  if(f[2]==='t'||f[2]==='a')return trimDot(v);
  return String(v)}
function fieldsText(list){var out=[];
  list.forEach(function(f){if(f[2]==='h'||!fvis(f))return;var v=fval(f,S);if(v==='')return;
    var m=f[1].match(/^(.*) \(([^)]*)\)$/);
    if(m&&(f[2]==='n'||(f[2]==='c'&&/^[\/%]/.test(m[2]))))out.push(m[1]+' '+String(v).replace('.',',')+(/^[\/%]/.test(m[2])?'':' ')+m[2]);
    else out.push(f[1]+': '+(f[2]==='s'||f[2]==='l'?lcw(v):v))});
  return out}
function sideWord(d,s){return d&&s?'bilateralmente':d?'a destra':s?'a sinistra':''}
function build(){
  var o={dati:[],an:[],ter:[],sx:[],cat:[],dec:[],cos:[],nc:[],mot:[],rif:[],ten:[],sens:[],seg:[],mar:[],scale:[],con:[]};
  var v=function(k){return S[k]},c=CAT[v('d.cat')],age=ageAt(P&&P.nascita,v('d.date'));
  // dati
  var tv=vt(v('d.tipo'),3);
  o.dati.push('Visita '+(tv?tv+' ':'')+'del '+dateIt(v('d.date')||td()));
  var who=[isFinite(age)?age+' anni':'',P&&P.sesso==='M'?'sesso maschile':P&&P.sesso==='F'?'sesso femminile':''].filter(Boolean);
  if(who.length)o.dati.push('Paziente di '+who.join(', '));
  if(c)o.dati.push('Patologia: '+c.l);
  if(v('d.dx'))o.dati.push('Diagnosi/intervento: '+v('d.dx'));
  if(v('d.dom'))o.dati.push('Lato dominante '+lc(v('d.dom')));
  // anamnesi, terapia, quadro clinico, patologia, decorso
  o.an=fieldsText(NX.ANAM);o.ter=fieldsText(NX.TER);
  if(v('q.sx'))o.sx.push(trimDot(v('q.sx')));
  o.sx=o.sx.concat(fieldsText(NX.QUAD.filter(function(f){return f[0]!=='q.sx'})));
  if(c)o.cat=fieldsText(c.f);
  if(showTab('dec'))o.dec=fieldsText(NX.DEC);
  // coscienza
  if(showTab('cos')){
  var p=[];['c.vig','c.coop','c.ori'].forEach(function(k){if(v(k))p.push(lc(v(k)))});
  if(p.length)o.cos.push(p.join(', '));
  if(v('c.ling'))o.cos.push(v('c.ling')==='Fluente'?'eloquio fluente':lc(v('c.ling')));
  if(v('c.dis'))o.cos.push(v('c.dis')==='Presente'?'disartria':'non disartria');
  if(v('c.neg')==='Presente')o.cos.push('neglect');
  if(v('c.e')&&v('c.v')&&v('c.m'))o.cos.push('GCS '+gcsTot(S)+' (E'+v('c.e')+' V'+v('c.v')+' M'+v('c.m')+')');
  if(v('c.nota'))o.cos.push(v('c.nota'));
  }
  // nervi cranici
  var alt=[],unex=[],exam=0;
  if(showTab('nc'))
  NC.forEach(function(r){var d=v('nc.'+r[0]+'.dx'),s=v('nc.'+r[0]+'.sx');
    if(!d&&!s){unex.push(r[3]);return}exam++;
    var da=d==='Alterato',sa=s==='Alterato';
    if(da||sa||v('nc.'+r[0]+'.n'))alt.push(r[1]+' ('+r[3].toLowerCase()+')'+(sideWord(da,sa)?' '+sideWord(da,sa):'')+(v('nc.'+r[0]+'.n')?': '+v('nc.'+r[0]+'.n'):''))});
  if(exam){if(alt.length)o.nc.push('alterazione di '+alt.join('; '));
    else o.nc.push(unex.length?'nervi cranici esaminati indenni (non valutati: '+unex.join(', ').toLowerCase()+')':'nervi cranici indenni')}
  // motorio
  if(v('m.tono'))o.mot.push('tono '+lc(v('m.tono')));
  if(v('m.trof'))o.mot.push('trofismo '+lc(v('m.trof')));
  var def=[],ex=0;
  var grp={},gk=[];
  visM().forEach(function(m){['dx','sx'].forEach(function(s){var x=v('f.'+m[0]+'.'+s);if(x===undefined)return;ex++;
    if(+x<5){var key=s+x;if(!grp[key]){grp[key]=[];gk.push(key)}grp[key].push(m[4])}})});
  gk.forEach(function(key){var it=grp[key],l=it.length>1?it.slice(0,-1).join(', ')+' e '+it[it.length-1]:it[0];
    def.push(l+(key[0]==='d'?' a destra':' a sinistra')+' ('+key.slice(2)+'/5)')});
  if(def.length)o.mot.push('alle prove segmentarie contro-resistenza: deficit di '+def.join('; ')+((v('m.antalg')||[]).length?' su base antalgica':''));
  else if(ex)o.mot.push(ex===visM().length*2?(T().ul?'forza 5/5 ai quattro arti':'forza 5/5 agli arti inferiori'):'forza 5/5 nei segmenti esaminati');
  var mg=[];(T().m1?[['m.m1','I'],['m.m2','II']]:[['m.m2','II']]).forEach(function(r){if(v(r[0]))mg.push('Mingazzini '+r[1]+(v(r[0])==='Mantenuto'?' mantenuto':' con '+v(r[0])))});
  if(mg.length)o.mot.push(mg.join(', '));
  // riflessi
  var cells=[];visR().forEach(function(r){['dx','sx'].forEach(function(s){var x=v('r.'+r[0]+'.'+s);if(x!==undefined)cells.push({r:r[1],s:s,x:x})})});
  if(cells.length){var vals=cells.map(function(c){return c.x}),sym=vals.every(function(x){return x===vals[0]})&&cells.length>=2;
    if(sym)o.rif.push(vals[0]==='2'?'ROT normoevocabili e simmetrici':'ROT '+GRTERM[vals[0]]+' simmetricamente');
    else o.rif.push('ROT: '+cells.map(function(c){return c.r.toLowerCase()+' '+c.s+' '+GR[+c.x][1]}).join(', '))}
  var pos=[],PY=[['bab','Babinski','Estensorio'],['hof','Hoffmann','Positivo'],['clo','clono','Presente']].filter(function(r){return r[0]!=='hof'||T().hof});
  PY.forEach(function(r){
    var d=v('r.'+r[0]+'.dx')===r[2],s=v('r.'+r[0]+'.sx')===r[2];if(d||s)pos.push(r[1]+' '+sideWord(d,s))});
  var anyPy=PY.map(function(r){return r[0]}).some(function(r){return v('r.'+r+'.dx')||v('r.'+r+'.sx')});
  if(pos.length)o.rif.push('segni piramidali: '+pos.join(', '));else if(anyPy)o.rif.push('non segni piramidali');
  // tensione
  (T().ten?[['Lasègue','las'],['Lasègue crociato','lcr'],['Wassermann','was']]:[]).forEach(function(r){
    var d=v('g.'+r[1]+'.dx'),s=v('g.'+r[1]+'.sx');if(!d&&!s)return;
    var dp=d==='Positivo',sp=s==='Positivo';
    if(!dp&&!sp){o.ten.push(r[0]+' negativo'+(d&&s?' bilateralmente':d?' a destra':' a sinistra'));return}
    var parts=[];if(dp)parts.push('a destra'+(v('g.'+r[1]+'.dxg')?' a '+v('g.'+r[1]+'.dxg')+'°':''));
    if(sp)parts.push('a sinistra'+(v('g.'+r[1]+'.sxg')?' a '+v('g.'+r[1]+'.sxg')+'°':''));
    o.ten.push(r[0]+' positivo '+parts.join(' e '))});
  // sensibilità
  var it=(v('s.sint')||[]).map(lc);if(v('s.sens')&&v('s.sens')!=='Conservata')it.push(lc(v('s.sens')));
  if(it.length)o.sens.push(it.join(', ')+(v('s.dist')?' ('+v('s.dist')+')':''));
  else if(v('s.sens'))o.sens.push('sensibilità tattile e dolorifica conservata');
  if(v('s.liv'))o.sens.push('livello sensitivo '+v('s.liv'));
  if(v('s.sella'))o.sens.push(v('s.sella')==='Conservata'?'sensibilità sellare conservata':lc(v('s.sella'))+' sellare');
  if(v('s.prop'))o.sens.push('propriocezione '+lc(v('s.prop')));
  if(v('s.coord'))o.sens.push(v('s.coord')==='Corrette'?'prove di coordinazione corrette':'dismetria');
  if(v('s.romb'))o.sens.push('Romberg '+lc(v('s.romb')));
  // meningei, sfinteri, rachide
  if(v('g.men'))o.seg.push(v('g.men')==='Assenti'?'non segni meningei':'segni meningei presenti');
  if(v('g.rach'))o.seg.push('rachide: '+v('g.rach'));
  if(v('g.sf'))o.seg.push('sfinteri '+v('g.sf'));
  if(v('g.ton'))o.seg.push('tono sfinterico anale '+lc(v('g.ton')));
  if(v('g.fer'))o.seg.push('ferita/drenaggi: '+v('g.fer'));
  // marcia
  if(v('s.marcia')){var mm=v('s.marcia')==='Autonoma'?'marcia autonoma':v('s.marcia')==='Con ausilio'?'marcia con ausilio':'marcia non possibile';
    if(v('s.punte'))mm+=v('s.punte')==='Possibile'?', possibile su punte e talloni':', non possibile su punte e talloni';
    var tp=(v('s.tipo')||[]).map(lc);if(tp.length)mm+='; andatura '+tp.join(', ');o.mar.push(mm)}
  else if((v('s.tipo')||[]).length)o.mar.push('andatura '+(v('s.tipo')).map(lc).join(', '));
  // scale
  var sc=SCALE.filter(function(s){return v('g.sc.'+s[0])}).map(function(s){return s[1].replace(/ \(.*\)/,'')+' '+v('g.sc.'+s[0])});
  if(sc.length)o.scale.push('scale: '+sc.join('; '));
  // conclusioni
  if(v('z.conf'))o.con.push('rispetto all\'esame precedente'+(v('z.confd')?' ('+dateIt(v('z.confd'))+')':'')+': '+lc(v('z.conf'))+(v('z.confn')?' — '+v('z.confn'):''));
  if(v('z.img'))o.con.push('correlazione con l\'imaging: '+trimDot(v('z.img')));
  if(v('z.ind'))o.con.push('indicazioni: '+trimDot(v('z.ind')));
  return o;
}
function gcsTot(s){if(!s['c.e']||!s['c.v']||!s['c.m'])return '';var t=(+s['c.e'])+(s['c.v']==='T'?1:+s['c.v'])+(+s['c.m']);return s['c.v']==='T'?t+'T':t}
var HEAD=[['dati','DATI GENERALI'],['an','ANAMNESI'],['ter','TERAPIA'],['sx','QUADRO CLINICO'],['cos','COSCIENZA E FUNZIONI SUPERIORI'],['nc','NERVI CRANICI'],
 ['mot','ESAME MOTORIO'],['rif','RIFLESSI E SEGNI PIRAMIDALI'],['ten','SEGNI DI TENSIONE RADICOLARE'],['sens','SENSIBILITÀ E COORDINAZIONE'],
 ['seg','SEGNI MENINGEI, SFINTERI, RACHIDE'],['mar','MARCIA'],['scale','SCALE'],['cat',''],['dec','DECORSO'],['con','CONCLUSIONI']];
function reports(){
  var o=build(),c=CAT[S['d.cat']],para=[];
  function sec(t,arr){if(arr.length)para.push(t+' – '+arr.map(function(s,i){return i?lcw(s):cap(s)}).join('; ')+'.')}
  sec('Anamnesi',o.an);sec('Terapia',o.ter);sec('Quadro clinico',o.sx);
  var ex=[];['cos','nc','mot','rif','ten','sens','seg','mar','scale'].forEach(function(g){o[g].forEach(function(s){ex.push(cap(s))})});
  if(ex.length)para.push('Esame neurologico – '+ex.join('. ')+'.');
  if(c)sec(c.l,o.cat);sec('Decorso',o.dec);sec('Conclusioni',o.con);
  var lines=[];
  HEAD.forEach(function(h){var arr=o[h[0]].slice(),t=h[1]||(c?c.l.toUpperCase():'');
    if(h[0]==='con'&&S['z.firma'])arr.push('firma: '+S['z.firma']);
    if(arr.length){lines.push(t);arr.forEach(function(s){lines.push('- '+cap(s))});lines.push('')}});
  return [para.length?para.join('\n\n'):EMPTY,lines.length>4?lines.join('\n').trim():EMPTY];
}
function withVisit(v,fn){var s0=S,p0=P;S=v.data;P=patOf(v.pid)||{};try{return fn()}finally{S=s0;P=p0}}

/* ---------- suggerimento termine ---------- */
function syndrome(){
  var aff=[];
  [['ul','dx'],['ul','sx'],['ll','dx'],['ll','sx']].forEach(function(l){
    var vals=visM().filter(function(m){return m[3]===l[0]}).map(function(m){return S['f.'+m[0]+'.'+l[1]]}).filter(function(x){return x!==undefined}).map(Number);
    if(vals.some(function(x){return x<5}))aff.push({r:l[0],s:l[1],pl:vals.every(function(x){return x===0})});
  });
  if(!aff.length)return '';
  var pl=aff.every(function(a){return a.pl}),b=pl?'plegia':'paresi',n=aff.length,sd=function(s){return s==='dx'?'destra':'sinistra'};
  var name,a=aff[0];
  if(n===1)name='Mono'+b+' '+(a.r==='ul'?'brachiale':'crurale')+' '+sd(a.s);
  else if(n===2&&aff[0].s===aff[1].s&&aff[0].r!==aff[1].r)name='Emi'+b+' '+sd(a.s);
  else if(n===2&&aff[0].r===aff[1].r)name=a.r==='ll'?'Para'+b:(pl?'Diplegia brachiale':'Diparesi brachiale');
  else if(n===2)name='Deficit crociato ('+aff.map(function(x){return (x.r==='ul'?'arto sup. ':'arto inf. ')+sd(x.s)}).join(' e ')+')';
  else if(n===3)name='Tri'+b+' ('+aff.map(function(x){return (x.r==='ul'?'sup. ':'inf. ')+sd(x.s)}).join(', ')+')';
  else name='Tetra'+b;
  return name;
}

/* ---------- visita: montaggio ---------- */
function paint(el){
  var k=el.dataset.k||'',x=el.value,alt=false;
  if(/^f\./.test(k)&&x!=='')alt=+x<5;
  if(/^r\.(bic|sr|tric|rot|ach)\./.test(k)&&x!=='')alt=x!=='2';
  el.classList.toggle('alt',alt);
}
function mountPanel(id){
  var el=$('p-'+id);el.innerHTML='<div class="card">'+VIEW[id]()+navHtml(id)+'</div>';
  el.querySelectorAll('select[data-k]').forEach(paint);
  if(id==='glo')drawGlossary('');
  if(id==='mot')updateSyn();
}
function navHtml(id){
  var Vs=PANELS.filter(function(p){return showTab(p.id)}),i=Vs.findIndex(function(p){return p.id===id}),prev=Vs[i-1],next=Vs[i+1];
  return '<div class="foot">'+(prev?'<button class="btn quiet" type="button" data-go="'+prev.id+'">← '+prev.label+'</button>':'<span></span>')+
    (next?'<button class="btn fill" type="button" data-go="'+next.id+'">'+next.label+' →</button>':'<span></span>')+'</div>';
}
function applyTpl(){
  var c=CAT[S['d.cat']];PANELS[4].label=c?c.s:'Patologia';
  $('t-cat').firstChild.nodeValue=PANELS[4].label;
  PANELS.forEach(function(p){$('t-'+p.id).hidden=!showTab(p.id);mountPanel(p.id)});
  var cur=PANELS.filter(function(p){return !$('p-'+p.id).hidden})[0];
  if(cur&&!showTab(cur.id))show('vis');
  updateDots();renderHeader('visit');
}
function show(id){
  PANELS.forEach(function(p){$('t-'+p.id).setAttribute('aria-selected',p.id===id);$('p-'+p.id).hidden=p.id!==id});
  if(id==='rep')fillReport();
  if(V)LASTTAB[V.id]=id;
  var t=$('t-'+id);if(t&&t.scrollIntoView)t.scrollIntoView({block:'nearest',inline:'center'});
  window.scrollTo(0,0);
}
function fillReport(){
  if(!V.repEdit){var r=reports();$('rep1').value=r[0];$('rep2').value=r[1]}
  else{$('rep1').value=V.rep1||'';$('rep2').value=V.rep2||''}
  $('repnote').hidden=!V.repEdit;
}
function updateDots(){
  PANELS.forEach(function(p){if(!p.re)return;var on=Object.keys(S).some(function(k){return p.re.test(k)&&S[k]!==''&&!(Array.isArray(S[k])&&!S[k].length)});
    var t=$('t-'+p.id),d=t.querySelector('.dot');if(on&&!d)t.insertAdjacentHTML('beforeend','<span class="dot" aria-label="compilato"></span>');if(!on&&d)d.remove()});
}
function updateSyn(){
  var el=$('syn');if(!el)return;var s=syndrome();
  el.hidden=!s;if(s)el.innerHTML='Termine suggerito dai valori inseriti: <b>'+esc(s)+'</b>. Verifica sempre con il quadro clinico (es. deficit su base antalgica).';
}
function updateCalc(){document.querySelectorAll('[data-calc]').forEach(function(el){var r=calcVal(FIELD[el.dataset.calc],S,P);el.textContent=r===''?'–':r})}
function changed(){saveV();updateDots();updateSyn();updateCalc()}
function drawGlossary(q){
  q=(q||'').toLowerCase();var rows=G.filter(function(r){return !q||(r[0]+' '+r[1]+' '+r[2]).toLowerCase().indexOf(q)>-1});
  if(!rows.length){$('gout').innerHTML='<p class="desc">Nessun termine trovato.</p>';return}
  var cat='',h='';rows.forEach(function(r){if(r[0]!==cat){if(cat)h+='</table>';cat=r[0];h+='<h3>'+esc(cat)+'</h3><table class="g-table">'}
    h+='<tr><td>'+esc(r[1])+'</td><td>'+esc(r[2])+'</td></tr>'});
  $('gout').innerHTML=h+'</table>';
}

/* ---------- schermate: pazienti, paziente, visita ---------- */
function renderHeader(view){
  var h='Neurosurg Clinic',sub='Visite neurochirurgiche. I dati restano solo su questo dispositivo.',btn='';
  if(view!=='home'&&P){var a=ageAt(P.nascita,view==='visit'?S['d.date']:td());h=pname(P);
    if(view==='pat'){sub=[P.code,isFinite(a)?a+' anni':'',P.sesso].filter(Boolean).join(' · ');
      btn='<button class="btn quiet" type="button" data-act="home">← Pazienti</button>'}
    else{var c=CAT[S['d.cat']];sub=[V.code,vt(S['d.tipo'],2),c?c.s:'',isFinite(a)?a+' anni':''].filter(Boolean).join(' · ');
      btn='<button class="btn quiet" type="button" data-act="openpat">← Paziente</button><button class="btn fill" type="button" data-go="rep">Fine visita → Report</button>'}}
  $('htitle').textContent=h;$('hsub').textContent=sub;$('hbtns').innerHTML=btn;
}
function go(view){flush();['home','pat','visit'].forEach(function(x){$('v-'+x).hidden=x!==view});$('tabs').hidden=view!=='visit';renderHeader(view);drawOpen(view);window.scrollTo(0,0)}
/* visite aperte: barra fissa per passare da un paziente all'altro durante il giro */
var OPEN=[],LASTTAB={},LASTY={};
try{OPEN=JSON.parse(localStorage.getItem('nx-open')||'[]')}catch(e){}
function saveOpen(){try{localStorage.setItem('nx-open',JSON.stringify(OPEN))}catch(e){}}
function visById(id){return DB.vis.filter(function(v){return v.id===id})[0]}
function drawOpen(view){
  var el=$('openbar');OPEN=OPEN.filter(visById);
  if(!OPEN.length&&view==='home'){el.hidden=true;el.innerHTML='';return}
  el.hidden=false;
  el.innerHTML=(view!=='home'?'<button type="button" class="chip home" data-act="home">☰ Pazienti</button>':'<span class="ob-l">Visite aperte</span>')+
    OPEN.map(function(id){var v=visById(id),p=patOf(v.pid)||{},on=V&&V.id===id;
      return '<span class="chip'+(on?' on':'')+'"><button type="button" data-act="openvis" data-id="'+id+'"'+(on?' aria-current="true"':'')+'>'+
        esc((p.cognome||'(senza nome)')+(p.nome?' '+p.nome.charAt(0)+'.':''))+' <small>'+esc(vt(v.data['d.tipo'],2))+'</small></button>'+
        '<button type="button" class="x" data-act="closevis" data-id="'+id+'" aria-label="Chiudi la visita dalla barra">×</button></span>'}).join('');
  var cur=el.querySelector('.chip.on');if(cur&&cur.scrollIntoView)cur.scrollIntoView({block:'nearest',inline:'nearest'});
}
function closeVis(id){OPEN=OPEN.filter(function(x){return x!==id});saveOpen();
  if(V&&V.id===id){var p=P.id;openPat(p)}else drawOpen($('v-home').hidden?($('v-pat').hidden?'visit':'pat'):'home')}
function dirty(){var b=meta().backup||0;return DB.pat.some(function(p){return (p.updated||0)>b})||DB.vis.some(function(v){return (v.updated||0)>b})}
function homeHtml(){var m=meta(),today=DB.vis.filter(function(v){return v.data['d.date']===td()}).sort(function(a,b){return b.updated-a.updated});
  var nx=DB.vis.filter(function(v){return (v.updated||0)>(m.xlsx||0)}).length;
  return (today.length?'<div class="card"><h2>Visite di oggi</h2><div class="plist">'+today.map(function(v){var p=patOf(v.pid)||{},c=CAT[v.data['d.cat']];
     return '<button type="button" class="prow" data-act="openvis" data-id="'+v.id+'"><span class="pn">'+esc(pname(p)+' · '+vt(v.data['d.tipo'],2))+'</span>'+
       '<span class="pm">'+esc([p.code,c?c.s:'',v.data['d.dx']].filter(Boolean).join(' · '))+'</span></button>'}).join('')+'</div></div>':'')+
   '<div class="card'+(today.length?' mt':'')+'"><div class="ph"><div><h2>Pazienti</h2><p class="desc">'+DB.pat.length+' pazienti · '+DB.vis.length+' visite su questo dispositivo.</p></div>'+
   '<button class="btn fill" type="button" data-act="newpat">+ Nuovo paziente</button></div>'+
   '<input type="search" id="psearch" placeholder="Cerca per cognome, nome, codice o n. cartella" aria-label="Cerca paziente" value="'+esc(HQ)+'">'+
   '<div id="plist" class="plist"></div></div>'+
   '<div class="card mt"><h2>Salvataggio ed Excel</h2>'+
   '<p class="desc" style="margin-bottom:10px">Ogni dato che inserisci viene <b>salvato automaticamente su questo iPad</b>: puoi chiudere l\'app o cambiare paziente senza perdere nulla. '+
   'L\'Excel invece <b>non si aggiorna da solo</b>: è una copia che crei con "Esporta Excel" e contiene tutti i dati fino a quel momento.</p>'+
   '<div class="hint'+(nx?' warn':'')+'">'+(m.xlsx?'Ultimo Excel esportato il '+new Date(m.xlsx).toLocaleString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'})+'. ':'Nessun Excel esportato finora. ')+
   (nx?nx+(nx===1?' visita nuova o modificata non è ancora':' visite nuove o modificate non sono ancora')+' nell\'Excel.':(DB.vis.length?'L\'Excel contiene tutte le visite.':''))+'</div>'+
   (dirty()?'<div class="hint warn">Ci sono dati non ancora salvati in un backup'+(m.backup?' (ultimo: '+new Date(m.backup).toLocaleDateString('it-IT')+')':'')+'. Se l\'app viene eliminata o l\'iPad si guasta, i dati si perdono: fai un backup e salvalo in File o OneDrive.</div>':'')+
   '<div class="acts"><button class="btn" type="button" data-act="xlsx">Esporta Excel</button><button class="btn" type="button" data-act="xlsx-anon">Excel senza anagrafica</button>'+
   '<button class="btn" type="button" data-act="backup">Backup (.json)</button><button class="btn quiet" type="button" data-act="restore">Ripristina backup</button></div>'+
   '<p class="desc" id="hmsg" role="status"></p>'+
   '<p class="note">Excel: un foglio Anagrafica, un foglio Visite con i dati comuni e l\'esame, e un foglio per ogni patologia. Nei fogli di ricerca il paziente compare solo con il codice; le scelte multiple diventano colonne 1/0. '+
   'Il backup .json serve a ripristinare l\'archivio, anche su un altro iPad; l\'Excel no.</p></div>';
}
function drawPList(){
  var q=HQ.toLowerCase(),el=$('plist');if(!el)return;
  var rows=DB.pat.map(function(p){var vs=visitsOf(p.id);return {p:p,vs:vs}})
    .filter(function(r){return !q||[r.p.cognome,r.p.nome,r.p.code,r.p.cartella].join(' ').toLowerCase().indexOf(q)>-1})
    .sort(function(a,b){var x=a.vs[0]?Math.max(a.vs[0].updated,a.p.updated):a.p.updated,y=b.vs[0]?Math.max(b.vs[0].updated,b.p.updated):b.p.updated;return y-x});
  if(!rows.length){el.innerHTML='<p class="desc" style="margin-top:12px">'+(DB.pat.length?'Nessun paziente trovato.':'Nessun paziente. Inizia con "Nuovo paziente".')+'</p>';return}
  el.innerHTML=rows.map(function(r){var p=r.p,l=r.vs[0],a=ageAt(p.nascita),c=l&&CAT[l.data['d.cat']];
    var m=[p.code,isFinite(a)?a+' anni':'',p.sesso,r.vs.length+(r.vs.length===1?' visita':' visite')].filter(Boolean).join(' · ');
    if(l)m+=' · ultima '+dateIt(l.data['d.date'])+(c?' '+c.s:'');
    return '<button type="button" class="prow" data-act="openpat" data-id="'+p.id+'"><span class="pn">'+esc(pname(p))+'</span><span class="pm">'+esc(m)+'</span></button>'}).join('');
}
function pin(k,label,type){return field(label,'<input type="'+(type||'text')+'" data-p="'+k+'" aria-label="'+esc(label)+'" value="'+esc(P[k]||'')+'">')}
function pseg(attr,cur,opts){return '<div class="seg" role="group" '+attr+'>'+opts.map(function(o){var a=typeof o==='string'?[o,o]:o;
  return '<button type="button" data-v="'+esc(a[0])+'" aria-pressed="'+(cur===a[0])+'">'+esc(a[1])+'</button>'}).join('')+'</div>'}
function patHtml(){var vs=visitsOf(P.id);
  return '<div class="card"><div class="ph"><div><h2>Anagrafica</h2><p class="desc">Codice '+esc(P.code)+': è l\'unico identificativo nei fogli di ricerca.</p></div></div>'+
   '<div class="grid">'+pin('cognome','Cognome')+pin('nome','Nome')+pin('nascita','Data di nascita','date')+field('Sesso',pseg('data-p="sesso"',P.sesso,['M','F']))+pin('cartella','N. cartella / nosologico')+'</div>'+
   '<div class="stack" style="margin-top:14px">'+pin('note','Note (es. medico inviante, recapito)')+'</div></div>'+
   '<div class="card mt"><h2>Nuova visita</h2><div class="stack" style="margin-top:10px">'+
   field('Tipo di visita',pseg('data-nv="tipo"',NEWV.tipo,NX.VT.map(function(t){return [t[0],t[1]]})))+
   field('Patologia',catSel('data-nv="cat"',NEWV.cat))+
   (vs.length?'<p class="desc">Anamnesi, terapia e dati della patologia vengono copiati dall\'ultima visita. L\'esame neurologico parte sempre vuoto.</p>':'')+
   '<div><button class="btn fill" type="button" data-act="newvis">Inizia visita →</button> <span class="desc" id="nvmsg" role="status"></span></div></div></div>'+
   '<div class="card mt"><h2>Visite</h2>'+(vs.length?'<div class="plist">'+vs.map(function(v){var c=CAT[v.data['d.cat']];
     return '<div class="vrow"><button type="button" class="prow" data-act="openvis" data-id="'+v.id+'"><span class="pn">'+esc(dateIt(v.data['d.date'])+' · '+vt(v.data['d.tipo'],1))+'</span>'+
       '<span class="pm">'+esc([v.code,c?c.l:'',v.data['d.dx']].filter(Boolean).join(' · '))+'</span></button>'+
       '<button type="button" class="btn quiet sm" data-act="delvis" data-arm data-id="'+v.id+'">Elimina</button></div>'}).join('')+'</div>':'<p class="desc">Nessuna visita.</p>')+'</div>'+
   '<div class="foot"><span></span><button type="button" class="btn danger" data-act="delpat" data-arm>Elimina paziente e visite</button></div>';
}
function cleanEmpty(){if(P&&!P.cognome&&!P.nome&&!P.nascita&&!visitsOf(P.id).length){DB.pat=DB.pat.filter(function(p){return p!==P});del('pat',P.id)}}
function openHome(){if(V)LASTY[V.id]=window.scrollY;flush();cleanEmpty();P=null;V=null;S={};$('v-home').innerHTML=homeHtml();drawPList();go('home')}
function openPat(id){if(V)LASTY[V.id]=window.scrollY;flush();V=null;S={};P=patOf(id);var l=visitsOf(id)[0];NEWV={tipo:l?'':'pre',cat:l?l.data['d.cat']:''};
  $('v-pat').innerHTML=patHtml();go('pat')}
function nextCode(){var m=0;DB.pat.forEach(function(p){var n=parseInt(String(p.code).replace(/\D/g,''),10);if(n>m)m=n});return 'NCH-'+String(m+1).padStart(4,'0')}
function newPat(){var t=Date.now();P={id:uid(),code:nextCode(),cognome:'',nome:'',nascita:'',sesso:'',cartella:'',note:'',created:t,updated:t};
  DB.pat.push(P);put('pat',P);openPat(P.id);var f=document.querySelector('[data-p="cognome"]');if(f)f.focus()}
function newVisit(){
  if(!NEWV.tipo||!NEWV.cat){$('nvmsg').textContent='Scegli tipo di visita e patologia.';return}
  var vs=visitsOf(P.id),prev=vs[0],c=CAT[NEWV.cat],d={'d.date':td(),'d.tipo':NEWV.tipo,'d.cat':NEWV.cat,'d.tpl':c.tpl},copied='';
  if(prev){var ps=prev.data;Object.keys(ps).forEach(function(k){var f=FIELD[k],same=ps['d.cat']===NEWV.cat;
    if((same&&(k==='d.dx'||k==='d.dom'))||(f&&fl(f).indexOf('k')>-1&&(k.indexOf('x.')!==0||same))){d[k]=Array.isArray(ps[k])?ps[k].slice():ps[k];copied=ps['d.date']||'?'}});
    if(ps['d.date'])d['z.confd']=ps['d.date']}
  var m=0;vs.forEach(function(v){var n=parseInt(String(v.code).split('-V')[1],10);if(n>m)m=n});
  var t=Date.now();V={id:uid(),pid:P.id,code:P.code+'-V'+(m+1),created:t,updated:t,data:d};if(copied)V.copied=copied;
  DB.vis.push(V);put('vis',V);openVisit(V.id,copied?'quad':'anam')}
function openVisit(id,tab){if(V)LASTY[V.id]=window.scrollY;flush();V=visById(id);if(!V)return;P=patOf(V.pid);S=V.data;
  if(OPEN.indexOf(id)<0){OPEN.push(id);if(OPEN.length>10)OPEN.shift();saveOpen()}
  $('tabs').innerHTML=PANELS.map(function(p){return '<button type="button" class="tab" role="tab" id="t-'+p.id+'" data-go="'+p.id+'" aria-selected="false">'+p.label+'</button>'}).join('');
  $('panels').innerHTML=PANELS.map(function(p){return '<section id="p-'+p.id+'" role="tabpanel" aria-labelledby="t-'+p.id+'" hidden></section>'}).join('');
  go('visit');applyTpl();show(tab||LASTTAB[id]||'vis');if(!tab&&LASTY[id])window.scrollTo(0,LASTY[id])}
function delVisit(id){DB.vis=DB.vis.filter(function(v){return v.id!==id});del('vis',id);openPat(P.id)}
function delPat(){visitsOf(P.id).forEach(function(v){del('vis',v.id)});DB.vis=DB.vis.filter(function(v){return v.pid!==P.id});
  DB.pat=DB.pat.filter(function(p){return p!==P});del('pat',P.id);P=null;openHome()}

/* ---------- esportazione Excel ---------- */
function cell(t,v){
  if(v===undefined||v===null||v==='')return '';
  if(t==='n'){var s=String(v).trim();return /^-?\d+([.,]\d+)?$/.test(s)?parseFloat(s.replace(',','.')):s}
  if(t==='d')return /^\d{4}-\d{2}-\d{2}$/.test(v)?{d:v}:String(v);
  return Array.isArray(v)?v.join('; '):String(v)}
function specCols(list){var C=[];
  list.forEach(function(f){var k=f[0],t=f[2];if(t==='h')return;
    if(t==='m')f[3].forEach(function(o){C.push({h:f[1]+': '+o,t:'n',fn:function(s){var a=s[k];return a&&a.length?(a.indexOf(o)>-1?1:0):''}})});
    else if(t==='c')C.push({h:f[1],t:'n',fn:function(s){return calcVal(f,s,P)}});
    else C.push({h:f[1],t:t==='n'?'n':t==='d'?'d':'s',k:k})});
  return C}
function multiCols(k,label,opts){return opts.map(function(o){return {h:label+': '+o,t:'n',fn:function(s){var a=s[k];return a&&a.length?(a.indexOf(o)>-1?1:0):''}}})}
function examCols(){var C=[];function f(k,h,t){C.push({k:k,h:h,t:t||'s'})}
  f('d.tpl','Modello esame');
  f('c.e','GCS E','n');f('c.v','GCS V','n');f('c.m','GCS M','n');C.push({h:'GCS totale',t:'n',fn:gcsTot});
  [['c.vig','Vigilanza'],['c.coop','Collaborazione'],['c.ori','Orientamento'],['c.ling','Linguaggio'],['c.dis','Disartria'],['c.neg','Neglect'],['c.nota','Note funzioni superiori']].forEach(function(x){f(x[0],x[1])});
  NC.forEach(function(r){f('nc.'+r[0]+'.dx',r[1]+' '+r[3]+' dx');f('nc.'+r[0]+'.sx',r[1]+' '+r[3]+' sx');f('nc.'+r[0]+'.n',r[1]+' dettaglio')});
  f('m.tono','Tono');f('m.trof','Trofismo');f('m.m1','Mingazzini I');f('m.m2','Mingazzini II');
  C.push({h:'Deficit di forza su base antalgica',t:'n',fn:function(s){return (s['m.antalg']||[]).length?1:''}});
  M.forEach(function(m){['dx','sx'].forEach(function(s){f('f.'+m[0]+'.'+s,'MRC '+m[1]+' '+m[2]+' '+s,'n')})});
  ROT.forEach(function(r){['dx','sx'].forEach(function(s){f('r.'+r[0]+'.'+s,'ROT '+r[1]+' '+s,'n')})});
  [['bab','Babinski'],['hof','Hoffmann'],['clo','Clono']].forEach(function(r){['dx','sx'].forEach(function(s){f('r.'+r[0]+'.'+s,r[1]+' '+s)})});
  f('s.sens','Sensibilità tattile/dolorifica');C=C.concat(multiCols('s.sint','Sintomi sensitivi',SINT));
  [['s.dist','Distribuzione sensitiva'],['s.liv','Livello sensitivo'],['s.sella','Sensibilità sellare'],['s.prop','Propriocezione'],['s.coord','Coordinazione'],['s.romb','Romberg'],['s.marcia','Deambulazione'],['s.punte','Punte e talloni']].forEach(function(x){f(x[0],x[1])});
  C=C.concat(multiCols('s.tipo','Marcia',GAIT));
  f('g.men','Segni meningei');
  [['las','Lasègue'],['lcr','Lasègue crociato'],['was','Wassermann']].forEach(function(r){['dx','sx'].forEach(function(s){f('g.'+r[0]+'.'+s,r[1]+' '+s);f('g.'+r[0]+'.'+s+'g',r[1]+' '+s+' gradi')})});
  [['g.sf','Sfinteri'],['g.ton','Tono sfinterico anale'],['g.rach','Rachide'],['g.fer','Ferita / drenaggi']].forEach(function(x){f(x[0],x[1])});
  SCALE.forEach(function(s){f('g.sc.'+s[0],s[1],s[0]==='asia'?'s':'n')});
  f('z.conf','Rispetto all\'esame precedente');f('z.confd','Data esame precedente','d');f('z.confn','Dettaglio confronto');f('z.img','Imaging');f('z.ind','Indicazioni');f('z.firma','Firma');
  return C}
function baseCols(withCat){return [
  {h:'Codice paziente',fn:function(s,v,p){return p.code}},{h:'ID visita',fn:function(s,v){return v.code}},
  {h:'Data visita',t:'d',k:'d.date'},{h:'Tipo visita',fn:function(s){return vt(s['d.tipo'],1)}}]
  .concat(withCat?[{h:'Patologia',fn:function(s){return CAT[s['d.cat']]?CAT[s['d.cat']].l:''}}]:[])
  .concat([{h:'Età alla visita',t:'n',fn:function(s,v,p){var a=ageAt(p.nascita,s['d.date']);return isFinite(a)?a:''}},
  {h:'Sesso',fn:function(s,v,p){return p.sesso||''}},{h:'Diagnosi / intervento',k:'d.dx'}])}
function rowOfV(cols,v){return withVisit(v,function(){return cols.map(function(c){return cell(c.t,c.fn?c.fn(S,v,P):S[c.k])})})}
function exportXlsx(anon){
  flush();
  var vs=DB.vis.slice().sort(function(a,b){return (a.data['d.date']||'').localeCompare(b.data['d.date']||'')||String(a.code).localeCompare(String(b.code))});
  var sheets=[];
  if(!anon)sheets.push({name:'Anagrafica',rows:[['Codice','Cognome','Nome','Data di nascita','Sesso','N. cartella','Note','N. visite']].concat(
    DB.pat.slice().sort(function(a,b){return String(a.code).localeCompare(String(b.code))}).map(function(p){
      return [p.code,p.cognome||'',p.nome||'',p.nascita?{d:p.nascita}:'',p.sesso||'',p.cartella||'',p.note||'',visitsOf(p.id).length]}))});
  var vc=baseCols(true).concat([{h:'Lato dominante',k:'d.dom'}],specCols(NX.QUAD),specCols(NX.ANAM),specCols(NX.TER),specCols(NX.DEC),examCols(),
    [{h:'Testo per cartella',fn:function(s,v){return v.repEdit?v.rep1:reports()[0]}}]);
  sheets.push({name:'Visite',rows:[vc.map(function(c){return c.h})].concat(vs.map(function(v){return rowOfV(vc,v)}))});
  NX.CAT.forEach(function(c){var cv=vs.filter(function(v){return v.data['d.cat']===c.id});if(!cv.length)return;
    var cc=baseCols(false).concat(specCols(c.f));
    sheets.push({name:c.s,rows:[cc.map(function(x){return x.h})].concat(cv.map(function(v){return rowOfV(cc,v)}))})});
  deliver(window.XLSXW.build(sheets),'NeurosurgClinic_'+(anon?'anonimo_':'')+td()+'.xlsx');
  if(!anon){setMeta({xlsx:Date.now()});
    setTimeout(function(){if(!$('v-home').hidden){$('v-home').innerHTML=homeHtml();drawPList();$('hmsg').textContent='Excel creato con '+vs.length+' visite: salvalo in File o OneDrive.'}},600)}
}
function backup(){
  flush();
  var data={app:'neurosurg-exam',version:1,exported:new Date().toISOString(),pat:DB.pat,vis:DB.vis};
  deliver(new Blob([JSON.stringify(data)],{type:'application/json'}),'NeurosurgClinic_backup_'+td()+'.json');
  setMeta({backup:Date.now()});
  setTimeout(function(){if(!$('v-home').hidden){$('v-home').innerHTML=homeHtml();drawPList();$('hmsg').textContent='Backup creato: salvalo in File o OneDrive.'}},600);
}
function restore(text){
  var d,np=0,nv=0;try{d=JSON.parse(text)}catch(e){}
  if(!d||d.app!=='neurosurg-exam'||!Array.isArray(d.pat)||!Array.isArray(d.vis)){$('hmsg').textContent='Il file non è un backup di Neurosurg Clinic.';return}
  function merge(list,arr,store){arr.forEach(function(r){if(!r||!r.id)return;var e=DB[list].filter(function(x){return x.id===r.id})[0];
    if(!e){DB[list].push(r);put(store,r);list==='pat'?np++:nv++}else if((r.updated||0)>(e.updated||0)){Object.assign(e,r);put(store,e);list==='pat'?np++:nv++}})}
  merge('pat',d.pat,'pat');merge('vis',d.vis.filter(function(v){return v&&v.data}),'vis');
  $('v-home').innerHTML=homeHtml();drawPList();
  $('hmsg').textContent='Ripristino completato: '+np+' pazienti e '+nv+' visite aggiunti o aggiornati.';
}
function deliver(blob,name){
  var f=null;try{f=new File([blob],name,{type:blob.type})}catch(e){}
  if(f&&navigator.maxTouchPoints>1&&navigator.canShare&&navigator.canShare({files:[f]})){
    navigator.share({files:[f],title:name}).catch(function(e){if(e&&e.name!=='AbortError')download(blob,name)});return}
  download(blob,name);
}
function download(blob,name){var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();
  setTimeout(function(){URL.revokeObjectURL(a.href);a.remove()},4000)}

/* ---------- eventi ---------- */
function pressed(g,on){g.querySelectorAll('button').forEach(function(x){x.setAttribute('aria-pressed',on(x.dataset.v))})}
function arm(el){var t=el.textContent,wasDanger=el.classList.contains('danger');el.dataset.armed='1';el.textContent='Conferma: elimina';el.classList.add('danger');
  setTimeout(function(){if(el.isConnected){el.dataset.armed='';el.textContent=t;if(!wasDanger)el.classList.remove('danger')}},4000)}
function doAct(el){var a=el.dataset.act,id=el.dataset.id;
  if(el.hasAttribute('data-arm')&&!el.dataset.armed){arm(el);return}
  if(a==='home')openHome();else if(a==='newpat')newPat();else if(a==='openpat')openPat(id||P.id);
  else if(a==='newvis')newVisit();else if(a==='openvis'){if(!V||V.id!==id)openVisit(id)}else if(a==='closevis')closeVis(id);else if(a==='delvis')delVisit(id);else if(a==='delpat')delPat();
  else if(a==='xlsx')exportXlsx(false);else if(a==='xlsx-anon')exportXlsx(true);else if(a==='backup')backup();
  else if(a==='restore')$('rfile').click();
}
document.addEventListener('click',function(e){
  var b=e.target.closest('.seg button');
  if(b){var g=b.parentElement,val=b.dataset.v;
    if(g.dataset.k){var k=g.dataset.k;
      if(g.dataset.multi){var a=S[k]||[],i=a.indexOf(val);if(i>-1)a.splice(i,1);else a.push(val);if(a.length)S[k]=a;else delete S[k]}
      else if(S[k]===val&&FIXED.indexOf(k)<0)delete S[k];else S[k]=val;
      pressed(g,function(x){return g.dataset.multi?(S[k]||[]).indexOf(x)>-1:S[k]===x});
      if(k==='d.tpl'||k==='d.tipo')applyTpl();
      changed();return}
    if(g.dataset.p){var pk=g.dataset.p;P[pk]=P[pk]===val?'':val;pressed(g,function(x){return P[pk]===x});saveP();renderHeader('pat');return}
    if(g.dataset.nv){var nk=g.dataset.nv;NEWV[nk]=val;pressed(g,function(x){return NEWV[nk]===x});return}
  }
  var go_=e.target.closest('[data-go]');if(go_){show(go_.dataset.go);return}
  var nm=e.target.closest('[data-norm]');if(nm){NORM[nm.dataset.norm]();mountPanel(nm.dataset.norm);changed();return}
  var cp=e.target.closest('[data-copy]');
  if(cp){var ta=$(cp.dataset.copy),orig=cp.textContent;
    var done=function(m){cp.textContent=m;setTimeout(function(){cp.textContent=orig},1800)};
    try{navigator.clipboard.writeText(ta.value).then(function(){done('Copiato')},function(){ta.select();done('Selezionato: copia con Ctrl/Cmd+C')})}catch(x){ta.select();done('Selezionato: copia con Ctrl/Cmd+C')}
    return}
  if(e.target.id==='regen'){V.repEdit=false;delete V.rep1;delete V.rep2;saveV();fillReport();return}
  var act=e.target.closest('[data-act]');if(act)doAct(act);
});
function onField(e){
  var t=e.target;
  if(t.id==='gsearch'){drawGlossary(t.value);return}
  if(t.id==='psearch'){HQ=t.value;drawPList();return}
  if(t.id==='rep1'||t.id==='rep2'){if(V.repEdit&&V[t.id]===t.value)return;V.repEdit=true;V.rep1=$('rep1').value;V.rep2=$('rep2').value;$('repnote').hidden=false;saveV();return}
  if(t.id==='rfile'){var fi=t.files&&t.files[0];if(!fi)return;var rd=new FileReader();rd.onload=function(){restore(String(rd.result))};rd.readAsText(fi);t.value='';return}
  if(t.dataset.p){if((P[t.dataset.p]||'')===t.value)return;P[t.dataset.p]=t.value;saveP();renderHeader('pat');return}
  if(t.dataset.nv){NEWV[t.dataset.nv]=t.value;return}
  if(!t.dataset.k||t.closest('.seg'))return;
  var k=t.dataset.k;if((S[k]||'')===t.value)return;
  if(t.value==='')delete S[k];else S[k]=t.value;
  if(t.tagName==='SELECT')paint(t);
  if(k==='d.cat'){var c=CAT[t.value];if(c)S['d.tpl']=c.tpl;changed();applyTpl();return}
  if(k==='d.date'){renderHeader('visit');var vs=$('p-vis');if(vs&&!vs.hidden)return changed()}
  changed();
}
document.addEventListener('input',onField);
document.addEventListener('change',onField);
document.addEventListener('visibilitychange',function(){if(document.hidden)flush()});
window.addEventListener('pagehide',flush);

/* ---------- avvio ---------- */
openDB().then(function(){return Promise.all([getAll('pat'),getAll('vis')])}).then(function(r){
  DB.pat=r[0];DB.vis=r[1];
  if(!idb)banner('Archivio non disponibile (es. navigazione privata): i dati non verranno conservati.');
  try{if(navigator.storage&&navigator.storage.persist)navigator.storage.persist()}catch(e){}
  openHome();
});
})();
