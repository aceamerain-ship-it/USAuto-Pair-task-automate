/**
 * USAUTO Pair Project Automation — Pure Node.js
 * Produces output matching the correct reference template exactly.
 */
'use strict';

const fs       = require('fs');
const path     = require('path');
const unzipper = require('unzipper');
const JSZip    = require('jszip');
const { SaxesParser } = require('saxes');

let INPUT_FILE  = process.argv[2];
let OUTPUT_FILE = process.argv[3];

const DATA_START = 2, DATA_END = 200;
const COL_PASTE_VAL  = 8;   // H
const SS_PASTE_HDR   = 1969891; // "PASTE AS VALUE HERE" already in input SS
const PATH_SHEET3    = 'xl/worksheets/sheet3.xml';
const PATH_SHEET4    = 'xl/worksheets/sheet4.xml';
const PATH_SHEET5    = 'xl/worksheets/sheet5.xml';
const PATH_SS        = 'xl/sharedStrings.xml';
const PATH_CA        = 'xl/worksheets/sheet2.xml';  // CA Data Combined
const CA_COLS = new Set(['B','C','H','J','K','L','N','O','P','R','S','X','Z','AA','AB','AD','AE','AF','AH','AI']);
const PW_EVAL_MAP = {C:'S', D:'B', E:'R', F:'AH', G:'AI', H:'J', I:'Z'};  // photowork col -> CA col
const DESC_SRC = {
  C:{src:'pw',ref:'C'}, D:{src:'pw',ref:'H'}, E:{src:'pw',ref:'I'},
  F:{src:'ca',ref:'H'}, G:{src:'ca',ref:'X'}, H:{src:'pw',ref:'D'},
  I:{src:'ca',ref:'R'}, J:{src:'ca',ref:'K'}, K:{src:'ca',ref:'L'},
  L:{src:'ca',ref:'AA'},M:{src:'ca',ref:'AB'},N:{src:'ca',ref:'N'},
  O:{src:'ca',ref:'O'}, P:{src:'ca',ref:'P'}, Q:{src:'ca',ref:'AD'},
  R:{src:'ca',ref:'AE'},S:{src:'ca',ref:'AF'},T:{src:'pw',ref:'F'},
  U:{src:'pw',ref:'G'},
};

// ── Helpers ──────────────────────────────────────────────────────────────────
function colLetter(n) {
  let r = '';
  while (n > 0) { const rem = (n-1)%26; r = String.fromCharCode(65+rem)+r; n = Math.floor((n-1)/26); }
  return r;
}
function colNum(l) { let n=0; for (const ch of l) n=n*26+(ch.charCodeAt(0)-64); return n; }
function xmlEsc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function formulaEsc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }  // no quote escaping
function stripPrefix(s) { return s.replace(/^ITEMIMAGEURL\d+=\s*/i,'').trim(); }
function hStyle(rn) { return rn===1?'61':rn===2?'76':'64'; }

function makeSsCell(ref, ssIdx, style) {
  const s = style ? ` s="${style}"` : '';
  return `<c r="${ref}"${s} t="s"><v>${ssIdx}</v></c>`;
}
function cleanCombined(raw) {
  // Remove N/A entries, strip duplicates, normalize Combined Partslink or OEM string
  if (!raw || raw === '#N/A') return null;
  // Decode HTML entities
  let decoded = raw.replace(/&#x2F;/g, '/').replace(/&#47;/g, '/').replace(/&amp;/g, '&');
  // Normalize all whitespace
  decoded = decoded.split(/\s+/).join(' ').trim();
  let prefix = '';
  let val = decoded;
  // Flexible prefix match: handles missing/extra space after colon
  for (const keyword of ['Partslink Numbers', 'OEM Numbers']) {
    const pat = decoded.match(new RegExp('^' + keyword.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + ':\\s*(.*)', 'i'));
    if (pat) { prefix = keyword + ': '; val = pat[1]; break; }
  }
  const parts = val.split(',').map(s => s.trim()).filter(s => s && !/^N\/A$/i.test(s.trim()));
  // Deduplicate using normalized key (case-insensitive, whitespace-normalized)
  const seen = new Set(); const unique = [];
  for (const p of parts) {
    const key = p.split(/\s+/).join(' ').toUpperCase();
    if (!seen.has(key)) { seen.add(key); unique.push(p.split(/\s+/).join(' ')); }
  }
  return unique.length ? prefix + unique.join(', ') : null;
}

function makeFormulaCell(col, style, formula, rn, cellType='str', cached='#N/A') {
  return `<c r="${col}${rn}" s="${style}" t="${cellType}"><f>${formulaEsc(formula)}</f><v>${cached}</v></c>`;
}
function makeBlank(ref, style) { return `<c r="${ref}" s="${style}"/>`; }

function colNum(l) { let n=0; for (const ch of l) n=n*26+(ch.charCodeAt(0)-64); return n; }

function sortRowCells(body) {
  // Sort all cells in a row body by ascending column number (Excel requires strict order)
  const cellRe = /<c r="([A-Z]+)\d+"(?:[^>]*\/>[^<]*|[^>]*>[\s\S]*?<\/c>)/g;
  const cells = [];
  for (const m of body.matchAll(cellRe)) {
    cells.push([colNum(m[1]), m[0]]);
  }
  cells.sort((a,b) => a[0]-b[0]);
  return cells.map(c=>c[1]).join('');
}

function removeCells(colsSet, rn, body) {
  const colRefs = [...colsSet].map(c=>`${c}${rn}`).join('|');
  return body.replace(
    new RegExp(`<c r="(?:${colRefs})"(?:[^>]*\\/>[^<]*|[^>]*>[\\s\\S]*?<\\/c>)`,'g'), '');
}

// ── Formula definitions ───────────────────────────────────────────────────────
const PHOTO_FORMULA_COLS = [
  {col:'C', style:'79', caCol:'S',  tmpl:r=>`INDEX('CA Data Combined'!S:S,MATCH(Photowork!B${r},'CA Data Combined'!C:C,0))`},
  {col:'D', style:'89', caCol:'B',  tmpl:r=>`INDEX('CA Data Combined'!B:B,MATCH(Photowork!B${r},'CA Data Combined'!C:C,0))`},
  {col:'E', style:'79', caCol:'R',  tmpl:r=>`INDEX('CA Data Combined'!R:R,MATCH(Photowork!B${r},'CA Data Combined'!C:C,0))`},
  {col:'F', style:'28', caCol:'AH', tmpl:r=>`INDEX('CA Data Combined'!AH:AH,MATCH(Photowork!B${r},'CA Data Combined'!C:C,0))`},
  {col:'G', style:'28', caCol:'AI', tmpl:r=>`INDEX('CA Data Combined'!AI:AI,MATCH(Photowork!B${r},'CA Data Combined'!C:C,0))`},
  {col:'H', style:'3',  caCol:'J',  tmpl:r=>`INDEX('CA Data Combined'!J:J,MATCH(Photowork!B${r},'CA Data Combined'!C:C,0))`},
  {col:'I', style:'3',  caCol:'Z',  tmpl:r=>`INDEX('CA Data Combined'!Z:Z,MATCH(Photowork!B${r},'CA Data Combined'!C:C,0))`},
  {col:'K', style:'15', caCol:null, tmpl:r=>`CONCATENATE(D${r},"_master_SET_v2.JPG")`},
  {col:'M', style:'15', caCol:null, tmpl:r=>`CONCATENATE(D${r},"_02_SET_v2.JPG")`},
  {col:'O', style:'15', caCol:null, tmpl:r=>`CONCATENATE(D${r},"_03_SET_v2.JPG")`},
  {col:'Q', style:'15', caCol:null, tmpl:r=>`CONCATENATE(D${r},"_04_SET_v2.JPG")`},
  {col:'S', style:'15', caCol:null, tmpl:r=>`CONCATENATE(D${r},"_05_SET_v2.JPG")`},
  {col:'U', style:'15', caCol:null, tmpl:r=>`CONCATENATE(D${r},"_06_SET_v2.JPG")`},
  {col:'W', style:'15', caCol:null, tmpl:r=>`CONCATENATE(D${r},"_07_SET_v2.JPG")`},
  {col:'Y', style:'15', caCol:null, tmpl:r=>`CONCATENATE(D${r},"_08_SET_v2.JPG")`},
];
const AA_STYLE_MAP = Object.assign({},
  ...[15,52,89,126,163,200].map(r=>({[r]:'17'})),
  ...[20,22,25,57,59,62,94,96,99,131,133,136,168,170,173].map(r=>({[r]:'78'})),
  ...[6,43,80,117,154,191].map(r=>({[r]:'80'}))
);
const HLOOKUP_AE_AL = ['AE','AF','AG','AH','AI','AJ','AK','AL'];
const HLOOKUP_HEADER = {
  AD:'_master', AE:'_01', AF:'_02', AG:'_03',
  AH:'_04',     AI:'_05', AJ:'_06', AK:'_07', AL:'_08',
};
const PHOTO_TOUCH = new Set([...PHOTO_FORMULA_COLS.map(x=>x.col),'B','J','AA','AB','AD',...HLOOKUP_AE_AL]);

const DESC_FORMULA_COLS = [
  {col:'C', style:'54', tmpl:r=>`Photowork!C${r}`},
  {col:'D', style:'35', tmpl:r=>`Photowork!H${r}`},
  {col:'E', style:'43', tmpl:r=>`Photowork!I${r}`},
  {col:'F', style:'21', tmpl:r=>`INDEX('CA Data Combined'!H:H,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'G', style:'49', tmpl:r=>`INDEX('CA Data Combined'!X:X,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'H', style:'20', tmpl:r=>`Photowork!D${r}`},
  {col:'I', style:'79', tmpl:r=>`INDEX('CA Data Combined'!R:R,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'J', style:'20', tmpl:r=>`INDEX('CA Data Combined'!K:K,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'K', style:'20', tmpl:r=>`INDEX('CA Data Combined'!L:L,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'L', style:'50', tmpl:r=>`INDEX('CA Data Combined'!AA:AA,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'M', style:'50', tmpl:r=>`INDEX('CA Data Combined'!AB:AB,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'N', style:'20', tmpl:r=>`INDEX('CA Data Combined'!N:N,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'O', style:'20', tmpl:r=>`INDEX('CA Data Combined'!O:O,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'P', style:'20', tmpl:r=>`INDEX('CA Data Combined'!P:P,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'Q', style:'50', tmpl:r=>`INDEX('CA Data Combined'!AD:AD,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'R', style:'50', tmpl:r=>`INDEX('CA Data Combined'!AE:AE,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'S', style:'50', tmpl:r=>`INDEX('CA Data Combined'!AF:AF,MATCH(Description!B${r},'CA Data Combined'!C:C,0))`},
  {col:'T', style:'21', tmpl:r=>`Photowork!F${r}`},
  {col:'U', style:'59', tmpl:r=>`Photowork!G${r}`},
];
const DESC_TOUCH = new Set([...DESC_FORMULA_COLS.map(x=>x.col),'A','B']);

// ── Main ─────────────────────────────────────────────────────────────────────

// ── Pair Description Processor (ported from EU-Pair_Description_Processor.html) ─
const INFO_FINISH = ["Textured","Primed","Chrome","Satin","Bright","Glossy","Gloss","Matte",
  "Dull","Oil Painted","Oil-Painted","Paint to Match","Painted","Plain","Powder-coated",
  "Powdercoated","Phosphate Coated","Shiny","Smoked","Smooth","Electroplated","Zinc Plated",
  "Zinc-Plated","Zinc Coated","Zinc-Coated","Paintable"];
const EXTRA_POS = ["Inner","Outer","Lower","Upper","Rearward","Frontward","Forward"];
const KEEP_LOWER_PD = new Set(["and","or","of","on","in","at"]);
const PREMIUM_AFTERMARKET_TEXT = "<br><br><b>Premium Aftermarket Replacement Parts:</b><br>- Every component undergoes a thorough evaluation and meticulous inspection to establish adherence to our strict safety standards and high-quality benchmarks.<br><br><b>Engineered For Longevity:</b><br>- Crafted from top-tier materials, this product has undergone rigorous testing to provide maximum durability.<br><br><b>Please confirm OEM # or Partslink # matches exactly before purchasing, otherwise, the item might not fit.";

const PAIR_DESC_COLS = {
  V:  {style:'2', key:'pair_description'},
  W:  {style:'2', key:'premium_aftermarket'},
  Y:  {style:'2', key:'bullet_point_1'},
  Z:  {style:'2', key:'bullet_point_2'},
  AA: {style:'2', key:'bullet_point_3'},
  AB: {style:'2', key:'bullet_point_4'},
  AC: {style:'2', key:'bullet_point_5'},
  AE: {style:'2', key:'attr1_part'},
  AF: {style:'2', key:'attr2_placement'},
  AG: {style:'2', key:'attr3_orientation'},
  AH: {style:'2', key:'attr4_color'},
  AI: {style:'2', key:'attr5_lens_color'},
  AJ: {style:'2', key:'attr6_finish'},
  AK: {style:'2', key:'attr7_material'},
  AL: {style:'2', key:'attr8_light_source'},
  AM: {style:'2', key:'attr9_housing_color'},
  AN: {style:'2', key:'attr10_cert'},
};
const PAIR_DESC_TOUCH = new Set(Object.keys(PAIR_DESC_COLS));
const ALWAYS_VALUE_COLS = new Set(['AG','AH','AI','AJ','AK','AL','AM','AN']);

function titleCasePos(str) {
  if (!str) return str;
  return str.split(/(\s+|,)/).map(tok => {
    if (tok === ',' || /^\s+$/.test(tok)) return tok;
    const lo = tok.toLowerCase();
    if (KEEP_LOWER_PD.has(lo)) return lo;
    return lo.charAt(0).toUpperCase() + lo.slice(1);
  }).join('').replace(/\s+,/g, ',');
}

function parsePosDet(line) {
  let txt = (line||'').trim().replace(/^-\s*/,'').replace(/^Position:\s*/i,'');
  const mains = [];
  if (/\bFront\b/i.test(txt)) mains.push('front');
  if (/\bRear\b/i.test(txt)) mains.push('rear');
  let side = '';
  if (/\bdriver\s*and\s*passenger\s*side\b/i.test(txt)) side='driver and passenger side';
  else if (/\bdriver\s*or\s*passenger\s*side\b/i.test(txt)) side='driver or passenger side';
  else if (/\bdriver\s*\/\s*passenger\s*side\b/i.test(txt)) side='driver or passenger side';
  else if (/\bdriver\s*side\b/i.test(txt)) side='driver side';
  else if (/\bpassenger\s*side\b/i.test(txt)) side='passenger side';
  const extras = EXTRA_POS.filter(e=>new RegExp('\\b'+e+'\\b','i').test(txt)).map(e=>e.toLowerCase());
  return {mains,side,extras};
}
function posScore(d){return (!d)?0:((d.mains||[]).length+(d.side?2:0)+(d.extras||[]).length);}
function buildPos(d){
  const mains=(d&&d.mains)||[]; const side=(d&&d.side)||''; const extras=(d&&d.extras)||[];
  const ms=new Set(mains.map(x=>x.toLowerCase()));
  let mStr=ms.has('front')&&ms.has('rear')?'front and rear':ms.has('front')?'front':ms.has('rear')?'rear':'';
  // Format extras
  const es=new Set(extras.map(x=>x.toLowerCase())); const eg=[];
  [['inner','outer'],['lower','upper'],['rearward','frontward']].forEach(([a,b])=>{
    if(es.has(a)&&es.has(b)){eg.push(`${a} and ${b}`);es.delete(a);es.delete(b);}
    else{if(es.has(a)){eg.push(a);es.delete(a);}if(es.has(b)){eg.push(b);es.delete(b);}}
  });
  if(es.has('forward')){eg.push('forward');es.delete('forward');}
  const exStr=eg.join(', '); const hasEx=exStr.trim()!=='';
  let pos='';
  if(mStr&&!side&&!hasEx) pos=mStr;
  else if(mStr&&side&&!hasEx) pos=`${mStr}, ${side}`;
  else if(!mStr&&side&&!hasEx) pos=side;
  else if(mStr&&side&&hasEx) pos=`${mStr} ${side}, ${exStr}`;
  else if(!mStr&&side&&hasEx) pos=`${side}, ${exStr}`;
  else if(mStr&&!side&&hasEx) pos=`${mStr}, ${exStr}`;
  else pos=exStr;
  return pos.trim().toLowerCase();
}
function mergePosDets(d1,d2){
  const a=d1||{mains:[],side:'',extras:[]}; const b=d2||{mains:[],side:'',extras:[]};
  const ms=new Set([...(a.mains||[]),...(b.mains||[])].map(x=>x.toLowerCase()));
  const mains=['front','rear'].filter(x=>ms.has(x));
  const exS=new Set([...(a.extras||[]),...(b.extras||[])].map(x=>x.toLowerCase()));
  const extras=EXTRA_POS.map(e=>e.toLowerCase()).filter(e=>exS.has(e));
  const sA=(a.side||'').toLowerCase(),sB=(b.side||'').toLowerCase();
  let side='';
  if(sA==='driver and passenger side'||sB==='driver and passenger side') side='driver and passenger side';
  else if(sA==='driver or passenger side'&&sB==='driver or passenger side') side='driver and passenger side';
  else if(new Set([sA,sB]).size===2&&new Set([sA,sB]).has('driver side')&&new Set([sA,sB]).has('passenger side')) side='driver and passenger side';
  else if(['driver or passenger side'].includes(sA)&&['driver side','passenger side'].includes(sB)) side='driver and passenger side';
  else if(['driver or passenger side'].includes(sB)&&['driver side','passenger side'].includes(sA)) side='driver and passenger side';
  else if(sA&&sB&&sA===sB) side=sA;
  else side=sA||sB||'';
  return {mains,side,extras};
}
function isStandaloneSide(l){return /^\s*-?\s*(Driver Side|Passenger Side|Driver or Passenger Side|Driver and Passenger Side)\s*$/i.test(l);}
function convertUnitsStr(s){
  return s.replace(/(\d+)\s*mm\.?(?!\w)/gi,'$1 millimeters')
           .replace(/(\d+)\s*in\.?(?!\w)/gi,'$1 inches')
           .replace(/(\d+)\s*inch\.?(?!\w)/gi,'$1 inches')
           .replace(/(\d+)\s*lb\.?(?!\w)/gi,'$1 Pounds')
           .replace(/(\d+)\s*"/g,'$1 inches')
           .replace(/\bx\b/g,'by');
}
function normHdrs(s){
  return s.replace(/<b>(PRODUCT INFO):?\s*<\/b>/gi,'<b>Product Info: </b>')
           .replace(/<b>(PRODUCT INTERCHANGE):?\s*<\/b>/gi,'<b>Product Interchange: </b>');
}
function normAttrLine(l){
  l=l.replace(/^(Lens\s*Color:\s*)(.+)$/i,(_,p1,p2)=>p1+p2.replace(/\s*lens\s*$/i,'').trim());
  l=l.replace(/^(Housing\s*Color:\s*)(.+)$/i,(_,p1,p2)=>p1+p2.replace(/\s*housing\s*$/i,'').trim());
  return l;
}
function pluralizePts(l){
  const t=l.trim().replace(/^[-\s]+/,'');
  if(t.toLowerCase().startsWith('part: ')){
    const part=t.slice(6).trim();
    if(part.toLowerCase()==='glass') return 'Part: Glasses';
    const words=part.split(' ');
    if(words.length&&!words[words.length-1].endsWith('s')) words[words.length-1]+='s';
    return 'Part: '+words.join(' ');
  }
  return l;
}
function cleanLens(v){return (!v||v==='N/A')?v||'N/A':v.replace(/\s*lens\s*$/i,'').trim();}
function cleanHousing(v){return (!v||v==='N/A')?v||'N/A':v.replace(/\s*housing\s*$/i,'').trim();}
function stripUsa(sku){ return (sku||'').trim().replace(/^USA-/i,''); }

function cleanInput(v){
  if(!v) return '';
  v=v.trim();
  if(['N/A','#N/A','#VALUE!','#REF!','#NAME?','#DIV/0!'].includes(v.toUpperCase())) return '';
  v=v.replace(/^(Partslink Numbers:|OEM Numbers:)\s*(?:N\/A(?:,\s*)?)+$/i,'').trim();
  return v;
}
function handleColorFinish(info){
  const parts=info.split(/[;,\/|]+/).map(s=>s.trim()).filter(Boolean);
  const finishes=[],colors=[],housingColors=[];
  for(const p of parts){
    let fnd=null;
    for(const f of INFO_FINISH){if(new RegExp('\\b'+f.replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&')+'\\b','i').test(p)){fnd=f;break;}}
    let rem=fnd?p.replace(new RegExp(fnd.replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&'),'i'),'').trim():p;
    if(/\bhousing\b/i.test(rem)){let hc=rem.replace(/\bhousing\b/i,'').replace(/\bcolor\b/i,'').trim();if(hc)housingColors.push(hc);if(fnd&&!finishes.includes(fnd))finishes.push(fnd);continue;}
    if(fnd){if(!finishes.includes(fnd))finishes.push(fnd);if(rem)colors.push(rem);}else colors.push(rem);
  }
  const res={};
  if(finishes.length) res.Finish='Finish: '+finishes.join(', ');
  if(colors.length) res.Color='Color: '+colors.join(', ');
  if(housingColors.length) res['Housing Color']='Housing Color: '+housingColors.join(', ');
  return res;
}


function cleanPairDescription(html) {
  if (!html) return html;
  // Remove residual #N/A lines (Excel error values that leaked into Product Interchange)
  html = html.replace(/(?:<br>- #N\/A)+(<br>|$)/g, '$1');
  // (trailing <br> stripped at return below)

  const COMPOUND_PUDDLE_AUTO = [
    /^Without Puddle Light,\s*Without Auto.?Dimm\w*$/i,
    /^No Puddle Light,\s*(Auto Dimming:\s*)?No Auto.?Dimm\w*$/i,
    /^Puddle Light Included:\s*(Without|No) Puddle Light,\s*Auto Dimming:\s*(Without|No) Auto.?Dimm\w*$/i,
    /^Puddle Light Included:\s*(Without|No) Puddle Light,\s*Notes:\s*(Without|No) Auto.?Dimm\w*$/i,
  ];
  const SIDE_CAM_POS = /^No Puddle Light,\s*No Auto.?Dimm\w*,\s*Side View Camera$/i;
  const SIDE_CAM_NEG = /^Notes:\s*(Without|No) Auto.?Dimm\w+,\s*(Without|No) Side View Camera$/i;
  const DEDUP = [
    {key:'signal_no',   pats:[/^Built In Signal Light:\s*(Without|No) Signal Light$/i, /^Notes:\s*Without Signal Light$/i, /^Without Signal Light$/i, /^No Signal Light$/i]},
    {key:'blind_no',    pats:[/^Blind Spot Detection:\s*(Without|No) Blind Spot (Feature|Detection)$/i, /^Without Blind Spot (Feature|Detection)$/i, /^No Blind Spot (Feature|Detection)$/i]},
    {key:'mem_no',      pats:[/^Memory Recall:\s*(Without|No) Memory$/i, /^Without Memory$/i, /^No Memory$/i]},
    {key:'puddle_no',   pats:[/^Puddle Light Included:\s*(Without|No) Puddle Light$/i, /^Without Puddle Light$/i, /^No Puddle Light$/i]},
    {key:'auto_no',     pats:[/^Auto.?Dimming:\s*(Without|No) Auto.?Dimm\w*$/i, /^Auto Dimming:\s*(Without|No) Auto.?Dimm\w*$/i, /^(Without|No) Auto.?Dimm\w*$/i, /^Notes:\s*(Without|No) Auto.?Dimm\w*$/i]},
    {key:'glass_pow',   pats:[/^Glass Adjustment Method:\s*Power( Adjust)?$/i, /^Power Adjust$/i, /^Power$/i]},
    {key:'color_black', pats:[/^Color:\s*Black Base$/i, /^Notes:\s*Black [Bb]ase$/i, /^Black [Bb]ase$/i]},
    {key:'heated_non',  pats:[/^Heated:\s*Non-Heated$/i, /^Non-Heated$/i]},
    {key:'pin_plug_3',  pats:[/^GTN 3 Pin Plug$/i, /^3 Pin Plug$/i]},
  ];
  function matchAny(line, pats){ return pats.some(p=>p.test(line)); }
  function groupHit(line){
    if(COMPOUND_PUDDLE_AUTO.some(p=>p.test(line))) return 'compound';
    if(SIDE_CAM_POS.test(line)) return 'sidecam_pos';
    if(SIDE_CAM_NEG.test(line)) return 'sidecam_neg';
    for(const {key,pats} of DEDUP) if(matchAny(line,pats)) return key;
    return null;
  }

  const parts = html.split(/(<br>)/);
  const segs = [];
  for(let i=0;i<parts.length;){
    const text=parts[i]; const br=(i+1<parts.length&&parts[i+1]==='<br>')?'<br>':'';
    segs.push([text,br]); i+=br?2:1;
  }
  const seen={}; const keep=segs.map(()=>true);
  let seenPuddle=false,seenAuto=false,seenBoth=false,seenSideCamPos=false;

  for(let idx=0;idx<segs.length;idx++){
    let [text,br]=segs[idx];
    const raw=text.trim();
    if(!raw.startsWith('- ')) continue;
    let t=raw.replace(/^-\s*/,'').trim();
    if(!t) continue;

    // Normalize "Heated: Heated" -> "Heated"
    if(/^Heated:\s*Heated$/i.test(t)){
      t='Heated'; segs[idx][0]=text.replace(raw,'- '+t);
    }
    // Dedup plain "Heated" standalone
    if(t==='Heated'){if(seen.heated){keep[idx]=false;continue;}seen.heated=true;}

    // Notes with extra info + auto-dim reference — keep but mark auto as seen
    if(/^Notes:/i.test(t)&&!/^Notes:\s*(Without|No) Auto.?Dimm/i.test(t)&&/(Without|No) Auto.?Dimm/i.test(t)){
      seenAuto=true; continue;
    }

    const grp=groupHit(t);
    if(!grp) continue;

    if(grp==='compound'){
      if(seenBoth||(seenPuddle&&seenAuto)){keep[idx]=false;}
      else{seenBoth=seenPuddle=seenAuto=true;}
    } else if(grp==='sidecam_pos'){
      seenSideCamPos=true;
    } else if(grp==='sidecam_neg'){
      if(seenSideCamPos) keep[idx]=false;
    } else if(grp==='puddle_no'){
      if(seenPuddle||seenBoth){keep[idx]=false;}else seenPuddle=true;
    } else if(grp==='auto_no'){
      if(seenAuto||seenBoth){keep[idx]=false;}else seenAuto=true;
    } else {
      if(seen[grp]!==undefined){keep[idx]=false;}else seen[grp]=idx;
    }
  }
  const raw = segs.filter((_,i)=>keep[i]).map(([t,b])=>t+b).join('');
  return raw.replace(/<br>$/, '');
}

function processPair(input1Xml, input2Xml, partslinkMerged, oemMerged, partsNameMerged='', compSkuDriver='', compSkuPassenger='') {
  const decode=s=>(s||'').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&').replace(/&quot;/g,'"');
  let i1=convertUnitsStr(decode(input1Xml).trim().replace(/ - /g,''));
  let i2=convertUnitsStr(decode(input2Xml).trim().replace(/ - /g,''));
  i1=normHdrs(i1); i2=normHdrs(i2);
  const partslink=cleanInput(partslinkMerged);
  const oem=cleanInput(oemMerged);
  const partsName=cleanInput(partsNameMerged);
  const skuD=stripUsa(compSkuDriver), skuP=stripUsa(compSkuPassenger);
  const skuParts=[skuD,skuP].filter(Boolean);
  const partNumbersStr=skuParts.join(', ');
  const useSkuFallback=(!partslink&&!oem)&&!!partNumbersStr;

  const pi1m=i1.match(/<b>Product Info: <\/b><br>(.+?)<br><br>/is);
  const pi2m=i2.match(/<b>Product Info: <\/b><br>(.+?)<br><br>/is);

  const lines=[]; const seenKeys=new Set(); let bp1=null,bp2=null;
  function processPiBlock(pim,bestPos){
    if(!pim) return bestPos;
    for(let line of pim[1].split('<br>')){
      line=line.replace(/Location:/gi,'Position:').trim();
      if(!line||line.includes('Quantity Sold:')) continue;
      const cand=parsePosDet(line);
      if(posScore(cand)>0&&(!bestPos||posScore(cand)>posScore(bestPos))) bestPos=cand;
      if(/^Color\s*Finish\s*:/i.test(line)){
        const info=line.split(':').slice(1).join(':').trim();
        const cf=handleColorFinish(info);
        ['Finish','Color','Housing Color'].forEach(k=>{if(cf[k]){const v=normAttrLine(pluralizePts(cf[k]));const key=v.trim().toLowerCase();if(!seenKeys.has(key)){seenKeys.add(key);lines.push(v);}}});
      } else {
        const v=normAttrLine(pluralizePts(line));
        const key=v.trim().toLowerCase();
        if(!seenKeys.has(key)){seenKeys.add(key);lines.push(v);}
      }
    }
    return bestPos;
  }
  bp1=processPiBlock(pi1m,bp1); bp2=processPiBlock(pi2m,bp2);

  const hasPos=posScore(bp1)>0||posScore(bp2)>0;
  let filteredLines=lines;
  if(hasPos){
    const mPos=buildPos(mergePosDets(bp1,bp2));
    filteredLines=lines.filter(l=>{const t=l.trim().replace(/^-\s*/,'');return !/^Position:/i.test(t)&&!isStandaloneSide(t);});
    if(mPos) filteredLines.push('Position: '+mPos);
  }

  const brandL=[],partL=[],posL=[],otherL=[];
  for(const l of filteredLines){
    const t=l.trim().replace(/^-\s*/,'');
    if(/^Brand:/i.test(t)&&!brandL.includes(t)) brandL.push(t);
    else if(/^Part:/i.test(t)&&!partL.includes(t)) partL.push(t);
    else if(/^Position:/i.test(t)){const p=buildPos(parsePosDet(t));if(posL.length===0)posL.push(p);}
    else if(isStandaloneSide(t)) {}
    else if(!otherL.includes(t)) otherL.push(t);
  }
  if(posL.length===0){
    for(const l of filteredLines){const t=l.trim().replace(/^-\s*/,'');if(isStandaloneSide(t)||/Driver|Passenger|Front|Rear/i.test(t)){const p=buildPos(parsePosDet(t));if(p){posL.push(p);break;}}}
    if(posL.length===0) posL.push('driver and passenger side');
  }
  // CAPA normalization
  let hasCapa=false;
  for(let i=0;i<otherL.length;i++){if(/\bcapa\b/i.test(otherL[i])){otherL[i]='CAPA Certified';hasCapa=true;}}
  if(hasCapa){const fi=otherL.findIndex(l=>l==='CAPA Certified');for(let i=otherL.length-1;i>=0;i--){if(otherL[i]==='CAPA Certified'&&i!==fi)otherL.splice(i,1);}}

  const ordered=[]; const seenF=new Set(); const partVals=new Set(); let brandKept=false;
  const addLine=l=>{const key=l.trim().toLowerCase();if(seenF.has(key))return;seenF.add(key);ordered.push(normAttrLine(l));};
  for(const l of brandL){if(!brandKept){brandKept=true;addLine(l);}}
  for(const l of partL){const v=l.replace(/^Part:\s*/i,'').trim().toLowerCase();const dup=[...partVals].some(p=>v===p||v+'s'===p||(p.endsWith('s')&&v===p.slice(0,-1)));if(!dup){partVals.add(v);addLine(l);}}
  for(const p of posL){const key='position:'+p.toLowerCase();if(!seenF.has(key)){seenF.add(key);ordered.push('Position: '+titleCasePos(p));}}
  for(const l of otherL){const t=l.trim().replace(/^-\s*/,'');if(/^Color\s*Finish:/i.test(t)) continue;if(/^capa certified$/i.test(t)){addLine('CAPA Certified');}else if(/^with bulbs$/i.test(t)){addLine('With Bulbs');}else{addLine(t);}}

  // Build HTML output
  let htmlOut='',plainOut='';
  if(ordered.length){
    htmlOut+='<b>Product Info: </b><br>';
    for(const l of ordered){const c=l.trim().replace(/^-\s*/,'');htmlOut+=`- ${c}<br>`;}
    htmlOut+='- Sold as Pair<br><br>';
  }
  htmlOut+='<b>Product Interchange:</b><br>';
  if(useSkuFallback){
    htmlOut+=`- Part Numbers: ${partNumbersStr}<br>`;
  } else {
    if(partslink) htmlOut+=`- ${partslink}<br>`;
    if(oem) htmlOut+=`- ${oem}<br>`;
    if(partsName) htmlOut+=`- Alternative Part Number: ${partsName}<br>`;
  }

  // Extract attributes
  const combined=htmlOut.replace(/<b>/g,'').replace(/<\/b>/g,'').replace(/<br>/g,'\n');
  const gv=(re,def='N/A')=>{const m=combined.match(re);return m?m[1].trim():def;};
  const partVal=gv(/Part:\s*([^\n<]+)/i);
  const posValRaw=gv(/Position:\s*([^\n<]+)/i);
  const posVal=posValRaw!=='N/A'?titleCasePos(posValRaw):'N/A';
  const attr4Val=gv(/(?:^|\n)- Color:\s*([^\n<]+)/i);
  const attr5Raw=gv(/Lens\s*Color:\s*([^\n<]+)/i);
  const attr6Val=gv(/Finish:\s*([^\n<]+)/i);
  const attr7Val=gv(/(?:^|\n)- Material:\s*([^\n<]+)/i);
  const attr8Val=gv(/Light\s*Source:\s*([^\n<]+)/i);
  const attr9Raw=gv(/Housing\s*Color:\s*([^\n<]+)/i);
  const attr5Val=cleanLens(attr5Raw);
  const attr9Val=cleanHousing(attr9Raw);
  hasCapa=hasCapa||/\bCAPA\b/i.test(combined);

  let bp2Val;
  if(useSkuFallback){
    bp2Val=partNumbersStr?`Part Numbers: ${partNumbersStr}`:'N/A';
  } else {
    const bp2parts=[];
    if(partslink) bp2parts.push(partslink);
    if(oem) bp2parts.push(oem);
    bp2Val=bp2parts.length?bp2parts.join(' / '):'N/A';
  }
  // BP3: extract detail lines from Product Info (exclude Brand/Part/Position/Sold/negative lines)
  const _bp3Exclude = /^(Brand:|Part:|Position:|Sold as Pair$|CAPA Certified$|DOT &|Without |No |Non-)/i;
  const _piMatch = htmlOut.match(/<b>Product Info: <\/b><br>([\s\S]*?)<br><br>/);
  let bp3Val;
  if (_piMatch) {
    const _piLines = _piMatch[1].split('<br>')
      .map(l => l.trim().replace(/^-\s*/, ''))
      .filter(l => l.startsWith('- ') || (_piMatch[1].split('<br>').indexOf(l) >= 0 && l));
    const _details = _piMatch[1].split('<br>')
      .map(l => l.trim().replace(/^-\s*/, '').trim())
      .filter(l => l && !_bp3Exclude.test(l) && l.toLowerCase() !== 'sold as pair');
    bp3Val = _details.length ? _details.join(', ') : 'Made From The Highest Quality Materials';
  } else {
    const bp3parts=[];
    if(attr4Val!=='N/A') bp3parts.push(attr4Val);
    if(attr5Val!=='N/A') bp3parts.push(attr5Val+' Lens');
    if(attr6Val!=='N/A') bp3parts.push(attr6Val);
    if(attr8Val!=='N/A') bp3parts.push(attr8Val);
    if(attr9Val!=='N/A') bp3parts.push(attr9Val+' Housing');
    bp3Val = bp3parts.length?bp3parts.join(', '):'Made From The Highest Quality Materials';
  }

  // If BP3 contains a 'Lens' detail, populate ATTR5 (Lens Color) from it
  function _extractLensFromBp3(bp3Str) {
    if (!bp3Str || bp3Str === 'Made From The Highest Quality Materials') return null;
    for (const item of bp3Str.split(',').map(s => s.trim())) {
      let m = item.match(/^Lens\s*(?:Color)?:\s*(.+)$/i);
      if (m) { const v=m[1].replace(/\s*lens\s*$/i,'').trim(); return v||null; }
      let m2 = item.match(/^(.+?)\s+Lens$/i);
      if (m2) return m2[1].trim();
      if (/\bLens\b/i.test(item)) {
        let v = item.replace(/\s*lens\b/i,'').replace(/^\s*(Color|Tint)?\s*:\s*/i,'').trim();
        return v || null;
      }
    }
    return null;
  }
  let attr5ValFinal = attr5Val;
  const _lensFromBp3 = _extractLensFromBp3(bp3Val);
  if (_lensFromBp3 && attr5Val === 'N/A') attr5ValFinal = _lensFromBp3;


  htmlOut=cleanPairDescription(htmlOut);
  return {
    pair_description: htmlOut,
    premium_aftermarket: PREMIUM_AFTERMARKET_TEXT,
    bullet_point_1: `${partVal}, ${posVal}`,
    bullet_point_2: bp2Val,
    bullet_point_3: bp3Val,
    bullet_point_4: hasCapa?'CAPA Certified':'DOT & SAE Compliant',
    bullet_point_5: '',
    attr1_part: partVal,
    attr2_placement: posVal,
    attr3_orientation: 'N/A',
    attr4_color: attr4Val,
    attr5_lens_color: attr5ValFinal,
    attr6_finish: attr6Val,
    attr7_material: attr7Val,
    attr8_light_source: attr8Val,
    attr9_housing_color: attr9Val,
    attr10_cert: hasCapa?'CAPA':'N/A',
  };
}

function patchPairDescription(xml, rowsDataRef, effectiveEnd=DATA_END) {
  return xml.replace(/<row([^>]+r=")(\d+)"([^>]*?)>([\s\S]*?)<\/row>/g,(full,p1,rnStr,p2,body)=>{
    const rn=parseInt(rnStr);
    if(rn<DATA_START||rn>effectiveEnd) return full;
    const rd=rowsDataRef[rn]||{}; if(!rd.sku) return full;

    // Read F, G, T, U from already-patched body
    const getV=col=>{
      const pos=body.indexOf(`<c r="${col}${rn}"`); if(pos===-1) return '';
      const tg=body.indexOf('>',pos); if(tg===-1||body[tg-1]==='/') return '';
      const end=body.indexOf('</c>',tg); if(end===-1) return '';
      const vm=body.slice(pos,end+4).match(/<v>([\s\S]*?)<\/v>/);
      return vm?vm[1]:'';
    };
    const fVal=getV('F'), gVal=getV('G'), tVal=getV('T'), uVal=getV('U');
    if(!fVal&&!gVal) return full;

    const hVal=getV('H'), iVal=getV('I');
    const result=processPair(fVal,gVal,tVal,uVal,'',hVal,iVal);
    let cleaned=removeCells(PAIR_DESC_TOUCH,rn,body);
    let nc='';
    for(const [col,{style,key}] of Object.entries(PAIR_DESC_COLS)){
      const val=result[key]||'';
      if(val&&val!==''){
        nc+=`<c r="${col}${rn}" s="${style}" t="str"><v>${xmlEsc(val)}</v></c>`;
      } else if(ALWAYS_VALUE_COLS.has(col)){
        nc+=`<c r="${col}${rn}" s="${style}" t="str"><v>N/A</v></c>`;
      } else {
        nc+=makeBlank(`${col}${rn}`,style);
      }
    }
    return full.replace(body, sortRowCells(cleaned+nc));
  });
}

async function run() {
  console.log('\n  ╔════════════════════════════════════════════════╗');
  console.log('  ║   USAUTO Pair Project Automation               ║');
  console.log('  ║   Pure Node.js — No Python Required            ║');
  console.log('  ╚════════════════════════════════════════════════╝\n');

  // Auto-detect input/output folders if no explicit args given
  if (!INPUT_FILE || !OUTPUT_FILE) {
    const scriptDir = __dirname;
    const inputDir  = path.join(scriptDir, 'input');
    const outputDir = path.join(scriptDir, 'output');
    if (!fs.existsSync(inputDir)) {
      console.error(`\n  ERROR: No input/ folder at: ${inputDir}`);
      console.error('         Create an input/ folder and place your .xlsx file inside.');
      console.error('         Or run: node usauto_pair.js <input.xlsx> <output.xlsx>\n');
      process.exit(1);
    }
    const xlsxFiles = fs.readdirSync(inputDir).filter(f=>f.toLowerCase().endsWith('.xlsx'));
    if (!xlsxFiles.length) { console.error(`\n  ERROR: No .xlsx file in ${inputDir}\n`); process.exit(1); }
    if (xlsxFiles.length > 1) console.log('  WARNING: Multiple .xlsx files found — using first:', xlsxFiles[0]);
    INPUT_FILE  = path.join(inputDir, xlsxFiles[0]);
    const base  = path.basename(xlsxFiles[0], '.xlsx');
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, {recursive:true});
    OUTPUT_FILE = path.join(outputDir, base + '_output.xlsx');
  }
  if (!fs.existsSync(INPUT_FILE)) { console.error(`\n  ERROR: Not found: ${INPUT_FILE}`); process.exit(1); }
  if (!fs.existsSync(INPUT_FILE)) { console.error(`  ERROR: Not found: ${INPUT_FILE}`); process.exit(1); }
  console.log(`  Input : ${INPUT_FILE}\n  Output: ${OUTPUT_FILE}\n`);

  // 1. Read target sheets
  console.log('  [1/5] Reading target sheets...');
  const sheets = {};
  await new Promise((res,rej) => {
    fs.createReadStream(INPUT_FILE).pipe(unzipper.Parse())
      .on('entry', e => {
        if ([PATH_SHEET3,PATH_SHEET4,PATH_SHEET5,PATH_CA].includes(e.path)) {
          const c=[]; e.on('data',d=>c.push(d)); e.on('end',()=>{ sheets[e.path]=Buffer.concat(c).toString('utf-8'); });
        } else e.autodrain();
      }).on('finish',res).on('error',rej);
  });

  // 2. Parse sheet3 for col A SS indices and col B URL values
  console.log('  [2/5] Parsing sheet3 and resolving SKU names...');
  const rowRe  = /<row[^>]+r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  const cellRe = /<c r="([A-Z]+)(\d+)"([^>]*?)>([\s\S]*?)<\/c>/g;
  const vRe    = /<v>([\s\S]*?)<\/v>/;

  const ssIndices = new Set();
  const rawRows   = {};

  for (const m of sheets[PATH_SHEET3].matchAll(/<row[^>]+r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const rn = parseInt(m[1]); if (rn<DATA_START||rn>DATA_END) continue;
    const body=m[2]; let ssIdx=null, imgSs=null, imgVal='';
    for (const cm of body.matchAll(/<c r="([A-Z]+)\d+"([^>]*)>([\s\S]*?)<\/c>/g)) {
      const colL=cm[1], attrs=cm[2], inner=cm[3];
      const vM=vRe.exec(inner); const val=vM?vM[1]:'';
      if (colL==='A'&&attrs.includes('t="s"')) { ssIdx=parseInt(val); ssIndices.add(ssIdx); }
      else if (colL==='B') {
        if (attrs.includes('t="s"')) { imgSs=parseInt(val); ssIndices.add(imgSs); }
        else imgVal=val;
      }
    }
    rawRows[rn]={ssIdx, imgSs, img:imgVal};
  }

  // 2b. Parse CA Data Combined
  console.log('        → Parsing CA Data Combined...');
  const caRaw={};   // {rn_col: ('ss',idx)|('val',str)}
  const caSsNeeded=new Set();
  for (const m of sheets[PATH_CA].matchAll(/<row[^>]+r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const rn=parseInt(m[1]); if(rn<2) continue;
    const body=m[2];
    for (const col of CA_COLS) {
      const pos=body.indexOf(`<c r="${col}${rn}"`);
      if(pos===-1) continue;
      // Find the > that closes the <c> opening tag (not inner tags like <f .../>)
      const tagClose=body.indexOf('>',pos);
      if(tagClose===-1) continue;
      if(body[tagClose-1]==='/') continue;  // true self-closing <c .../> — skip
      const end=body.indexOf('</c>',tagClose);
      if(end===-1) continue;
      const cell=body.slice(pos,end+4);
      const tM=cell.match(/t="([^"]+)"/);
      const vM=cell.match(/<v>([\s\S]*?)<\/v>/);
      if(vM) {
        if(tM&&tM[1]==='s'){const idx=parseInt(vM[1]);ssIndices.add(idx);caSsNeeded.add(idx);caRaw[`${rn}_${col}`]=['ss',idx];}
        else caRaw[`${rn}_${col}`]=['val',vM[1]];
      }
    }
  }

  // Get SS uniqueCount from header
  let inputSsCount=0, ssMap={};
  await new Promise((res,rej) => {
    let headerRead=false, buf=Buffer.alloc(0);
    fs.createReadStream(INPUT_FILE).pipe(unzipper.Parse())
      .on('entry', e => {
        if (e.path===PATH_SS) {
          e.on('data',chunk=>{ if(!headerRead){buf=Buffer.concat([buf,chunk]); if(buf.length>=500){
            const h=buf.slice(0,500).toString('utf-8');
            const uc=h.match(/uniqueCount="(\d+)"/); if(uc) inputSsCount=parseInt(uc[1]);
            headerRead=true;
          }}});
          // SAX parse to resolve needed indices
          const parser=new SaxesParser(); let siCount=-1, inT=false, capturing=false, curText='', stopped=false;
          const maxNeeded=ssIndices.size>0?Math.max(...ssIndices):-1;
          parser.on('opentag',node=>{ if(node.name==='si'){siCount++;curText='';capturing=ssIndices.has(siCount);} if(node.name==='t') inT=true; });
          parser.on('text',txt=>{ if(capturing&&inT) curText+=txt; });
          parser.on('closetag',node=>{ if(node.name==='t') inT=false; if(node.name==='si'){ if(capturing) ssMap[siCount]=curText; capturing=false; if(siCount>=maxNeeded&&Object.keys(ssMap).length>=ssIndices.size) stopped=true; }});
          e.on('data',chunk=>{ if(!stopped) parser.write(chunk.toString('utf-8')); });
          e.on('end',()=>res());
        } else e.autodrain();
      }).on('error',rej);
  });

  console.log(`        → Input SS uniqueCount: ${inputSsCount}`);

  // 3. Build rowsData
  // Auto-detect effective end from last row with actual SKU SS ref
  const detectedEnd = Math.max(...Object.keys(rawRows).filter(r => rawRows[r].ssIdx != null).map(Number), DATA_START);
  const EFFECTIVE_END = detectedEnd;

  console.log(`        → Auto-detected DATA_END = ${EFFECTIVE_END}`);


  const rowsData={};
  for (let rn=DATA_START;rn<=EFFECTIVE_END;rn++) {
    const rd=rawRows[rn]||{}; const ssIdx=rd.ssIdx;
    const sku=ssIdx!=null?ssMap[ssIdx]:null;
    // Use SS-resolved image string if available (new input format), else cached formula value
    const img = rd.imgSs!=null ? (ssMap[rd.imgSs]||rd.img||'') : (rd.img||'');
    const tokens=img.split(',').map(s=>stripPrefix(s)).filter(s=>s.length>0);
    rowsData[rn]={sku, ssIdx, tokens};
  }
  const dataCount=Object.values(rowsData).filter(r=>r.sku).length;
  const maxTokens=Math.max(...Object.values(rowsData).map(r=>r.tokens.length),0);
  console.log(`        → ${dataCount} rows with SKU | max ${maxTokens} URLs/row`);

  // 4. Build URL→SS index map
  const urlToSs={}, newSsUrls=[];
  for (let rn=DATA_START;rn<=EFFECTIVE_END;rn++)
    for (const url of rowsData[rn].tokens)
      if (url&&!urlToSs[url]) { urlToSs[url]=inputSsCount+newSsUrls.length; newSsUrls.push(url); }
  // 3b. Build CA lookup and pre-evaluate Photowork values
  const caLookup={};  // clxSkuDriver -> {col: value}
  for (const key of Object.keys(caRaw)) {
    const parts=key.split('_'), rn=parseInt(parts[0]), col=parts.slice(1).join('_');
    const keyEntry=caRaw[`${rn}_C`];
    if(!keyEntry) continue;
    const skuKey=keyEntry[0]==='ss'?ssMap[keyEntry[1]]:keyEntry[1];
    if(!skuKey) continue;
    if(!caLookup[skuKey]) caLookup[skuKey]={};
    const [kind,val]=caRaw[key];
    const resolved=kind==='ss'?ssMap[val]:val;
    if(resolved) caLookup[skuKey][col]=resolved;
  }
  console.log(`        → CA lookup: ${Object.keys(caLookup).length} SKU keys`);

  // Evaluation helpers
  const PW_TO_CA = {C:'S', D:'B', E:'R', F:'AH', G:'AI', H:'J', I:'Z'};
  function pwEval(rn, pwCol) {
    const sku = rowsData[rn]?.sku; if(!sku) return null;
    const caCol = PW_TO_CA[pwCol]; if(!caCol) return null;
    return caLookup[sku]?.[caCol] || null;
  }
  function descEval(rn, descCol) {
    const sku = rowsData[rn]?.sku; if(!sku) return null;
    const PW_BACKED = {C:'C', D:'H', E:'I', H:'D', T:'F', U:'G'};
    const CA_BACKED = {F:'H',G:'X',I:'R',J:'K',K:'L',L:'AA',M:'AB',N:'N',O:'O',P:'P',Q:'AD',R:'AE',S:'AF'};
    if (PW_BACKED[descCol]) return pwEval(rn, PW_BACKED[descCol]);
    if (CA_BACKED[descCol]) return caLookup[sku]?.[CA_BACKED[descCol]] || null;
    return null;
  }


  const pwVals={};  // rn -> {col: value}
  for (let rn=DATA_START;rn<=EFFECTIVE_END;rn++) {
    const sku=rowsData[rn]?.sku;
    const ca=sku&&caLookup[sku]?caLookup[sku]:{};
    pwVals[rn]={};
    for (const [pwCol,caCol] of Object.entries(PW_EVAL_MAP)) {
      pwVals[rn][pwCol]=ca[caCol]||null;
    }
  }

  console.log(`        → ${newSsUrls.length} unique URLs → SS indices ${inputSsCount}–${inputSsCount+newSsUrls.length-1}`);

  function getRowUrlSs(rn) { return (rowsData[rn]||{tokens:[]}).tokens.map(u=>urlToSs[u]).filter(x=>x!=null); }

  // 5. Patch sheets
  console.log('  [3/5] Patching sheets...');

  // Sheet3
  function patchSheet3(xml) {
    return xml.replace(/<row([^>]+r=")(\d+)"([^>]*?)>([\s\S]*?)<\/row>/g, (full,p1,rnStr,p2,body)=>{
      const rn=parseInt(rnStr);
      if (rn===1) {
        const cleaned=body.replace(/<c r="([A-Z]+)1"(?:[^>]*\/>[^<]*|[^>]*>[\s\S]*?<\/c>)/g,
          m=>{ const l=m.match(/^<c r="([A-Z]+)/)[1]; return colNum(l)>=COL_PASTE_VAL?'':m; });
        return full.replace(/spans="[^"]*"/,'spans="1:8"').replace(body, cleaned+makeSsCell('H1',SS_PASTE_HDR,'61'));
      }
      if (rn>=DATA_START&&rn<=EFFECTIVE_END) {
        const urlSs=getRowUrlSs(rn);
        const cleaned=body.replace(new RegExp(`<c r="([A-Z]+)${rn}"(?:[^>]*\\/>[^<]*|[^>]*>[\\s\\S]*?<\\/c>)`,'g'),
          m=>{ const l=m.match(/^<c r="([A-Z]+)/)[1]; return colNum(l)>=COL_PASTE_VAL?'':m; });
        if (!urlSs.length) return full.replace(body, sortRowCells(cleaned));
        let nc='';
        urlSs.forEach((ssIdx,t)=>{ nc+=makeSsCell(colLetter(COL_PASTE_VAL+t)+rn, ssIdx, t===0?hStyle(rn):null); });
        const maxC=COL_PASTE_VAL+urlSs.length-1;
        return full.replace(/spans="[^"]*"/,`spans="1:${maxC}"`).replace(body, sortRowCells(cleaned+nc));
      }
      return full;
    });
  }

  // Sheet4
  function patchPhotowork(xml, effectiveEnd=DATA_END) {
    return xml.replace(/<row([^>]+r=")(\d+)"([^>]*?)>([\s\S]*?)<\/row>/g,(full,p1,rnStr,p2,body)=>{
      const rn=parseInt(rnStr);
      if (rn<DATA_START||rn>effectiveEnd) return full;
      const cleaned=removeCells(PHOTO_TOUCH,rn,body);
      const rd=rowsData[rn]||{sku:null,ssIdx:null,tokens:[]}; const urlSs=getRowUrlSs(rn);
      let nc='';
      if (rd.ssIdx!=null) nc+=`<c r="B${rn}" s="98" t="s"><v>${rd.ssIdx}</v></c>`;
      const pv=pwVals[rn]||{};
      const CONCAT_SUFFIX={K:'_master_SET_v2.JPG',M:'_02_SET_v2.JPG',O:'_03_SET_v2.JPG',
                            Q:'_04_SET_v2.JPG',S:'_05_SET_v2.JPG',U:'_06_SET_v2.JPG',
                            W:'_07_SET_v2.JPG',Y:'_08_SET_v2.JPG'};
      for (const {col,style,tmpl} of PHOTO_FORMULA_COLS) {
        let cached='#N/A', ctype='e';
        if (['C','D','H','I','F','G'].includes(col)) {
          const rawV=pv[col];
          const v=(col==='F'||col==='G') ? (cleanCombined(rawV)||'N/A') : rawV;
          if(v&&v!=='#N/A'){cached=v;ctype='str';}
          if((col==='F'||col==='G')&&cached==='#N/A'){cached='N/A';ctype='str';}
        } else if (CONCAT_SUFFIX[col]) {
          const compSku=pv['D'];
          if(compSku){cached=compSku+CONCAT_SUFFIX[col];ctype='str';}
        }
        nc+=makeFormulaCell(col,style,tmpl(rn),rn,ctype,cached);
      }
      nc+=makeBlank(`J${rn}`,'13');
      nc+=makeBlank(`AA${rn}`,AA_STYLE_MAP[rn]||'76');
      nc+=makeBlank(`AB${rn}`,'80');
      const tokens=(rowsData[rn]||{tokens:[]}).tokens;
      for (const col of ['AD',...HLOOKUP_AE_AL]) {
        const formula=`HLOOKUP("*"&${col}$1&"*",'Separate Image URLs Here'!H${rn}:P${rn},1,FALSE)`;
        const header=HLOOKUP_HEADER[col];
        const matchUrl=tokens.find(u=>u.toLowerCase().includes(header.toLowerCase()))||null;
        if (col==='AD'&&rn===DATA_START&&urlSs.length) {
          // Row 2 AD: SS ref (no formula)
          nc+=`<c r="AD${rn}" s="3" t="s"><v>${urlSs[0]}</v></c>`;
        } else if (matchUrl) {
          nc+=makeFormulaCell(col,'3',formula,rn,'str',matchUrl);
        } else {
          nc+=makeFormulaCell(col,'3',formula,rn,'e','#N/A');
        }
      }
      return full.replace(body, sortRowCells(cleaned+nc));
    });
  }

  // Sheet5
  function patchDescription(xml, effectiveEnd=DATA_END) {
    return xml.replace(/<row([^>]+r=")(\d+)"([^>]*?)>([\s\S]*?)<\/row>/g,(full,p1,rnStr,p2,body)=>{
      const rn=parseInt(rnStr);
      if (rn<DATA_START||rn>effectiveEnd) return full;
      const cleaned=removeCells(DESC_TOUCH,rn,body);
      const rd=rowsData[rn]||{sku:null,ssIdx:null,tokens:[]};
      const sku=rd?.sku;
      const ca=(sku&&caLookup[sku])?caLookup[sku]:{};
      const pv2=pwVals[rn]||{};
      let nc=makeBlank(`A${rn}`,'14');
      if (rd.ssIdx!=null) nc+=`<c r="B${rn}" s="98" t="s"><v>${rd.ssIdx}</v></c>`;
      for (const {col,style,tmpl} of DESC_FORMULA_COLS) {
        const {src,ref}=DESC_SRC[col]||{};
        const rawVal=src==='pw'?pv2[ref]:ca[ref];
        const cleanedVal=(col==='T'||col==='U')?cleanCombined(rawVal):rawVal;
        const cached=(col==='T'||col==='U')?(cleanedVal||'N/A'):((cleanedVal&&cleanedVal!=='#N/A')?cleanedVal:'#N/A');
        const ctype=(col==='T'||col==='U')?'str':((cleanedVal&&cleanedVal!=='#N/A')?'str':'e');
        nc+=makeFormulaCell(col,style,tmpl(rn),rn,ctype,cached);
      }
      return full.replace(body, sortRowCells(cleaned+nc));
    });
  }

  const patched3=patchSheet3(sheets[PATH_SHEET3]);
  const patched4=patchPhotowork(sheets[PATH_SHEET4], EFFECTIVE_END);
  const patched5raw2=patchDescription(sheets[PATH_SHEET5]);
  const patched5=patchPairDescription(patched5raw2,rowsData,EFFECTIVE_END);
  console.log('        → All 3 sheets patched ✓');

  // 6. Rebuild zip (stream SS patch)
  console.log('  [4/5] Writing output (streaming SS)...');
  const CLOSE_TAG=Buffer.from('</sst>');
  const newSiXml=newSsUrls.map(u=>`<si><t>${xmlEsc(u)}</t></si>`).join('');
  const newTotal=inputSsCount+newSsUrls.length;
  const PATCHED_SMALL={
    [PATH_SHEET3]:Buffer.from(patched3,'utf-8'),
    [PATH_SHEET4]:Buffer.from(patched4,'utf-8'),
    [PATH_SHEET5]:Buffer.from(patched5,'utf-8'),
  };

  const zip=new JSZip();
  await new Promise((res,rej)=>{
    fs.createReadStream(INPUT_FILE).pipe(unzipper.Parse())
      .on('entry',e=>{
        const fn=e.path;
        if (PATCHED_SMALL[fn]) { zip.file(fn,PATCHED_SMALL[fn]); e.autodrain(); }
        else if (fn==='xl/calcChain.xml') { e.autodrain(); }
        else if (fn===PATH_SS) {
          const chunks=[]; let headerPatched=false, buf=Buffer.alloc(0);
          e.on('data',chunk=>{
            buf=Buffer.concat([buf,chunk]);
            if (!headerPatched&&buf.length>=500) {
              let h=buf.slice(0,500).toString('utf-8');
              h=h.replace(/uniqueCount="\d+"/,`uniqueCount="${newTotal}"`).replace(/count="\d+"/,`count="${newTotal}"`);
              buf=Buffer.concat([Buffer.from(h,'utf-8'),buf.slice(500)]);
              headerPatched=true;
            }
            const idx=buf.indexOf(CLOSE_TAG);
            if (idx!==-1) {
              chunks.push(buf.slice(0,idx));
              chunks.push(Buffer.from(newSiXml+`</sst>`,'utf-8'));
              buf=Buffer.alloc(0);
            } else {
              const keep=CLOSE_TAG.length-1;
              if(buf.length>keep){chunks.push(buf.slice(0,buf.length-keep));buf=buf.slice(buf.length-keep);}
            }
          });
          e.on('end',()=>{ if(buf.length>0)chunks.push(buf); zip.file(fn,Buffer.concat(chunks)); });
        }
        else { const c=[]; e.on('data',d=>c.push(d)); e.on('end',()=>zip.file(fn,Buffer.concat(c))); }
      }).on('finish',res).on('error',rej);
  });

  await new Promise((res,rej)=>{
    const out=fs.createWriteStream(OUTPUT_FILE);
    zip.generateNodeStream({type:'nodebuffer',streamFiles:true}).pipe(out).on('finish',res).on('error',rej);
  });

  console.log('  [5/5] Done.\n');
  console.log(`  Output: ${OUTPUT_FILE}\n`);
}

run().catch(err=>{ console.error('\n  ERROR:',err.message); process.exit(1); });
