/* ECO-KAKAO · Datos privados (Procesadores + Importadores)
   - Los datos con licencia (Euromonitor, Penta) se guardan CIFRADOS en private.enc.json.
   - Solo se leen con la contraseña del equipo. La contraseña nunca está en el código.
   - "Subir Excel" lee el archivo en el navegador, actualiza, cifra y publica en GitHub. */
(function(){
const REPO = "rivadisa/ecokakao-intel", FILE = "private.enc.json";
const SHEETJS = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
const EURO_NAMES = {"Turkey":"TUR","Netherlands":"NLD","Germany":"DEU","Belgium":"BEL","France":"FRA","Italy":"ITA","Spain":"ESP","United Kingdom":"GBR","Switzerland":"CHE","Poland":"POL","Estonia":"EST","Austria":"AUT","Czech Republic":"CZE","Sweden":"SWE","Denmark":"DNK","Finland":"FIN","Ireland":"IRL","Portugal":"PRT","Greece":"GRC","Ukraine":"UKR","Russia":"RUS","USA":"USA","Canada":"CAN","Mexico":"MEX","Brazil":"BRA","Argentina":"ARG","Colombia":"COL","Chile":"CHL","Japan":"JPN","China":"CHN","South Korea":"KOR","Malaysia":"MYS","Indonesia":"IDN","Singapore":"SGP","India":"IND","Australia":"AUS","Saudi Arabia":"SAU","United Arab Emirates":"ARE","Egypt":"EGY","South Africa":"ZAF"};
let PRIV = null;            // datos descifrados {procesadores, importadores, volumen}
let PW = null;              // contraseña de la sesión
let FILE_STATE = "unknown"; // unknown | missing | present
let MSG = "";
const ss = {get:k=>{try{return sessionStorage.getItem(k)}catch(e){return null}}, set:(k,v)=>{try{sessionStorage.setItem(k,v)}catch(e){}}};
const ls = {get:k=>{try{return localStorage.getItem(k)}catch(e){return null}}, set:(k,v)=>{try{localStorage.setItem(k,v)}catch(e){}}};
const fmt1 = n => (Math.round(n*10)/10).toLocaleString("es-EC",{minimumFractionDigits:1,maximumFractionDigits:1});
const fmtInt = n => Math.round(n).toLocaleString("es-EC");

/* ---------- cifrado (AES-GCM + PBKDF2) ---------- */
const enc = new TextEncoder(), dec = new TextDecoder();
const b64 = u8 => { let s=""; for(let i=0;i<u8.length;i+=32768) s+=String.fromCharCode.apply(null,u8.subarray(i,i+32768)); return btoa(s); };
const unb64 = s => Uint8Array.from(atob(s), c=>c.charCodeAt(0));
async function keyFrom(pw, salt){ const base=await crypto.subtle.importKey("raw",enc.encode(pw),"PBKDF2",false,["deriveKey"]);
  return crypto.subtle.deriveKey({name:"PBKDF2",salt,iterations:310000,hash:"SHA-256"},base,{name:"AES-GCM",length:256},false,["encrypt","decrypt"]); }
async function encryptObj(obj, pw){ const salt=crypto.getRandomValues(new Uint8Array(16)), iv=crypto.getRandomValues(new Uint8Array(12));
  const ct=new Uint8Array(await crypto.subtle.encrypt({name:"AES-GCM",iv},await keyFrom(pw,salt),enc.encode(JSON.stringify(obj))));
  return {v:1, note:"Datos privados de ECO-KAKAO, cifrados. No contienen nada legible sin la contraseña del equipo.", salt:b64(salt), iv:b64(iv), ct:b64(ct)}; }
async function decryptObj(file, pw){ const pt=await crypto.subtle.decrypt({name:"AES-GCM",iv:unb64(file.iv)},await keyFrom(pw,unb64(file.salt)),unb64(file.ct)); return JSON.parse(dec.decode(pt)); }

async function fetchFile(){ try{ const r=await fetch(FILE+"?t="+Date.now(),{cache:"no-store"}); if(!r.ok){FILE_STATE="missing"; return null;} FILE_STATE="present"; return await r.json(); }catch(e){ FILE_STATE="missing"; return null; } }
async function unlock(pw){ const f=await fetchFile(); if(!f) return false; PRIV=await decryptObj(f,pw); PW=pw; ss.set("ek_pw",pw); return true; }
async function tryAutoUnlock(){ const pw=ss.get("ek_pw"); if(pw){ try{ await unlock(pw); }catch(e){} } else { await fetchFile(); } }

/* ---------- lectura de Excel ---------- */
function loadSheetJS(){ return window.XLSX ? Promise.resolve() : new Promise((res,rej)=>{ const s=document.createElement("script"); s.src=SHEETJS; s.onload=res; s.onerror=()=>rej(new Error("No se pudo cargar el lector de Excel. Revisa tu conexión.")); document.head.appendChild(s); }); }
async function readBook(file){ await loadSheetJS(); const buf=await file.arrayBuffer(); return XLSX.read(buf,{type:"array"}); }
const rowsOf = ws => XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
const num = v => { if(v==null||v==="-"||v==="") return null; if(typeof v==="number") return v; const s=String(v).replace(/\s/g,"").replace(/[^\d,.\-]/g,""); if(!s) return null;
  const t = s.includes(",")&&s.includes(".") ? (s.lastIndexOf(",")>s.lastIndexOf(".") ? s.replace(/\./g,"").replace(",",".") : s.replace(/,/g,"")) : s.replace(",","."); const n=parseFloat(t); return isNaN(n)?null:n; };

function parseEuromonitor(wb){
  let sheet=null, hdr=-1, rows=null;
  for(const name of wb.SheetNames){ const r=rowsOf(wb.Sheets[name]); const i=r.findIndex(x=>x&&String(x[0]||"").trim()==="Geography"&&x.some(c=>/Company/i.test(String(c||"")))); if(i>=0){sheet=name;hdr=i;rows=r;break;} }
  if(!rows) throw new Error("No encuentro la tabla de cuotas (columnas Geography / Company Name). ¿Es el Excel de Euromonitor?");
  const H=rows[hdr].map(c=>String(c||"").trim()), gi=H.indexOf("Geography"), ci=H.findIndex(h=>/Company/i.test(h)), cati=H.indexOf("Category"), dti=H.findIndex(h=>/Data Type/i.test(h));
  const yearCols=H.map((h,i)=>/^(19|20)\d\d$/.test(h)?i:-1).filter(i=>i>=0); const years=yearCols.map(i=>H[i]);
  const out={}; let category="", dataType="";
  for(const r of rows.slice(hdr+1)){ if(!r||!r[gi]||!r[ci]) continue; const iso=EURO_NAMES[String(r[gi]).trim()]; if(!iso) continue;
    const name=String(r[ci]).trim(); category=category||r[cati]; dataType=dataType||r[dti];
    const sh={}; yearCols.forEach((c,k)=>{ const v=num(r[c]); if(v!=null) sh[years[k]]=v; });
    (out[iso]=out[iso]||{rows:[]}).rows.push({name, shares:sh}); }
  let volumen={};
  for(const name of wb.SheetNames){ const r=rowsOf(wb.Sheets[name]); const i=r.findIndex(x=>x&&String(x[0]||"").trim()==="Geography"&&x.some(c=>/Unit/i.test(String(c||"")))); if(i<0) continue;
    const H2=r[i].map(c=>String(c||"").trim()), ui=H2.indexOf("Unit"), yc=H2.map((h,k)=>/^(19|20)\d\d$/.test(h)?k:-1).filter(k=>k>=0);
    for(const x of r.slice(i+1)){ if(!x||!x[0]) continue; const iso=EURO_NAMES[String(x[0]).trim()]; if(!iso) continue; volumen[iso]={unit:x[ui], years:Object.fromEntries(yc.map(k=>[H2[k],num(x[k])]).filter(e=>e[1]!=null))}; } }
  const exported=(rows.find(x=>x&&/Date Exported/i.test(String(x[0]||"")))||[""])[0];
  return {markets:out, years, category, dataType, exported:String(exported||"").replace(/Date Exported \(GMT\):\s*/i,""), volumen};
}

const COLS = {
  name:/comprador|buyer|importador|importer|consignee|consignatario|raz[oó]n social|^empresa|company/i,
  usd:/u\$s|us\$|usd|fob|cif|valor|value|monto/i,
  exporter:/^exportador|exporter|shipper|embarcador/i,
  dest:/destino|destination/i,
  kg:/peso neto|net weight|peso|kg|kilo|cantidad|quantity|toneladas|tons/i,
  date:/fecha|date|per[ií]odo|^mes|month/i,
  origin:/pa[ií]s de origen|^origen$|^origin$|^country of origin|pa[ií]s de procedencia|^procedencia$/i,
  hs:/posici[oó]n|partida|ncm|hs|arancel|nandina|tariff/i
};
const normC = t => String(t||"").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^a-z ]/g," ").replace(/\s+/g," ").trim();
const DEST_ALIAS = {"republica checa":"CZE","chequia":"CZE","czech republic":"CZE","holanda":"NLD","paises bajos":"NLD","netherlands":"NLD","estados unidos de america":"USA","estados unidos":"USA","eeuu":"USA","united states":"USA","usa":"USA","corea del sur":"KOR","corea republica de":"KOR","corea sur":"KOR","republica de corea":"KOR","south korea":"KOR","korea":"KOR","federacion de rusia":"RUS","rusia":"RUS","russia":"RUS","turquia":"TUR","turkiye":"TUR","turkey":"TUR","reino unido":"GBR","inglaterra":"GBR","gran bretana":"GBR","united kingdom":"GBR","emiratos arabes unidos":"ARE","emiratos arabes":"ARE","arabia saudita":"SAU","arabia saudi":"SAU","sudafrica":"ZAF","africa del sur":"ZAF","singapur":"SGP","japon":"JPN","china":"CHN","china republica popular":"CHN","hong kong":"HKG","hong kong china":"HKG","rae de hong kong china":"HKG","taiwan":"TWN"};
function destIso(name){ const n=normC(name); if(!n) return null; if(DEST_ALIAS[n]) return DEST_ALIAS[n];
  const m=MARKETS.find(x=>normC(x.name)===n); if(m) return m.iso3; const e=Object.keys(EURO_NAMES).find(k=>normC(k)===n); return e?EURO_NAMES[e]:null; }
function findTable(wb){
  for(const name of wb.SheetNames){ const r=rowsOf(wb.Sheets[name]);
    for(let i=0;i<Math.min(r.length,40);i++){ const H=(r[i]||[]).map(c=>String(c||"").trim()); const idx={};
      for(const [k,re] of Object.entries(COLS)){ const j=H.findIndex((h,jj)=>h&&re.test(h)&&!Object.values(idx).includes(jj)); if(j>=0) idx[k]=j; }
      if(idx.name!=null&&(idx.usd!=null||idx.kg!=null)) return {name,hdr:i,H,idx,rows:r}; } }
  throw new Error("No encuentro una columna de importador o comprador y otra de valor o peso. Revisa que el Excel tenga encabezados como \"Comprador\" y \"U$S FOB\" o \"Peso neto\".");
}
function aggregate(rows, idx, H){
  const agg={}; let tot=0, totKg=0, n=0, minD=null, maxD=null;
  for(const r of rows){ const nm=String(r[idx.name]||"").trim(); if(!nm||/^total/i.test(nm)) continue;
    const usd=idx.usd!=null?num(r[idx.usd])||0:0, kg=idx.kg!=null?num(r[idx.kg])||0:0; if(!usd&&!kg) continue;
    const words=nm.toUpperCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^A-Z0-9 ]/g," ").split(/\s+/).filter(Boolean);
    const key=words.slice(0,2).join(" ")||nm.toUpperCase(); const a=agg[key]=agg[key]||{usd:0,kg:0,ops:0,origins:{},hs:{},exporters:{},names:{},dests:{}};
    a.names[nm]=(a.names[nm]||0)+1;
    if(idx.exporter!=null){ const ex=String(r[idx.exporter]||"").trim(); if(ex) a.exporters[ex]=(a.exporters[ex]||0)+(usd||kg); }
    if(idx.dest!=null){ const dd=String(r[idx.dest]||"").trim(); if(dd) a.dests[dd]=(a.dests[dd]||0)+(usd||kg); }
    a.usd+=usd; a.kg+=kg; a.ops++; tot+=usd; totKg+=kg; n++;
    if(idx.origin!=null){ const o=String(r[idx.origin]||"").trim(); if(o) a.origins[o]=(a.origins[o]||0)+(usd||kg); }
    if(idx.hs!=null){ const h=String(r[idx.hs]||"").replace(/\D/g,"").slice(0,4); if(h) a.hs[h]=1; }
    if(idx.date!=null){ let d=r[idx.date]; if(typeof d==="number"&&d>20000&&d<80000){ d=new Date(Math.round((d-25569)*864e5)).toISOString().slice(0,10); } d=d?String(d).slice(0,10):null; if(d){ if(!minD||d<minD)minD=d; if(!maxD||d>maxD)maxD=d; } } }
  const by = tot>0 ? "usd" : "kg", T = tot>0 ? tot : totKg;
  const list=Object.values(agg).sort((a,b)=>b[by]-a[by]).slice(0,15).map(a=>{ const ot=Object.values(a.origins).reduce((x,y)=>x+y,0)||1;
    const et=Object.values(a.exporters).reduce((x,y)=>x+y,0)||1; const nmBest=Object.entries(a.names).sort((x,y)=>y[1]-x[1]||y[0].length-x[0].length)[0][0];
    return {name:nmBest, usd:Math.round(a.usd), kg:Math.round(a.kg), ops:a.ops, share:T?a[by]/T*100:0, hs:Object.keys(a.hs).sort(),
      dests:(()=>{const dt=Object.values(a.dests).reduce((x,y)=>x+y,0)||1; return Object.entries(a.dests).sort((x,y)=>y[1]-x[1]).slice(0,3).map(([o,v])=>({name:o,share:v/dt*100}));})(),
      exporters:Object.entries(a.exporters).sort((x,y)=>y[1]-x[1]).slice(0,3).map(([o,v])=>({name:o,share:v/et*100})),
      origins:Object.entries(a.origins).sort((x,y)=>y[1]-x[1]).slice(0,3).map(([o,v])=>({name:o,share:v/ot*100}))}; });
  return {rows:list, by, mode:idx.exporter!=null?"ec_export":"import", total:T, totalUsd:Math.round(tot), totalKg:Math.round(totKg), lines:n, from:minD, to:maxD, columns:Object.fromEntries(Object.entries(idx).map(([k,j])=>[k,H[j]])), importers:Object.keys(agg).length};
}
const HQ_EXTRA={USA:["USA","U.S.A","UNITED STATES"],GBR:["UNITED KINGDOM","ENGLAND"],NLD:["NETHERLANDS","HOLLAND"],KOR:["KOREA"],ARE:["UAE","DUBAI"],HKG:["HONG KONG"],LVA:["LATVIA","RIGA"],PHL:["PHILIPPINES","MANILA"],SGP:["SINGAPORE"],RUS:["RUSSIA","RUSSIAN FEDERATION"],TUR:["TURKEY","TURKIYE"],CZE:["CZECH"]};
function hqPatterns(){ return MARKETS.map(m=>{ const names=[...new Set([enName(m.iso2),esName(m.iso2),m.name,...(HQ_EXTRA[m.iso3]||[])].map(x=>normC(x).toUpperCase()).filter(x=>x&&x.length>2))];
  return {iso:m.iso3, re:new RegExp("\\b("+names.map(x=>x.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")).join("|")+")\\b","g")}; }); }
function hqIso(text,pats){ const T=normC(text).toUpperCase(); let best=null,pos=-1;
  for(const p of pats){ p.re.lastIndex=0; let mm; while((mm=p.re.exec(T))){ if(mm.index>pos){pos=mm.index;best=p.iso;} } } return best; }
/* Devuelve {byIso:{ISO3:resultado}, unmatched:{destino:filas}, single:true|false} */
function parsePenta(wb, fallbackIso){
  const t=findTable(wb), data=t.rows.slice(t.hdr+1).filter(r=>r&&r.some(c=>c!=null&&c!==""));
  if(t.idx.dest==null) return {byIso:{[fallbackIso]:aggregate(data,t.idx,t.H)}, unmatched:{}, single:true};
  const groups={}, unmatched={};
  for(const r of data){ const d=String(r[t.idx.dest]||"").trim(); const iso=destIso(d); if(iso&&BY3[iso]) (groups[iso]=groups[iso]||[]).push(r); else if(d) unmatched[d]=(unmatched[d]||0)+1; }
  const byIso={}; for(const [iso,rs] of Object.entries(groups)) byIso[iso]=aggregate(rs,t.idx,t.H);
  /* Sede del comprador según su dirección (para quien compra desde un país y recibe en otro) */
  const ac=[t.idx.name-1,t.idx.name+1].filter(j=>j>=0&&/direcci|address|domicilio/i.test(t.H[j]||""));
  const hq={};
  if(ac.length){ const pats=hqPatterns();
    for(const r of data){ const iso=hqIso(ac.map(j=>r[j]).join(" | "),pats); const di=destIso(String(r[t.idx.dest]||"")); if(iso&&iso!==di) (hq[iso]=hq[iso]||[]).push(r); } }
  const hqAgg={}; for(const [iso,rs] of Object.entries(hq)) hqAgg[iso]=aggregate(rs,t.idx,t.H);
  return {byIso, hq:hqAgg, unmatched, single:Object.keys(byIso).length<=1};
}

/* ---------- publicar en GitHub ---------- */
async function publish(obj){
  const encd=await encryptObj(obj,PW); const content=JSON.stringify(encd);
  const token=ls.get("ek_gh_token");
  if(!token) return {ok:false, encd, reason:"sin token"};
  const api=`https://api.github.com/repos/${REPO}/contents/${FILE}`, h={Authorization:"Bearer "+token, Accept:"application/vnd.github+json"};
  let sha; const g=await fetch(api,{headers:h}); if(g.ok) sha=(await g.json()).sha; else if(g.status!==404) return {ok:false, encd, reason:`GitHub respondió ${g.status} al leer el archivo. Revisa el token.`};
  const p=await fetch(api,{method:"PUT",headers:h,body:JSON.stringify({message:"Datos privados actualizados desde la app",content:b64(enc.encode(content)),sha})});
  if(!p.ok) return {ok:false, encd, reason:`GitHub respondió ${p.status} al guardar. Revisa que el token tenga permiso de escritura (Contents) en ${REPO}.`};
  FILE_STATE="present"; return {ok:true};
}
function downloadFile(encd){ const a=document.createElement("a"); a.href=URL.createObjectURL(new Blob([JSON.stringify(encd)],{type:"application/json"})); a.download=FILE; document.body.appendChild(a); a.click(); a.remove(); }

/* ---------- vistas ---------- */
const KIND = {procesadores:{t:"Procesadores", lead:"Principales empresas del mercado de chocolate por cuota de mercado (% del valor de venta minorista).", src:"Euromonitor International", file:"el Excel de Euromonitor (cuotas de mercado por país)"},
              importadores:{t:"Importadores", lead:"Principales importadores del país en HS 1801–1806, según registros de aduana.", src:"Penta-Transaction", file:"el Excel exportado de Penta-Transaction (importaciones del país)"}};
function lockView(kind){
  const first = FILE_STATE==="missing";
  return `<div class="notice info"><b>Datos privados del equipo.</b> ${first?"Todavía no hay datos cargados. Crea la contraseña del equipo y sube el primer Excel.":"Escribe la contraseña del equipo para ver las cuotas e importadores."} La licencia de estos datos no permite publicarlos, por eso se guardan cifrados.</div>
  <form id="pv-unlock" class="pv-form">
    <label for="pv-pw" class="src"><b>${first?"Nueva contraseña del equipo":"Contraseña del equipo"}</b></label>
    <input id="pv-pw" type="password" autocomplete="current-password" required minlength="${first?10:1}" class="pv-in">
    ${first?`<label for="pv-pw2" class="src"><b>Repite la contraseña</b> (mínimo 10 caracteres; compártela solo con el equipo)</label><input id="pv-pw2" type="password" autocomplete="new-password" required class="pv-in">`:""}
    <button class="cta" type="submit">${first?"Crear contraseña y continuar":"Entrar"} <span>→</span></button>
    ${MSG?`<p class="src" style="color:var(--cacao-red)">${esc(MSG)}</p>`:""}
  </form>`;
}
function uploadBox(kind){
  const tok=!!ls.get("ek_gh_token");
  return `<details class="pv-up" ${(!PRIV||!PRIV[kind]||!Object.keys(PRIV[kind]).length)?"open":""}><summary><b>Subir Excel y actualizar datos</b></summary>
    <form id="pv-upload" class="pv-form" data-kind="${kind}">
      <p class="src">Sube ${KIND[kind].file}.${kind==="procesadores"?" Un solo archivo actualiza los 40 mercados.":""}</p>
      ${kind==="importadores"?`<p class="src"><b>Recomendado:</b> en Penta, Ecuador → Exportación → partidas 1801–1806 → todos los destinos. Un solo Excel actualiza todos los mercados; la app lo reparte por "País de Destino".</p><label for="pv-country" class="src"><b>País del archivo</b> (solo si el Excel no trae columna de destino)</label><select id="pv-country" class="pv-in">${MARKETS.map(m=>`<option value="${m.iso3}" ${m.iso3===state.market.iso3?"selected":""}>${esc(m.name)}</option>`).join("")}</select>`:""}
      <label for="pv-file" class="src"><b>Archivo Excel</b> (.xls, .xlsx o .csv)</label>
      <input id="pv-file" type="file" accept=".xls,.xlsx,.csv" required class="pv-in">
      <label for="pv-token" class="src"><b>Token de GitHub</b> ${tok?"(ya guardado en este navegador; déjalo vacío)":"(solo la primera vez; queda guardado en este navegador)"}</label>
      <input id="pv-token" type="password" autocomplete="off" placeholder="${tok?"••••••••••":"github_pat_…"}" class="pv-in">
      <button class="cta" type="submit">Subir y actualizar <span>↑</span></button>
      <p class="src" id="pv-msg">${MSG?esc(MSG):""}</p>
    </form></details>`;
}
function bar(p,max,col){ return `<div class="bar"><span style="width:${Math.max(1,p/max*100)}%;background:${col}"></span></div>`; }
function viewRanking(){
  const kind=state.rkind, m=state.market, K=KIND[kind];
  const head=`<div><div class="eyebrow">${esc(m.name)} · ${K.t}</div><h2>${K.t}</h2><p class="lead">${K.lead}</p></div>
    <div class="filters">${Object.keys(KIND).map(k=>`<button data-rkind="${k}" class="${k===kind?"on":""}">${KIND[k].t}</button>`).join("")}</div>`;
  if(!PRIV) return head+lockView(kind);
  const d=(PRIV[kind]||{})[m.iso3];
  let body="";
  if(kind==="procesadores"){
    if(!d) body=`<p class="empty">No hay cuotas cargadas para ${esc(m.name)}.</p>`;
    else{
      const ys=PRIV.meta_proc.years, y=ys[ys.length-1], yp=ys[ys.length-2];
      const firms=d.rows.filter(r=>!/^(total|others|private label)/i.test(r.name)&&r.shares[y]!=null).sort((a,b)=>b.shares[y]-a.shares[y]).slice(0,10);
      const pl=d.rows.find(r=>/^private label/i.test(r.name)), max=Math.max(...firms.map(f=>f.shares[y]),1);
      const vol=(PRIV.volumen||{})[m.iso3];
      const top=firms.reduce((a,f)=>a+f.shares[y],0);
      body=`<div class="kpis"><div class="kpi"><div class="l">Top 10 · cuota conjunta ${y}</div><div class="v">${fmt1(top)} % <small>del valor</small></div></div>
        <div class="kpi"><div class="l">Mercado de chocolate ${vol?Object.keys(vol.years).slice(-1)[0]:""}</div><div class="v">${vol?fmtInt(Object.values(vol.years).slice(-1)[0]):"—"} <small>${vol?"mil t":"sin dato"}</small></div></div></div>
      <div class="list">${firms.map((f,i)=>{ const dlt=yp&&f.shares[yp]!=null?f.shares[y]-f.shares[yp]:null;
        return `<button class="item pv-row" data-pco="procesadores|${i}" style="grid-template-columns:28px minmax(0,1fr) auto">
          <span class="n">${String(i+1).padStart(2,"0")}</span>
          <span><div class="t">${esc(f.name)}</div>${bar(f.shares[y],max,"var(--cacao-orange)")}</span>
          <span style="text-align:right"><div class="mono" style="font-size:16px">${fmt1(f.shares[y])} %</div><div class="src mono" style="color:${dlt==null?"var(--dim)":dlt>0?"var(--cacao-green)":dlt<0?"var(--cacao-red)":"var(--dim)"}">${dlt==null?"—":(dlt>0?"+":"")+fmt1(dlt)+" pp"}</div></span></button>`;}).join("")}</div>
      ${pl?`<p class="src">Marca blanca (private label): ${fmt1(pl.shares[y]||0)} % del mercado.</p>`:""}
      <p class="src"><b>Fuente:</b> ${K.src}, ${esc(PRIV.meta_proc.category||"Chocolate Confectionery")}, cuota de ${esc(PRIV.meta_proc.dataType||"Retail Value RSP")} ${ys.join("–")}. Exportado: ${esc(PRIV.meta_proc.exported||"—")}. Cargado: ${esc(PRIV.meta_proc.loaded||"—")}.</p>
      <div class="def">Estas cuotas miden ventas de chocolate al consumidor. Indican quién domina el mercado, no cuánto cacao en grano compra cada empresa.</div>`;
      d._top=firms;
    }
  } else {
    const anyEc=Object.values(PRIV.importadores||{}).find(x=>x&&x.mode==="ec_export"&&!x.hqOnly);
    if(!d||d.hqOnly) body = anyEc ? `<div class="notice info"><b>Sin envíos directos desde Ecuador</b> hacia ${esc(m.name)} entre ${esc(anyEc.from||"—")} y ${esc(anyEc.to||"—")}, según la aduana de Ecuador (Penta).</div>` : `<p class="empty">Todavía no hay datos de Penta para ${esc(m.name)}. Exporta el Excel del país en Penta y súbelo aquí abajo.</p>`;
    else{
      const by=d.by, max=Math.max(...d.rows.slice(0,10).map(r=>r.share),1), ec=d.mode==="ec_export";
      body=`${ec?`<div class="notice info"><b>Compradores de cacao ecuatoriano.</b> Exportaciones de Ecuador con destino ${esc(m.name)}, según la aduana de Ecuador. Cada empresa ya compra origen Ecuador.</div>`:""}<div class="kpis"><div class="kpi"><div class="l">${ec?"Compradores":"Importadores"} en el archivo</div><div class="v">${fmtInt(d.importers)} <small>${fmtInt(d.lines)} operaciones</small></div></div>
        <div class="kpi"><div class="l">Total del período</div><div class="v">${by==="usd"?fmt1(d.totalUsd/1e6)+" <small>USD M</small>":fmtInt(d.totalKg/1000)+" <small>t</small>"}</div></div></div>
      <div class="list">${d.rows.slice(0,10).map((r,i)=>`<button class="item pv-row" data-pco="importadores|${i}" style="grid-template-columns:28px minmax(0,1fr) auto">
          <span class="n">${String(i+1).padStart(2,"0")}</span>
          <span><div class="t">${esc(r.name)}</div>${bar(r.share,max,"var(--cacao-yellow)")}<div class="s">${(r.exporters&&r.exporters.length)?"Le venden: "+r.exporters.map(x=>`${esc(x.name)} ${fmt1(x.share)} %`).join(" · "):""}${r.origins.map(o=>`${o.name==="ECUADOR"||/ecuador/i.test(o.name)?"<b style=\"color:var(--cacao-green)\">"+esc(o.name)+"</b>":esc(o.name)} ${fmt1(o.share)} %`).join(" · ")||""}${r.hs.length?` · HS ${r.hs.join(", ")}`:""}</div></span>
          <span style="text-align:right"><div class="mono" style="font-size:16px">${fmt1(r.share)} %</div><div class="src mono">${by==="usd"?"USD "+fmtInt(r.usd/1000)+" mil":fmtInt(r.kg/1000)+" t"}</div></span></button>`).join("")}</div>
      <p class="src"><b>Fuente:</b> ${K.src}${ec?" (Ecuador, exportaciones)":""}. Período: ${esc(d.from||"—")} a ${esc(d.to||"—")}. Cuota sobre el total del archivo, en ${by==="usd"?"valor":"peso"}. Cargado: ${esc(d.loaded||"—")}.</p>
      <p class="src">Columnas leídas: ${Object.entries(d.columns).map(([k,v])=>`${k} = "${esc(v)}"`).join(" · ")}</p>`;
    }
    if(d&&d.hq&&d.hq.rows.length){ const h=d.hq, mx=Math.max(...h.rows.slice(0,10).map(r=>r.share),1);
      body+=`<section><h3 style="margin:8px 0 4px">Compradores con sede aquí que reciben en otro país</h3>
      <p class="src" style="margin:0 0 8px">Empresas cuya dirección en la aduana está en ${esc(m.name)}, aunque el cacao se envía a otro destino. Suelen ser traders o casas matrices.</p>
      <div class="list">${h.rows.slice(0,10).map((r,i)=>`<button class="item pv-row" data-pco="hq|${i}" style="grid-template-columns:28px minmax(0,1fr) auto">
          <span class="n">${String(i+1).padStart(2,"0")}</span>
          <span><div class="t">${esc(r.name)}</div>${bar(r.share,mx,"var(--cacao-green)")}<div class="s">Recibe en: ${r.dests.map(o=>`${esc(o.name)} ${fmt1(o.share)} %`).join(" · ")||"—"}${r.exporters&&r.exporters.length?" · Le venden: "+r.exporters.map(x=>esc(x.name)).join(", "):""}</div></span>
          <span style="text-align:right"><div class="mono" style="font-size:16px">${r.ops}</div><div class="src mono">operaciones</div></span></button>`).join("")}</div>
      <p class="src"><b>Fuente:</b> Penta-Transaction (Ecuador, exportaciones), dirección del comprador. Período: ${esc(h.from||"—")} a ${esc(h.to||"—")}. La sede se deduce del texto de la dirección; revisa la ficha antes de contactar.</p></section>`; }
  }
  return head+body+uploadBox(kind);
}

function rankCompany(kind,i){
  const m=state.market, d=(PRIV[kind]||{})[m.iso3];
  if(kind==="procesadores"){ const ys=PRIV.meta_proc.years, y=ys[ys.length-1], f=d._top[i];
    return {id:`${m.iso3}-P-${i}`, name:f.name, type:"procesador", city:m.name, what:`Cuota del mercado de chocolate en ${m.name}: ${fmt1(f.shares[y])} % en ${y} (${ys.map(yy=>f.shares[yy]!=null?yy+": "+fmt1(f.shares[yy])+" %":"").filter(Boolean).join(", ")}).`, source:"Euromonitor International", sourceDate:PRIV.meta_proc.exported}; }
  if(kind==="hq"){ const r=PRIV.importadores[m.iso3].hq.rows[i];
    return {id:`${m.iso3}-H-${i}`, name:r.name, type:"importador", city:m.name, what:`Comprador de cacao ecuatoriano con sede en ${m.name} según su dirección en aduana; la mercancía se envía a ${r.dests.map(o=>o.name+" "+fmt1(o.share)+" %").join(", ")||"otro país"} (${r.ops} operaciones${r.usd?", USD "+fmtInt(r.usd):""}). ${r.exporters&&r.exporters.length?"Exportadores ecuatorianos que le venden: "+r.exporters.map(o=>o.name+" "+fmt1(o.share)+" %").join(", ")+".":""}`, hs:r.hs.length?r.hs.filter(h=>HS.some(x=>x.code===h)):undefined, source:"Penta-Transaction (aduana de Ecuador, exportaciones)", sourceDate:`${r.from||PRIV.importadores[m.iso3].hq.from||""} a ${PRIV.importadores[m.iso3].hq.to||""}`}; }
  const r=d.rows[i];
  return {id:`${m.iso3}-I-${i}`, name:r.name, type:"importador", city:m.name, what:`${fmt1(r.share)} % de las importaciones del archivo (${r.ops} operaciones${r.usd?", USD "+fmtInt(r.usd):""}${r.kg?", "+fmtInt(r.kg/1000)+" t":""}). ${r.exporters&&r.exporters.length?"Exportadores ecuatorianos que le venden: "+r.exporters.map(o=>o.name+" "+fmt1(o.share)+" %").join(", "):"Orígenes: "+(r.origins.map(o=>o.name+" "+fmt1(o.share)+" %").join(", ")||"sin dato")}.`, hs:r.hs.length?r.hs.filter(h=>HS.some(x=>x.code===h)):undefined, source:d.mode==="ec_export"?"Penta-Transaction (aduana de Ecuador, exportaciones)":"Penta-Transaction", sourceDate:`${d.from||""} a ${d.to||""}`};
}

/* ---------- integración con la app ---------- */
const st=document.createElement("style");
st.textContent=`.pv-form{display:flex;flex-direction:column;gap:8px}.pv-in{background:var(--surface-2);color:var(--fg);border:1px solid var(--line-2);border-radius:6px;padding:9px 10px;font:inherit;font-size:13px;width:100%}
.pv-up{border:1px solid var(--line);border-radius:8px;padding:12px 14px}.pv-up summary{cursor:pointer;font-size:13px}.pv-up[open] summary{margin-bottom:10px}.pv-row .bar{height:5px;background:var(--surface-2);border-radius:3px;overflow:hidden;margin:6px 0 4px;max-width:260px}.pv-row .bar span{display:block;height:100%}`;
document.head.appendChild(st);

const _viewMarketReal = viewMarketReal;
viewMarketReal = function(){
  const html=_viewMarketReal();
  const cta=`<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
    <button class="cta" data-rank="procesadores">Procesadores <span>→</span></button>
    <button class="cta" data-rank="importadores" style="background:var(--cacao-yellow)">Importadores <span>→</span></button></div>
    <p class="src" style="margin-top:-12px">Procesadores: cuota de mercado (Euromonitor). Importadores: registros de aduana (Penta). Datos privados del equipo.</p>`;
  return html.replace(/<button class="cta" data-act="companies">[\s\S]*?<\/button>|<div class="notice info"><b>Empresas:<\/b>[\s\S]*?<\/div>/, cta);
};
const _viewMarketPending = viewMarketPending;
viewMarketPending = function(){ return _viewMarketPending().replace(/<button class="ghost" data-act="pilot">[\s\S]*?<\/button>/, `<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px"><button class="cta" data-rank="procesadores">Procesadores <span>→</span></button><button class="cta" data-rank="importadores" style="background:var(--cacao-yellow)">Importadores <span>→</span></button></div>`); };
const _render = render;
render = function(){
  if(state.view==="companies"){ state.view="ranking"; state.rkind=state.rkind||(state.company&&state.company.type==="importador"?"importadores":"procesadores"); }
  if(state.view==="ranking"){
    const m=state.market;
    $("#crumbs").innerHTML=`<button data-go="world">Mundo</button><span class="sep">›</span><button data-go="market">${esc(m.name)}</button><span class="sep">›</span><span class="cur">${KIND[state.rkind].t}</span>`;
    $("#crumbs").querySelectorAll("[data-go]").forEach(b=>b.onclick=()=>{ if(b.dataset.go==="world") closePanel(); else {state.view="market"; render();} });
    const b=$("#pbody"); b.innerHTML=viewRanking(); b.scrollTop=0; bind(); bindPriv(); return;
  }
  _render(); bindPriv();
  if(state.view==="company"||state.view==="contact"){ const cr=$("#crumbs").querySelector('[data-go="companies"]'); if(cr) cr.textContent=KIND[state.rkind||"procesadores"].t; }
};
const _viewCompany = viewCompany;
viewCompany = function(){
  let h=_viewCompany(); const c=state.company;
  if(c&&c.source==="Euromonitor International") h=h.replace(/<dt>Encaje estimado<\/dt><dd>[\s\S]*?<\/dd>/,'<dt>Encaje</dt><dd>Por confirmar. La cuota mide ventas de chocolate al consumidor, no compras de cacao en grano.</dd>');
  return h;
};
function bindPriv(){
  const b=$("#pbody");
  b.querySelectorAll("[data-rank]").forEach(x=>x.onclick=()=>{ state.rkind=x.dataset.rank; state.view="ranking"; MSG=""; render(); });
  b.querySelectorAll("[data-rkind]").forEach(x=>x.onclick=()=>{ state.rkind=x.dataset.rkind; MSG=""; render(); });
  b.querySelectorAll("[data-pco]").forEach(x=>x.onclick=()=>{ const [k,i]=x.dataset.pco.split("|"); state.company=rankCompany(k,+i); state.view="company"; state.tab="perfil"; render(); });
  const u=$("#pv-unlock"); if(u) u.onsubmit=async e=>{ e.preventDefault(); const pw=$("#pv-pw").value;
    if(FILE_STATE==="missing"){ if(pw!==$("#pv-pw2").value){ MSG="Las contraseñas no coinciden."; return render(); } if(pw.length<10){ MSG="Usa al menos 10 caracteres."; return render(); }
      PW=pw; ss.set("ek_pw",pw); PRIV={procesadores:{},importadores:{},volumen:{}}; MSG=""; return render(); }
    try{ const ok=await unlock(pw); MSG=ok?"":"No se encontró el archivo de datos."; }catch(err){ MSG="Contraseña incorrecta."; } render(); };
  const f=$("#pv-upload"); if(f) f.onsubmit=async e=>{ e.preventDefault(); const kind=f.dataset.kind, file=$("#pv-file").files[0], tok=$("#pv-token").value.trim(), msg=$("#pv-msg");
    if(tok) ls.set("ek_gh_token",tok);
    try{ msg.textContent="Leyendo el Excel…"; const wb=await readBook(file); const now=new Date().toISOString().slice(0,10);
      if(kind==="procesadores"){ const p=parseEuromonitor(wb); PRIV.procesadores=p.markets; PRIV.volumen=p.volumen; PRIV.meta_proc={years:p.years,category:p.category,dataType:p.dataType,exported:p.exported,loaded:now,file:file.name}; MSG=`Leídos ${Object.keys(p.markets).length} mercados (${p.years.join(", ")}).`; }
      else { const res=parsePenta(wb,$("#pv-country").value); const isos=Object.keys(res.byIso);
        if(!isos.length) throw new Error("El Excel no trae envíos hacia ninguno de los 40 mercados.");
        if(!res.single) PRIV.importadores={};
        for(const iso of isos){ const p=res.byIso[iso]; p.loaded=now; p.file=file.name; PRIV.importadores[iso]=p; }
        if(!res.single) for(const [iso,h] of Object.entries(res.hq||{})){ const e=PRIV.importadores[iso]||(PRIV.importadores[iso]={hqOnly:true,loaded:now,file:file.name,from:h.from,to:h.to}); e.hq=h; }
        const lines=isos.reduce((a,i)=>a+res.byIso[i].lines,0), un=Object.keys(res.unmatched);
        MSG = res.single ? `Leídos ${lines} registros y ${res.byIso[isos[0]].importers} compradores (${BY3[isos[0]].name}).` : `Leídos ${lines} registros: ${isos.length} mercados actualizados (${isos.map(i=>BY3[i].name).join(", ")}).`;
        if(un.length) MSG+=` Destinos fuera de los 40 mercados, no incluidos: ${un.slice(0,8).join(", ")}${un.length>8?"…":""}.`;
        if(!isos.includes(state.market.iso3)) state.market=BY3[isos[0]]; }
      msg.textContent=MSG+" Cifrando y publicando…";
      const clean=JSON.parse(JSON.stringify(PRIV,(k,v)=>k==="_top"?undefined:v)); const r=await publish(clean);
      if(r.ok) MSG+=" Publicado: el equipo lo verá en 1–2 minutos.";
      else { downloadFile(r.encd); MSG+=` No se pudo publicar en GitHub (${r.reason}). Se descargó "${FILE}": súbelo a la raíz del repositorio en GitHub.`; }
    }catch(err){ MSG="Error: "+err.message; }
    render(); };
}

/* ---------- lista de mercados automática (calculada por el robot en imports.json) ---------- */
const REG = {
  "Europa":"AD AL AT BA BE BG BY CH CY CZ DE DK EE ES FI FR GB GR HR HU IE IS IT LI LT LU LV MC MD ME MK MT NL NO PL PT RO RS SE SI SK SM VA XK",
  "Europa y Asia Central":"TR UA RU GE AM AZ KZ KG TJ TM UZ",
  "Américas":"US CA MX GT BZ SV HN NI CR PA CU DO HT JM TT BS BB PR CO VE EC PE BO CL AR UY PY BR GY SR",
  "Asia-Pacífico":"CN JP KR KP TW HK MO MN VN TH MY SG ID PH KH LA MM BN TL IN PK BD LK NP AU NZ PG FJ",
  "Oriente Medio y África":"SA AE QA KW BH OM YE IQ IR IL JO LB SY PS EG LY TN DZ MA SD ZA NG GH CI CM KE ET TZ UG SN AO MZ ZW ZM NA BW MG MU"
};
const regionOf = iso2 => { for(const [r,l] of Object.entries(REG)) if(l.split(" ").includes(iso2)) return r; return "Otros mercados"; };
let SELECTION = null;
const esName = iso2 => { if(iso2==="HK") return "Hong Kong"; try{ return new Intl.DisplayNames(["es"],{type:"region"}).of(iso2); }catch(e){ return iso2; } };
const enName = iso2 => { try{ return new Intl.DisplayNames(["en"],{type:"region"}).of(iso2); }catch(e){ return iso2; } };
function centroidOf(iso2, iso3){ try{ const f=(features||[]).find(f=>f.properties.iso2===iso2||f.properties.iso3===iso3); if(f&&window.d3){ const c=d3.geoCentroid(f); if(isFinite(c[0])) return c; } }catch(e){} return null; }
function applySelection(){
  if(!SELECTION||!SELECTION.markets||!SELECTION.markets.length) return;
  const old=Object.fromEntries(MARKETS.map(m=>[m.iso3,m]));
  const list=SELECTION.markets.filter(x=>x.iso2&&x.iso3).map(x=>{ const o=old[x.iso3]; const c=o?[o.lon,o.lat]:centroidOf(x.iso2,x.iso3);
    return {iso2:x.iso2, iso3:x.iso3, name:o?o.name:esName(x.iso2), region:o?o.region:regionOf(x.iso2), lon:c?c[0]:null, lat:c?c[1]:null, pilot:false, avgUsd:x.avgUsd, rank:x.rank}; });
  if(!list.length) return;
  MARKETS.splice(0, MARKETS.length, ...list);
  Object.keys(BY2).forEach(k=>delete BY2[k]); Object.keys(BY3).forEach(k=>delete BY3[k]);
  MARKETS.forEach(m=>{ BY2[m.iso2]=m; BY3[m.iso3]=m; EURO_NAMES[enName(m.iso2)]=EURO_NAMES[enName(m.iso2)]||m.iso3; });
  if(typeof features!=="undefined" && features.length && document.querySelector("#map svg")){
    MARKETS.forEach(m=>{ if(m.lon==null){ const c=centroidOf(m.iso2,m.iso3); if(c){m.lon=c[0];m.lat=c[1];} } });
    $("#map").innerHTML=""; drawMap(); }
}
const _drawMap = drawMap;
drawMap = function(){ MARKETS.forEach(m=>{ if(m.lon==null){ const c=centroidOf(m.iso2,m.iso3); m.lon=c?c[0]:0; m.lat=c?c[1]:-89; } }); _drawMap(); };
const _renderMarkets = renderMarkets;
renderMarkets = function(){ _renderMarkets(); const n=$("#mbody .notice");
  if(n&&SELECTION) n.innerHTML=`<b>Top ${MARKETS.length} destinos del cacao ecuatoriano.</b> Elegidos automáticamente con datos de UN Comtrade (exportaciones de Ecuador, HS 1801–1806, promedio ${esc((SELECTION.years||[]).join(", "))}). Se recalcula solo cada mes. Actualizado: ${esc(SELECTION.generated||"—")}.`;
  if(SELECTION) $("#mbody").querySelectorAll(".mrow").forEach(b=>{ const m=BY2[b.dataset.m]; if(m&&m.rank){ const sm=b.querySelector("small,.chip"); b.firstElementChild.textContent=`${m.rank}. ${m.name}`; } }); };
fetch("imports.json",{cache:"no-store"}).then(r=>r.ok?r.json():null).then(j=>{ if(j&&j.selection){ SELECTION=j.selection; applySelection(); } }).catch(()=>{});


/* Sin señalización de "piloto" */
MARKETS.forEach(m=>m.pilot=false);
(function(){ const h=document.querySelector("#hero p"); if(h) h.textContent="Busca un país o selecciona uno de los 40 mercados prioritarios.";
  document.querySelectorAll("#hero .legend span").forEach(s=>{ if(/Piloto/.test(s.textContent)) s.remove(); }); })();
/* ===== Búsqueda de países + foto pequeña ===== */
const norm = t => (t||"").normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase().trim();
(function addSearchUI(){
  const st=document.createElement("style"); st.textContent=`
  .srch{position:relative;width:min(260px,34vw)}
  .srch input{width:100%;background:var(--surface);border:1px solid var(--line);border-radius:999px;padding:7px 14px 7px 32px;font:inherit;font-size:13px;color:inherit}
  .srch input:focus{outline:none;border-color:var(--cacao-orange)}
  .srch .ic{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--muted);font-size:13px;pointer-events:none}
  .sres{position:absolute;top:calc(100% + 6px);left:0;right:0;background:var(--surface);border:1px solid var(--line);border-radius:12px;overflow:hidden auto;max-height:320px;box-shadow:0 12px 30px rgba(0,0,0,.5)}
  .sres button{display:flex;justify-content:space-between;gap:8px;width:100%;text-align:left;background:none;border:0;border-bottom:1px solid var(--line);padding:9px 12px;font-size:13px;color:inherit}
  .sres button:last-child{border-bottom:0}
  .sres button:hover,.sres button.on{background:var(--surface-2)}
  .sres small{color:var(--dim);font-size:11px;white-space:nowrap}
  .sres small.mk{color:var(--cacao-orange)}
  .cphoto{float:right;margin:0 0 8px 14px;width:112px}
  .cphoto img{display:block;width:112px;height:80px;object-fit:cover;border-radius:10px;border:1px solid var(--line);background:var(--surface-2)}
  .cphoto figcaption{font-size:10px;color:var(--dim);margin-top:4px;line-height:1.3}
  .cphoto figcaption a{color:inherit}
  .stoast{position:fixed;left:50%;bottom:28px;transform:translateX(-50%);z-index:30;background:var(--surface);border:1px solid var(--line);border-radius:999px;padding:8px 16px;font-size:13px}
  @media (max-width:720px){ .top .brand span{display:none} .srch{width:auto;flex:1;min-width:0} .top .spacer{display:none} }`;
  document.head.appendChild(st);
  const box=document.createElement("div"); box.className="srch";
  box.innerHTML=`<span class="ic">⌕</span><input type="search" id="srchIn" placeholder="Buscar país…" aria-label="Buscar país" autocomplete="off"><div class="sres" id="srchRes" hidden></div>`;
  const top=document.querySelector(".top"), anchor=document.getElementById("btnMarkets"); top.insertBefore(box, anchor);
  const inp=box.querySelector("input"), res=box.querySelector(".sres");
  let items=[], sel=0;
  function allCountries(){
    const out=new Map();
    MARKETS.forEach(m=>out.set(m.iso2,{iso2:m.iso2,name:m.name,market:m}));
    (features||[]).forEach(f=>{ const i2=f.properties.iso2; if(!i2||out.has(i2)) return; const n=esName(i2); out.set(i2,{iso2:i2,name:(n&&n!==i2)?n:f.properties.name,feature:f}); });
    return [...out.values()];
  }
  function show(){
    const q=norm(inp.value); if(!q){ res.hidden=true; return; }
    items=allCountries().map(c=>({c, n:norm(c.name), e:norm(enName(c.iso2))})).filter(x=>x.n.includes(q)||x.e.includes(q))
      .sort((a,b)=>(b.n.startsWith(q)-a.n.startsWith(q))||(!!b.c.market-!!a.c.market)||a.n.localeCompare(b.n)).slice(0,8).map(x=>x.c);
    sel=0;
    res.innerHTML = items.length ? items.map((c,i)=>`<button data-i="${i}" class="${i===0?"on":""}"><span>${esc(c.name)}</span>${c.market?`<small class="mk">${c.market.rank?"#"+c.market.rank+" · ":""}Mercado prioritario</small>`:"<small>Fuera de los 40</small>"}</button>`).join("") : `<button disabled><span>Sin resultados</span></button>`;
    res.hidden=false;
    res.querySelectorAll("[data-i]").forEach(b=>b.onmousedown=e=>{ e.preventDefault(); pick(items[+b.dataset.i]); });
  }
  function pick(c){
    if(!c) return; inp.value=""; res.hidden=true; inp.blur();
    if(c.market){ openMarket(c.market); return; }
    if(typeof closePanel==="function" && $("#panel").classList.contains("open")) $("#panel").classList.remove("open");
    focusFeature(c.feature);
    toast(`${c.name} no está entre los 40 mercados prioritarios.`);
  }
  inp.addEventListener("input",show);
  inp.addEventListener("focus",show);
  inp.addEventListener("blur",()=>setTimeout(()=>res.hidden=true,150));
  inp.addEventListener("keydown",e=>{
    if(e.key==="Enter"){ e.preventDefault(); pick(items[sel]); }
    else if(e.key==="Escape"){ inp.value=""; res.hidden=true; }
    else if(e.key==="ArrowDown"||e.key==="ArrowUp"){ e.preventDefault(); if(!items.length) return; sel=(sel+(e.key==="ArrowDown"?1:-1)+items.length)%items.length; res.querySelectorAll("[data-i]").forEach((b,i)=>b.classList.toggle("on",i===sel)); }
  });
})();
function focusFeature(f){
  if(!svg||!f) return;
  g.selectAll(".country").classed("active",x=>x===f);
  const el=$("#map"); const w=el.clientWidth,h=el.clientHeight;
  let [[x0,y0],[x1,y1]]=path.bounds(f);
  if((x1-x0)>w*0.5){ const p=projection(d3.geoCentroid(f)); x0=p[0]-90;x1=p[0]+90;y0=p[1]-60;y1=p[1]+60; }
  const k=Math.min(8,0.55/Math.max((x1-x0)/w,(y1-y0)/h)); const cx=(x0+x1)/2, cy=(y0+y1)/2;
  svg.transition().duration(750).call(zoom.transform,d3.zoomIdentity.translate(w/2-k*cx,h/2-k*cy).scale(k));
}
let toastT; function toast(t){ let e=document.querySelector(".stoast"); if(!e){ e=document.createElement("div"); e.className="stoast"; document.body.appendChild(e); } e.textContent=t; e.hidden=false; clearTimeout(toastT); toastT=setTimeout(()=>e.hidden=true,3500); }
/* Foto pequeña del país: imagen de su capital en Wikimedia Commons (vía Wikidata) */
const PHOTO={};
function photoFor(iso2){
  if(PHOTO[iso2]) return PHOTO[iso2];
  const q=`SELECT ?img ?cimg ?capLabel WHERE { ?c wdt:P297 "${iso2}". OPTIONAL{ ?c wdt:P36 ?cap. ?cap wdt:P18 ?img. } OPTIONAL{ ?c wdt:P18 ?cimg. } SERVICE wikibase:label { bd:serviceParam wikibase:language "es,en". } } LIMIT 5`;
  PHOTO[iso2]=fetch("https://query.wikidata.org/sparql?format=json&query="+encodeURIComponent(q),{headers:{Accept:"application/sparql-results+json"}})
    .then(r=>r.json()).then(j=>{ const b=j.results.bindings; const x=b.find(r=>r.img)||b.find(r=>r.cimg); if(!x) return null;
      const u=(x.img||x.cimg).value.replace(/^http:/,"https:"); const file=decodeURIComponent(u.split("/Special:FilePath/")[1]||"");
      return {src:u+"?width=260", page:"https://commons.wikimedia.org/wiki/File:"+encodeURIComponent(file), cap:x.img&&x.capLabel?x.capLabel.value:""}; })
    .catch(()=>null);
  return PHOTO[iso2];
}
const _render2=render;
render=function(){ _render2(); if(state.view!=="market"||!state.market) return; const m=state.market;
  photoFor(m.iso2).then(p=>{ if(!p||state.view!=="market"||state.market!==m) return; const head=$("#pbody").firstElementChild; if(!head||head.querySelector(".cphoto")) return;
    const fig=document.createElement("figure"); fig.className="cphoto";
    fig.innerHTML=`<img src="${esc(p.src)}" alt="${esc(p.cap||m.name)}" loading="lazy"><figcaption>${p.cap?esc(p.cap)+" · ":""}<a href="${esc(p.page)}" target="_blank" rel="noopener">Wikimedia Commons</a></figcaption>`;
    fig.querySelector("img").onerror=()=>fig.remove(); head.insertBefore(fig, head.firstChild); }); };

/* ===== Contactos reales vía Apollo (se activa al poner APOLLO_URL) ===== */
const APOLLO_URL = "";   // URL del intermediario de Cloudflare. Vacío = botones ocultos.
const APL = {};
async function teamHash(){ const d=await crypto.subtle.digest("SHA-256", new TextEncoder().encode(PW||"")); return [...new Uint8Array(d)].map(b=>b.toString(16).padStart(2,"0")).join(""); }
async function apl(path, body){
  const r=await fetch(APOLLO_URL.replace(/\/$/,"")+path,{method:"POST",headers:{"content-type":"application/json","x-team":await teamHash()},body:JSON.stringify(body||{})});
  const j=await r.json().catch(()=>({})); if(!r.ok) throw new Error(j.error||("Error "+r.status)); return j;
}
const aplOn = c => APOLLO_URL && PW && c && /Euromonitor|Penta/.test(c.source||"");
const _viewCompany3 = viewCompany;
viewCompany = function(){
  let h=_viewCompany3(); const c=state.company; if(!aplOn(c)) return h;
  h=h.replace(/<section><h3>Contactos clave<\/h3>[\s\S]*?<\/section>/, `<section><h3>Contactos de compras</h3><p class="lead" style="margin:0 0 10px">Personas de compras y abastecimiento según Apollo. Ver la lista no gasta créditos.</p><button class="cta" data-tab="contactos">Ver personas <span>→</span></button></section>`);
  if(state.tab==="contactos") h=h.replace(/(<div class="tabs" role="tablist">[\s\S]*?<\/div>)[\s\S]*$/, `$1<div id="apl"><p class="lead">Buscando personas en Apollo…</p></div>`);
  return h;
};
function aplPerson(p, rv){
  const mail = rv&&rv.emailStatus ? (rv.email?`<a href="mailto:${esc(rv.email)}" style="color:var(--cacao-yellow)">${esc(rv.email)}</a> <button class="ghost" data-copy="${esc(rv.email)}" style="padding:2px 8px;font-size:11px">Copiar</button>`:`<span style="color:var(--dim)">Apollo no tiene su email</span>`)
    : `<button class="cta" data-reveal="${esc(p.id)}" style="padding:7px 12px;font-size:13px">Revelar email <span style="font-size:11px;opacity:.8">· 1 crédito</span></button>`;
  let phone="";
  if(rv&&rv.emailStatus){
    if(rv.phoneStatus==="ok") phone=(rv.phones||[]).map(n=>`<a href="tel:${esc(n)}" style="color:var(--cacao-yellow)">${esc(n)}</a>`).join(" · ");
    else if(rv.phoneStatus==="pending") phone=`<span style="color:var(--dim)">Buscando número… puede tardar unos minutos.</span>`;
    else if(rv.phoneStatus==="none") phone=`<span style="color:var(--dim)">Apollo no tiene su número</span>`;
    else phone=`<button class="ghost" data-rphone="${esc(p.id)}" style="padding:4px 10px;font-size:12px">Revelar número</button> <small style="color:var(--dim)">+8 créditos</small>`;
  }
  return `<div class="pv-row" style="border-bottom:1px solid var(--line);padding:12px 0">
    <div style="font-weight:500">${esc(p.name)}${p.linkedin?` <a href="${esc(p.linkedin)}" target="_blank" rel="noopener" style="color:var(--muted);font-size:12px">LinkedIn ↗</a>`:""}</div>
    <div style="color:var(--muted);font-size:13px">${esc(p.title||"Cargo sin dato")}${p.city?" · "+esc(p.city):""}</div>
    <div style="margin-top:8px;display:flex;flex-direction:column;gap:6px;align-items:flex-start">${mail}${phone?`<div>${phone}</div>`:""}</div></div>`;
}
function aplDraw(){
  const box=$("#apl"); if(!box) return; const c=state.company, k=c.name+"|"+state.market.iso2, d=APL[k];
  if(!d) return; if(d.error){ box.innerHTML=`<div class="notice">No se pudo consultar Apollo: ${esc(d.error)}</div>`; return; }
  if(!d.org){ box.innerHTML=`<div class="notice info">Apollo no encontró esta empresa con ese nombre. Prueba buscarla a mano en <a href="https://app.apollo.io/#/companies?qOrganizationName=${encodeURIComponent(c.name)}" target="_blank" rel="noopener" style="color:var(--cacao-yellow)">Apollo ↗</a>.</div>`; return; }
  box.innerHTML=`<div class="def">Empresa en Apollo: <b style="color:var(--fg)">${esc(d.org.name)}</b>${d.org.domain?" · "+esc(d.org.domain.replace(/^https?:\/\//,"")):""}. Revela solo a quien vayas a contactar: el plan gratis trae 75 créditos al mes.</div>
    ${d.people.length?d.people.map(p=>aplPerson(p,d.rv[p.id])).join(""):'<p class="lead">No hay personas de compras registradas para esta empresa.</p>'}
    <p class="src">Fuente: Apollo.io. Solo datos profesionales. Lo ya revelado queda guardado para el equipo y no se vuelve a pagar.</p>`;
  box.querySelectorAll("[data-reveal]").forEach(b=>b.onclick=()=>aplReveal(b.dataset.reveal,false,b));
  box.querySelectorAll("[data-rphone]").forEach(b=>b.onclick=()=>aplReveal(b.dataset.rphone,true,b));
  box.querySelectorAll("[data-copy]").forEach(b=>b.onclick=()=>{ try{navigator.clipboard.writeText(b.dataset.copy); b.textContent="Copiado";}catch(e){} });
}
async function aplLoad(){
  const c=state.company, k=c.name+"|"+state.market.iso2;
  if(!APL[k]){ try{ const j=await apl("/people",{company:c.name,country:enName(state.market.iso2)}); APL[k]={org:j.org,people:j.people||[],rv:{}};
      await Promise.all((j.people||[]).map(async p=>{ try{ const r=await apl("/contact",{id:p.id}); if(r&&r.emailStatus) APL[k].rv[p.id]=r; }catch(e){} }));
    }catch(e){ APL[k]={error:e.message}; } }
  aplDraw(); aplPoll();
}
async function aplReveal(id, phone, btn){
  const k=state.company.name+"|"+state.market.iso2; btn.disabled=true; btn.textContent=phone?"Pidiendo número…":"Revelando…";
  try{ APL[k].rv[id]=await apl("/reveal",{id,phone}); }catch(e){ toast("Apollo: "+e.message); }
  aplDraw(); aplPoll();
}
let aplT; function aplPoll(){
  clearTimeout(aplT); const c=state.company; if(!c||!$("#apl")) return; const k=c.name+"|"+state.market.iso2, d=APL[k]; if(!d||!d.rv) return;
  const pend=Object.entries(d.rv).filter(([,v])=>v.phoneStatus==="pending"); if(!pend.length) return;
  aplT=setTimeout(async()=>{ for(const [id] of pend){ try{ d.rv[id]=await apl("/contact",{id}); }catch(e){} } aplDraw(); aplPoll(); },20000);
}
const _bindPriv2 = bindPriv;
bindPriv = function(){ _bindPriv2(); if($("#apl")) aplLoad(); };
tryAutoUnlock();
})();
