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
  name:/importador|importer|consignee|consignatario|raz[oó]n social|^empresa|company/i,
  usd:/(valor|value|fob|cif|monto).*(us|usd|\$|d[oó]lar)|^(us\$|usd)|valor (fob|cif)|^fob|^cif|valor/i,
  kg:/peso neto|net weight|peso|kg|kilo|cantidad|quantity|toneladas|tons/i,
  date:/fecha|date|per[ií]odo|^mes|month/i,
  origin:/pa[ií]s de origen|origen|origin|procedencia|pa[ií]s de procedencia/i,
  hs:/posici[oó]n|partida|ncm|hs|arancel|nandina|tariff/i
};
function parsePenta(wb){
  let best=null;
  for(const name of wb.SheetNames){ const r=rowsOf(wb.Sheets[name]);
    for(let i=0;i<Math.min(r.length,40);i++){ const H=(r[i]||[]).map(c=>String(c||"").trim()); const idx={};
      for(const [k,re] of Object.entries(COLS)){ const j=H.findIndex((h,jj)=>h&&re.test(h)&&!Object.values(idx).includes(jj)); if(j>=0) idx[k]=j; }
      if(idx.name!=null&&(idx.usd!=null||idx.kg!=null)){ best={name,hdr:i,H,idx,rows:r}; break; } }
    if(best) break; }
  if(!best) throw new Error("No encuentro una columna de importador y otra de valor o peso. Revisa que el Excel tenga encabezados como \"Importador\" y \"Valor FOB US$\" o \"Peso neto\".");
  const {idx,H,rows}=best, agg={}; let tot=0, totKg=0, n=0, minD=null, maxD=null;
  for(const r of rows.slice(best.hdr+1)){ if(!r) continue; const nm=String(r[idx.name]||"").trim(); if(!nm||/^total/i.test(nm)) continue;
    const usd=idx.usd!=null?num(r[idx.usd])||0:0, kg=idx.kg!=null?num(r[idx.kg])||0:0; if(!usd&&!kg) continue;
    const key=nm.toUpperCase().replace(/[.,]/g,"").replace(/\s+/g," "); const a=agg[key]=agg[key]||{name:nm,usd:0,kg:0,ops:0,origins:{},hs:{}};
    a.usd+=usd; a.kg+=kg; a.ops++; tot+=usd; totKg+=kg; n++;
    if(idx.origin!=null){ const o=String(r[idx.origin]||"").trim(); if(o) a.origins[o]=(a.origins[o]||0)+(usd||kg); }
    if(idx.hs!=null){ const h=String(r[idx.hs]||"").replace(/\D/g,"").slice(0,4); if(h) a.hs[h]=1; }
    if(idx.date!=null){ let d=r[idx.date]; if(typeof d==="number"&&d>20000&&d<80000){ d=new Date(Math.round((d-25569)*864e5)).toISOString().slice(0,10); } d=d?String(d).slice(0,10):null; if(d){ if(!minD||d<minD)minD=d; if(!maxD||d>maxD)maxD=d; } } }
  const by = tot>0 ? "usd" : "kg", T = tot>0 ? tot : totKg;
  const list=Object.values(agg).sort((a,b)=>b[by]-a[by]).slice(0,15).map(a=>{ const ot=Object.values(a.origins).reduce((x,y)=>x+y,0)||1;
    return {name:a.name, usd:Math.round(a.usd), kg:Math.round(a.kg), ops:a.ops, share:T?a[by]/T*100:0, hs:Object.keys(a.hs).sort(),
      origins:Object.entries(a.origins).sort((x,y)=>y[1]-x[1]).slice(0,3).map(([o,v])=>({name:o,share:v/ot*100}))}; });
  return {rows:list, by, total:T, totalUsd:Math.round(tot), totalKg:Math.round(totKg), lines:n, from:minD, to:maxD, columns:Object.fromEntries(Object.entries(idx).map(([k,j])=>[k,H[j]])), importers:Object.keys(agg).length};
}

/* ---------- publicar en GitHub ---------- */
async function publish(obj){
  const encd=await encryptObj(obj,PW); const content=JSON.stringify(encd);
  const token=ls.get("ek_gh_token");
  if(!token) return {ok:false, encd, reason:"sin token"};
  const api=`https://api.github.com/repos/${REPO}/contents/${FILE}`, h={Authorization:"Bearer "+token, Accept:"application/vnd.github+json"};
  let sha; const g=await fetch(api,{headers:h}); if(g.ok) sha=(await g.json()).sha; else if(g.status!==404) return {ok:false, encd, reason:`GitHub respondió ${g.status} al leer el archivo. Revisa el token.`};
  const p=await fetch(api,{method:"PUT",headers:h,body:JSON.stringify({message:"Datos privados actualizados desde la app",content:btoa(content),sha})});
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
      ${kind==="importadores"?`<label for="pv-country" class="src"><b>País del archivo</b></label><select id="pv-country" class="pv-in">${MARKETS.map(m=>`<option value="${m.iso3}" ${m.iso3===state.market.iso3?"selected":""}>${esc(m.name)}</option>`).join("")}</select>`:""}
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
    if(!d) body=`<p class="empty">Todavía no hay datos de Penta para ${esc(m.name)}. Exporta el Excel del país en Penta y súbelo aquí abajo.</p>`;
    else{
      const by=d.by, max=Math.max(...d.rows.slice(0,10).map(r=>r.share),1);
      body=`<div class="kpis"><div class="kpi"><div class="l">Importadores en el archivo</div><div class="v">${fmtInt(d.importers)} <small>${fmtInt(d.lines)} operaciones</small></div></div>
        <div class="kpi"><div class="l">Total del período</div><div class="v">${by==="usd"?fmt1(d.totalUsd/1e6)+" <small>USD M</small>":fmtInt(d.totalKg/1000)+" <small>t</small>"}</div></div></div>
      <div class="list">${d.rows.slice(0,10).map((r,i)=>`<button class="item pv-row" data-pco="importadores|${i}" style="grid-template-columns:28px minmax(0,1fr) auto">
          <span class="n">${String(i+1).padStart(2,"0")}</span>
          <span><div class="t">${esc(r.name)}</div>${bar(r.share,max,"var(--cacao-yellow)")}<div class="s">${r.origins.map(o=>`${o.name==="ECUADOR"||/ecuador/i.test(o.name)?"<b style=\"color:var(--cacao-green)\">"+esc(o.name)+"</b>":esc(o.name)} ${fmt1(o.share)} %`).join(" · ")||""}${r.hs.length?` · HS ${r.hs.join(", ")}`:""}</div></span>
          <span style="text-align:right"><div class="mono" style="font-size:16px">${fmt1(r.share)} %</div><div class="src mono">${by==="usd"?"USD "+fmtInt(r.usd/1000)+" mil":fmtInt(r.kg/1000)+" t"}</div></span></button>`).join("")}</div>
      <p class="src"><b>Fuente:</b> ${K.src}. Período: ${esc(d.from||"—")} a ${esc(d.to||"—")}. Cuota sobre el total del archivo, en ${by==="usd"?"valor":"peso"}. Cargado: ${esc(d.loaded||"—")}.</p>
      <p class="src">Columnas leídas: ${Object.entries(d.columns).map(([k,v])=>`${k} = "${esc(v)}"`).join(" · ")}</p>`;
    }
  }
  return head+body+uploadBox(kind);
}

function rankCompany(kind,i){
  const m=state.market, d=PRIV[kind][m.iso3];
  if(kind==="procesadores"){ const ys=PRIV.meta_proc.years, y=ys[ys.length-1], f=d._top[i];
    return {id:`${m.iso3}-P-${i}`, name:f.name, type:"procesador", city:m.name, what:`Cuota del mercado de chocolate en ${m.name}: ${fmt1(f.shares[y])} % en ${y} (${ys.map(yy=>f.shares[yy]!=null?yy+": "+fmt1(f.shares[yy])+" %":"").filter(Boolean).join(", ")}).`, source:"Euromonitor International", sourceDate:PRIV.meta_proc.exported}; }
  const r=d.rows[i];
  return {id:`${m.iso3}-I-${i}`, name:r.name, type:"importador", city:m.name, what:`${fmt1(r.share)} % de las importaciones del archivo (${r.ops} operaciones${r.usd?", USD "+fmtInt(r.usd):""}${r.kg?", "+fmtInt(r.kg/1000)+" t":""}). Orígenes: ${r.origins.map(o=>o.name+" "+fmt1(o.share)+" %").join(", ")||"sin dato"}.`, hs:r.hs.length?r.hs.filter(h=>HS.some(x=>x.code===h)):undefined, source:"Penta-Transaction", sourceDate:`${d.from||""} a ${d.to||""}`};
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
      else { const iso=$("#pv-country").value, p=parsePenta(wb); p.loaded=now; p.file=file.name; PRIV.importadores[iso]=p; MSG=`Leídos ${p.lines} registros y ${p.importers} importadores.`; if(iso!==state.market.iso3) state.market=BY3[iso]; }
      msg.textContent=MSG+" Cifrando y publicando…";
      const clean=JSON.parse(JSON.stringify(PRIV,(k,v)=>k==="_top"?undefined:v)); const r=await publish(clean);
      if(r.ok) MSG+=" Publicado: el equipo lo verá en 1–2 minutos.";
      else { downloadFile(r.encd); MSG+=` No se pudo publicar en GitHub (${r.reason}). Se descargó "${FILE}": súbelo a la raíz del repositorio en GitHub.`; }
    }catch(err){ MSG="Error: "+err.message; }
    render(); };
}
tryAutoUnlock();
})();
