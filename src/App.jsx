import { useState, useEffect, useMemo, useCallback } from "react";
import { BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend } from "recharts";
import * as XLSX from "xlsx";

const GAS_URL = "https://script.google.com/macros/s/AKfycbyNb_zx1ZnASY78XPNAF8RTMMgSjoe8GG9yJJ_SMsNVigiUmzLrwAL4tNyu1iBXOj1TEg/exec";

// ── 預設類別 ──
const DEFAULT_CATS = [
  { key:"income",   label:"薪資收入", icon:"💼", color:"#4ade80", sign: 1, signed:false, desc:"薪水、兼職、獎金" },
  { key:"invest",   label:"股票/ETF", icon:"📊", color:"#38bdf8", sign: 0, signed:true,  desc:"+獲利/入金  −虧損/出金" },
  { key:"living",   label:"生活費用", icon:"🛒", color:"#fb923c", sign:-1, signed:false, desc:"餐飲、購物、交通" },
  { key:"learning", label:"學習花費", icon:"📚", color:"#a78bfa", sign:-1, signed:false, desc:"課程、書籍、工具" },
  { key:"loan",     label:"貸款",     icon:"🏦", color:"#fbbf24", sign:-1, signed:false, desc:"房貸、車貸、分期" },
  { key:"card",     label:"信用卡費", icon:"💳", color:"#f472b6", sign:-1, signed:false, desc:"信用卡帳單" },
];

const PALETTE = ["#4ade80","#38bdf8","#fb923c","#a78bfa","#fbbf24","#f472b6","#f87171","#34d399","#818cf8","#e879f9","#facc15","#60a5fa","#2dd4bf","#c084fc"];
const ICONS   = ["💼","📊","🛒","📚","🏦","💳","🏠","🚗","✈️","🍽️","💊","🎮","📱","💡","🎁","💰","📈","📉","🏋️","🎓","🏥","☕","🐶","👶","🌿","⛪","✝️","🕌","🙏","⚕️","💉","🧘","🎵","📷","🎨","🏖️","⛺","🎯","🏆","💎","🌟","🔑","🛡️"];

// ── Storage helpers ──
const SK  = "fin_v10";
const UK  = "fin_user_v10";
const CK  = (u) => `fin_cats_${u}`;
const loadRecs  = (u) => { try { return JSON.parse(localStorage.getItem(`${SK}_${u}`)||"[]"); } catch { return []; } };
const saveRecs  = (u,r) => { try { localStorage.setItem(`${SK}_${u}`,JSON.stringify(r)); } catch {} };
const loadCats  = (u) => { try { const c=JSON.parse(localStorage.getItem(CK(u))||"null"); return c||DEFAULT_CATS; } catch { return DEFAULT_CATS; } };
const saveCats  = (u,c) => { try { localStorage.setItem(CK(u),JSON.stringify(c)); } catch {} };
const loadUser  = () => { try { return localStorage.getItem(UK)||""; } catch { return ""; } };
const saveUser  = (u) => { try { localStorage.setItem(UK,u); } catch {} };

// ── Cloud sync ──
async function pushCloud(user, records) {
  try { await fetch(GAS_URL,{method:"POST",mode:"no-cors",headers:{"Content-Type":"application/json"},body:JSON.stringify({action:"save",user,records})}); return true; }
  catch { return false; }
}
async function pullCloud(user) {
  try { const r=await fetch(`${GAS_URL}?user=${encodeURIComponent(user)}&t=${Date.now()}`); const j=await r.json(); return j.ok&&Array.isArray(j.records)?j.records:null; }
  catch { return null; }
}

// ── Utils ──
const fmtN  = (n) => Math.abs(Math.round(n)).toLocaleString("zh-TW");
const fmtS  = (n) => (n>=0?"+":"-")+fmtN(n);
const today = () => new Date().toISOString().slice(0,10);
const nowYM = () => { const n=new Date(); return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}`; };
const ymLbl = (ym) => { const [y,m]=ym.split("-"); return `${y}年${parseInt(m)}月`; };

// ── CSS ──
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
html,body{background:#04060d}
input,button,select,textarea{font-family:inherit;outline:none}
::-webkit-scrollbar{width:4px} ::-webkit-scrollbar-track{background:transparent} ::-webkit-scrollbar-thumb{background:#1e293b;border-radius:99px}

/* glass card - lighter navy */
.gc{background:rgba(30,45,80,.75);backdrop-filter:blur(24px);-webkit-backdrop-filter:blur(24px);border:1px solid rgba(255,255,255,.12);border-radius:22px;padding:20px;position:relative;overflow:hidden}
.gc::before{content:'';position:absolute;top:0;left:15%;right:15%;height:1px;background:linear-gradient(90deg,transparent,rgba(255,255,255,.15),transparent)}

/* field */
.fi{width:100%;background:rgba(15,25,55,.8);border:1.5px solid rgba(255,255,255,.15);border-radius:13px;padding:12px 15px;color:#f1f5f9;font-size:15px;transition:border-color .2s,box-shadow .2s;-webkit-appearance:none}
.fi:focus{border-color:#4ade80;box-shadow:0 0 0 3px rgba(74,222,128,.15)}
.fi::placeholder{color:#475569}

/* tabs */
.tabs{display:flex;gap:3px;background:rgba(15,25,55,.7);border:1px solid rgba(255,255,255,.1);border-radius:15px;padding:4px}
.tab{flex:1;border:none;background:transparent;color:#64748b;font-size:12px;font-weight:600;padding:9px 4px;cursor:pointer;border-radius:11px;transition:all .22s;white-space:nowrap;letter-spacing:.02em}
.tab.on{color:#fff;box-shadow:0 2px 12px rgba(0,0,0,.4)}
.tab.on.t0{background:linear-gradient(135deg,#166534,#1d4ed8)}
.tab.on.t1{background:linear-gradient(135deg,#1d4ed8,#4338ca)}
.tab.on.t2{background:linear-gradient(135deg,#6d28d9,#1d4ed8)}
.tab.on.t3{background:linear-gradient(135deg,#b45309,#1d4ed8)}
.tab.on.t4{background:linear-gradient(135deg,#0e7490,#1d4ed8)}

/* buttons */
.btn-primary{width:100%;border:none;border-radius:14px;padding:15px;font-size:16px;font-weight:700;cursor:pointer;letter-spacing:.04em;transition:transform .1s,box-shadow .2s,opacity .2s}
.btn-primary:active{transform:scale(.97)}
.btn-primary:disabled{opacity:.4;cursor:not-allowed;box-shadow:none!important}
.btn-sm{background:rgba(30,45,80,.6);border:1.5px solid rgba(255,255,255,.15);border-radius:10px;padding:7px 13px;font-size:11px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:5px;transition:all .2s;color:#94a3b8}
.btn-sm:hover{border-color:rgba(255,255,255,.3);color:#f1f5f9}
.btn-sm:disabled{opacity:.3;cursor:not-allowed}
.btn-ghost{background:none;border:1.5px solid rgba(255,255,255,.15);border-radius:11px;padding:9px 16px;font-size:12px;font-weight:600;cursor:pointer;transition:all .2s}
.btn-danger{background:none;border:1.5px solid rgba(248,113,113,.2);color:rgba(248,113,113,.6);border-radius:12px;padding:10px;font-size:12px;cursor:pointer;width:100%;margin-top:14px;font-family:inherit;transition:all .2s}
.btn-danger:hover{border-color:rgba(248,113,113,.5);color:rgba(248,113,113,.9)}

/* sign toggle */
.sign-toggle{display:flex;border-radius:13px;overflow:hidden;border:1.5px solid rgba(255,255,255,.12);background:rgba(15,25,55,.6)}
.sign-btn{flex:1;border:none;padding:12px;font-size:13px;font-weight:700;cursor:pointer;transition:all .2s;background:transparent;display:flex;align-items:center;justify-content:center;gap:7px}

/* category grid */
.cat-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.cat-btn{border:1.5px solid rgba(255,255,255,.1);background:rgba(15,25,55,.6);border-radius:14px;padding:12px 6px 10px;font-size:11px;font-weight:600;color:#64748b;cursor:pointer;transition:all .18s;text-align:center;display:flex;flex-direction:column;align-items:center;gap:5px;line-height:1.3}
.cat-btn:hover{border-color:rgba(255,255,255,.25);color:#94a3b8}
.cat-ico{font-size:24px;line-height:1.1}

/* record row */
.rr{display:flex;align-items:center;gap:11px;padding:13px 0;border-bottom:1px solid rgba(255,255,255,.07)}
.rr:last-child{border-bottom:none}
.rr:hover{background:rgba(255,255,255,.04);margin:0 -8px;padding:13px 8px;border-radius:10px;border-bottom-color:transparent}
.del{background:none;border:none;color:#f87171;padding:7px;opacity:.3;border-radius:8px;cursor:pointer;transition:opacity .2s;flex-shrink:0;font-size:13px}
.del:hover{opacity:1}

/* bar */
.bar-t{height:6px;background:rgba(255,255,255,.08);border-radius:99px;overflow:hidden;margin-top:9px}
.bar-f{height:100%;border-radius:99px;transition:width 1s cubic-bezier(.4,0,.2,1)}

/* pill toggle */
.pill{display:inline-flex;background:rgba(15,25,55,.7);border:1.5px solid rgba(255,255,255,.12);border-radius:11px;overflow:hidden}
.pill button{border:none;padding:8px 16px;font-size:12px;color:#64748b;background:transparent;cursor:pointer;font-weight:600;transition:all .15s}
.pill button.on{background:rgba(29,78,216,.5);color:#f1f5f9}

/* chip */
.chip{font-size:10px;padding:3px 9px;border-radius:99px;font-weight:700;letter-spacing:.03em}
.mono{font-family:'DM Mono',monospace}

/* toast */
.toast{position:fixed;top:22px;left:50%;transform:translateX(-50%);background:rgba(20,35,70,.97);border:1px solid rgba(74,222,128,.5);color:#f1f5f9;padding:11px 24px;border-radius:99px;font-size:13px;z-index:999;box-shadow:0 8px 40px rgba(0,0,0,.5);white-space:nowrap;animation:fu .22s ease;backdrop-filter:blur(20px)}
.toast.err{border-color:rgba(248,113,113,.5)}.toast.warn{border-color:rgba(251,191,36,.5)}
@keyframes fu{from{opacity:0;transform:translateX(-50%) translateY(-10px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}

/* spin */
.spin{animation:sp 1s linear infinite;display:inline-block}
@keyframes sp{to{transform:rotate(360deg)}}

/* cat manager */
.cm-row{display:flex;align-items:center;gap:10px;padding:12px 0;border-bottom:1px solid rgba(255,255,255,.08)}
.cm-row:last-child{border-bottom:none}
.icon-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:6px;max-height:200px;overflow-y:auto;padding:2px}
.icon-btn{border:1.5px solid rgba(255,255,255,.08);background:rgba(15,25,55,.6);border-radius:9px;padding:8px;font-size:20px;cursor:pointer;transition:all .15s;text-align:center}
.icon-btn:hover{background:rgba(255,255,255,.08)}
.icon-btn.on{border-color:#4ade80;background:rgba(74,222,128,.2)}
.color-grid{display:flex;gap:7px;flex-wrap:wrap}
.color-btn{width:28px;height:28px;border-radius:99px;border:2.5px solid transparent;cursor:pointer;transition:all .15s;flex-shrink:0}
.color-btn.on{border-color:#fff;transform:scale(1.25);box-shadow:0 0 8px rgba(255,255,255,.3)}

/* month nav */
.mnav{display:flex;align-items:center;background:rgba(15,25,55,.7);border-radius:13px;border:1.5px solid rgba(255,255,255,.12);overflow:hidden}
.mnav button{border:none;background:transparent;color:#64748b;padding:9px 16px;font-size:18px;cursor:pointer;transition:color .15s}
.mnav button:hover{color:#94a3b8}
.mnav button:disabled{opacity:.25;cursor:not-allowed}

/* filter pills */
.f-pill{border-radius:99px;padding:6px 14px;font-size:11px;font-weight:700;cursor:pointer;transition:all .15s;display:flex;align-items:center;gap:5px;border:1.5px solid rgba(255,255,255,.1);background:rgba(15,25,55,.5);color:#64748b}
.f-pill.on{background:rgba(74,222,128,.15);border-color:rgba(74,222,128,.5);color:#4ade80}
.f-pill:hover{border-color:rgba(255,255,255,.2);color:#94a3b8}

/* glow */
.glow-g{box-shadow:0 0 50px rgba(74,222,128,.1)} .glow-r{box-shadow:0 0 50px rgba(248,113,113,.1)}

/* modal */
.modal-bg{position:fixed;inset:0;z-index:200;display:flex;align-items:flex-end;justify-content:center}
.modal-bd{position:absolute;inset:0;background:rgba(0,0,0,.65);backdrop-filter:blur(10px)}
.modal-sheet{position:relative;width:100%;max-width:520px;background:linear-gradient(160deg,#1a2d5a,#131f42);border:1px solid rgba(255,255,255,.15);border-top:1px solid rgba(255,255,255,.2);border-radius:24px 24px 0 0;padding:24px 20px 40px;z-index:1;max-height:92vh;overflow-y:auto}
.handle{width:40px;height:4px;background:rgba(255,255,255,.2);border-radius:99px;margin:0 auto 20px}

@media(max-width:360px){.cat-grid{grid-template-columns:repeat(2,1fr)}}
`;

export default function App() {
  const [user,      setUser]      = useState(loadUser);
  const [nameInput, setNameInput] = useState("");
  const [cats,      setCats]      = useState(()=> loadUser() ? loadCats(loadUser()) : DEFAULT_CATS);
  const [recs,      setRecs]      = useState(()=> loadUser() ? loadRecs(loadUser()) : []);
  const [tab,       setTab]       = useState("add");
  const [ym,        setYm]        = useState(nowYM);
  const [rMode,     setRMode]     = useState("month");
  const [range,     setRange]     = useState(()=>{ const n=new Date(); return {from:`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}-01`,to:today()}; });
  const [checked,   setChecked]   = useState(()=> new Set(loadUser() ? loadCats(loadUser()).map(c=>c.key) : DEFAULT_CATS.map(c=>c.key)));
  const [syncing,   setSyncing]   = useState(false);
  const [syncOk,    setSyncOk]    = useState(null);
  const [lastSync,  setLastSync]  = useState(null);
  const [toast,     setToast]     = useState(null);
  const [form,      setForm]      = useState({date:today(), cat:"income", rawAmt:"", note:""});
  const [sSign,     setSSign]     = useState(1);
  const [editRec,   setEditRec]   = useState(null);
  const [editForm,  setEditForm]  = useState({date:"",cat:"income",rawAmt:"",note:""});
  const [editSSign, setEditSSign] = useState(1);

  // ── Category Manager state ──
  const [showCatMgr, setShowCatMgr] = useState(false);
  const [editCat,    setEditCat]    = useState(null); // null | "new" | cat object
  const [catForm,    setCatForm]    = useState({label:"",icon:"💡",color:PALETTE[0],sign:-1,signed:false,desc:""});

  const showToast = useCallback((msg,type="ok")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),2500); },[]);

  useEffect(()=>{ if(user){ saveRecs(user,recs); } },[recs,user]);
  useEffect(()=>{ if(user){ saveCats(user,cats); } },[cats,user]);

  const catMap = useMemo(()=> Object.fromEntries(cats.map(c=>[c.key,c])), [cats]);
  const allKeys = useMemo(()=> cats.map(c=>c.key), [cats]);

  // PWA: inject manifest dynamically
  useEffect(()=>{
    if(document.getElementById("pwa-manifest")) return;
    const manifest = { name:"財務記帳", short_name:"記帳", start_url:"/", display:"standalone", background_color:"#04060d", theme_color:"#04060d", icons:[{src:"https://fav.farm/💹",sizes:"192x192",type:"image/png"}] };
    const blob = new Blob([JSON.stringify(manifest)],{type:"application/json"});
    const url = URL.createObjectURL(blob);
    const link = document.createElement("link"); link.id="pwa-manifest"; link.rel="manifest"; link.href=url;
    document.head.appendChild(link);
    const meta = document.createElement("meta"); meta.name="apple-mobile-web-app-capable"; meta.content="yes"; document.head.appendChild(meta);
    const meta2 = document.createElement("meta"); meta2.name="apple-mobile-web-app-status-bar-style"; meta2.content="black-translucent"; document.head.appendChild(meta2);
    const meta3 = document.createElement("meta"); meta3.name="apple-mobile-web-app-title"; meta3.content="財務記帳"; document.head.appendChild(meta3);
    const meta4 = document.createElement("meta"); meta4.name="theme-color"; meta4.content="#04060d"; document.head.appendChild(meta4);
  },[]);

  // Login
  function doLogin() {
    const name=nameInput.trim();
    if(!name){ showToast("⚠️ 請輸入你的名字","warn"); return; }
    saveUser(name); setUser(name);
    const local=loadRecs(name); setRecs(local);
    const userCats=loadCats(name); setCats(userCats);
    setChecked(new Set(userCats.map(c=>c.key)));
    setForm(f=>({...f,cat:userCats[0]?.key||"income"}));
    // Background sync
    (async()=>{
      setSyncing(true);
      const cloud=await pullCloud(name);
      setSyncing(false);
      if(cloud){
        const ids=new Set(cloud.map(r=>String(r.id)));
        const lo=local.filter(r=>!ids.has(String(r.id)));
        const merged=[...cloud,...lo].sort((a,b)=>String(b.date).localeCompare(String(a.date)));
        setRecs(merged); saveRecs(name,merged);
        setSyncOk(true); setLastSync(new Date().toISOString());
        showToast(`👋 歡迎，${name}！`);
      } else { setSyncOk(false); showToast(`👋 ${name}，本機模式`); }
    })();
  }

  function doLogout(){
    if(!window.confirm(`確定登出「${user}」？`))return;
    saveUser(""); setUser(""); setRecs([]); setCats(DEFAULT_CATS); setNameInput(""); setSyncOk(null); setLastSync(null);
  }

  async function doPush(){ setSyncing(true); const ok=await pushCloud(user,recs); setSyncing(false); if(ok){setSyncOk(true);setLastSync(new Date().toISOString());showToast("☁️ 已推送到雲端！");}else{setSyncOk(false);showToast("❌ 同步失敗","err");} }
  async function doPull(){ setSyncing(true); const cloud=await pullCloud(user); setSyncing(false); if(cloud){setRecs(cloud.sort((a,b)=>String(b.date).localeCompare(String(a.date))));setSyncOk(true);setLastSync(new Date().toISOString());showToast("⬇️ 已載入最新資料！");}else{setSyncOk(false);showToast("❌ 載入失敗","err");} }

  const allYMs = useMemo(()=>{ const s=new Set(recs.map(r=>r.date.slice(0,7))); s.add(nowYM()); return [...s].sort().reverse(); },[recs]);
  const datFlt = useMemo(()=>{ if(rMode==="month")return recs.filter(r=>r.date.slice(0,7)===ym); return recs.filter(r=>r.date>=range.from&&r.date<=range.to); },[recs,rMode,ym,range]);
  const fltRecs= useMemo(()=> datFlt.filter(r=>checked.has(r.cat)),[datFlt,checked]);

  const sum = useMemo(()=>{
    let tin=0, tout=0;
    const by={};
    allKeys.forEach(k=>{ by[k]=0; });
    fltRecs.forEach(r=>{
      const a=Number(r.amt);
      const cat=catMap[r.cat];
      if(!(r.cat in by)) by[r.cat]=0;
      if(!cat){
        if(a>=0) tin+=a; else tout+=Math.abs(a);
        by[r.cat]+=a; return;
      }
      if(cat.signed){
        if(a>=0){ tin+=a; by[r.cat]+=a; }
        else { tout+=Math.abs(a); by[r.cat]+=a; }
      } else if(cat.sign===1){
        tin+=a; by[r.cat]+=a;
      } else {
        tout+=Math.abs(a); by[r.cat]+=Math.abs(a);
      }
    });
    return {...by, totalIn:Math.round(tin), totalOut:Math.round(tout), net:Math.round(tin-tout)};
  },[fltRecs, catMap, allKeys]);

  const trend = useMemo(()=>[...allYMs].reverse().slice(-6).map(m=>{
    const r=recs.filter(x=>x.date.slice(0,7)===m&&checked.has(x.cat));
    let inn=0,out=0;
    r.forEach(x=>{ const a=Number(x.amt),cat=catMap[x.cat]; if(!cat)return;
      if(cat.signed){if(a>=0)inn+=a;else out+=Math.abs(a);}
      else if(cat.sign===1)inn+=a; else out+=a; });
    return {m:`${parseInt(m.split("-")[1])}月`,inn:Math.round(inn),out:Math.round(out),net:Math.round(inn-out)};
  }),[recs,allYMs,checked,catMap]);

  // 圓餅圖：收入類別用綠色系，支出類別用原色，signed類別分拆成獲利/虧損兩格
  const pie = useMemo(()=>{
    const result=[];
    [...checked].forEach(k=>{
      const c=catMap[k]; if(!c) return;
      const raw=sum[k]||0;
      if(c.signed){
        // 分拆：獲利部分 & 虧損部分
        const earn=fltRecs.filter(r=>r.cat===k&&Number(r.amt)>=0).reduce((a,r)=>a+Number(r.amt),0);
        const loss=fltRecs.filter(r=>r.cat===k&&Number(r.amt)<0).reduce((a,r)=>a+Math.abs(Number(r.amt)),0);
        if(earn>0) result.push({name:`${c.label}獲利`,val:Math.round(earn),color:c.color});
        if(loss>0) result.push({name:`${c.label}虧損`,val:Math.round(loss),color:"#f87171"});
      } else {
        const val=Math.round(Math.abs(raw));
        if(val>0) result.push({name:c.label,val,color:c.color});
      }
    });
    return result;
  },[sum,checked,catMap,fltRecs]);

  async function addRec(){
    const n=parseFloat(form.rawAmt);
    if(!form.rawAmt||isNaN(n)||n<=0){showToast("⚠️ 請輸入正確金額","warn");return;}
    const cat=catMap[form.cat]; if(!cat)return;
    const amt=cat.signed?n*sSign:n;
    const r={id:Date.now(),date:form.date,cat:form.cat,amt,note:form.note.trim()};
    const upd=[r,...recs].sort((a,b)=>String(b.date).localeCompare(String(a.date)));
    setRecs(upd); setForm(f=>({...f,rawAmt:"",note:""}));
    showToast(`✅ 已新增：${cat.label}`);
    setSyncing(true); const ok=await pushCloud(user,upd); setSyncing(false);
    if(ok){setSyncOk(true);setLastSync(new Date().toISOString());}
    else{setSyncOk(false);showToast("⚠️ 本機已存，雲端同步失敗","warn");}
  }

  async function delRec(id){
    if(!window.confirm("確定刪除？"))return;
    const upd=recs.filter(r=>r.id!==id); setRecs(upd);
    setSyncing(true); await pushCloud(user,upd); setSyncing(false); showToast("🗑 已刪除");
  }

  function openEdit(r){
    const a=Number(r.amt),cat=catMap[r.cat];
    setEditForm({date:r.date,cat:r.cat,rawAmt:String(Math.abs(a)),note:r.note||""});
    setEditSSign(cat?.signed?(a>=0?1:-1):1); setEditRec(r);
  }

  async function saveEdit(){
    const n=parseFloat(editForm.rawAmt);
    if(!editForm.rawAmt||isNaN(n)||n<=0){showToast("⚠️ 請輸入正確金額","warn");return;}
    const cat=catMap[editForm.cat]; if(!cat)return;
    const amt=cat.signed?n*editSSign:n;
    const upd=recs.map(r=>r.id===editRec.id?{...r,date:editForm.date,cat:editForm.cat,amt,note:editForm.note.trim()}:r).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
    setRecs(upd); setEditRec(null); showToast("✏️ 已儲存修改");
    setSyncing(true); const ok=await pushCloud(user,upd); setSyncing(false);
    if(ok){setSyncOk(true);setLastSync(new Date().toISOString());}
    else showToast("⚠️ 本機已存，雲端同步失敗","warn");
  }

  function exportXlsx(){
    const data=recs.map(r=>{ const c=catMap[r.cat]; const a=Number(r.amt); return {日期:r.date,類別:c?.label||r.cat,金額:Math.round(a),備註:r.note||""}; });
    const ws=XLSX.utils.json_to_sheet(data); ws["!cols"]=[{wch:12},{wch:14},{wch:14},{wch:28}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"記帳明細");
    XLSX.writeFile(wb,`財務_${new Date().toLocaleDateString("zh-TW").replace(/\//g,"-")}.xlsx`);
    showToast("📊 已匯出");
  }

  // ── Category Manager ──
  function openNewCat(){ setCatForm({label:"",icon:"💡",color:PALETTE[Math.floor(Math.random()*PALETTE.length)],sign:-1,signed:false,desc:""}); setEditCat("new"); }
  function openEditCat(c){ setCatForm({label:c.label,icon:c.icon,color:c.color,sign:c.sign,signed:c.signed,desc:c.desc||""}); setEditCat(c); }
  function saveCatForm(){
    if(!catForm.label.trim()){showToast("⚠️ 請輸入類別名稱","warn");return;}
    if(editCat==="new"){
      const key=`cat_${Date.now()}`;
      const nc={key,label:catForm.label.trim(),icon:catForm.icon,color:catForm.color,sign:catForm.signed?0:catForm.sign,signed:catForm.signed,desc:catForm.desc};
      const newCats=[...cats,nc]; setCats(newCats);
      setChecked(prev=>new Set([...prev,key]));
      showToast(`✅ 已新增「${nc.label}」`);
    } else {
      const newCats=cats.map(c=>c.key===editCat.key?{...c,...catForm,sign:catForm.signed?0:catForm.sign}:c);
      setCats(newCats); showToast("✏️ 已儲存");
    }
    setEditCat(null);
  }
  function deleteCat(key){
    if(cats.length<=1){showToast("⚠️ 至少需要一個類別","warn");return;}
    if(!window.confirm("確定刪除此類別？（已有的記錄不受影響）"))return;
    setCats(prev=>prev.filter(c=>c.key!==key));
    setChecked(prev=>{ const s=new Set(prev); s.delete(key); return s; });
    showToast("🗑 已刪除類別");
  }

  function toggleCat(k){ setChecked(prev=>{ const s=new Set(prev); if(s.has(k)){if(s.size===1)return s; s.delete(k);}else s.add(k); return s; }); }

  const isPos=sum.net>=0;
  const netPct=sum.totalOut>0?Math.abs((sum.net/sum.totalOut)*100).toFixed(1):"0.0";
  const totalFlow=(sum.totalIn+sum.totalOut)||1;
  const lastSyncStr=lastSync?new Date(lastSync).toLocaleTimeString("zh-TW",{hour:"2-digit",minute:"2-digit"}):null;
  const tt={background:"rgba(10,15,26,.97)",border:"1px solid rgba(255,255,255,.12)",borderRadius:12,fontSize:13,color:"#e2e8f0",padding:"10px 16px",boxShadow:"0 8px 32px rgba(0,0,0,.6)"};
  const curCat=catMap[form.cat];

  const MNav=()=>(
    <div className="mnav">
      <button onClick={()=>{const i=allYMs.indexOf(ym);if(i<allYMs.length-1)setYm(allYMs[i+1]);}} disabled={allYMs.indexOf(ym)>=allYMs.length-1}>‹</button>
      <div style={{flex:1,textAlign:"center",fontSize:14,fontWeight:600,color:"#e2e8f0",minWidth:100}}>{ymLbl(ym)}</div>
      <button onClick={()=>{const i=allYMs.indexOf(ym);if(i>0)setYm(allYMs[i-1]);}} disabled={allYMs.indexOf(ym)<=0}>›</button>
    </div>
  );

  const FilterBar=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        <button className={`f-pill${checked.size===allKeys.length?" on":""}`} onClick={()=>setChecked(new Set(allKeys))}>◎ 全部</button>
        {cats.filter(c=>c.sign===1||c.signed).map(c=>{
          const only=checked.size===1&&checked.has(c.key);
          return <button key={c.key} className={`f-pill${only?" on":""}`} onClick={()=>setChecked(new Set([c.key]))} style={only?{borderColor:c.color+"88",background:c.color+"15",color:c.color}:{}}>{c.icon} {c.label}</button>;
        })}
        {cats.filter(c=>c.sign===-1&&!c.signed).length>1&&(
          <button className={`f-pill${cats.filter(c=>c.sign===-1&&!c.signed).every(c=>checked.has(c.key))&&checked.size===cats.filter(c=>c.sign===-1&&!c.signed).length?" on":""}`}
            onClick={()=>setChecked(new Set(cats.filter(c=>c.sign===-1&&!c.signed).map(c=>c.key)))}>🧾 支出</button>
        )}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:7}}>
        {cats.map(c=>{
          const on=checked.has(c.key);
          return (
            <button key={c.key} onClick={()=>toggleCat(c.key)}
              style={{border:`1.5px solid ${on?c.color:"rgba(255,255,255,.07)"}`,background:on?c.color+"15":"rgba(0,0,0,.3)",
                color:on?c.color:"#334155",borderRadius:12,padding:"8px 6px",fontSize:11,fontWeight:600,
                cursor:"pointer",transition:"all .15s",display:"flex",alignItems:"center",gap:5}}>
              <span style={{fontSize:13,opacity:on?1:.3}}>{on?"☑":"☐"}</span>
              <span>{c.icon} {c.label}</span>
            </button>
          );
        })}
      </div>
      <div style={{fontSize:10,color:"#1e3a5f",fontWeight:500}}>已選 {checked.size} 個類別</div>
    </div>
  );

  const RangeBar=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:12}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
        <div className="pill">
          <button className={rMode==="month"?"on":""} onClick={()=>setRMode("month")}>月份</button>
          <button className={rMode==="custom"?"on":""} onClick={()=>setRMode("custom")}>自訂</button>
        </div>
        {rMode==="month"&&<MNav/>}
      </div>
      {rMode==="custom"&&(
        <div style={{display:"flex",gap:10}}>
          <div style={{flex:1}}><div style={{fontSize:10,color:"#334155",letterSpacing:".1em",marginBottom:6,fontWeight:700}}>開始</div><input type="date" className="fi" style={{fontSize:13}} value={range.from} onChange={e=>setRange(r=>({...r,from:e.target.value}))}/></div>
          <div style={{flex:1}}><div style={{fontSize:10,color:"#334155",letterSpacing:".1em",marginBottom:6,fontWeight:700}}>結束</div><input type="date" className="fi" style={{fontSize:13}} value={range.to} onChange={e=>setRange(r=>({...r,to:e.target.value}))}/></div>
        </div>
      )}
    </div>
  );

  // ── Login Screen ──
  if(!user) return (
    <div style={{fontFamily:"'Noto Sans TC',sans-serif",minHeight:"100vh",background:"#04060d",display:"flex",alignItems:"center",justifyContent:"center",padding:"24px",position:"relative",overflow:"hidden"}}>
      <style>{CSS}</style>
      {/* bg glow */}
      <div style={{position:"absolute",top:"20%",left:"50%",transform:"translateX(-50%)",width:400,height:400,background:"radial-gradient(circle,rgba(74,222,128,.06) 0%,transparent 70%)",pointerEvents:"none"}}/>
      {toast&&<div className={`toast${toast.type==="err"?" err":toast.type==="warn"?" warn":""}`}>{toast.msg}</div>}
      <div style={{width:"100%",maxWidth:360,display:"flex",flexDirection:"column",alignItems:"center"}}>
        <div style={{width:76,height:76,background:"linear-gradient(135deg,#4ade80,#38bdf8)",borderRadius:24,display:"flex",alignItems:"center",justifyContent:"center",fontSize:38,marginBottom:20,boxShadow:"0 8px 40px rgba(74,222,128,.25)"}}>💹</div>
        <div style={{fontSize:26,fontWeight:700,background:"linear-gradient(90deg,#4ade80,#38bdf8)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",marginBottom:6}}>財務記帳</div>
        <div style={{fontSize:11,color:"#1e3a5f",letterSpacing:".18em",fontWeight:600,marginBottom:40}}>PERSONAL FINANCE</div>
        <div className="gc" style={{width:"100%",padding:"28px 24px"}}>
          <div style={{fontSize:13,fontWeight:600,color:"#64748b",textAlign:"center",marginBottom:4}}>輸入你的名字開始使用</div>
          <div style={{fontSize:11,color:"#1e3a5f",textAlign:"center",marginBottom:18}}>每個人資料獨立，互不影響</div>
          <input className="fi" placeholder="例：小明、爸爸、Daniel..." value={nameInput}
            style={{fontSize:17,textAlign:"center",letterSpacing:".05em",marginBottom:16}}
            onChange={e=>setNameInput(e.target.value)}
            onKeyDown={e=>{if(e.key==="Enter")doLogin()}} autoFocus/>
          <button className="btn-primary" onClick={doLogin} disabled={syncing}
            style={{background:"linear-gradient(135deg,#16a34a,#0284c7)",color:"#fff",boxShadow:"0 4px 24px rgba(74,222,128,.2)"}}>
            {syncing?"載入中...":"▶ 進入我的帳本"}
          </button>
        </div>
        <div style={{marginTop:20,fontSize:11,color:"#1e3a5f",textAlign:"center",lineHeight:2}}>
          同一名字 → 跨裝置同步 ☁️<br/>不同名字 → 獨立帳本 🔒
        </div>
      </div>
    </div>
  );

  return (
    <div style={{fontFamily:"'Noto Sans TC',sans-serif",minHeight:"100vh",background:"#04060d",color:"#e2e8f0",position:"relative"}}>
      <style>{CSS}</style>
      {/* Background glow */}
      <div style={{position:"fixed",top:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:600,height:300,background:"radial-gradient(ellipse at top,rgba(74,222,128,.04) 0%,transparent 70%)",pointerEvents:"none",zIndex:0}}/>

      {toast&&<div className={`toast${toast.type==="err"?" err":toast.type==="warn"?" warn":""}`}>{toast.msg}</div>}

      {/* ── Header ── */}
      <div style={{background:"rgba(4,6,13,.92)",backdropFilter:"blur(28px)",WebkitBackdropFilter:"blur(28px)",borderBottom:"1px solid rgba(255,255,255,.05)",padding:"13px 18px 10px",position:"sticky",top:0,zIndex:50}}>
        <div style={{maxWidth:520,margin:"0 auto"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:36,height:36,background:"linear-gradient(135deg,#4ade80,#38bdf8)",borderRadius:11,display:"flex",alignItems:"center",justifyContent:"center",fontSize:19,flexShrink:0,boxShadow:"0 4px 14px rgba(74,222,128,.25)"}}>💹</div>
              <div>
                <div style={{display:"flex",alignItems:"center",gap:7}}>
                  <span style={{fontSize:15,fontWeight:700,background:"linear-gradient(90deg,#4ade80,#38bdf8)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>財務記帳</span>
                  <span style={{background:"rgba(74,222,128,.12)",border:"1px solid rgba(74,222,128,.25)",borderRadius:99,padding:"2px 9px",fontSize:10,fontWeight:700,color:"#4ade80"}}>👤 {user}</span>
                </div>
                <div style={{fontSize:9,color:"#1e293b",letterSpacing:".14em",fontWeight:600}}>PERSONAL FINANCE</div>
              </div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:5}}>
              {lastSyncStr&&<div style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:6,height:6,borderRadius:99,background:syncOk?"#4ade80":syncOk===false?"#f87171":"#fbbf24"}}/><span style={{fontSize:10,color:"#334155"}}>{lastSyncStr}</span></div>}
              <button className="btn-sm" onClick={doPull} disabled={syncing} style={{color:"#38bdf8"}}><span className={syncing?"spin":""}>{syncing?"⟳":"⬇"}</span><span>{syncing?"...":"拉取"}</span></button>
              <button className="btn-sm" onClick={doPush} disabled={syncing} style={{color:"#4ade80"}}><span>⬆</span><span>推送</span></button>
              <button className="btn-sm" onClick={doLogout} style={{color:"#f87171"}}><span>⏏</span></button>
            </div>
          </div>
          <div className="tabs">
            {[["add","📝 記帳","t0"],["summary","📊 統計","t1"],["chart","📈 圖表","t2"],["records","🗂 明細","t3"],["settings","⚙️ 類別","t4"]].map(([v,l,c])=>(
              <button key={v} className={`tab${tab===v?` on ${c}`:""}`} onClick={()=>setTab(v)}>{l}</button>
            ))}
          </div>
        </div>
      </div>

      <div style={{maxWidth:520,margin:"0 auto",padding:"18px 14px 60px",display:"flex",flexDirection:"column",gap:14,position:"relative",zIndex:1}}>

        {/* ══ ADD ══ */}
        {tab==="add"&&<>
          <div className="gc">
            <div style={{fontSize:15,fontWeight:700,marginBottom:20,background:"linear-gradient(90deg,#e2e8f0,#64748b)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>新增一筆記錄</div>
            <div style={{display:"flex",flexDirection:"column",gap:15}}>
              <div><div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>📅 日期</div><input type="date" className="fi" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/></div>
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>🏷 類別</div>
                <div className="cat-grid">
                  {cats.map(c=>(
                    <button key={c.key} className="cat-btn" style={form.cat===c.key?{borderColor:c.color,background:c.color+"14",color:c.color,transform:"scale(1.04)"}:{}} onClick={()=>setForm(f=>({...f,cat:c.key}))}>
                      <span className="cat-ico">{c.icon}</span>
                      <span style={{fontSize:11,fontWeight:700}}>{c.label}</span>
                      <span style={{fontSize:9,color:form.cat===c.key?c.color+"99":"#1e3a5f",lineHeight:1.3}}>{c.desc}</span>
                    </button>
                  ))}
                </div>
              </div>
              {curCat?.signed&&(
                <div>
                  <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>📌 方向</div>
                  <div className="sign-toggle">
                    <button className="sign-btn" style={sSign===1?{background:"linear-gradient(135deg,#0c4a6e,#14532d)",color:"#4ade80"}:{color:"#334155"}} onClick={()=>setSSign(1)}><span style={{fontSize:18}}>📈</span><div style={{textAlign:"left"}}><div style={{fontSize:13}}>獲利 / 入金</div><div style={{fontSize:10,opacity:.6}}>賺錢、存入</div></div></button>
                    <div style={{width:1,background:"rgba(255,255,255,.06)"}}/>
                    <button className="sign-btn" style={sSign===-1?{background:"linear-gradient(135deg,#7c1d1d,#1c1917)",color:"#f87171"}:{color:"#334155"}} onClick={()=>setSSign(-1)}><span style={{fontSize:18}}>📉</span><div style={{textAlign:"left"}}><div style={{fontSize:13}}>虧損 / 出金</div><div style={{fontSize:10,opacity:.6}}>賠錢、取出</div></div></button>
                  </div>
                </div>
              )}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>
                  💵 金額（元）{curCat?.signed&&<span style={{marginLeft:8,fontSize:10,color:sSign===1?"#4ade80":"#f87171",fontWeight:700}}>{sSign===1?"＋":"－"}</span>}
                </div>
                <div style={{position:"relative"}}>
                  {curCat?.signed&&<div style={{position:"absolute",left:15,top:"50%",transform:"translateY(-50%)",fontSize:22,fontWeight:700,color:sSign===1?"#4ade80":"#f87171",fontFamily:"'DM Mono',monospace",pointerEvents:"none",zIndex:1}}>{sSign===1?"+":"−"}</div>}
                  <input type="number" className="fi mono" placeholder="0" value={form.rawAmt} min="0" inputMode="decimal"
                    style={{fontSize:24,letterSpacing:"-.03em",color:curCat?.color||"#e2e8f0",paddingLeft:curCat?.signed?"40px":"15px"}}
                    onChange={e=>setForm(f=>({...f,rawAmt:e.target.value}))} onKeyDown={e=>{if(e.key==="Enter")addRec()}}/>
                </div>
              </div>
              <div><div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>📝 備註（選填）</div><input className="fi" placeholder="備註..." value={form.note} onChange={e=>setForm(f=>({...f,note:e.target.value}))} onKeyDown={e=>{if(e.key==="Enter")addRec()}}/></div>
              <button className="btn-primary" onClick={addRec} disabled={syncing}
                style={{background:curCat?`linear-gradient(135deg,${curCat.color}cc,${curCat.color}88)`:"linear-gradient(135deg,#4ade80,#38bdf8)",color:"#fff",boxShadow:`0 4px 20px ${curCat?.color||"#4ade80"}22`,marginTop:4}}>
                {syncing?"同步中...":curCat?.signed?(sSign===1?`＋ 新增 ${curCat.label} 獲利/入金`:`－ 新增 ${curCat.label} 虧損/出金`):`＋ 新增${curCat?.label}`}
              </button>
            </div>
          </div>
          {recs.slice(0,5).length>0&&(
            <div className="gc">
              <div style={{fontSize:11,color:"#1e3a5f",letterSpacing:".1em",fontWeight:700,marginBottom:14}}>最近記錄</div>
              {recs.slice(0,5).map(r=>{ const c=catMap[r.cat]; const a=Number(r.amt); const isIn=c?.sign===1||(c?.signed&&a>=0); return (
                <div key={r.id} className="rr">
                  <div style={{width:40,height:40,background:(c?.color||"#666")+"15",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:19,flexShrink:0}}>{c?.icon||"?"}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:500,color:"#cbd5e1",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c?.signed?(a>=0?"📈 獲利/入金":"📉 虧損/出金"):c?.label||r.cat}{r.note?` · ${r.note}`:""}</div>
                    <div style={{fontSize:11,color:"#1e3a5f",marginTop:3}}>{r.date}</div>
                  </div>
                  <span className="mono" style={{fontSize:15,fontWeight:600,color:isIn?"#4ade80":"#f87171",flexShrink:0}}>{isIn?"+":"-"}{fmtN(Math.abs(a))}</span>
                </div>
              ); })}
            </div>
          )}
          <div style={{background:"rgba(0,0,0,.3)",border:"1px solid rgba(255,255,255,.05)",borderRadius:16,padding:"13px 16px",display:"flex",alignItems:"center",gap:12}}>
            <div style={{width:34,height:34,background:"rgba(56,189,248,.1)",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>☁️</div>
            <div style={{flex:1}}>
              <div style={{fontSize:12,fontWeight:600,color:"#38bdf8"}}>雲端同步</div>
              <div style={{fontSize:11,color:"#1e3a5f",marginTop:2}}>{syncing?"同步中...":syncOk===true?`✓ ${lastSyncStr||"剛才"}`:syncOk===false?"未連線，資料暫存本機":"啟動中..."}</div>
            </div>
            <div style={{width:9,height:9,borderRadius:99,background:syncing?"#fbbf24":syncOk?"#4ade80":"#f87171"}}/>
          </div>
        </>}

        {/* ══ SUMMARY ══ */}
        {tab==="summary"&&<>
          <div className="gc" style={{padding:"16px 18px",gap:12,display:"flex",flexDirection:"column"}}>
            <RangeBar/>
            <div style={{borderTop:"1px solid rgba(255,255,255,.05)",paddingTop:12}}><div style={{fontSize:10,color:"#334155",letterSpacing:".1em",marginBottom:10,fontWeight:700}}>篩選類別</div><FilterBar/></div>
          </div>

          <div className={`${isPos?"glow-g":"glow-r"}`} style={{background:isPos?"linear-gradient(150deg,rgba(5,25,15,.9),rgba(4,8,22,.9))":"linear-gradient(150deg,rgba(25,5,10,.9),rgba(4,8,22,.9))",border:`1px solid ${isPos?"rgba(74,222,128,.15)":"rgba(248,113,113,.15)"}`,borderRadius:22,padding:"28px 22px",textAlign:"center"}}>
            <div style={{fontSize:10,letterSpacing:".2em",fontWeight:700,marginBottom:12,color:isPos?"rgba(74,222,128,.5)":"rgba(248,113,113,.5)"}}>已選 {checked.size} 類 · 淨損益</div>
            <div className="mono" style={{fontSize:52,fontWeight:500,color:isPos?"#4ade80":"#f87171",letterSpacing:"-.04em",lineHeight:1}}>{isPos?"+":""}{fmtN(sum.net)}<span style={{fontSize:16,fontWeight:400,marginLeft:8,opacity:.5}}>元</span></div>
            <div style={{marginTop:16}}><span style={{background:isPos?"rgba(74,222,128,.1)":"rgba(248,113,113,.1)",border:`1px solid ${isPos?"rgba(74,222,128,.25)":"rgba(248,113,113,.25)"}`,color:isPos?"#4ade80":"#f87171",borderRadius:99,padding:"6px 20px",fontSize:13,fontWeight:700}}>{isPos?"📈 結餘":"📉 透支"} {netPct}%</span></div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              {[{l:"總收入",v:sum.totalIn,c:"#4ade80",s:"+"},{l:"總支出",v:sum.totalOut,c:"#f87171",s:"-"},{l:"筆數",v:fltRecs.length,c:"#38bdf8",s:""}].map(({l,v,c,s})=>(
                <div key={l} style={{flex:1,background:c+"0c",border:`1px solid ${c}18`,borderRadius:14,padding:"14px 8px"}}>
                  <div style={{fontSize:9,color:c+"77",letterSpacing:".1em",fontWeight:700}}>{l}</div>
                  <div className="mono" style={{fontSize:16,color:c,marginTop:6}}>{s}{l==="筆數"?v:fmtN(v)}</div>
                </div>
              ))}
            </div>
          </div>

          <div className="gc">
            <div style={{fontSize:14,fontWeight:700,marginBottom:18,color:"#f1f5f9"}}>各類別明細</div>
            {cats.filter(c=>checked.has(c.key)).map(c=>{
              const raw=sum[c.key]||0; const val=Math.abs(raw); const pct=((val/totalFlow)*100).toFixed(1);
              const isIn=c.sign===1||(c.signed&&raw>0);
              return (
                <div key={c.key} style={{marginBottom:18}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{width:34,height:34,background:c.color+"15",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17}}>{c.icon}</div>
                      <div><div style={{fontSize:13,fontWeight:600,color:"#cbd5e1"}}>{c.label}</div>{c.signed&&<div style={{fontSize:10,color:"#334155"}}>{raw>=0?"獲利/入金":"虧損/出金"}</div>}</div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:10}}>
                      <span style={{fontSize:10,color:"#1e3a5f",fontWeight:600,background:"rgba(0,0,0,.4)",borderRadius:6,padding:"2px 7px"}}>{pct}%</span>
                      <span className="mono" style={{fontSize:15,fontWeight:600,color:c.color,minWidth:88,textAlign:"right"}}>{isIn?"+":"-"}{fmtN(val)}</span>
                    </div>
                  </div>
                  <div className="bar-t"><div className="bar-f" style={{width:`${pct}%`,background:`linear-gradient(90deg,${c.color}44,${c.color})`}}/></div>
                </div>
              );
            })}
          </div>
          <div style={{display:"flex",justifyContent:"flex-end"}}><button className="btn-ghost" onClick={exportXlsx} style={{color:"#4ade80",borderColor:"rgba(74,222,128,.2)"}}>📊 匯出 Excel</button></div>
        </>}

        {/* ══ CHART ══ */}
        {tab==="chart"&&<>
          <div className="gc" style={{padding:"16px 18px",gap:12,display:"flex",flexDirection:"column"}}>
            <div style={{fontSize:10,color:"#334155",letterSpacing:".1em",marginBottom:8,fontWeight:700}}>篩選類別</div><FilterBar/>
            <div style={{borderTop:"1px solid rgba(255,255,255,.05)",paddingTop:12}}><RangeBar/></div>
          </div>
          <div className="gc">
            <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",marginBottom:4}}>近 6 個月趨勢</div>
            <div style={{fontSize:11,color:"#1e3a5f",marginBottom:18}}>已選 {checked.size} 類 · 收支比較</div>
            {trend.length===0?<div style={{textAlign:"center",color:"#1e3a5f",padding:"40px 0"}}>尚無資料</div>:
              <ResponsiveContainer width="100%" height={210}><BarChart data={trend} barGap={3} barCategoryGap="32%">
                <XAxis dataKey="m" tick={{fill:"#1e3a5f",fontSize:11}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:"#1e3a5f",fontSize:10}} axisLine={false} tickLine={false} width={50} tickFormatter={v=>v>=10000?`${(v/10000).toFixed(0)}萬`:String(v)}/>
                <Tooltip contentStyle={tt} labelStyle={{color:"#64748b"}} formatter={(v,n)=>[`${v.toLocaleString()} 元`,n==="inn"?"收入":n==="out"?"支出":"淨損益"]}/>
                <Legend formatter={v=>v==="inn"?"收入":v==="out"?"支出":"淨損益"} wrapperStyle={{fontSize:11,color:"#475569"}}/>
                <Bar dataKey="inn" fill="#4ade80" radius={[5,5,0,0]}/><Bar dataKey="out" fill="#f87171" radius={[5,5,0,0]}/><Bar dataKey="net" fill="#38bdf8" radius={[5,5,0,0]}/>
              </BarChart></ResponsiveContainer>
            }
          </div>
          <div className="gc">
            <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",marginBottom:4}}>類別結構</div>
            <div style={{fontSize:11,color:"#1e3a5f",marginBottom:14}}>投資類別會分拆獲利／虧損</div>
            {pie.length===0?<div style={{textAlign:"center",color:"#1e3a5f",padding:"40px 0"}}>此期間無資料</div>:
              <ResponsiveContainer width="100%" height={260}>
                <PieChart margin={{top:10,right:10,bottom:10,left:10}}>
                  <Pie data={pie} cx="50%" cy="50%" outerRadius={85} innerRadius={38} dataKey="val" paddingAngle={3}
                    label={({name,percent,x,y,cx:cx2})=>(
                      <text x={x} y={y} textAnchor={x>cx2?"start":"end"} dominantBaseline="central" fill="#64748b" fontSize={10} fontFamily="Noto Sans TC">
                        {name} {(percent*100).toFixed(0)}%
                      </text>
                    )}
                    labelLine={{stroke:"#1e3a5f",strokeWidth:1}}>
                    {pie.map((d,i)=><Cell key={i} fill={d.color}/>)}
                  </Pie>
                  <Tooltip
                    contentStyle={{background:"rgba(10,15,26,.97)",border:"1px solid rgba(255,255,255,.12)",borderRadius:12,fontSize:13,color:"#e2e8f0",padding:"10px 16px"}}
                    itemStyle={{color:"#e2e8f0"}}
                    labelStyle={{display:"none"}}
                    formatter={(v,name)=>[`${v.toLocaleString()} 元`, name]}
                  />
                </PieChart>
              </ResponsiveContainer>
            }
          </div>
          <div className="gc">
            <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",marginBottom:4}}>淨損益走勢</div>
            <div style={{fontSize:11,color:"#1e3a5f",marginBottom:18}}>每月結算</div>
            {trend.length<2?<div style={{textAlign:"center",color:"#1e3a5f",padding:"30px 0"}}>至少需要 2 個月資料</div>:
              <ResponsiveContainer width="100%" height={170}><LineChart data={trend}>
                <XAxis dataKey="m" tick={{fill:"#1e3a5f",fontSize:11}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:"#1e3a5f",fontSize:10}} axisLine={false} tickLine={false} width={50} tickFormatter={v=>v>=10000?`${(v/10000).toFixed(0)}萬`:String(v)}/>
                <Tooltip contentStyle={tt} formatter={v=>[`${v.toLocaleString()} 元`,"淨損益"]}/>
                <Line type="monotone" dataKey="net" stroke="#38bdf8" strokeWidth={2.5} dot={{fill:"#38bdf8",r:5,strokeWidth:0}} activeDot={{r:7}}/>
              </LineChart></ResponsiveContainer>
            }
          </div>
        </>}

        {/* ══ RECORDS ══ */}
        {tab==="records"&&(
          <div className="gc">
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
              <div><div style={{fontSize:15,fontWeight:700,color:"#f1f5f9"}}>所有記錄</div><div style={{fontSize:11,color:"#1e3a5f",marginTop:2}}>{recs.length} 筆</div></div>
              <button className="btn-ghost" style={{padding:"8px 14px",fontSize:12,color:"#4ade80",borderColor:"rgba(74,222,128,.2)"}} onClick={exportXlsx}>📊 匯出</button>
            </div>
            <div style={{marginBottom:14}}><FilterBar/></div>
            {fltRecs.length===0&&<div style={{textAlign:"center",padding:"48px 0"}}><div style={{fontSize:36,marginBottom:10}}>📭</div><div style={{color:"#334155"}}>此類別無記錄</div></div>}
            {fltRecs.map(r=>{ const c=catMap[r.cat]; const a=Number(r.amt); const isIn=c?.sign===1||(c?.signed&&a>=0); return (
              <div key={r.id} className="rr" style={{cursor:"pointer"}} onClick={()=>openEdit(r)}>
                <div style={{width:40,height:40,background:(c?.color||"#666")+"15",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:19,flexShrink:0}}>{c?.icon||"?"}</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                    <span className="chip" style={{background:(c?.color||"#666")+"20",color:c?.color||"#666"}}>{c?.signed?(a>=0?"📈 獲利":"📉 虧損"):c?.label||r.cat}</span>
                    {r.note&&<span style={{fontSize:11,color:"#334155",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:120}}>{r.note}</span>}
                  </div>
                  <div style={{fontSize:11,color:"#1e293b",marginTop:4,fontWeight:600}}>{r.date}</div>
                </div>
                <span className="mono" style={{fontSize:15,fontWeight:600,color:isIn?"#4ade80":"#f87171",flexShrink:0}}>{isIn?"+":"-"}{fmtN(Math.abs(a))}</span>
                <button className="del" onClick={e=>{e.stopPropagation();delRec(r.id);}}>✕</button>
              </div>
            ); })}
            {recs.length>0&&<button className="btn-danger" onClick={async()=>{ if(!window.confirm("確定清空所有記錄？"))return; setRecs([]); saveRecs(user,[]); setSyncing(true); await pushCloud(user,[]); setSyncing(false); showToast("🗑 已清空"); }}>🗑 清空所有記錄</button>}
          </div>
        )}

        {/* ══ SETTINGS (Category Manager) ══ */}
        {tab==="settings"&&(
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <div className="gc">
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
                <div><div style={{fontSize:15,fontWeight:700,color:"#f1f5f9"}}>類別管理</div><div style={{fontSize:11,color:"#1e3a5f",marginTop:2}}>新增、編輯、刪除類別</div></div>
                <button onClick={openNewCat} style={{background:"linear-gradient(135deg,#14532d,#0c4a6e)",border:"none",color:"#4ade80",borderRadius:11,padding:"9px 16px",fontSize:13,fontWeight:700,cursor:"pointer"}}>＋ 新增</button>
              </div>
              {cats.map(c=>(
                <div key={c.key} className="cm-row">
                  <div style={{width:40,height:40,background:c.color+"18",borderRadius:11,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{c.icon}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:600,color:"#cbd5e1"}}>{c.label}</div>
                    <div style={{fontSize:10,color:"#334155",marginTop:2}}>{c.signed?"正負數輸入（獲利/虧損）":c.sign===1?"收入":"支出"} · {c.desc||"—"}</div>
                  </div>
                  <div style={{display:"flex",gap:6}}>
                    <button onClick={()=>openEditCat(c)} style={{background:"rgba(255,255,255,.05)",border:"1px solid rgba(255,255,255,.08)",color:"#94a3b8",borderRadius:8,padding:"6px 11px",fontSize:12,cursor:"pointer"}}>✏️</button>
                    <button onClick={()=>deleteCat(c.key)} style={{background:"rgba(248,113,113,.08)",border:"1px solid rgba(248,113,113,.15)",color:"#f87171",borderRadius:8,padding:"6px 11px",fontSize:12,cursor:"pointer"}}>🗑</button>
                  </div>
                </div>
              ))}
            </div>
            <div className="gc" style={{background:"rgba(56,189,248,.05)",borderColor:"rgba(56,189,248,.1)"}}>
              <div style={{fontSize:12,fontWeight:600,color:"#38bdf8",marginBottom:8}}>💡 使用說明</div>
              <div style={{fontSize:12,color:"#334155",lineHeight:1.9}}>
                • <b style={{color:"#94a3b8"}}>新增類別</b>：點右上角「＋ 新增」<br/>
                • <b style={{color:"#94a3b8"}}>正負數輸入</b>：適合股票、ETF 等有獲利也有虧損的類別<br/>
                • <b style={{color:"#94a3b8"}}>收入 / 支出</b>：單向記錄<br/>
                • 類別刪除後，已有的記錄不受影響
              </div>
            </div>
          </div>
        )}

      </div>

      {/* ══ Edit Modal ══ */}
      {editRec&&(
        <div className="modal-bg" onClick={e=>{if(e.target===e.currentTarget)setEditRec(null);}}>
          <div className="modal-bd" onClick={()=>setEditRec(null)}/>
          <div className="modal-sheet">
            <div className="handle"/>
            <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:20}}>✏️ 修改記錄</div>
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              <div><div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>📅 日期</div><input type="date" className="fi" value={editForm.date} onChange={e=>setEditForm(f=>({...f,date:e.target.value}))}/></div>
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>🏷 類別</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:7}}>
                  {cats.map(c=>(
                    <button key={c.key} style={{border:`1.5px solid ${editForm.cat===c.key?c.color:"rgba(255,255,255,.07)"}`,background:editForm.cat===c.key?c.color+"14":"rgba(0,0,0,.3)",color:editForm.cat===c.key?c.color:"#334155",borderRadius:12,padding:"10px 6px",fontSize:11,fontWeight:600,cursor:"pointer",transition:"all .15s",display:"flex",flexDirection:"column",alignItems:"center",gap:4}} onClick={()=>setEditForm(f=>({...f,cat:c.key}))}>
                      <span style={{fontSize:20}}>{c.icon}</span><span>{c.label}</span>
                    </button>
                  ))}
                </div>
              </div>
              {catMap[editForm.cat]?.signed&&(
                <div>
                  <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>📌 方向</div>
                  <div className="sign-toggle">
                    <button className="sign-btn" style={editSSign===1?{background:"linear-gradient(135deg,#0c4a6e,#14532d)",color:"#4ade80"}:{color:"#334155"}} onClick={()=>setEditSSign(1)}><span>📈</span><span style={{fontSize:13}}>獲利/入金</span></button>
                    <div style={{width:1,background:"rgba(255,255,255,.06)"}}/>
                    <button className="sign-btn" style={editSSign===-1?{background:"linear-gradient(135deg,#7c1d1d,#1c1917)",color:"#f87171"}:{color:"#334155"}} onClick={()=>setEditSSign(-1)}><span>📉</span><span style={{fontSize:13}}>虧損/出金</span></button>
                  </div>
                </div>
              )}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>💵 金額</div>
                <input type="number" className="fi mono" placeholder="0" value={editForm.rawAmt} min="0" inputMode="decimal" style={{fontSize:22,color:catMap[editForm.cat]?.color||"#e2e8f0"}} onChange={e=>setEditForm(f=>({...f,rawAmt:e.target.value}))}/>
              </div>
              <div><div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>📝 備註</div><input className="fi" placeholder="備註..." value={editForm.note} onChange={e=>setEditForm(f=>({...f,note:e.target.value}))} onKeyDown={e=>{if(e.key==="Enter")saveEdit()}}/></div>
              <div style={{display:"flex",gap:10,marginTop:4}}>
                <button onClick={()=>setEditRec(null)} style={{flex:1,background:"none",border:"1.5px solid rgba(255,255,255,.08)",color:"#475569",borderRadius:12,padding:"13px",fontSize:14,fontWeight:600,cursor:"pointer"}}>取消</button>
                <button onClick={saveEdit} disabled={syncing} style={{flex:2,background:"linear-gradient(135deg,#166534,#0c4a6e)",color:"#fff",border:"none",borderRadius:12,padding:"13px",fontSize:15,fontWeight:700,cursor:"pointer",opacity:syncing?.5:1}}>{syncing?"儲存中...":"✓ 儲存修改"}</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ══ Category Edit Modal ══ */}
      {editCat&&(
        <div className="modal-bg" onClick={e=>{if(e.target===e.currentTarget)setEditCat(null);}}>
          <div className="modal-bd" onClick={()=>setEditCat(null)}/>
          <div className="modal-sheet">
            <div className="handle"/>
            <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:20}}>{editCat==="new"?"➕ 新增類別":"✏️ 編輯類別"}</div>
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              <div><div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>名稱</div><input className="fi" placeholder="例：ETF、旅遊..." value={catForm.label} onChange={e=>setCatForm(f=>({...f,label:e.target.value}))}/></div>
              <div><div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>描述（選填）</div><input className="fi" placeholder="簡短說明..." value={catForm.desc} onChange={e=>setCatForm(f=>({...f,desc:e.target.value}))}/></div>
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>類型</div>
                <div style={{display:"flex",gap:8}}>
                  {[{l:"收入",v:1,signed:false},{l:"支出",v:-1,signed:false},{l:"正負數（獲利/虧損）",v:0,signed:true}].map(o=>(
                    <button key={o.l} onClick={()=>setCatForm(f=>({...f,sign:o.v,signed:o.signed}))}
                      style={{flex:1,border:`1.5px solid ${catForm.signed===o.signed&&catForm.sign===o.v?"#4ade80":"rgba(255,255,255,.07)"}`,background:catForm.signed===o.signed&&catForm.sign===o.v?"rgba(74,222,128,.12)":"rgba(0,0,0,.3)",color:catForm.signed===o.signed&&catForm.sign===o.v?"#4ade80":"#475569",borderRadius:10,padding:"9px 6px",fontSize:11,fontWeight:600,cursor:"pointer"}}>
                      {o.l}
                    </button>
                  ))}
                </div>
              </div>
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>圖示</div>
                <div className="icon-grid">
                  {ICONS.map(ic=><button key={ic} className={`icon-btn${catForm.icon===ic?" on":""}`} onClick={()=>setCatForm(f=>({...f,icon:ic}))}>{ic}</button>)}
                </div>
              </div>
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>顏色</div>
                <div className="color-grid">
                  {PALETTE.map(cl=><button key={cl} className={`color-btn${catForm.color===cl?" on":""}`} style={{background:cl}} onClick={()=>setCatForm(f=>({...f,color:cl}))}/>)}
                </div>
              </div>
              <div style={{display:"flex",gap:10,marginTop:6}}>
                <button onClick={()=>setEditCat(null)} style={{flex:1,background:"none",border:"1.5px solid rgba(255,255,255,.08)",color:"#475569",borderRadius:12,padding:"13px",fontSize:14,fontWeight:600,cursor:"pointer"}}>取消</button>
                <button onClick={saveCatForm} style={{flex:2,background:"linear-gradient(135deg,#166534,#0c4a6e)",color:"#fff",border:"none",borderRadius:12,padding:"13px",fontSize:15,fontWeight:700,cursor:"pointer"}}>✓ 儲存</button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
