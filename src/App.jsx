import { useState, useEffect, useMemo, useCallback } from "react";
import { BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend } from "recharts";
import * as XLSX from "xlsx";

const GAS_URL = "https://script.google.com/macros/s/AKfycbyNb_zx1ZnASY78XPNAF8RTMMgSjoe8GG9yJJ_SMsNVigiUmzLrwAL4tNyu1iBXOj1TEg/exec";

// ── 類別定義 ──
// stock: 正數=獲利/入金, 負數=虧損/出金
// 其餘: 永遠是支出(負)
// income: 工作薪資收入(正)
const CATEGORIES = {
  income:   { label: "薪資收入", icon: "💼", color: "#4ade80",  sign:  1, inputMode: "positive", desc: "薪水、兼職、獎金" },
  stock:    { label: "股票",     icon: "📊", color: "#38bdf8",  sign:  0, inputMode: "signed",   desc: "+入金/獲利  −出金/虧損" },
  living:   { label: "生活費用", icon: "🛒", color: "#fb923c",  sign: -1, inputMode: "positive", desc: "餐飲、購物、交通、娛樂" },
  learning: { label: "學習花費", icon: "📚", color: "#a78bfa",  sign: -1, inputMode: "positive", desc: "課程、書籍、工具" },
  loan:     { label: "貸款",     icon: "🏦", color: "#fbbf24",  sign: -1, inputMode: "positive", desc: "房貸、車貸、分期" },
  card:     { label: "信用卡費", icon: "💳", color: "#f472b6",  sign: -1, inputMode: "positive", desc: "信用卡帳單" },
};

const CAT_GROUPS = {
  all:      { label: "全部",  icon: "◎",  keys: Object.keys(CATEGORIES) },
  stock:    { label: "股票",  icon: "📊", keys: ["stock"] },
  living:   { label: "生活",  icon: "🛒", keys: ["living"] },
  learning: { label: "學習",  icon: "📚", keys: ["learning"] },
  debt:     { label: "負債",  icon: "🏦", keys: ["loan","card"] },
  income:   { label: "收入",  icon: "💼", keys: ["income"] },
};

const STORAGE_KEY = "finance_v7";
const USER_KEY    = "finance_user_v7";
const load    = (user) => { try { return JSON.parse(localStorage.getItem(`${STORAGE_KEY}_${user}`)||"[]"); } catch { return []; } };
const persist = (user, r) => { try { localStorage.setItem(`${STORAGE_KEY}_${user}`, JSON.stringify(r)); } catch {} };
const loadUser = () => { try { return localStorage.getItem(USER_KEY)||""; } catch { return ""; } };
const saveUser = (u) => { try { localStorage.setItem(USER_KEY, u); } catch {} };
const fmtN    = (n) => Math.abs(Math.round(n)).toLocaleString("zh-TW");
const fmtS    = (n) => (n>=0?"+":"")+Math.round(n).toLocaleString("zh-TW");
const today   = () => new Date().toISOString().slice(0,10);
const nowYM   = () => { const n=new Date(); return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}`; };
const ymLabel = (ym) => { const [y,m]=ym.split("-"); return `${y}年${parseInt(m)}月`; };

async function syncToCloud(user, records) {
  try {
    await fetch(GAS_URL,{method:"POST",mode:"no-cors",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({action:"save", user, records})});
    return true;
  } catch { return false; }
}
async function fetchFromCloud(user) {
  try {
    const r = await fetch(`${GAS_URL}?user=${encodeURIComponent(user)}&t=${Date.now()}`);
    const j = await r.json();
    return j.ok && Array.isArray(j.records) ? j.records : null;
  } catch { return null; }
}

// 計算某筆 stock 記錄的實際 sign
function recSign(r) {
  if (r.cat !== "stock") return CATEGORIES[r.cat]?.sign ?? -1;
  return r.amt >= 0 ? 1 : -1;
}
function recAmt(r) { return Math.abs(Number(r.amt)); }

const ALL_KEYS = Object.keys(CATEGORIES);

export default function App() {
  // ── 登入狀態 ──
  const [user,    setUser]    = useState(loadUser);  // "" = 未登入
  const [nameInput, setNameInput] = useState("");

  const [recs,    setRecs]    = useState(()=> loadUser() ? load(loadUser()) : []);
  const [tab,     setTab]     = useState("add");
  const [checked, setChecked] = useState(()=>new Set(ALL_KEYS));
  const [ym,      setYm]      = useState(nowYM);
  const [rMode,   setRMode]   = useState("month");
  const [range,   setRange]   = useState(()=>{
    const n=new Date();
    return {from:`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}-01`,to:today()};
  });
  const [syncing,  setSyncing]  = useState(false);
  const [lastSync, setLastSync] = useState(null);
  const [syncOk,   setSyncOk]   = useState(null);
  const [toast,    setToast]    = useState(null);
  const [form, setForm] = useState({date:today(), cat:"income", rawAmt:"", note:""});
  const [stockSign, setStockSign] = useState(1);

  const showToast = useCallback((msg,type="ok")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),2500); },[]);

  // 登入
  function doLogin() {
    const name = nameInput.trim();
    if (!name) { showToast("⚠️ 請輸入你的名字","warn"); return; }
    saveUser(name); setUser(name);
    const local = load(name);
    setRecs(local);
    // 自動從雲端拉取
    (async()=>{
      setSyncing(true);
      const cloud = await fetchFromCloud(name);
      setSyncing(false);
      if (cloud) {
        const ids = new Set(cloud.map(r=>String(r.id)));
        const localOnly = local.filter(r=>!ids.has(String(r.id)));
        const merged = [...cloud,...localOnly].sort((a,b)=>String(b.date).localeCompare(String(a.date)));
        setRecs(merged); persist(name, merged);
        setSyncOk(true); setLastSync(new Date().toISOString());
        showToast(`👋 歡迎回來，${name}！`);
      } else {
        setSyncOk(false);
        showToast(`👋 ${name}，已載入本機資料`);
      }
    })();
  }

  // 登出
  function doLogout() {
    if (!window.confirm(`確定登出「${user}」？`)) return;
    saveUser(""); setUser(""); setRecs([]); setNameInput("");
    setSyncOk(null); setLastSync(null);
  }

  useEffect(()=>{ if(user) persist(user, recs); },[recs, user]);

  const allYMs = useMemo(()=>{
    const s=new Set(recs.map(r=>r.date.slice(0,7))); s.add(nowYM());
    return [...s].sort().reverse();
  },[recs]);

  const datFiltered = useMemo(()=>{
    if (rMode==="month") return recs.filter(r=>r.date.slice(0,7)===ym);
    return recs.filter(r=>r.date>=range.from&&r.date<=range.to);
  },[recs,rMode,ym,range]);

  const filtered = useMemo(()=>{
    return datFiltered.filter(r=>checked.has(r.cat));
  },[datFiltered,checked]);

  // Summary: totalIn = income + stock gains; totalOut = all expenses + stock losses
  const sum = useMemo(()=>{
    let totalIn=0, totalOut=0;
    const by = Object.fromEntries(Object.keys(CATEGORIES).map(k=>[k,0]));
    filtered.forEach(r=>{
      const a = Number(r.amt);
      if (r.cat==="stock") {
        if (a>=0) { totalIn+=a; by.stock+= a; }
        else       { totalOut+=Math.abs(a); by.stock+=a; }
      } else if (CATEGORIES[r.cat]?.sign===1) {
        totalIn+=a; by[r.cat]+=a;
      } else {
        totalOut+=a; by[r.cat]+=a;
      }
    });
    return {...by, totalIn, totalOut, net:totalIn-totalOut};
  },[filtered]);

  // Stock breakdown
  const stockBreak = useMemo(()=>{
    const rows = datFiltered.filter(r=>r.cat==="stock");
    const earn = rows.filter(r=>Number(r.amt)>=0).reduce((a,r)=>a+Number(r.amt),0);
    const loss = rows.filter(r=>Number(r.amt)<0).reduce((a,r)=>a+Math.abs(Number(r.amt)),0);
    return {earn:Math.round(earn), loss:Math.round(loss), net:Math.round(earn-loss)};
  },[datFiltered]);

  const trend = useMemo(()=>{
    return [...allYMs].reverse().slice(-6).map(m=>{
      const r=recs.filter(x=>x.date.slice(0,7)===m&&checked.has(x.cat));
      let inn=0,out=0;
      r.forEach(x=>{
        const a=Number(x.amt);
        if(x.cat==="stock"){ if(a>=0)inn+=a; else out+=Math.abs(a); }
        else if(CATEGORIES[x.cat]?.sign===1) inn+=a;
        else out+=a;
      });
      return {m:`${parseInt(m.split("-")[1])}月`,inn:Math.round(inn),out:Math.round(out),net:Math.round(inn-out)};
    });
  },[recs,allYMs,checked]);

  const pie = useMemo(()=>{
    return [...checked].map(k=>{
      const v=CATEGORIES[k];
      const val = k==="stock" ? Math.abs(Math.round(sum[k]||0)) : Math.round(Math.abs(sum[k]||0));
      return {name:v.label, val, color:v.color};
    }).filter(d=>d.val>0);
  },[sum,checked]);

  async function addRecord() {
    const rawNum = parseFloat(form.rawAmt);
    if (!form.rawAmt||isNaN(rawNum)||rawNum<=0) { showToast("⚠️ 請輸入正確金額","warn"); return; }
    // for stock: apply sign toggle
    const finalAmt = form.cat==="stock" ? rawNum*stockSign : rawNum;
    const r = {id:Date.now(), date:form.date, cat:form.cat, amt:finalAmt, note:form.note.trim()};
    const updated=[r,...recs].sort((a,b)=>String(b.date).localeCompare(String(a.date)));
    setRecs(updated); setForm(f=>({...f,rawAmt:"",note:""}));
    const label = form.cat==="stock" ? (stockSign>0?"📈 入金/獲利":"📉 出金/虧損") : CATEGORIES[form.cat]?.label;
    showToast(`✅ 已新增：${label}`);
    setSyncing(true);
    const ok=await syncToCloud(user, updated); setSyncing(false);
    if(ok){setSyncOk(true);setLastSync(new Date().toISOString());}
    else{setSyncOk(false);showToast("⚠️ 本機已存，雲端同步失敗","warn");}
  }

  async function delRecord(id) {
    if(!window.confirm("確定刪除？"))return;
    const updated=recs.filter(r=>r.id!==id); setRecs(updated);
    setSyncing(true); await syncToCloud(user, updated); setSyncing(false);
    showToast("🗑 已刪除");
  }

  // ── 編輯 ──
  const [editRec, setEditRec] = useState(null); // null = 關閉，否則 = 正在編輯的record
  const [editForm, setEditForm] = useState({date:"", cat:"income", rawAmt:"", note:""});
  const [editStockSign, setEditStockSign] = useState(1);

  function openEdit(r) {
    const a = Number(r.amt);
    setEditForm({ date: r.date, cat: r.cat, rawAmt: String(Math.abs(a)), note: r.note||"" });
    setEditStockSign(r.cat==="stock" ? (a>=0?1:-1) : 1);
    setEditRec(r);
  }

  async function saveEdit() {
    const rawNum = parseFloat(editForm.rawAmt);
    if (!editForm.rawAmt||isNaN(rawNum)||rawNum<=0) { showToast("⚠️ 請輸入正確金額","warn"); return; }
    const finalAmt = editForm.cat==="stock" ? rawNum*editStockSign : rawNum;
    const updated = recs.map(r => r.id===editRec.id
      ? {...r, date:editForm.date, cat:editForm.cat, amt:finalAmt, note:editForm.note.trim()}
      : r
    ).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
    setRecs(updated); setEditRec(null);
    showToast("✏️ 已儲存修改");
    setSyncing(true); const ok=await syncToCloud(user, updated); setSyncing(false);
    if(ok){setSyncOk(true);setLastSync(new Date().toISOString());}
    else showToast("⚠️ 本機已存，雲端同步失敗","warn");
  }

  async function pushToCloud(){
    setSyncing(true); const ok=await syncToCloud(user, recs); setSyncing(false);
    if(ok){setSyncOk(true);setLastSync(new Date().toISOString());showToast("☁️ 已推送到雲端！");}
    else{setSyncOk(false);showToast("❌ 同步失敗","err");}
  }
  async function pullFromCloud(){
    setSyncing(true); const cloud=await fetchFromCloud(user); setSyncing(false);
    if(cloud){
      setRecs(cloud.sort((a,b)=>String(b.date).localeCompare(String(a.date))));
      setSyncOk(true);setLastSync(new Date().toISOString());showToast("⬇️ 已載入最新資料！");
    } else{setSyncOk(false);showToast("❌ 載入失敗","err");}
  }

  function exportXlsx(){
    const data=recs.map(r=>{
      const cat=CATEGORIES[r.cat];
      const a=Number(r.amt);
      let label=cat?.label||r.cat;
      if(r.cat==="stock") label=a>=0?"股票入金/獲利":"股票出金/虧損";
      return {日期:r.date,類別:label,金額:Math.round(a),備註:r.note||""};
    });
    const ws=XLSX.utils.json_to_sheet(data); ws["!cols"]=[{wch:12},{wch:16},{wch:14},{wch:28}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"記帳明細");
    XLSX.writeFile(wb,`財務_${new Date().toLocaleDateString("zh-TW").replace(/\//g,"-")}.xlsx`);
    showToast("📊 已匯出");
  }

  const isPos=sum.net>=0;
  const netPct=sum.totalOut>0?Math.abs((sum.net/sum.totalOut)*100).toFixed(1):"0.0";
  const totalFlow=(sum.totalIn+sum.totalOut)||1;
  const lastSyncStr=lastSync?new Date(lastSync).toLocaleTimeString("zh-TW",{hour:"2-digit",minute:"2-digit"}):null;
  const tt={background:"#0a0f1a",border:"1px solid #1e293b",borderRadius:12,fontSize:12,color:"#e2e8f0"};

  const MNav=()=>(
    <div style={{display:"flex",alignItems:"center",background:"#0a0f1a",borderRadius:12,border:"1px solid #1e293b",overflow:"hidden"}}>
      <button onClick={()=>{const i=allYMs.indexOf(ym);if(i<allYMs.length-1)setYm(allYMs[i+1]);}} disabled={allYMs.indexOf(ym)>=allYMs.length-1}
        style={{border:"none",background:"transparent",color:"#475569",padding:"9px 16px",fontSize:18,cursor:"pointer",opacity:allYMs.indexOf(ym)>=allYMs.length-1?.2:1}}>‹</button>
      <div style={{flex:1,textAlign:"center",fontSize:14,fontWeight:600,color:"#e2e8f0",minWidth:96}}>{ymLabel(ym)}</div>
      <button onClick={()=>{const i=allYMs.indexOf(ym);if(i>0)setYm(allYMs[i-1]);}} disabled={allYMs.indexOf(ym)<=0}
        style={{border:"none",background:"transparent",color:"#475569",padding:"9px 16px",fontSize:18,cursor:"pointer",opacity:allYMs.indexOf(ym)<=0?.2:1}}>›</button>
    </div>
  );

  // 切換單一類別
  function toggleCat(k){
    setChecked(prev=>{
      const s=new Set(prev);
      if(s.has(k)){ if(s.size===1)return s; s.delete(k); } else s.add(k);
      return s;
    });
  }
  // 快速預設
  const PRESETS = [
    {label:"全部",  keys: ALL_KEYS},
    {label:"股票",  keys:["stock"]},
    {label:"支出",  keys:["living","learning","loan","card"]},
    {label:"收支",  keys:["income","living","living","learning","loan","card"]},
  ];
  function applyPreset(keys){ setChecked(new Set(keys)); }
  const isAll = checked.size===ALL_KEYS.length;

  const CatFilter=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      {/* Quick presets */}
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        <button onClick={()=>setChecked(new Set(ALL_KEYS))}
          style={{border:`1.5px solid ${isAll?"#4ade80":"#1e293b"}`,background:isAll?"#4ade8015":"#0a0f1a",color:isAll?"#4ade80":"#475569",borderRadius:99,padding:"5px 13px",fontSize:11,fontWeight:700,cursor:"pointer",transition:"all .15s"}}>
          ◎ 全部
        </button>
        {[{l:"股票",k:["stock"]},{l:"生活+卡費",k:["living","card"]},{l:"負債",k:["loan","card"]},{l:"收入",k:["income"]}].map(({l,k})=>{
          const active=k.length===checked.size&&k.every(x=>checked.has(x));
          return (
            <button key={l} onClick={()=>applyPreset(k)}
              style={{border:`1.5px solid ${active?"#38bdf8":"#1e293b"}`,background:active?"#38bdf815":"#0a0f1a",color:active?"#38bdf8":"#475569",borderRadius:99,padding:"5px 13px",fontSize:11,fontWeight:700,cursor:"pointer",transition:"all .15s"}}>
              {l}
            </button>
          );
        })}
      </div>
      {/* Individual checkboxes */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:7}}>
        {Object.entries(CATEGORIES).map(([k,v])=>{
          const on=checked.has(k);
          return (
            <button key={k} onClick={()=>toggleCat(k)}
              style={{border:`1.5px solid ${on?v.color:"#1e293b"}`,background:on?v.color+"14":"#0a0f1a",color:on?v.color:"#334155",borderRadius:12,padding:"8px 6px",fontSize:11,fontWeight:600,cursor:"pointer",transition:"all .15s",display:"flex",alignItems:"center",gap:5}}>
              <span style={{fontSize:14,lineHeight:1,opacity:on?1:.4}}>
                {on?"☑":"☐"}
              </span>
              <span>{v.icon} {v.label}</span>
            </button>
          );
        })}
      </div>
      <div style={{fontSize:10,color:"#1e3a5f",fontWeight:500}}>
        已選 {checked.size} 個類別 · 可多選組合
      </div>
    </div>
  );

  const curCat = CATEGORIES[form.cat];

  // ── 登入畫面 ──
  if (!user) return (
    <div style={{fontFamily:"'Noto Sans TC',sans-serif",minHeight:"100vh",background:"#060910",color:"#e2e8f0",display:"flex",alignItems:"center",justifyContent:"center",padding:"24px"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600;700&family=DM+Mono:wght@400;500&display=swap');*{box-sizing:border-box;margin:0;padding:0}input,button{font-family:inherit;outline:none}.lfi{width:100%;background:#0a0f1a;border:1.5px solid #1e293b;border-radius:14px;padding:14px 18px;color:#e2e8f0;font-size:18px;text-align:center;letter-spacing:.05em;transition:border-color .2s}.lfi:focus{border-color:#4ade80;box-shadow:0 0 0 3px #4ade8015}.lfi::placeholder{color:#1e3a5f;font-size:15px}`}</style>
      {toast&&<div style={{position:"fixed",top:20,left:"50%",transform:"translateX(-50%)",background:"#0d1526",border:"1px solid #fbbf2466",color:"#e2e8f0",padding:"11px 24px",borderRadius:99,fontSize:13,zIndex:999,whiteSpace:"nowrap"}}>{toast.msg}</div>}
      <div style={{width:"100%",maxWidth:360,display:"flex",flexDirection:"column",alignItems:"center",gap:0}}>
        {/* Logo */}
        <div style={{width:72,height:72,background:"linear-gradient(135deg,#4ade80,#38bdf8)",borderRadius:22,display:"flex",alignItems:"center",justifyContent:"center",fontSize:36,marginBottom:20,boxShadow:"0 8px 32px #4ade8030"}}>💹</div>
        <div style={{fontSize:24,fontWeight:700,background:"linear-gradient(90deg,#4ade80,#38bdf8)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",marginBottom:6}}>財務記帳</div>
        <div style={{fontSize:12,color:"#1e3a5f",letterSpacing:".14em",fontWeight:600,marginBottom:36}}>PERSONAL FINANCE</div>

        {/* Login card */}
        <div style={{width:"100%",background:"linear-gradient(150deg,#0e1826,#080d16)",border:"1px solid #1a2540",borderRadius:24,padding:"28px 24px",display:"flex",flexDirection:"column",gap:20}}>
          <div>
            <div style={{fontSize:13,fontWeight:700,color:"#94a3b8",marginBottom:6,textAlign:"center"}}>輸入你的名字開始使用</div>
            <div style={{fontSize:11,color:"#1e3a5f",textAlign:"center",marginBottom:16}}>每個人的資料獨立儲存，不互相影響</div>
            <input className="lfi" placeholder="例：小明、爸爸、王小華..." value={nameInput}
              onChange={e=>setNameInput(e.target.value)}
              onKeyDown={e=>{if(e.key==="Enter")doLogin()}}
              autoFocus />
          </div>
          <button onClick={doLogin}
            style={{width:"100%",background:"linear-gradient(135deg,#166534,#0c4a6e)",color:"#fff",border:"none",borderRadius:14,padding:"15px",fontSize:16,fontWeight:700,cursor:"pointer",letterSpacing:".03em",boxShadow:"0 4px 20px #4ade8025",transition:"transform .1s"}}
            onMouseDown={e=>e.currentTarget.style.transform="scale(.97)"}
            onMouseUp={e=>e.currentTarget.style.transform="scale(1)"}>
            {syncing ? "載入中..." : "▶ 進入我的帳本"}
          </button>
        </div>

        <div style={{marginTop:20,fontSize:11,color:"#1e3a5f",textAlign:"center",lineHeight:1.8}}>
          不同名字 = 不同帳本<br/>
          同一名字在手機和電腦都能同步
        </div>
      </div>
    </div>
  );

  return (
    <div style={{fontFamily:"'Noto Sans TC',sans-serif",minHeight:"100vh",background:"#060910",color:"#e2e8f0"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        input,button{font-family:inherit;outline:none}
        ::-webkit-scrollbar{width:0}
        .fi{width:100%;background:#0a0f1a;border:1.5px solid #1e293b;border-radius:14px;padding:13px 16px;color:#e2e8f0;font-size:15px;transition:border-color .2s,box-shadow .2s;-webkit-appearance:none}
        .fi:focus{border-color:#4ade80;box-shadow:0 0 0 3px #4ade8015}
        .fi::placeholder{color:#1e3a5f}
        .card{background:linear-gradient(150deg,#0e1826 0%,#080d16 100%);border:1px solid #1a2540;border-radius:22px;padding:22px;position:relative;overflow:hidden}
        .card::after{content:'';position:absolute;top:0;left:10%;right:10%;height:1px;background:linear-gradient(90deg,transparent,#ffffff08,transparent)}
        .tab-bar{display:flex;gap:3px;background:#080d16;border-radius:14px;padding:4px;border:1px solid #1a2540}
        .tab{flex:1;border:none;background:transparent;color:#334155;font-size:12px;font-weight:600;padding:9px 4px;cursor:pointer;border-radius:10px;transition:all .2s;white-space:nowrap;letter-spacing:.02em}
        .tab.on{color:#fff;box-shadow:0 2px 10px #00000055}
        .tab.on.t0{background:linear-gradient(135deg,#166534,#0c4a6e)}
        .tab.on.t1{background:linear-gradient(135deg,#0c4a6e,#1e1b4b)}
        .tab.on.t2{background:linear-gradient(135deg,#4c1d95,#1e3a5f)}
        .tab.on.t3{background:linear-gradient(135deg,#7c2d12,#1e3a5f)}
        .cat-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:9px}
        .cat-btn{border:1.5px solid #1a2540;background:#0a0f1a;border-radius:14px;padding:13px 8px 11px;font-size:11px;font-weight:600;color:#334155;cursor:pointer;transition:all .18s;text-align:center;display:flex;flex-direction:column;align-items:center;gap:5px;line-height:1.3}
        .cat-btn:hover{border-color:#334155;color:#64748b}
        .cat-ico{font-size:24px;line-height:1}
        .add-btn{width:100%;border:none;border-radius:14px;padding:15px;font-size:16px;font-weight:700;cursor:pointer;letter-spacing:.04em;transition:transform .1s,box-shadow .2s}
        .add-btn:active{transform:scale(.97)}
        .add-btn:disabled{opacity:.4;cursor:not-allowed;box-shadow:none!important}
        .sign-toggle{display:flex;border-radius:12px;overflow:hidden;border:1.5px solid #1e293b;background:#0a0f1a}
        .sign-btn{flex:1;border:none;padding:11px;font-size:13px;font-weight:700;cursor:pointer;transition:all .18s;background:transparent;display:flex;align-items:center;justify-content:center;gap:6px}
        .bar-t{height:5px;background:#0a1020;border-radius:99px;overflow:hidden;margin-top:9px}
        .bar-f{height:100%;border-radius:99px;transition:width .9s cubic-bezier(.4,0,.2,1)}
        .toast{position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#0d1526;border:1px solid #4ade8055;color:#e2e8f0;padding:11px 24px;border-radius:99px;font-size:13px;z-index:999;box-shadow:0 8px 40px #000c;white-space:nowrap;animation:fu .22s ease}
        .toast.err{border-color:#f8717166}.toast.warn{border-color:#fbbf2466}
        @keyframes fu{from{opacity:0;transform:translateX(-50%) translateY(-8px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}
        .rr{display:flex;align-items:center;gap:12px;padding:13px 0;border-bottom:1px solid #0d1526}
        .rr:last-child{border-bottom:none}
        .del{background:none;border:none;color:#f87171;padding:7px;opacity:.22;border-radius:8px;cursor:pointer;transition:opacity .2s;flex-shrink:0;font-size:13px}
        .del:hover{opacity:1}
        .pill{display:inline-flex;background:#0a0f1a;border:1.5px solid #1e293b;border-radius:10px;overflow:hidden}
        .pill button{border:none;padding:8px 16px;font-size:12px;color:#334155;background:transparent;cursor:pointer;font-weight:600;transition:all .15s}
        .pill button.on{background:#0d2035;color:#e2e8f0}
        .exp-btn{background:none;border:1.5px solid #1e293b;color:#4ade80;border-radius:12px;padding:9px 16px;font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px;transition:all .2s}
        .exp-btn:hover{background:#4ade8010;border-color:#4ade8044}
        .sync-btn{background:#080d16;border:1.5px solid #1a2540;border-radius:10px;padding:7px 12px;font-size:11px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:5px;transition:all .2s}
        .sync-btn:hover{border-color:#334155}
        .sync-btn:disabled{opacity:.3;cursor:not-allowed}
        .chip{font-size:10px;padding:3px 9px;border-radius:99px;font-weight:700}
        .mono{font-family:'DM Mono',monospace}
        .spin{animation:sp 1s linear infinite;display:inline-block}
        @keyframes sp{from{transform:rotate(0)}to{transform:rotate(360deg)}}
        .scard{border-radius:16px;padding:15px;display:flex;flex-direction:column;gap:5px}
        @media(max-width:360px){.cat-grid{grid-template-columns:repeat(2,1fr)}}
      `}</style>

      {toast&&<div className={`toast${toast.type==="err"?" err":toast.type==="warn"?" warn":""}`}>{toast.msg}</div>}

      {/* ── Header ── */}
      <div style={{background:"rgba(6,9,16,.95)",backdropFilter:"blur(24px)",WebkitBackdropFilter:"blur(24px)",borderBottom:"1px solid #0d1526",padding:"14px 18px 10px",position:"sticky",top:0,zIndex:50}}>
        <div style={{maxWidth:520,margin:"0 auto"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:38,height:38,background:"linear-gradient(135deg,#4ade80,#38bdf8)",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0,boxShadow:"0 4px 16px #4ade8030"}}>💹</div>
              <div>
                <div style={{display:"flex",alignItems:"center",gap:7}}>
                  <div style={{fontSize:15,fontWeight:700,background:"linear-gradient(90deg,#4ade80,#38bdf8)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>財務記帳</div>
                  <div style={{background:"#4ade8018",border:"1px solid #4ade8033",borderRadius:99,padding:"2px 10px",fontSize:11,fontWeight:700,color:"#4ade80"}}>👤 {user}</div>
                </div>
                <div style={{fontSize:9,color:"#1e3a5f",letterSpacing:".14em",fontWeight:600}}>PERSONAL FINANCE</div>
              </div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:5}}>
              {lastSyncStr&&(
                <div style={{display:"flex",alignItems:"center",gap:4}}>
                  <div style={{width:6,height:6,borderRadius:99,background:syncOk?"#4ade80":syncOk===false?"#f87171":"#fbbf24"}}/>
                  <span style={{fontSize:10,color:"#334155",fontWeight:500}}>{lastSyncStr}</span>
                </div>
              )}
              <button className="sync-btn" onClick={pullFromCloud} disabled={syncing} style={{color:"#38bdf8"}}>
                <span className={syncing?"spin":""} style={{fontSize:12}}>⬇</span><span>{syncing?"...":"拉取"}</span>
              </button>
              <button className="sync-btn" onClick={pushToCloud} disabled={syncing} style={{color:"#4ade80"}}>
                <span style={{fontSize:12}}>⬆</span><span>推送</span>
              </button>
              <button className="sync-btn" onClick={doLogout} style={{color:"#f87171",borderColor:"#f8717122"}}>
                <span style={{fontSize:12}}>⏏</span><span>登出</span>
              </button>
            </div>
          </div>
          <div className="tab-bar">
            {[["add","📝 記帳","t0"],["summary","📊 統計","t1"],["chart","📈 圖表","t2"],["records","🗂 明細","t3"]].map(([v,l,c])=>(
              <button key={v} className={`tab${tab===v?` on ${c}`:""}`} onClick={()=>setTab(v)}>{l}</button>
            ))}
          </div>
        </div>
      </div>

      <div style={{maxWidth:520,margin:"0 auto",padding:"18px 14px 48px",display:"flex",flexDirection:"column",gap:14}}>

        {/* ══ ADD ══ */}
        {tab==="add"&&<>
          <div className="card">
            <div style={{fontSize:15,fontWeight:700,marginBottom:20,background:"linear-gradient(90deg,#e2e8f0,#64748b)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>新增一筆記錄</div>
            <div style={{display:"flex",flexDirection:"column",gap:16}}>

              {/* Date */}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>📅 日期</div>
                <input type="date" className="fi" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/>
              </div>

              {/* Category */}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>🏷 類別</div>
                <div className="cat-grid">
                  {Object.entries(CATEGORIES).map(([k,v])=>(
                    <button key={k} className="cat-btn"
                      style={form.cat===k?{borderColor:v.color,background:v.color+"14",color:v.color,transform:"scale(1.04)"}:{}}
                      onClick={()=>setForm(f=>({...f,cat:k}))}>
                      <span className="cat-ico">{v.icon}</span>
                      <span style={{fontSize:11,fontWeight:700}}>{v.label}</span>
                      <span style={{fontSize:9,color:form.cat===k?v.color+"bb":"#1e3a5f",lineHeight:1.3}}>{v.desc}</span>
                    </button>
                  ))}
                </div>
              </div>

              {/* Stock sign toggle — only shown when stock selected */}
              {form.cat==="stock"&&(
                <div>
                  <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>📌 方向</div>
                  <div className="sign-toggle">
                    <button className="sign-btn"
                      style={stockSign===1?{background:"linear-gradient(135deg,#0c4a6e,#14532d)",color:"#4ade80"}:{color:"#334155"}}
                      onClick={()=>setStockSign(1)}>
                      <span style={{fontSize:18}}>📈</span>
                      <div style={{textAlign:"left"}}>
                        <div style={{fontSize:13,fontWeight:700}}>入金 / 獲利</div>
                        <div style={{fontSize:10,opacity:.7}}>資金存入、賺錢</div>
                      </div>
                    </button>
                    <div style={{width:"1px",background:"#1e293b"}}/>
                    <button className="sign-btn"
                      style={stockSign===-1?{background:"linear-gradient(135deg,#7c1d1d,#1c1917)",color:"#f87171"}:{color:"#334155"}}
                      onClick={()=>setStockSign(-1)}>
                      <span style={{fontSize:18}}>📉</span>
                      <div style={{textAlign:"left"}}>
                        <div style={{fontSize:13,fontWeight:700}}>出金 / 虧損</div>
                        <div style={{fontSize:10,opacity:.7}}>資金取出、賠錢</div>
                      </div>
                    </button>
                  </div>
                </div>
              )}

              {/* Amount */}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>
                  💵 金額（元）
                  {form.cat==="stock"&&<span style={{marginLeft:8,fontSize:10,color:stockSign===1?"#4ade80":"#f87171",fontWeight:700}}>{stockSign===1?"＋ 入金/獲利":"－ 出金/虧損"}</span>}
                </div>
                <div style={{position:"relative"}}>
                  {form.cat==="stock"&&(
                    <div style={{position:"absolute",left:16,top:"50%",transform:"translateY(-50%)",fontSize:22,fontWeight:700,color:stockSign===1?"#4ade80":"#f87171",fontFamily:"'DM Mono',monospace",pointerEvents:"none",zIndex:1}}>
                      {stockSign===1?"+":"−"}
                    </div>
                  )}
                  <input type="number" className="fi mono" placeholder="0" value={form.rawAmt} min="0" inputMode="decimal"
                    style={{fontSize:24,letterSpacing:"-.03em",color:curCat?.color||"#e2e8f0",paddingLeft:form.cat==="stock"?"40px":"16px"}}
                    onChange={e=>setForm(f=>({...f,rawAmt:e.target.value}))}
                    onKeyDown={e=>{if(e.key==="Enter")addRecord()}}/>
                </div>
              </div>

              {/* Note */}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:8,fontWeight:700}}>📝 備註（選填）</div>
                <input className="fi" placeholder={
                  form.cat==="stock"?"例：台積電、大盤、ETF 0050...":
                  form.cat==="learning"?"例：Udemy 課程、程式設計書...":
                  form.cat==="loan"?"例：房貸、車貸...":
                  form.cat==="card"?"例：國泰世華帳單...":
                  "備註..."
                } value={form.note}
                  onChange={e=>setForm(f=>({...f,note:e.target.value}))}
                  onKeyDown={e=>{if(e.key==="Enter")addRecord()}}/>
              </div>

              {/* Submit */}
              <button className="add-btn" onClick={addRecord} disabled={syncing}
                style={{
                  background: form.cat==="stock"
                    ? stockSign===1
                      ? "linear-gradient(135deg,#166534,#0c4a6e)"
                      : "linear-gradient(135deg,#991b1b,#1c1917)"
                    : `linear-gradient(135deg,${curCat?.color}cc,${curCat?.color}88)`,
                  color:"#fff",
                  boxShadow:`0 4px 20px ${curCat?.color}25`
                }}>
                {syncing?"同步中...":
                  form.cat==="stock"
                    ? stockSign===1?"＋ 新增入金 / 獲利":"－ 新增出金 / 虧損"
                    : `＋ 新增${curCat?.label}`
                }
              </button>
            </div>
          </div>

          {/* Recent 5 */}
          {recs.slice(0,5).length>0&&(
            <div className="card">
              <div style={{fontSize:11,color:"#1e3a5f",letterSpacing:".1em",fontWeight:700,marginBottom:14}}>最近記錄</div>
              {recs.slice(0,5).map(r=>{
                const c=CATEGORIES[r.cat];
                const a=Number(r.amt);
                const isIn=r.cat==="income"||(r.cat==="stock"&&a>=0);
                return (
                  <div key={r.id} className="rr">
                    <div style={{width:40,height:40,background:c?.color+"15",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{c?.icon}</div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:13,fontWeight:500,color:"#cbd5e1",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {r.cat==="stock"?(a>=0?"📈 入金/獲利":"📉 出金/虧損"):c?.label}
                        {r.note?` · ${r.note}`:""}
                      </div>
                      <div style={{fontSize:11,color:"#1e3a5f",marginTop:3,fontWeight:500}}>{r.date}</div>
                    </div>
                    <span className="mono" style={{fontSize:15,fontWeight:600,color:isIn?"#4ade80":"#f87171",flexShrink:0}}>
                      {isIn?"+":"-"}{fmtN(a)}
                    </span>
                  </div>
                );
              })}
            </div>
          )}

          {/* Sync status */}
          <div style={{background:"#080d16",border:"1px solid #1a2540",borderRadius:16,padding:"13px 16px",display:"flex",alignItems:"center",gap:12}}>
            <div style={{width:34,height:34,background:"#38bdf812",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>☁️</div>
            <div style={{flex:1}}>
              <div style={{fontSize:12,fontWeight:600,color:"#38bdf8"}}>雲端同步</div>
              <div style={{fontSize:11,color:"#1e3a5f",marginTop:2}}>{syncing?"同步中...":syncOk===true?`✓ 同步完成 ${lastSyncStr||""}`:syncOk===false?"未連線，資料暫存本機":"啟動中..."}</div>
            </div>
            <div style={{width:9,height:9,borderRadius:99,background:syncing?"#fbbf24":syncOk?"#4ade80":"#f87171"}}/>
          </div>
        </>}

        {/* ══ SUMMARY ══ */}
        {tab==="summary"&&<>
          {/* Controls */}
          <div className="card" style={{padding:"16px 18px",gap:12,display:"flex",flexDirection:"column"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
              <div className="pill">
                <button className={rMode==="month"?"on":""} onClick={()=>setRMode("month")}>月份</button>
                <button className={rMode==="custom"?"on":""} onClick={()=>setRMode("custom")}>自訂</button>
              </div>
              {rMode==="month"&&<MNav/>}
            </div>
            {rMode==="custom"&&(
              <div style={{display:"flex",gap:10}}>
                <div style={{flex:1}}>
                  <div style={{fontSize:10,color:"#334155",letterSpacing:".1em",marginBottom:6,fontWeight:700}}>開始</div>
                  <input type="date" className="fi" style={{fontSize:13}} value={range.from} onChange={e=>setRange(r=>({...r,from:e.target.value}))}/>
                </div>
                <div style={{flex:1}}>
                  <div style={{fontSize:10,color:"#334155",letterSpacing:".1em",marginBottom:6,fontWeight:700}}>結束</div>
                  <input type="date" className="fi" style={{fontSize:13}} value={range.to} onChange={e=>setRange(r=>({...r,to:e.target.value}))}/>
                </div>
              </div>
            )}
            <div>
              <div style={{fontSize:10,color:"#334155",letterSpacing:".1em",marginBottom:8,fontWeight:700}}>篩選類別</div>
              <CatFilter/>
            </div>
          </div>

          {/* Net */}
          <div style={{background:isPos?"linear-gradient(150deg,#041510,#060d18)":"linear-gradient(150deg,#150408,#060d18)",border:`1px solid ${isPos?"#4ade8020":"#f8717120"}`,borderRadius:22,padding:"28px 22px",textAlign:"center",boxShadow:isPos?"0 0 32px #4ade8010":"0 0 32px #f8717110"}}>
            <div style={{fontSize:10,color:isPos?"#4ade8055":"#f8717155",letterSpacing:".2em",fontWeight:700,marginBottom:12}}>
              已選 {checked.size} 個類別 · 淨損益
            </div>
            <div className="mono" style={{fontSize:52,fontWeight:500,color:isPos?"#4ade80":"#f87171",letterSpacing:"-.04em",lineHeight:1}}>
              {isPos?"+":""}{fmtN(sum.net)}<span style={{fontSize:16,fontWeight:400,marginLeft:8,opacity:.55}}>元</span>
            </div>
            <div style={{marginTop:16}}>
              <span style={{background:isPos?"#4ade8015":"#f8717115",border:`1px solid ${isPos?"#4ade8030":"#f8717130"}`,color:isPos?"#4ade80":"#f87171",borderRadius:99,padding:"6px 20px",fontSize:13,fontWeight:700}}>
                {isPos?"📈 結餘":"📉 透支"} {netPct}%
              </span>
            </div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              {[{l:"總收入",v:sum.totalIn,c:"#4ade80",s:"+"},{l:"總支出",v:sum.totalOut,c:"#f87171",s:"-"},{l:"筆數",v:filtered.length,c:"#38bdf8",s:""}].map(({l,v,c,s})=>(
                <div key={l} style={{flex:1,background:c+"0c",border:`1px solid ${c}18`,borderRadius:14,padding:"13px 8px"}}>
                  <div style={{fontSize:9,color:c+"77",letterSpacing:".1em",fontWeight:700}}>{l}</div>
                  <div className="mono" style={{fontSize:16,color:c,marginTop:6}}>{s}{l==="筆數"?v:fmtN(v)}</div>
                </div>
              ))}
            </div>
          </div>

          {/* Stock special card */}
          {checked.has("stock")&&(
            <div style={{background:"linear-gradient(150deg,#060d18,#080d16)",border:"1px solid #1a2d4a",borderRadius:20,padding:"18px"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#38bdf8",marginBottom:14,display:"flex",alignItems:"center",gap:6}}>
                <span>📊</span> 股票專區
              </div>
              <div style={{display:"flex",gap:10}}>
                {[{l:"入金/獲利",v:stockBreak.earn,c:"#4ade80",s:"+"},{l:"出金/虧損",v:stockBreak.loss,c:"#f87171",s:"-"},{l:"股票合計",v:stockBreak.net,c:stockBreak.net>=0?"#4ade80":"#f87171",s:fmtS(stockBreak.net)[0]}].map(({l,v,c,s})=>(
                  <div key={l} className="scard" style={{flex:1,background:c+"0c",border:`1px solid ${c}18`}}>
                    <div style={{fontSize:9,color:c+"77",letterSpacing:".08em",fontWeight:700}}>{l}</div>
                    <div className="mono" style={{fontSize:16,color:c,fontWeight:600}}>{s==="+"?"+":s==="-"?"-":""}{fmtN(Math.abs(v))}</div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Breakdown */}
          <div className="card">
            <div style={{fontSize:14,fontWeight:700,marginBottom:18,color:"#f1f5f9"}}>各類別明細</div>
            {Object.entries(CATEGORIES)
              .filter(([k])=>checked.has(k))
              .map(([k,v])=>{
                const raw=sum[k]||0;
                const val=Math.abs(raw);
                const pct=((val/totalFlow)*100).toFixed(1);
                const isInCat=k==="income"||(k==="stock"&&raw>0);
                return (
                  <div key={k} style={{marginBottom:18}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:34,height:34,background:v.color+"15",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17}}>{v.icon}</div>
                        <div>
                          <div style={{fontSize:13,fontWeight:600,color:"#cbd5e1"}}>{v.label}</div>
                          {k==="stock"&&<div style={{fontSize:10,color:"#334155"}}>{raw>=0?"入金/獲利":"出金/虧損"}</div>}
                        </div>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:10}}>
                        <span style={{fontSize:10,color:"#1e3a5f",fontWeight:600,background:"#0a1020",borderRadius:6,padding:"2px 7px"}}>{pct}%</span>
                        <span className="mono" style={{fontSize:15,fontWeight:600,color:v.color,minWidth:90,textAlign:"right"}}>
                          {isInCat?"+":"-"}{fmtN(val)}
                        </span>
                      </div>
                    </div>
                    <div className="bar-t"><div className="bar-f" style={{width:`${pct}%`,background:`linear-gradient(90deg,${v.color}44,${v.color})`}}/></div>
                  </div>
                );
              })
            }
          </div>

          <div style={{display:"flex",justifyContent:"flex-end"}}>
            <button className="exp-btn" onClick={exportXlsx}>📊 匯出 Excel</button>
          </div>
        </>}

        {/* ══ CHART ══ */}
        {tab==="chart"&&<>
          <div className="card" style={{padding:"14px 18px",gap:12,display:"flex",flexDirection:"column"}}>
            <div>
              <div style={{fontSize:10,color:"#334155",letterSpacing:".1em",marginBottom:8,fontWeight:700}}>篩選類別</div>
              <CatFilter/>
            </div>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
              <div className="pill">
                <button className={rMode==="month"?"on":""} onClick={()=>setRMode("month")}>月份</button>
                <button className={rMode==="custom"?"on":""} onClick={()=>setRMode("custom")}>自訂</button>
              </div>
              {rMode==="month"&&<MNav/>}
            </div>
          </div>

          <div className="card">
            <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",marginBottom:3}}>近 6 個月趨勢</div>
            <div style={{fontSize:11,color:"#1e3a5f",marginBottom:18}}>已選 {checked.size} 類 · 收支比較</div>
            {trend.length===0
              ?<div style={{textAlign:"center",color:"#1e3a5f",padding:"40px 0"}}>尚無足夠資料</div>
              :<ResponsiveContainer width="100%" height={210}>
                <BarChart data={trend} barGap={3} barCategoryGap="32%">
                  <XAxis dataKey="m" tick={{fill:"#1e3a5f",fontSize:11}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:"#1e3a5f",fontSize:10}} axisLine={false} tickLine={false} width={50}
                    tickFormatter={v=>v>=10000?`${(v/10000).toFixed(0)}萬`:String(v)}/>
                  <Tooltip contentStyle={tt} labelStyle={{color:"#64748b"}}
                    formatter={(v,n)=>[`${v.toLocaleString()} 元`,n==="inn"?"收入":n==="out"?"支出":"淨損益"]}/>
                  <Legend formatter={v=>v==="inn"?"收入":v==="out"?"支出":"淨損益"} wrapperStyle={{fontSize:11,color:"#475569"}}/>
                  <Bar dataKey="inn" fill="#4ade80" radius={[5,5,0,0]}/>
                  <Bar dataKey="out" fill="#f87171" radius={[5,5,0,0]}/>
                  <Bar dataKey="net" fill="#38bdf8" radius={[5,5,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            }
          </div>

          <div className="card">
            <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",marginBottom:14}}>已選類別結構</div>
            {pie.length===0
              ?<div style={{textAlign:"center",color:"#1e3a5f",padding:"40px 0"}}>此期間無資料</div>
              :<ResponsiveContainer width="100%" height={220}>
                <PieChart>
                  <Pie data={pie} cx="50%" cy="50%" outerRadius={85} innerRadius={36} dataKey="val" paddingAngle={4}
                    label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} labelLine={false} fontSize={11}>
                    {pie.map((d,i)=><Cell key={i} fill={d.color}/>)}
                  </Pie>
                  <Tooltip formatter={v=>`${v.toLocaleString()} 元`} contentStyle={tt}/>
                </PieChart>
              </ResponsiveContainer>
            }
          </div>

          <div className="card">
            <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",marginBottom:3}}>淨損益走勢</div>
            <div style={{fontSize:11,color:"#1e3a5f",marginBottom:18}}>已選 {checked.size} 類 · 每月結算</div>
            {trend.length<2
              ?<div style={{textAlign:"center",color:"#1e3a5f",padding:"30px 0"}}>至少需要 2 個月資料</div>
              :<ResponsiveContainer width="100%" height={170}>
                <LineChart data={trend}>
                  <XAxis dataKey="m" tick={{fill:"#1e3a5f",fontSize:11}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:"#1e3a5f",fontSize:10}} axisLine={false} tickLine={false} width={50}
                    tickFormatter={v=>v>=10000?`${(v/10000).toFixed(0)}萬`:String(v)}/>
                  <Tooltip contentStyle={tt} formatter={v=>[`${v.toLocaleString()} 元`,"淨損益"]}/>
                  <Line type="monotone" dataKey="net" stroke="#38bdf8" strokeWidth={2.5}
                    dot={{fill:"#38bdf8",r:5,strokeWidth:0}} activeDot={{r:7}}/>
                </LineChart>
              </ResponsiveContainer>
            }
          </div>
        </>}

        {/* ══ RECORDS ══ */}
        {tab==="records"&&(
          <div className="card">
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:"#f1f5f9"}}>所有記錄</div>
                <div style={{fontSize:11,color:"#1e3a5f",marginTop:2,fontWeight:500}}>{recs.length} 筆</div>
              </div>
              <button className="exp-btn" style={{padding:"8px 14px",fontSize:12}} onClick={exportXlsx}>📊 匯出</button>
            </div>
            <div style={{marginBottom:14}}><CatFilter/></div>

            {filtered.length===0&&(
              <div style={{textAlign:"center",padding:"48px 0"}}>
                <div style={{fontSize:36,marginBottom:10}}>📭</div>
                <div style={{color:"#334155",fontSize:14}}>此類別無記錄</div>
              </div>
            )}
            {filtered.map(r=>{
              const c=CATEGORIES[r.cat];
              const a=Number(r.amt);
              const isIn=r.cat==="income"||(r.cat==="stock"&&a>=0);
              return (
                <div key={r.id} className="rr" style={{cursor:"pointer"}} onClick={()=>openEdit(r)}>
                  <div style={{width:40,height:40,background:c?.color+"15",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontSize:19,flexShrink:0}}>{c?.icon}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                      <span className="chip" style={{background:c?.color+"20",color:c?.color}}>
                        {r.cat==="stock"?(a>=0?"📈 獲利":"📉 虧損"):c?.label}
                      </span>
                      {r.note&&<span style={{fontSize:11,color:"#334155",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:120}}>{r.note}</span>}
                    </div>
                    <div style={{fontSize:11,color:"#1e293b",marginTop:4,fontWeight:600}}>{r.date}</div>
                  </div>
                  <span className="mono" style={{fontSize:15,fontWeight:600,color:isIn?"#4ade80":"#f87171",flexShrink:0}}>
                    {isIn?"+":"-"}{fmtN(Math.abs(a))}
                  </span>
                  <button className="del" onClick={e=>{e.stopPropagation();delRecord(r.id);}}>✕</button>
                </div>
              );
            })}

            {recs.length>0&&(
              <button onClick={async()=>{
                if(!window.confirm("確定清空所有記錄？"))return;
                setRecs([]); persist(user, []);
                setSyncing(true); await syncToCloud(user, []); setSyncing(false);
                showToast("🗑 已清空");
              }} style={{marginTop:16,width:"100%",background:"none",border:"1.5px solid #f8717115",color:"#f8717140",borderRadius:12,padding:"11px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>
                🗑 清空所有記錄
              </button>
            )}
          </div>
        )}

      </div>

      {/* ══ 編輯 Modal ══ */}
      {editRec&&(
        <div style={{position:"fixed",inset:0,zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"center"}}
          onClick={e=>{if(e.target===e.currentTarget)setEditRec(null);}}>
          {/* Backdrop */}
          <div style={{position:"absolute",inset:0,background:"rgba(0,0,0,.7)",backdropFilter:"blur(6px)"}}
            onClick={()=>setEditRec(null)}/>
          {/* Sheet */}
          <div style={{position:"relative",width:"100%",maxWidth:520,background:"linear-gradient(160deg,#0e1826,#080d16)",border:"1px solid #1a2540",borderTop:"1px solid #2a3f60",borderRadius:"24px 24px 0 0",padding:"24px 20px 36px",zIndex:1,maxHeight:"92vh",overflowY:"auto"}}>
            {/* Handle bar */}
            <div style={{width:40,height:4,background:"#1e3a5f",borderRadius:99,margin:"0 auto 20px"}}/>

            <div style={{fontSize:16,fontWeight:700,color:"#f1f5f9",marginBottom:20}}>✏️ 修改記錄</div>

            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              {/* Date */}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>📅 日期</div>
                <input type="date" className="fi" value={editForm.date} onChange={e=>setEditForm(f=>({...f,date:e.target.value}))}/>
              </div>

              {/* Category */}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>🏷 類別</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:7}}>
                  {Object.entries(CATEGORIES).map(([k,v])=>(
                    <button key={k}
                      style={{border:`1.5px solid ${editForm.cat===k?v.color:"#1a2540"}`,background:editForm.cat===k?v.color+"14":"#0a0f1a",
                        color:editForm.cat===k?v.color:"#334155",borderRadius:12,padding:"10px 6px",fontSize:11,fontWeight:600,
                        cursor:"pointer",transition:"all .15s",display:"flex",flexDirection:"column",alignItems:"center",gap:4}}
                      onClick={()=>setEditForm(f=>({...f,cat:k}))}>
                      <span style={{fontSize:20}}>{v.icon}</span>
                      <span>{v.label}</span>
                    </button>
                  ))}
                </div>
              </div>

              {/* Stock sign toggle */}
              {editForm.cat==="stock"&&(
                <div>
                  <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>📌 方向</div>
                  <div style={{display:"flex",borderRadius:12,overflow:"hidden",border:"1.5px solid #1e293b",background:"#0a0f1a"}}>
                    <button style={{flex:1,border:"none",padding:"11px",fontSize:13,fontWeight:700,cursor:"pointer",transition:"all .18s",
                      background:editStockSign===1?"linear-gradient(135deg,#0c4a6e,#14532d)":"transparent",
                      color:editStockSign===1?"#4ade80":"#334155",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}
                      onClick={()=>setEditStockSign(1)}>
                      <span>📈</span><div style={{textAlign:"left"}}><div style={{fontSize:13}}>入金／獲利</div></div>
                    </button>
                    <div style={{width:1,background:"#1e293b"}}/>
                    <button style={{flex:1,border:"none",padding:"11px",fontSize:13,fontWeight:700,cursor:"pointer",transition:"all .18s",
                      background:editStockSign===-1?"linear-gradient(135deg,#7c1d1d,#1c1917)":"transparent",
                      color:editStockSign===-1?"#f87171":"#334155",display:"flex",alignItems:"center",justifyContent:"center",gap:6}}
                      onClick={()=>setEditStockSign(-1)}>
                      <span>📉</span><div style={{textAlign:"left"}}><div style={{fontSize:13}}>出金／虧損</div></div>
                    </button>
                  </div>
                </div>
              )}

              {/* Amount */}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>
                  💵 金額（元）
                  {editForm.cat==="stock"&&<span style={{marginLeft:8,fontSize:10,color:editStockSign===1?"#4ade80":"#f87171",fontWeight:700}}>{editStockSign===1?"＋ 入金/獲利":"－ 出金/虧損"}</span>}
                </div>
                <div style={{position:"relative"}}>
                  {editForm.cat==="stock"&&(
                    <div style={{position:"absolute",left:16,top:"50%",transform:"translateY(-50%)",fontSize:22,fontWeight:700,
                      color:editStockSign===1?"#4ade80":"#f87171",fontFamily:"'DM Mono',monospace",pointerEvents:"none",zIndex:1}}>
                      {editStockSign===1?"+":"−"}
                    </div>
                  )}
                  <input type="number" className="fi mono" placeholder="0" value={editForm.rawAmt} min="0" inputMode="decimal"
                    style={{fontSize:24,letterSpacing:"-.03em",color:CATEGORIES[editForm.cat]?.color||"#e2e8f0",
                      paddingLeft:editForm.cat==="stock"?"40px":"16px"}}
                    onChange={e=>setEditForm(f=>({...f,rawAmt:e.target.value}))}/>
                </div>
              </div>

              {/* Note */}
              <div>
                <div style={{fontSize:10,color:"#334155",letterSpacing:".12em",marginBottom:7,fontWeight:700}}>📝 備註</div>
                <input className="fi" placeholder="備註..." value={editForm.note}
                  onChange={e=>setEditForm(f=>({...f,note:e.target.value}))}
                  onKeyDown={e=>{if(e.key==="Enter")saveEdit()}}/>
              </div>

              {/* Buttons */}
              <div style={{display:"flex",gap:10,marginTop:4}}>
                <button onClick={()=>setEditRec(null)}
                  style={{flex:1,background:"none",border:"1.5px solid #1a2540",color:"#475569",borderRadius:12,padding:"13px",fontSize:14,fontWeight:600,cursor:"pointer"}}>
                  取消
                </button>
                <button onClick={saveEdit} disabled={syncing}
                  style={{flex:2,background:"linear-gradient(135deg,#166534,#0c4a6e)",color:"#fff",border:"none",
                    borderRadius:12,padding:"13px",fontSize:15,fontWeight:700,cursor:"pointer",
                    boxShadow:"0 4px 16px #4ade8020",opacity:syncing?.5:1}}>
                  {syncing?"儲存中...":"✓ 儲存修改"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
