import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, onValue, set, update, remove, get } from "firebase/database";
import firebaseConfig from "./firebaseConfig";

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

const CURRENCIES = [
  { code:"HKD", symbol:"HK$", name:"港幣" },
  { code:"CNY", symbol:"¥",   name:"人民幣" },
  { code:"TWD", symbol:"NT$", name:"台幣" },
  { code:"USD", symbol:"US$", name:"美元" },
  { code:"EUR", symbol:"€",   name:"歐元" },
  { code:"GBP", symbol:"£",   name:"英鎊" },
  { code:"AUD", symbol:"A$",  name:"澳元" },
  { code:"JPY", symbol:"¥",   name:"日元" },
  { code:"KRW", symbol:"₩",   name:"韓元" },
];
const BASE = "HKD";
const DEFAULT_RATES = { CNY:1.07, TWD:0.24, USD:7.78, EUR:8.45, GBP:9.85, AUD:4.92, JPY:0.052, KRW:0.0057 };
const COLORS = ["#f87171","#fb923c","#fbbf24","#a3e635","#34d399","#22d3ee","#60a5fa","#a78bfa","#f472b6","#94a3b8"];
const MY_ID_KEY = "splitpay_myid";
const MY_TRIPS_KEY = "splitpay_trips";

const uid = () => Math.random().toString(36).slice(2,9);
const fmt = (n,d=1) => (isNaN(n)?0:+n).toFixed(d);
const getMyId = () => {
  let id = localStorage.getItem(MY_ID_KEY);
  if (!id) { id = uid(); localStorage.setItem(MY_ID_KEY, id); }
  return id;
};
const loadMyTrips = () => { try { return JSON.parse(localStorage.getItem(MY_TRIPS_KEY)||"[]"); } catch { return []; } };
const saveMyTrips = ids => localStorage.setItem(MY_TRIPS_KEY, JSON.stringify(ids));

// ── Settlement ─────────────────────────────────────────────────────────────
function settleMixed(members, expenses, rates) {
  const hkdBal = {};
  members.forEach(m => hkdBal[m.id] = 0);
  const curNet = {};

  (expenses||[]).forEach(exp => {
    if (!exp.splits?.length) return;
    const useHKD = exp.settleMode !== "original" || exp.currency === BASE;
    if (useHKD) {
      const r = exp.currency === BASE ? 1 : (rates[exp.currency]||1);
      const total = exp.amount * r, share = total / exp.splits.length;
      exp.splits.forEach(sid => { if (sid in hkdBal) hkdBal[sid] -= share; });
      if (exp.paidBy in hkdBal) hkdBal[exp.paidBy] += total;
    } else {
      const share = exp.amount / exp.splits.length;
      exp.splits.forEach(sid => {
        if (sid === exp.paidBy) return;
        const [a,b] = [sid, exp.paidBy].sort(), key = `${a}|${b}|${exp.currency}`;
        if (!curNet[key]) curNet[key] = { a, b, currency: exp.currency, net: 0 };
        if (a === sid) curNet[key].net += share; else curNet[key].net -= share;
      });
    }
  });

  const cred = [], debt = [];
  Object.entries(hkdBal).forEach(([id,b]) => {
    if (b > 0.01) cred.push({id, amt:b}); else if (b < -0.01) debt.push({id, amt:-b});
  });
  cred.sort((a,b)=>b.amt-a.amt); debt.sort((a,b)=>b.amt-a.amt);
  const hkdTxns = []; let ci=0,di=0;
  while (ci<cred.length && di<debt.length) {
    const c=cred[ci],d=debt[di],a=Math.min(c.amt,d.amt);
    hkdTxns.push({from:d.id, to:c.id, currency:"HKD", amount:+a.toFixed(2)});
    c.amt-=a; d.amt-=a; if(c.amt<0.01)ci++; if(d.amt<0.01)di++;
  }
  const curTxns = Object.values(curNet).filter(x=>Math.abs(x.net)>0.01).map(x => x.net>0
    ? {from:x.a, to:x.b, currency:x.currency, amount:+x.net.toFixed(2)}
    : {from:x.b, to:x.a, currency:x.currency, amount:+(-x.net).toFixed(2)});

  const balances = {};
  members.forEach(m => {
    const paid = (expenses||[]).reduce((s,e)=>e.paidBy===m.id?s+e.amount*(e.currency===BASE?1:(rates[e.currency]||1)):s,0);
    const owe  = (expenses||[]).reduce((s,e)=>e.splits?.includes(m.id)?s+e.amount*(e.currency===BASE?1:(rates[e.currency]||1))/(e.splits.length||1):s,0);
    balances[m.id] = paid - owe;
  });
  return { balances, allTxns:[...hkdTxns,...curTxns] };
}

// ── Styles ─────────────────────────────────────────────────────────────────
const C = { bg:"#080c18", card:"rgba(255,255,255,0.04)", border:"rgba(255,255,255,0.09)", accent:"#38bdf8", gold:"#fbbf24", green:"#34d399", red:"#f87171", text:"#e2e8f0", muted:"#64748b", input:"rgba(255,255,255,0.07)" };
const inp = { background:C.input, border:`1px solid ${C.border}`, borderRadius:10, color:C.text, padding:"9px 13px", fontSize:14, outline:"none", width:"100%", boxSizing:"border-box", fontFamily:"inherit" };
const pill = (active, color="#38bdf8") => ({ border:`1px solid ${active?color:C.border}`, borderRadius:20, padding:"5px 13px", fontSize:12, cursor:"pointer", background:active?color+"22":"transparent", color:active?color:C.muted, fontFamily:"inherit", transition:"all .15s" });
const primaryBtn = { background:"linear-gradient(135deg,#0ea5e9,#6366f1)", border:"none", borderRadius:12, color:"#fff", padding:"11px 20px", fontSize:14, fontWeight:700, cursor:"pointer", fontFamily:"inherit", boxShadow:"0 2px 16px rgba(14,165,233,.35)" };
const ghostBtn = { background:"rgba(255,255,255,0.06)", border:`1px solid ${C.border}`, borderRadius:10, color:C.text, padding:"9px 16px", fontSize:13, cursor:"pointer", fontFamily:"inherit" };
const dangerBtn = { background:"rgba(248,113,113,0.12)", border:"1px solid rgba(248,113,113,0.25)", borderRadius:10, color:C.red, padding:"7px 13px", fontSize:12, cursor:"pointer", fontFamily:"inherit" };

// ── UI Helpers ─────────────────────────────────────────────────────────────
function Modal({title, onClose, children, wide}) {
  return (
    <div style={{position:"fixed",inset:0,zIndex:200,background:"rgba(0,0,0,.8)",backdropFilter:"blur(6px)",display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#0f172a",borderRadius:"22px 22px 0 0",padding:"24px 20px 44px",width:"100%",maxWidth:wide?560:440,maxHeight:"92vh",overflowY:"auto",border:`1px solid ${C.border}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <span style={{fontWeight:700,fontSize:17,color:C.text}}>{title}</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:C.muted,fontSize:24,cursor:"pointer",lineHeight:1}}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}
function Tag({color,name}) { return <span style={{display:"inline-flex",alignItems:"center",background:color+"22",border:`1px solid ${color}44`,borderRadius:20,padding:"2px 10px",fontSize:12,color}}>{name}</span>; }
function Avatar({member,size=36}) { return <div style={{width:size,height:size,borderRadius:"50%",background:member.color+"33",border:`2px solid ${member.color}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*.38,fontWeight:700,color:member.color,flexShrink:0}}>{member.name[0]}</div>; }

// ── Calculator ─────────────────────────────────────────────────────────────
function Calculator({ onConfirm, onClose }) {
  const [display, setDisplay] = useState("0");
  const [prev, setPrev]       = useState(null);
  const [op, setOp]           = useState(null);
  const [fresh, setFresh]     = useState(true);

  const press = (val) => {
    if (val === "C") { setDisplay("0"); setPrev(null); setOp(null); setFresh(true); return; }
    if (val === "⌫") { setDisplay(d => d.length>1 ? d.slice(0,-1) : "0"); return; }
    if (["+","-","×","÷"].includes(val)) {
      setPrev(parseFloat(display)); setOp(val); setFresh(true); return;
    }
    if (val === "=") {
      if (prev===null||!op) return;
      const a=prev, b=parseFloat(display);
      const res = op==="+"?a+b:op==="-"?a-b:op==="×"?a*b:op==="÷"&&b!==0?a/b:display;
      setDisplay(String(+res.toFixed(4))); setPrev(null); setOp(null); setFresh(true); return;
    }
    if (val === "." ) {
      if (fresh) { setDisplay("0."); setFresh(false); return; }
      if (!display.includes(".")) setDisplay(d=>d+"."); return;
    }
    if (fresh) { setDisplay(String(val)); setFresh(false); }
    else setDisplay(d => d==="0" ? String(val) : d+String(val));
  };

  const rows = [["C","⌫","÷","×"],[7,8,9,"-"],[4,5,6,"+"],[1,2,3,"="],[0,".","✓",""]];
  return (
    <div style={{background:"#1e293b",borderRadius:16,padding:16,marginBottom:16,border:`1px solid ${C.border}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
        <span style={{fontSize:12,color:C.muted}}>計算機</span>
        <button onClick={onClose} style={{background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:16}}>×</button>
      </div>
      <div style={{background:"rgba(0,0,0,.4)",borderRadius:10,padding:"10px 14px",marginBottom:12,textAlign:"right"}}>
        <div style={{fontSize:11,color:C.muted,height:16}}>{prev!==null?`${prev} ${op||""}`:""}</div>
        <div style={{fontSize:28,fontWeight:700,color:C.text,fontFamily:"monospace"}}>{display}</div>
      </div>
      {rows.map((row,ri)=>(
        <div key={ri} style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:6}}>
          {row.map((k,ki)=>{
            if(k==="")return<div key={ki}/>;
            const isOp=["+","-","×","÷"].includes(k);
            const isEq=k==="=";
            const isChk=k==="✓";
            const isDel=k==="⌫"||k==="C";
            return(
              <button key={ki} onClick={()=>isChk?onConfirm(parseFloat(display)||0):press(k)} style={{
                border:"none",borderRadius:10,padding:"14px 0",fontSize:isChk?18:16,fontWeight:600,cursor:"pointer",fontFamily:"monospace",
                background:isChk?"linear-gradient(135deg,#10b981,#3b82f6)":isEq?"rgba(56,189,248,.25)":isOp?"rgba(56,189,248,.12)":isDel?"rgba(248,113,113,.12)":"rgba(255,255,255,.07)",
                color:isChk?"#fff":isOp||isEq?C.accent:isDel?C.red:C.text,
              }}>{k}</button>
            );
          })}
        </div>
      ))}
      <div style={{fontSize:11,color:C.muted,textAlign:"center",marginTop:4}}>按 ✓ 填入金額</div>
    </div>
  );
}



// ── Export ─────────────────────────────────────────────────────────────────
function exportExcel(trip, members, expenses, rates) {
  const {balances,allTxns} = settleMixed(members,expenses,rates);
  const mm = Object.fromEntries(members.map(m=>[m.id,m]));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet((expenses||[]).map(e=>({
    "日期": e.date||new Date(e.createdAt).toLocaleDateString("zh-TW"),
    "描述": e.desc, "金額": e.amount, "貨幣": e.currency,
    "港幣": +((e.amount*(e.currency===BASE?1:(rates[e.currency]||1))).toFixed(2)),
    "付款人": mm[e.paidBy]?.name||"",
    "分攤": (e.splits||[]).map(s=>mm[s]?.name||"").join("、"),
    "結算": e.settleMode==="original"?"原幣":"港幣",
  }))), "帳目");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(members.map(m=>({
    "成員":m.name, "餘額HKD":+fmt(balances[m.id]||0,2),
    "狀態":(balances[m.id]||0)>0.01?"待收":(balances[m.id]||0)<-0.01?"待付":"已清"
  }))), "餘額");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allTxns.map(t=>({
    "付款": mm[t.from]?.name||"", "收款": mm[t.to]?.name||"", "幣": t.currency, "金額": +t.amount.toFixed(2)
  }))), "結算");
  XLSX.writeFile(wb, `${trip.name}_結帳.xlsx`);
}

function exportPDF(trip, members, expenses, rates) {
  const {balances,allTxns} = settleMixed(members,expenses,rates);
  const mm = Object.fromEntries(members.map(m=>[m.id,m]));
  const cs = c => CURRENCIES.find(x=>x.code===c);
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${trip.name}</title>
  <style>body{font-family:'Noto Sans TC',sans-serif;padding:32px;color:#1e293b;max-width:800px;margin:0 auto}h1{font-size:22px;color:#0ea5e9}h2{font-size:15px;color:#334155;border-bottom:2px solid #e2e8f0;padding-bottom:5px;margin-top:24px}table{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px}th{background:#f1f5f9;padding:7px 9px;text-align:left;color:#475569;font-weight:600}td{padding:7px 9px;border-bottom:1px solid #f1f5f9}.g{color:#059669;font-weight:700}.r{color:#dc2626;font-weight:700}.tag{display:inline-block;background:#e0f2fe;color:#0369a1;border-radius:8px;padding:1px 7px;font-size:11px;margin:1px}@media print{body{padding:12px}}</style>
  </head><body>
  <h1>💸 ${trip.name}</h1><p style="color:#64748b;font-size:12px">生成：${new Date().toLocaleString("zh-TW")}</p>
  <h2>帳目（${(expenses||[]).length}筆）</h2>
  <table><thead><tr><th>日期</th><th>描述</th><th>金額</th><th>港幣</th><th>付款</th><th>分攤</th></tr></thead><tbody>
  ${(expenses||[]).sort((a,b)=>(b.date||"").localeCompare(a.date||"")).map(e=>{const r=e.currency===BASE?1:(rates[e.currency]||1);return`<tr><td>${e.date||new Date(e.createdAt).toLocaleDateString("zh-TW")}</td><td>${e.desc}</td><td>${cs(e.currency)?.symbol}${e.amount.toLocaleString()}</td><td>HK$${fmt(e.amount*r)}</td><td><span class="tag">${mm[e.paidBy]?.name||""}</span></td><td>${(e.splits||[]).map(s=>`<span class="tag">${mm[s]?.name||""}</span>`).join("")}</td></tr>`;}).join("")}
  </tbody></table>
  <h2>個人餘額</h2>
  <table><thead><tr><th>成員</th><th>餘額(HKD)</th><th>狀態</th></tr></thead><tbody>
  ${members.map(m=>{const b=balances[m.id]||0;return`<tr><td>${m.name}</td><td class="${b>0.01?"g":b<-0.01?"r":""}">${b>0.01?"+":""}HK$${fmt(b)}</td><td>${b>0.01?"💰待收":b<-0.01?"💸待付":"✅已清"}</td></tr>`;}).join("")}
  </tbody></table>
  <h2>結算（${allTxns.length}筆）</h2>
  <table><thead><tr><th>付款方</th><th></th><th>收款方</th><th>金額</th></tr></thead><tbody>
  ${allTxns.map(t=>`<tr><td><b>${mm[t.from]?.name||""}</b></td><td style="color:#94a3b8">→</td><td><b>${mm[t.to]?.name||""}</b></td><td class="r">${cs(t.currency)?.symbol||""}${t.amount.toLocaleString()}</td></tr>`).join("")}
  ${allTxns.length===0?`<tr><td colspan="4" style="text-align:center;color:#059669;padding:14px">🎉已結清！</td></tr>`:""}
  </tbody></table></body></html>`;
  const w = window.open("","_blank"); w.document.write(html); w.document.close(); setTimeout(()=>w.print(),500);
}

// ── Expense Form (shared by Add & Edit) ───────────────────────────────────
function ExpenseForm({initial, members, rates, onSubmit, submitLabel}) {
  const [desc, setDesc] = useState(initial?.desc||"");
  const [amount, setAmount] = useState(initial?.amount?String(initial.amount):"");
  const [currency, setCurrency] = useState(initial?.currency||"HKD");
  const [paidBy, setPaidBy] = useState(initial?.paidBy||(members[0]?.id||""));
  const [splits, setSplits] = useState(initial?.splits||members.map(m=>m.id));
  const [date, setDate] = useState(initial?.date||new Date().toISOString().split("T")[0]);
  const [showCalc, setShowCalc] = useState(false);
  const toggle = id => setSplits(s=>s.includes(id)?s.filter(x=>x!==id):[...s,id]);
  const r = currency===BASE?1:(rates[currency]||1);
  const hkd = (parseFloat(amount)||0)*r;

  const submit = () => {
    if (!desc.trim()||!amount||splits.length===0) { alert("請填寫描述、金額，並選擇分攤成員"); return; }
    onSubmit({ desc:desc.trim(), amount:parseFloat(amount), currency, paidBy, splits, date, settleMode: initial?.settleMode||"hkd" });
  };

  return (
    <>
      <div style={{display:"grid",gridTemplateColumns:"1fr 130px",gap:10,marginBottom:14}}>
        <div>
          <div style={{fontSize:12,color:C.muted,marginBottom:6}}>描述</div>
          <input style={inp} placeholder="晚餐、門票、Uber…" value={desc} onChange={e=>setDesc(e.target.value)}/>
        </div>
        <div>
          <div style={{fontSize:12,color:C.muted,marginBottom:6}}>日期</div>
          <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{...inp}}/>
        </div>
      </div>
      {showCalc && <Calculator onClose={()=>setShowCalc(false)} onConfirm={v=>{setAmount(String(v));setShowCalc(false);}}/>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:4}}>
        <div>
          <div style={{fontSize:12,color:C.muted,marginBottom:6,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span>金額</span>
            <button onClick={()=>setShowCalc(s=>!s)} style={{background:"rgba(56,189,248,.1)",border:"none",borderRadius:6,padding:"2px 8px",fontSize:11,color:C.accent,cursor:"pointer",fontFamily:"inherit"}}>🧮 計算機</button>
          </div>
          <input type="number" min="0" step="0.01" style={inp} placeholder="0.00" value={amount} onChange={e=>setAmount(e.target.value)}/>
        </div>
        <div>
          <div style={{fontSize:12,color:C.muted,marginBottom:6}}>貨幣</div>
          <select value={currency} onChange={e=>setCurrency(e.target.value)} style={{...inp,cursor:"pointer"}}>
            {CURRENCIES.map(c=><option key={c.code} value={c.code}>{c.code} {c.name}</option>)}
          </select>
        </div>
      </div>
      {currency!==BASE && hkd>0 && <div style={{fontSize:12,color:C.muted,marginBottom:14,padding:"7px 12px",background:"rgba(56,189,248,.08)",borderRadius:8}}>≈ HK${fmt(hkd)}（每人約 HK${splits.length>0?fmt(hkd/splits.length):"0"}）</div>}
      {currency===BASE && <div style={{marginBottom:14}}/>}
      <div style={{marginBottom:14}}>
        <div style={{fontSize:12,color:C.muted,marginBottom:8}}>由誰付款</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:7}}>
          {members.map(m=><button key={m.id} onClick={()=>setPaidBy(m.id)} style={pill(paidBy===m.id,m.color)}>{m.name}</button>)}
        </div>
      </div>
      <div style={{marginBottom:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <div style={{fontSize:12,color:C.muted}}>誰需要分攤（{splits.length}人）</div>
          <div style={{display:"flex",gap:10}}>
            <span style={{fontSize:12,color:C.accent,cursor:"pointer"}} onClick={()=>setSplits(members.map(m=>m.id))}>全選</span>
            <span style={{fontSize:12,color:C.muted,cursor:"pointer"}} onClick={()=>setSplits([])}>清空</span>
          </div>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:7}}>
          {members.map(m=><button key={m.id} onClick={()=>toggle(m.id)} style={pill(splits.includes(m.id),m.color)}>{splits.includes(m.id)?"✓ ":""}{m.name}</button>)}
        </div>
      </div>
      <button onClick={submit} style={{...primaryBtn,width:"100%"}}>{submitLabel}</button>
    </>
  );
}

// ── Main App ───────────────────────────────────────────────────────────────
export default function App() {
  const myId = useRef(getMyId()).current;
  const [screen, setScreen]     = useState("home");
  const [tripId, setTripId]     = useState(null);
  const [trip, setTrip]         = useState(null);
  const [loading, setLoading]   = useState(false);
  const [syncing, setSyncing]   = useState(false);
  const [myTripIds, setMyTripIds] = useState(loadMyTrips);
  const [tripPreviews, setTripPreviews] = useState({});

  const addMyTrip = id => {
    const ids = [...new Set([...myTripIds, id])];
    setMyTripIds(ids); saveMyTrips(ids);
  };

  // Subscribe to trip previews for home screen
  useEffect(() => {
    const unsubs = myTripIds.map(id => {
      const r = ref(db, `trips/${id}/meta`);
      return onValue(r, snap => {
        if (snap.exists()) setTripPreviews(p=>({...p,[id]:snap.val()}));
        else setTripPreviews(p=>{const n={...p};delete n[id];return n;});
      });
    });
    return () => unsubs.forEach(u=>u());
  }, [myTripIds.join(",")]);

  // Subscribe to current trip
  useEffect(() => {
    if (!tripId) return;
    setLoading(true);
    const r = ref(db, `trips/${tripId}`);
    const unsub = onValue(r, snap => {
      if (snap.exists()) {
        const data = snap.val();
        setTrip({
          ...data.meta, id:tripId,
          members:  data.members  ? Object.values(data.members)  : [],
          expenses: data.expenses ? Object.values(data.expenses) : [],
          rates: data.rates || DEFAULT_RATES,
        });
      } else setTrip(null);
      setLoading(false);
    });
    return () => unsub();
  }, [tripId]);

  // Firebase helpers
  const fbSet    = async (path,val) => { setSyncing(true); await set(ref(db,path),val);    setSyncing(false); };
  const fbUpdate = async (path,val) => { setSyncing(true); await update(ref(db,path),val); setSyncing(false); };
  const fbRemove = async path       => { setSyncing(true); await remove(ref(db,path));      setSyncing(false); };

  // Trip actions
  const createTrip = async (name, memberName) => {
    const tid = uid();
    const me = { id:myId, name:memberName, color:COLORS[0] };
    await fbSet(`trips/${tid}`, { meta:{name,createdAt:Date.now(),archived:false}, members:{[myId]:me}, rates:DEFAULT_RATES });
    addMyTrip(tid); setTripId(tid); setScreen("trip");
  };

  // FIX: Join uses onValue with onlyOnce:true — most reliable way to fetch once
  const joinTrip = async (code, memberName) => {
    if (!code.trim()) { alert("請輸入行程代碼"); return; }
    const tripRef = ref(db, `trips/${code.trim()}`);
    let resolved = false;
    const timeout = setTimeout(() => {
      if (!resolved) alert("連線逾時，請檢查代碼是否正確");
    }, 8000);
    onValue(tripRef, async snap => {
      if (resolved) return;
      resolved = true;
      clearTimeout(timeout);
      if (!snap.exists()) { alert("找不到此行程代碼，請確認是否正確"); return; }
      const data = snap.val();
      if (data.meta?.archived) { alert("此行程已封存，無法加入"); return; }
      const existingMembers = data.members ? Object.keys(data.members) : [];
      const alreadyIn = existingMembers.includes(myId);
      if (!alreadyIn) {
        const color = COLORS[existingMembers.length % COLORS.length];
        await fbUpdate(`trips/${code.trim()}/members`, { [myId]: { id:myId, name:memberName, color } });
      }
      addMyTrip(code.trim());
      setTripId(code.trim());
      setScreen("trip");
    }, { onlyOnce: true });
  };

  const archiveTrip = async tid => {
    if (!window.confirm("確定封存此行程？封存後不可再新增帳目。")) return;
    await fbUpdate(`trips/${tid}/meta`, { archived:true });
  };
  const deleteTrip = async tid => {
    if (!window.confirm("確定刪除？此操作不可復原！")) return;
    await fbRemove(`trips/${tid}`);
    const ids = myTripIds.filter(x=>x!==tid);
    setMyTripIds(ids); saveMyTrips(ids);
  };

  // Expense actions — all members can edit/delete any expense
  const addExpense    = async exp => fbUpdate(`trips/${tripId}/expenses`, {[exp.id]: exp});
  const updateExpense = async exp => fbUpdate(`trips/${tripId}/expenses`, {[exp.id]: exp});
  const deleteExpense = async eid => fbRemove(`trips/${tripId}/expenses/${eid}`);

  // Member actions
  const addMember    = async name => {
    const me = { id:uid(), name, color:COLORS[(trip?.members?.length||0)%COLORS.length] };
    await fbUpdate(`trips/${tripId}/members`, {[me.id]: me});
  };
  const updateMember = async (mid, name) => fbUpdate(`trips/${tripId}/members/${mid}`, {name});
  const deleteMember = async mid => {
    if (!window.confirm("確定刪除此成員？與其相關的帳目記錄不受影響。")) return;
    await fbRemove(`trips/${tripId}/members/${mid}`);
  };

  const updateRates  = async rates => fbSet(`trips/${tripId}/rates`, rates);

  // UI state
  const [tab, setTab]               = useState("expenses");
  const [showNewTrip, setShowNewTrip] = useState(false);
  const [showJoin, setShowJoin]     = useState(false);
  const [showAddExp, setShowAddExp] = useState(false);
  const [editingExp, setEditingExp] = useState(null);
  const [editingMember, setEditingMember] = useState(null);
  const [showRates, setShowRates]   = useState(false);
  const [newTripName, setNewTripName] = useState("");
  const [myName, setMyName]         = useState("");
  const [joinCode, setJoinCode]     = useState("");
  const [newMemberName, setNewMemberName] = useState("");
  const [editMemberName, setEditMemberName] = useState("");
  const [ratesLoading, setRatesLoading] = useState(false);
  const [ratesDone, setRatesDone]   = useState(false);
  const [joining, setJoining]       = useState(false);

  const members   = trip?.members  || [];
  const expenses  = trip?.expenses || [];
  const rates     = trip?.rates    || DEFAULT_RATES;
  const memberMap = Object.fromEntries(members.map(m=>[m.id,m]));
  const myMember  = memberMap[myId];
  const {balances, allTxns} = trip ? settleMixed(members,expenses,rates) : {balances:{},allTxns:[]};
  const myBalance = balances[myId] || 0;

  // Sort expenses: newest date first, then by createdAt
  const sortedExpenses = [...expenses].sort((a,b) => {
    const da = a.date||"", db2 = b.date||"";
    if (da !== db2) return db2.localeCompare(da);
    return (b.createdAt||0) - (a.createdAt||0);
  });

  const refreshRates = async () => {
    setRatesLoading(true);
    try {
      const res  = await fetch("https://api.exchangerate-api.com/v4/latest/HKD");
      const json = await res.json(); const r = json.rates;
      await updateRates({CNY:+(1/r.CNY).toFixed(4),TWD:+(1/r.TWD).toFixed(4),USD:+(1/r.USD).toFixed(4),EUR:+(1/r.EUR).toFixed(4),GBP:+(1/r.GBP).toFixed(4),AUD:+(1/r.AUD).toFixed(4),JPY:+(1/r.JPY).toFixed(5),KRW:+(1/r.KRW).toFixed(6)});
      setRatesDone(true); setTimeout(()=>setRatesDone(false),2500);
    } catch {}
    setRatesLoading(false);
  };

  // ── HOME ────────────────────────────────────────────────────────────────
  if (screen === "home") return (
    <div style={{minHeight:"100vh",background:C.bg,padding:24,fontFamily:"'Noto Sans TC',sans-serif"}}>
      <div style={{maxWidth:480,margin:"0 auto"}}>
        <div style={{textAlign:"center",padding:"32px 0 28px"}}>
          <div style={{fontSize:48,marginBottom:8}}>✈️</div>
          <div style={{fontFamily:"'Playfair Display',serif",fontSize:30,color:C.accent,fontWeight:700}}>SplitPay</div>
          <div style={{fontSize:13,color:C.muted,marginTop:4}}>旅行記帳・即時同步・智能分帳</div>
        </div>
        <div style={{display:"flex",gap:10,marginBottom:28}}>
          <button onClick={()=>setShowNewTrip(true)} style={{...primaryBtn,flex:1}}>＋ 新增行程</button>
          <button onClick={()=>setShowJoin(true)} style={{...ghostBtn,flex:1}}>🔗 加入行程</button>
        </div>
        {myTripIds.length===0 && <div style={{textAlign:"center",color:C.muted,padding:"40px 0",fontSize:14}}>還沒有行程，點上方按鈕開始！</div>}
        {[false,true].map(archived => {
          const list = myTripIds.filter(id => tripPreviews[id]?.archived === archived);
          if (!list.length) return null;
          return (
            <div key={String(archived)}>
              <div style={{fontSize:11,color:C.muted,marginBottom:8,marginTop:archived?20:0,letterSpacing:1,textTransform:"uppercase"}}>{archived?"📦 已封存":"📍 進行中"}</div>
              {list.map(id => {
                const meta = tripPreviews[id];
                if (!meta) return <div key={id} style={{background:C.card,borderRadius:14,padding:14,marginBottom:8,color:C.muted,fontSize:13}}>載入中…</div>;
                return (
                  <div key={id} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:16,marginBottom:10}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                      <div style={{flex:1}}>
                        <div style={{fontSize:16,fontWeight:700,color:archived?C.muted:C.text}}>{meta.name}</div>
                        <div style={{fontSize:11,color:C.muted,marginTop:2}}>代碼：<span style={{fontFamily:"monospace",letterSpacing:.5}}>{id}</span></div>
                      </div>
                      <div style={{display:"flex",gap:6,flexWrap:"wrap",justifyContent:"flex-end",marginLeft:8}}>
                        {!archived && <button onClick={()=>{setTripId(id);setScreen("trip");setTab("expenses");}} style={{...primaryBtn,padding:"7px 14px",fontSize:12}}>進入</button>}
                        {!archived && <button onClick={()=>archiveTrip(id)} style={{...ghostBtn,padding:"7px 12px",fontSize:12}}>封存</button>}
                        <button onClick={()=>deleteTrip(id)} style={{...dangerBtn,padding:"7px 12px",fontSize:12}}>刪除</button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>

      {showNewTrip && (
        <Modal title="新增行程" onClose={()=>setShowNewTrip(false)}>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>行程名稱</div>
            <input style={inp} placeholder="例：202607 台北" value={newTripName} onChange={e=>setNewTripName(e.target.value)}/>
          </div>
          <div style={{marginBottom:20}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>你的名字</div>
            <input style={inp} placeholder="例：Alexis" value={myName} onChange={e=>setMyName(e.target.value)}/>
          </div>
          <button onClick={async()=>{
            if (!newTripName.trim()||!myName.trim()) return;
            await createTrip(newTripName.trim(), myName.trim());
            setShowNewTrip(false); setNewTripName(""); setMyName("");
          }} style={{...primaryBtn,width:"100%"}}>建立行程 →</button>
        </Modal>
      )}

      {showJoin && (
        <Modal title="加入行程" onClose={()=>{if(!joining){setShowJoin(false);}}}>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>你的名字</div>
            <input style={inp} placeholder="例：Alexis" value={myName} onChange={e=>setMyName(e.target.value)}/>
          </div>
          <div style={{marginBottom:6}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>行程代碼</div>
            <input style={inp} placeholder="朋友分享的代碼（注意大小寫）" value={joinCode} onChange={e=>setJoinCode(e.target.value)}/>
          </div>
          <div style={{fontSize:11,color:C.muted,marginBottom:20}}>代碼在行程頁「👥成員」分頁找到</div>
          {joining && <div style={{textAlign:"center",color:C.gold,marginBottom:14,fontSize:13}}>⏳ 連接中，請稍候…</div>}
          <button onClick={async()=>{
            if (!myName.trim()||!joinCode.trim()) return;
            setJoining(true);
            await joinTrip(joinCode.trim(), myName.trim());
            setJoining(false); setShowJoin(false); setJoinCode(""); setMyName("");
          }} disabled={joining} style={{...primaryBtn,width:"100%",opacity:joining?0.6:1}}>加入行程 →</button>
        </Modal>
      )}
    </div>
  );

  // Loading
  if (loading && !trip) return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",color:C.muted,fontFamily:"sans-serif",flexDirection:"column",gap:12}}>
      <div style={{fontSize:32}}>⏳</div><div>載入行程資料…</div>
    </div>
  );

  const tabs = [{key:"expenses",icon:"📋",label:"帳目"},{key:"balances",icon:"💰",label:"結算"},{key:"members",icon:"👥",label:"成員"},{key:"export",icon:"📤",label:"匯出"}];

  return (
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Noto Sans TC',sans-serif",paddingBottom:80}}>
      {/* Header */}
      <div style={{position:"sticky",top:0,zIndex:50,background:"rgba(8,12,24,.94)",backdropFilter:"blur(12px)",borderBottom:`1px solid ${C.border}`}}>
        <div style={{maxWidth:480,margin:"0 auto",padding:"14px 16px 0"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <button onClick={()=>{setScreen("home");setTrip(null);setTripId(null);}} style={{background:"none",border:"none",color:C.muted,fontSize:20,cursor:"pointer",padding:0}}>←</button>
              <div>
                <div style={{fontFamily:"'Playfair Display',serif",fontSize:17,color:C.accent,fontWeight:700}}>{trip?.name||"…"}</div>
                <div style={{fontSize:10,color:C.muted,display:"flex",gap:6,alignItems:"center"}}>
                  {syncing  && <span style={{color:C.gold}}>⏳ 同步中</span>}
                  {!syncing && <span style={{color:C.green}}>● 即時同步</span>}
                  {trip?.archived && <span style={{color:C.gold,background:"rgba(251,191,36,.15)",borderRadius:6,padding:"1px 6px"}}>已封存</span>}
                </div>
              </div>
            </div>
            <button onClick={()=>setShowRates(true)} style={{...ghostBtn,padding:"6px 11px",fontSize:12}}>💱 匯率</button>
          </div>
          <div style={{margin:"8px 0",padding:"9px 14px",borderRadius:12,background:myBalance>0.01?"rgba(52,211,153,.1)":myBalance<-0.01?"rgba(248,113,113,.1)":C.card,border:`1px solid ${myBalance>0.01?"rgba(52,211,153,.25)":myBalance<-0.01?"rgba(248,113,113,.25)":C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontSize:12,color:C.muted}}>我的結餘 ({myMember?.name||"我"})</span>
            <span style={{fontSize:16,fontWeight:700,fontFamily:"monospace",color:myBalance>0.01?C.green:myBalance<-0.01?C.red:C.muted}}>
              {myBalance>0.01?`+HK$${fmt(myBalance)} 待收`:myBalance<-0.01?`-HK$${fmt(Math.abs(myBalance))} 待付`:"✅ 已結清"}
            </span>
          </div>
          <div style={{display:"flex"}}>
            {tabs.map(t=>(
              <button key={t.key} onClick={()=>setTab(t.key)} style={{flex:1,border:"none",background:"none",cursor:"pointer",padding:"8px 4px 10px",fontSize:11,fontFamily:"inherit",color:tab===t.key?C.accent:C.muted,borderBottom:`2px solid ${tab===t.key?C.accent:"transparent"}`,transition:"all .15s"}}>
                {t.icon} {t.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div style={{maxWidth:480,margin:"0 auto",padding:"16px 16px 0"}}>

        {/* ── EXPENSES ── */}
        {tab==="expenses" && <>
          {!trip?.archived && <button onClick={()=>setShowAddExp(true)} style={{...primaryBtn,width:"100%",marginBottom:16}}>＋ 新增帳目</button>}
          {sortedExpenses.length===0 && <div style={{textAlign:"center",color:C.muted,padding:"50px 0",fontSize:14}}>還沒有帳目</div>}
          {sortedExpenses.map(exp => {
            const payer   = memberMap[exp.paidBy];
            const curInfo = CURRENCIES.find(c=>c.code===exp.currency);
            const r       = exp.currency===BASE ? 1 : (rates[exp.currency]||1);
            const hkd     = exp.amount * r;
            return (
              <div key={exp.id} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:"14px 16px",marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                  <div style={{flex:1}}>
                    <div style={{fontSize:15,fontWeight:600,color:C.text}}>{exp.desc}</div>
                    <div style={{fontSize:11,color:C.muted,marginTop:2}}>{exp.date||new Date(exp.createdAt).toLocaleDateString("zh-TW")}</div>
                  </div>
                  <div style={{textAlign:"right",flexShrink:0,marginLeft:10}}>
                    <div style={{fontSize:16,fontWeight:700,color:C.accent,fontFamily:"monospace"}}>{curInfo?.symbol}{exp.amount.toLocaleString()}</div>
                    {exp.currency!==BASE && <div style={{fontSize:11,color:C.muted}}>≈ HK${fmt(hkd)}</div>}
                    <div style={{fontSize:10,color:C.muted}}>每人 HK${fmt(hkd/(exp.splits?.length||1))}</div>
                  </div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap",marginBottom:6}}>
                  <span style={{fontSize:11,color:C.muted}}>付款：</span>
                  {payer && <Tag color={payer.color} name={payer.name}/>}
                  <span style={{fontSize:11,color:C.muted}}>｜分攤 {exp.splits?.length||0} 人</span>
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:8}}>
                  {(exp.splits||[]).map(sid => memberMap[sid] ? <Tag key={sid} color={memberMap[sid].color} name={memberMap[sid].name}/> : null)}
                </div>
                {exp.currency!==BASE && (
                  <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
                    <span style={{fontSize:11,color:C.muted}}>結算幣值：</span>
                    <div style={{display:"flex",gap:4,background:"rgba(255,255,255,0.05)",borderRadius:8,padding:2}}>
                      {[["hkd","HK$ 港幣"],["original",`${curInfo?.symbol} 原幣`]].map(([m,l])=>(
                        <button key={m} onClick={()=>{if(!trip?.archived)updateExpense({...exp,settleMode:m});}}
                          style={{border:"none",borderRadius:6,padding:"3px 9px",fontSize:11,cursor:trip?.archived?"default":"pointer",fontFamily:"inherit",
                            background:(exp.settleMode||"hkd")===m?"rgba(56,189,248,.3)":"transparent",
                            color:(exp.settleMode||"hkd")===m?C.accent:C.muted,transition:"all .15s"}}>{l}</button>
                      ))}
                    </div>
                  </div>
                )}
                {!trip?.archived && (
                  <div style={{display:"flex",gap:8,marginTop:4}}>
                    <button onClick={()=>setEditingExp(exp)} style={{...ghostBtn,padding:"6px 12px",fontSize:12,flex:1}}>✏️ 修改</button>
                    <button onClick={()=>{if(window.confirm(`確定刪除「${exp.desc}」？`))deleteExpense(exp.id);}} style={{...dangerBtn,padding:"6px 12px",fontSize:12,flex:1}}>🗑 刪除</button>
                  </div>
                )}
              </div>
            );
          })}
        </>}

        {/* ── BALANCES ── */}
        {tab==="balances" && <>
          <div style={{fontSize:12,color:C.muted,marginBottom:10,fontWeight:600}}>個人餘額</div>
          {members.map(m => {
            const b    = balances[m.id]||0;
            const paid = expenses.reduce((s,e)=>e.paidBy===m.id?s+e.amount*(e.currency===BASE?1:(rates[e.currency]||1)):s,0);
            const owe  = expenses.reduce((s,e)=>e.splits?.includes(m.id)?s+e.amount*(e.currency===BASE?1:(rates[e.currency]||1))/(e.splits.length||1):s,0);
            return (
              <div key={m.id} style={{background:C.card,border:`1px solid ${m.id===myId?m.color+"44":C.border}`,borderRadius:14,padding:"14px 16px",marginBottom:10}}>
                <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:10}}>
                  <Avatar member={m}/>
                  <div style={{flex:1}}><div style={{fontSize:15,fontWeight:600,color:C.text}}>{m.name}{m.id===myId&&" （我）"}</div></div>
                  <div style={{fontFamily:"monospace",fontWeight:700,fontSize:16,color:b>0.01?C.green:b<-0.01?C.red:C.muted}}>
                    {b>0.01?`+HK$${fmt(b)}`:b<-0.01?`-HK$${fmt(Math.abs(b))}`:"✅"}
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                  {[["💰 實際花費",`HK$${fmt(owe)}`],["💳 已付出",`HK$${fmt(paid)}`]].map(([l,v])=>(
                    <div key={l} style={{background:"rgba(255,255,255,0.03)",borderRadius:8,padding:"8px 10px"}}>
                      <div style={{fontSize:10,color:C.muted,marginBottom:2}}>{l}</div>
                      <div style={{fontSize:13,fontWeight:600,color:C.text,fontFamily:"monospace"}}>{v}</div>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
          <div style={{borderTop:`1px solid ${C.border}`,paddingTop:16,marginTop:8}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
              <div style={{fontSize:13,color:C.text,fontWeight:600}}>💡 結算清單</div>
              <div style={{fontSize:11,color:C.muted}}>{allTxns.length} 筆轉帳</div>
            </div>
            <div style={{fontSize:11,color:C.muted,marginBottom:12}}>HKD 帳目智能合併；原幣帳目按幣種列出</div>
            {allTxns.length===0 && <div style={{textAlign:"center",color:C.green,padding:"20px 0",fontSize:14}}>🎉 大家已結清！</div>}
            {CURRENCIES.map(cur => {
              const group = allTxns.filter(t=>t.currency===cur.code);
              if (!group.length) return null;
              const isHKD = cur.code===BASE;
              return (
                <div key={cur.code} style={{marginBottom:16}}>
                  <div style={{fontSize:11,fontWeight:600,marginBottom:7}}>
                    <span style={{background:isHKD?"rgba(52,211,153,.15)":"rgba(251,191,36,.15)",borderRadius:8,padding:"2px 9px",color:isHKD?C.green:C.gold}}>
                      {cur.symbol} {cur.code} {cur.name}{isHKD?" · 最少轉帳":""}
                    </span>
                  </div>
                  {group.map((t,i) => {
                    const from=memberMap[t.from], to=memberMap[t.to], isMe=t.from===myId;
                    return (
                      <div key={i} style={{background:isMe?"rgba(248,113,113,.08)":C.card,border:`1px solid ${isMe?"rgba(248,113,113,.2)":C.border}`,borderRadius:12,padding:"11px 14px",marginBottom:6,display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                        <Tag color={from?.color||"#888"} name={from?.name||"?"}/>
                        <span style={{color:C.muted,fontSize:13}}>→</span>
                        <Tag color={to?.color||"#888"} name={to?.name||"?"}/>
                        <span style={{marginLeft:"auto",fontFamily:"monospace",fontWeight:700,fontSize:15,color:isMe?C.red:C.text}}>{cur.symbol}{t.amount.toLocaleString()}</span>
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        </>}

        {/* ── MEMBERS ── */}
        {tab==="members" && <>
          {!trip?.archived && (
            <div style={{display:"flex",gap:8,marginBottom:16}}>
              <input style={{...inp,flex:1}} placeholder="新增成員名字" value={newMemberName} onChange={e=>setNewMemberName(e.target.value)}
                onKeyDown={e=>e.key==="Enter"&&newMemberName.trim()&&addMember(newMemberName.trim()).then(()=>setNewMemberName(""))}/>
              <button onClick={()=>newMemberName.trim()&&addMember(newMemberName.trim()).then(()=>setNewMemberName(""))} style={{...primaryBtn,padding:"9px 16px",whiteSpace:"nowrap"}}>加入</button>
            </div>
          )}
          {members.map(m => (
            <div key={m.id} style={{display:"flex",alignItems:"center",gap:12,background:C.card,border:`1px solid ${m.id===myId?m.color+"44":C.border}`,borderRadius:12,padding:"12px 14px",marginBottom:8}}>
              <Avatar member={m}/>
              <div style={{flex:1}}>
                <div style={{fontSize:14,color:C.text,fontWeight:600}}>{m.name}{m.id===myId&&" （我）"}</div>
              </div>
              {!trip?.archived && (
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>{setEditingMember(m);setEditMemberName(m.name);}} style={{...ghostBtn,padding:"5px 10px",fontSize:12}}>✏️</button>
                  {m.id!==myId && <button onClick={()=>deleteMember(m.id)} style={{...dangerBtn,padding:"5px 10px",fontSize:12}}>🗑</button>}
                </div>
              )}
            </div>
          ))}
          <div style={{marginTop:16,padding:14,background:"rgba(56,189,248,.07)",border:"1px solid rgba(56,189,248,.2)",borderRadius:12}}>
            <div style={{fontSize:13,color:C.accent,marginBottom:6,fontWeight:600}}>🔗 邀請朋友加入</div>
            <div style={{fontSize:12,color:C.muted,marginBottom:8}}>把以下代碼分享給朋友，他們選「加入行程」輸入即可即時同步。</div>
            <div style={{fontFamily:"monospace",fontSize:13,color:C.text,background:"rgba(0,0,0,.3)",padding:"10px 12px",borderRadius:8,wordBreak:"break-all",letterSpacing:.5}}>{tripId}</div>
            <button onClick={()=>navigator.clipboard?.writeText(tripId)} style={{...ghostBtn,marginTop:10,width:"100%",fontSize:12}}>複製代碼</button>
          </div>
        </>}

        {/* ── EXPORT ── */}
        {tab==="export" && <>
          <div style={{fontSize:13,color:C.muted,marginBottom:16}}>匯出帳目明細、個人餘額、結算清單</div>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <button onClick={()=>exportExcel(trip,members,expenses,rates)} style={{...primaryBtn,display:"flex",alignItems:"center",justifyContent:"center",gap:10,fontSize:15,padding:16}}>
              <span style={{fontSize:24}}>📊</span>
              <div style={{textAlign:"left"}}><div>匯出 Excel</div><div style={{fontSize:11,opacity:.7,fontWeight:400}}>3個工作表</div></div>
            </button>
            <button onClick={()=>exportPDF(trip,members,expenses,rates)} style={{...ghostBtn,display:"flex",alignItems:"center",justifyContent:"center",gap:10,fontSize:15,padding:16,borderRadius:12}}>
              <span style={{fontSize:24}}>📄</span>
              <div style={{textAlign:"left"}}><div>匯出 / 列印 PDF</div><div style={{fontSize:11,opacity:.6,fontWeight:400}}>開啟瀏覽器列印</div></div>
            </button>
          </div>
          <div style={{marginTop:24,padding:14,background:"rgba(251,191,36,.07)",border:"1px solid rgba(251,191,36,.2)",borderRadius:12}}>
            <div style={{fontSize:13,color:C.gold,fontWeight:600,marginBottom:6}}>行程資訊</div>
            <div style={{fontSize:12,color:C.muted,lineHeight:2}}>
              行程：{trip?.name}<br/>成員：{members.length} 人<br/>帳目：{expenses.length} 筆<br/>
              總金額：HK${fmt(expenses.reduce((s,e)=>s+e.amount*(e.currency===BASE?1:(rates[e.currency]||1)),0))}<br/>
              狀態：{trip?.archived?"已封存":"進行中"}
            </div>
          </div>
        </>}
      </div>

      {/* ── Modals ── */}
      {showAddExp && (
        <Modal title="新增帳目" onClose={()=>setShowAddExp(false)} wide>
          <ExpenseForm members={members} rates={rates} submitLabel="確認新增"
            onSubmit={async fields => {
              await addExpense({id:uid(), ...fields, createdAt:Date.now()});
              setShowAddExp(false);
            }}/>
        </Modal>
      )}

      {editingExp && (
        <Modal title="修改帳目" onClose={()=>setEditingExp(null)} wide>
          <ExpenseForm initial={editingExp} members={members} rates={rates} submitLabel="儲存修改"
            onSubmit={async fields => {
              await updateExpense({...editingExp, ...fields});
              setEditingExp(null);
            }}/>
        </Modal>
      )}

      {editingMember && (
        <Modal title="修改成員名字" onClose={()=>setEditingMember(null)}>
          <div style={{marginBottom:20}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>名字</div>
            <input style={inp} value={editMemberName} onChange={e=>setEditMemberName(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&editMemberName.trim()&&updateMember(editingMember.id,editMemberName.trim()).then(()=>setEditingMember(null))}/>
          </div>
          <button onClick={async()=>{
            if (!editMemberName.trim()) return;
            await updateMember(editingMember.id, editMemberName.trim());
            setEditingMember(null);
          }} style={{...primaryBtn,width:"100%"}}>儲存</button>
        </Modal>
      )}

      {showRates && (
        <Modal title="匯率（1外幣 = ? HKD）" onClose={()=>setShowRates(false)} wide>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontSize:12,color:C.muted}}>可手動修改（如信用卡匯率）</div>
            <button onClick={refreshRates} style={{...ghostBtn,padding:"6px 12px",fontSize:12}}>{ratesLoading?"更新中…":ratesDone?"✅ 已更新":"🔄 即時匯率"}</button>
          </div>
          {CURRENCIES.filter(c=>c.code!==BASE).map(c=>(
            <div key={c.code} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
              <div style={{minWidth:70}}><div style={{fontSize:14,fontWeight:600,color:C.text}}>{c.code}</div><div style={{fontSize:11,color:C.muted}}>{c.name}</div></div>
              <span style={{color:C.muted}}>=</span>
              <div style={{position:"relative",flex:1}}>
                <input type="number" step="0.0001" min="0" value={rates[c.code]||""} onChange={e=>updateRates({...rates,[c.code]:parseFloat(e.target.value)||0})} style={{...inp,paddingRight:50,textAlign:"right"}}/>
                <span style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",fontSize:11,color:C.muted,pointerEvents:"none"}}>HKD</span>
              </div>
            </div>
          ))}
        </Modal>
      )}
    </div>
  );
}
