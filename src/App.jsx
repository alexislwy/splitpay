import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Constants ──────────────────────────────────────────────────────────────
const CURRENCIES = [
  { code: "HKD", symbol: "HK$", name: "港幣" },
  { code: "CNY", symbol: "¥",   name: "人民幣" },
  { code: "TWD", symbol: "NT$", name: "台幣" },
  { code: "USD", symbol: "US$", name: "美元" },
  { code: "EUR", symbol: "€",   name: "歐元" },
  { code: "GBP", symbol: "£",   name: "英鎊" },
  { code: "AUD", symbol: "A$",  name: "澳元" },
  { code: "JPY", symbol: "¥",   name: "日元" },
  { code: "KRW", symbol: "₩",   name: "韓元" },
];
const BASE = "HKD";
const DEFAULT_RATES = { CNY:1.07, TWD:0.24, USD:7.78, EUR:8.45, GBP:9.85, AUD:4.92, JPY:0.052, KRW:0.0057 };
const MEMBER_COLORS = ["#f87171","#fb923c","#fbbf24","#a3e635","#34d399","#22d3ee","#60a5fa","#a78bfa","#f472b6","#94a3b8"];

const uid = () => Math.random().toString(36).slice(2,9);
const fmt = (n, dec=1) => n.toFixed(dec);
const STORE = "tripay_v2";

function load() { try { return JSON.parse(localStorage.getItem(STORE)) || { trips:[], activeTrip:null }; } catch { return { trips:[], activeTrip:null }; } }
function save(data) { try { localStorage.setItem(STORE, JSON.stringify(data)); } catch {} }

// Per-currency settlement respecting per-expense settleMode
// settleMode: "hkd" = convert to HKD, "original" = keep original currency
function settleMixed(members, expenses, rates) {
  // HKD pool for expenses marked settleInHKD (or HKD currency)
  const hkdBal = {};
  members.forEach(m => hkdBal[m.id] = 0);

  // Original-currency pool: net[a|b|cur] = { a, b, currency, net }
  const curNet = {};

  expenses.forEach(exp => {
    if (!exp.splits.length) return;
    const useHKD = exp.settleMode !== "original" || exp.currency === BASE;
    if (useHKD) {
      // settle in HKD
      const r = exp.currency === BASE ? 1 : (rates[exp.currency] || 1);
      const total = exp.amount * r;
      const share = total / exp.splits.length;
      exp.splits.forEach(sid => { if (sid in hkdBal) hkdBal[sid] -= share; });
      if (exp.paidBy in hkdBal) hkdBal[exp.paidBy] += total;
    } else {
      // settle in original currency
      const share = exp.amount / exp.splits.length;
      exp.splits.forEach(sid => {
        if (sid === exp.paidBy) return;
        const [a,b] = [sid, exp.paidBy].sort();
        const key = `${a}|${b}|${exp.currency}`;
        if (!curNet[key]) curNet[key] = { a, b, currency: exp.currency, net: 0 };
        if (a === sid) curNet[key].net += share;
        else curNet[key].net -= share;
      });
    }
  });

  // HKD transactions (minimize transfers)
  const cred = [], debt = [];
  Object.entries(hkdBal).forEach(([id,b]) => {
    if (b > 0.01) cred.push({id, amt:b});
    else if (b < -0.01) debt.push({id, amt:-b});
  });
  cred.sort((a,b)=>b.amt-a.amt); debt.sort((a,b)=>b.amt-a.amt);
  const hkdTxns = [];
  let ci=0,di=0;
  while(ci<cred.length && di<debt.length){
    const c=cred[ci],d=debt[di],a=Math.min(c.amt,d.amt);
    hkdTxns.push({from:d.id,to:c.id,currency:"HKD",amount:+a.toFixed(2)});
    c.amt-=a; d.amt-=a;
    if(c.amt<0.01)ci++; if(d.amt<0.01)di++;
  }

  // Original-currency transactions
  const curTxns = Object.values(curNet)
    .filter(x => Math.abs(x.net) > 0.01)
    .map(x => x.net > 0
      ? { from: x.a, to: x.b, currency: x.currency, amount: +x.net.toFixed(2) }
      : { from: x.b, to: x.a, currency: x.currency, amount: +(-x.net).toFixed(2) }
    );

  // All transactions grouped by currency
  const allTxns = [...hkdTxns, ...curTxns];

  // HKD balances (for summary)
  const balances = {};
  members.forEach(m => {
    const r_paid = expenses.reduce((s,e)=> e.paidBy===m.id ? s+e.amount*(e.currency===BASE?1:(rates[e.currency]||1)) : s, 0);
    const r_owe  = expenses.reduce((s,e)=> e.splits.includes(m.id) ? s+e.amount*(e.currency===BASE?1:(rates[e.currency]||1))/e.splits.length : s, 0);
    balances[m.id] = r_paid - r_owe;
  });

  return { balances, hkdTxns, curTxns, allTxns };
}



// ── Styles ─────────────────────────────────────────────────────────────────
const C = {
  bg:"#080c18", card:"rgba(255,255,255,0.04)", border:"rgba(255,255,255,0.09)",
  accent:"#38bdf8", gold:"#fbbf24", green:"#34d399", red:"#f87171",
  text:"#e2e8f0", muted:"#64748b", input:"rgba(255,255,255,0.07)",
};
const inp = {
  background:C.input, border:`1px solid ${C.border}`, borderRadius:10,
  color:C.text, padding:"9px 13px", fontSize:14, outline:"none",
  width:"100%", boxSizing:"border-box", fontFamily:"inherit",
};
const pill = (active,color="#38bdf8") => ({
  border:`1px solid ${active ? color : C.border}`,
  borderRadius:20, padding:"5px 13px", fontSize:12, cursor:"pointer",
  background: active ? color+"22" : "transparent",
  color: active ? color : C.muted, fontFamily:"inherit", transition:"all .15s",
});
const primaryBtn = {
  background:"linear-gradient(135deg,#0ea5e9,#6366f1)",
  border:"none", borderRadius:12, color:"#fff", padding:"11px 20px",
  fontSize:14, fontWeight:700, cursor:"pointer", fontFamily:"inherit",
  boxShadow:"0 2px 16px rgba(14,165,233,.35)", transition:"opacity .15s",
};
const ghostBtn = {
  background:"rgba(255,255,255,0.06)", border:`1px solid ${C.border}`,
  borderRadius:10, color:C.text, padding:"9px 16px",
  fontSize:13, cursor:"pointer", fontFamily:"inherit",
};
const dangerBtn = {
  background:"rgba(248,113,113,0.12)", border:"1px solid rgba(248,113,113,0.25)",
  borderRadius:10, color:C.red, padding:"7px 13px",
  fontSize:12, cursor:"pointer", fontFamily:"inherit",
};

// ── Modal ──────────────────────────────────────────────────────────────────
function Modal({title,onClose,children,wide}){
  return (
    <div style={{position:"fixed",inset:0,zIndex:200,background:"rgba(0,0,0,.75)",backdropFilter:"blur(6px)",display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={onClose}>
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

// ── Chip Tag ───────────────────────────────────────────────────────────────
function Tag({color,name}){
  return <span style={{display:"inline-flex",alignItems:"center",background:color+"22",border:`1px solid ${color}44`,borderRadius:20,padding:"2px 10px",fontSize:12,color}}>{name}</span>;
}

// ── Avatar ─────────────────────────────────────────────────────────────────
function Avatar({member,size=36}){
  return <div style={{width:size,height:size,borderRadius:"50%",background:member.color+"33",border:`2px solid ${member.color}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*.38,fontWeight:700,color:member.color,flexShrink:0}}>{member.name[0]}</div>;
}

// ── Export helpers ─────────────────────────────────────────────────────────
function exportExcel(trip, members, expenses, rates) {
  const {balances,transactions} = settle(members,expenses,rates);
  const memberMap = Object.fromEntries(members.map(m=>[m.id,m]));

  // Sheet 1: Expenses
  const expRows = expenses.map(e=>{
    const cur = CURRENCIES.find(c=>c.code===e.currency);
    const r = e.currency===BASE?1:(rates[e.currency]||1);
    return {
      "日期": e.date || new Date(e.createdAt).toLocaleDateString("zh-TW"),
      "描述": e.desc,
      "金額": e.amount,
      "貨幣": e.currency,
      "港幣換算": +(e.amount*r).toFixed(2),
      "付款人": memberMap[e.paidBy]?.name||"",
      "分攤成員": e.splits.map(s=>memberMap[s]?.name||"").join("、"),
      "人數": e.splits.length,
    };
  });

  // Sheet 2: Balances
  const balRows = members.map(m=>({
    "成員": m.name,
    "總餘額(HKD)": +fmt(balances[m.id]||0,2),
    "狀態": (balances[m.id]||0)>0.01?"待收款":(balances[m.id]||0)<-0.01?"待付款":"已結清",
  }));

  // Sheet 3: Settlements
  const settleRows = transactions.map(t=>({
    "付款方": memberMap[t.from]?.name||"",
    "收款方": memberMap[t.to]?.name||"",
    "金額(HKD)": +fmt(t.amount,2),
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(expRows), "帳目明細");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(balRows), "個人餘額");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settleRows), "結算清單");
  XLSX.writeFile(wb, `${trip.name}_結帳.xlsx`);
}

function exportPDF(trip, members, expenses, rates) {
  const {balances,transactions} = settle(members,expenses,rates);
  const memberMap = Object.fromEntries(members.map(m=>[m.id,m]));
  const cur = c => CURRENCIES.find(x=>x.code===c);

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${trip.name} 結帳報告</title>
  <style>
    body{font-family:'Noto Sans TC',sans-serif;padding:32px;color:#1e293b;max-width:800px;margin:0 auto}
    h1{font-size:24px;color:#0ea5e9;margin-bottom:4px} .sub{color:#64748b;font-size:13px;margin-bottom:28px}
    h2{font-size:16px;color:#334155;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin-top:28px}
    table{width:100%;border-collapse:collapse;font-size:13px;margin-top:10px}
    th{background:#f1f5f9;padding:8px 10px;text-align:left;color:#475569;font-weight:600}
    td{padding:8px 10px;border-bottom:1px solid #f1f5f9}
    .green{color:#059669;font-weight:700} .red{color:#dc2626;font-weight:700}
    .tag{display:inline-block;background:#e0f2fe;color:#0369a1;border-radius:10px;padding:1px 8px;font-size:11px;margin:1px}
    .arrow{color:#94a3b8;margin:0 6px}
    @media print{body{padding:16px}}
  </style></head><body>
  <h1>💸 ${trip.name}</h1>
  <div class="sub">生成時間：${new Date().toLocaleString("zh-TW")} ｜ 基準貨幣：HKD</div>

  <h2>帳目明細（${expenses.length} 筆）</h2>
  <table><thead><tr><th>描述</th><th>金額</th><th>港幣</th><th>付款</th><th>分攤</th></tr></thead><tbody>
  ${expenses.map(e=>{
    const r=e.currency===BASE?1:(rates[e.currency]||1);
    return `<tr>
      <td>${e.desc}<br><small style="color:#94a3b8">${e.date || new Date(e.createdAt).toLocaleDateString("zh-TW")}</small></td>
      <td>${cur(e.currency)?.symbol}${e.amount.toLocaleString()}</td>
      <td>HK$${fmt(e.amount*r,1)}</td>
      <td><span class="tag">${memberMap[e.paidBy]?.name||""}</span></td>
      <td>${e.splits.map(s=>`<span class="tag">${memberMap[s]?.name||""}</span>`).join("")}</td>
    </tr>`;
  }).join("")}
  </tbody></table>

  <h2>個人餘額</h2>
  <table><thead><tr><th>成員</th><th>餘額 (HKD)</th><th>狀態</th></tr></thead><tbody>
  ${members.map(m=>{
    const b=balances[m.id]||0;
    return `<tr><td>${m.name}</td>
    <td class="${b>0.01?"green":b<-0.01?"red":""}">${b>0.01?"+":""}HK$${fmt(b,1)}</td>
    <td>${b>0.01?"💰 待收款":b<-0.01?"💸 待付款":"✅ 已結清"}</td></tr>`;
  }).join("")}
  </tbody></table>

  <h2>結算清單（最少 ${transactions.length} 筆轉帳）</h2>
  <table><thead><tr><th>付款方</th><th></th><th>收款方</th><th>金額 (HKD)</th></tr></thead><tbody>
  ${transactions.map(t=>`<tr>
    <td><b>${memberMap[t.from]?.name||""}</b></td>
    <td class="arrow">→</td>
    <td><b>${memberMap[t.to]?.name||""}</b></td>
    <td class="red">HK$${fmt(t.amount,1)}</td>
  </tr>`).join("")}
  ${transactions.length===0?`<tr><td colspan="4" style="text-align:center;color:#059669;padding:16px">🎉 所有帳目已結清！</td></tr>`:""}
  </tbody></table>
  <div style="margin-top:40px;color:#94a3b8;font-size:11px;text-align:center">由 SplitPay 自動生成</div>
  </body></html>`;

  const w = window.open("","_blank");
  w.document.write(html);
  w.document.close();
  setTimeout(()=>w.print(),500);
}

// ── Main App ───────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData] = useState(load);
  const [screen, setScreen] = useState("home"); // home | trip
  const [tripId, setTripId] = useState(null);
  const [myId, setMyId] = useState(null);

  const persist = (d) => { setData(d); save(d); };
  const trip = data.trips.find(t=>t.id===tripId);
  const members = trip?.members||[];
  const expenses = trip?.expenses||[];
  const rates = trip?.rates||DEFAULT_RATES;

  // ── Home ────────────────────────────────────────────────────────────────
  const [showNewTrip, setShowNewTrip] = useState(false);
  const [showJoin, setShowJoin] = useState(false);
  const [newTripName, setNewTripName] = useState("");
  const [myName, setMyName] = useState("");
  const [joinCode, setJoinCode] = useState("");

  const createTrip = () => {
    if (!newTripName.trim()||!myName.trim()) return;
    const me = {id:uid(),name:myName.trim(),color:MEMBER_COLORS[0]};
    const t = {id:uid(),name:newTripName.trim(),createdAt:Date.now(),archived:false,members:[me],expenses:[],rates:{...DEFAULT_RATES}};
    const d = {...data,trips:[...data.trips,t]};
    persist(d); setTripId(t.id); setMyId(me.id); setScreen("trip");
    setShowNewTrip(false); setNewTripName(""); setMyName("");
  };

  const joinTrip = () => {
    const t = data.trips.find(x=>x.id===joinCode.trim());
    if (!t) { alert("找不到此行程代碼"); return; }
    if (t.archived) { alert("此行程已封存"); return; }
    const me = {id:uid(),name:myName.trim(),color:MEMBER_COLORS[t.members.length%MEMBER_COLORS.length]};
    const updated = {...t, members:[...t.members,me]};
    const d = {...data,trips:data.trips.map(x=>x.id===t.id?updated:x)};
    persist(d); setTripId(t.id); setMyId(me.id); setScreen("trip");
    setShowJoin(false); setJoinCode(""); setMyName("");
  };

  const archiveTrip = (tid) => {
    if (!window.confirm("確定封存此行程？封存後不可再新增帳目。")) return;
    persist({...data,trips:data.trips.map(t=>t.id===tid?{...t,archived:true}:t)});
  };
  const deleteTrip = (tid) => {
    if (!window.confirm("確定刪除此行程及所有帳目？此操作不可復原！")) return;
    persist({...data,trips:data.trips.filter(t=>t.id!==tid)});
  };

  // ── Trip screen state ───────────────────────────────────────────────────
  const [tab,setTab] = useState("expenses");
  const [showAddExp,setShowAddExp] = useState(false);
  const [showRates,setShowRates] = useState(false);
  const [showAddMember,setShowAddMember] = useState(false);
  const [newMemberName,setNewMemberName] = useState("");
  const [ratesLoading,setRatesLoading] = useState(false);
  const [ratesDone,setRatesDone] = useState(false);

  const updateTrip = (updated) => persist({...data,trips:data.trips.map(t=>t.id===updated.id?updated:t)});

  const refreshRates = async () => {
    setRatesLoading(true);
    try {
      const res = await fetch("https://api.exchangerate-api.com/v4/latest/HKD");
      const json = await res.json();
      const r = json.rates;
      const newRates = {CNY:+(1/r.CNY).toFixed(4),TWD:+(1/r.TWD).toFixed(4),USD:+(1/r.USD).toFixed(4),EUR:+(1/r.EUR).toFixed(4),GBP:+(1/r.GBP).toFixed(4),AUD:+(1/r.AUD).toFixed(4),JPY:+(1/r.JPY).toFixed(5),KRW:+(1/r.KRW).toFixed(6)};
      updateTrip({...trip,rates:newRates});
      setRatesDone(true); setTimeout(()=>setRatesDone(false),2500);
    } catch {}
    setRatesLoading(false);
  };

  const addMember = () => {
    if (!newMemberName.trim()) return;
    const me = {id:uid(),name:newMemberName.trim(),color:MEMBER_COLORS[members.length%MEMBER_COLORS.length]};
    updateTrip({...trip,members:[...members,me]});
    setNewMemberName("");
  };

  const deleteExpense = (eid) => updateTrip({...trip,expenses:expenses.filter(e=>e.id!==eid)});

  const memberMap = Object.fromEntries(members.map(m=>[m.id,m]));
  const {balances,hkdTxns,curTxns,allTxns} = trip ? settleMixed(members,expenses,rates) : {balances:{},hkdTxns:[],curTxns:[],allTxns:[]};
  const [settleMode,setSettleMode] = useState("mixed"); // "mixed"
  const myBalance = balances[myId]||0;

  // ── HOME SCREEN ─────────────────────────────────────────────────────────
  if (screen==="home") return (
    <div style={{minHeight:"100vh",background:C.bg,padding:24,fontFamily:"'Noto Sans TC',sans-serif"}}>
      <div style={{maxWidth:480,margin:"0 auto"}}>
        <div style={{textAlign:"center",padding:"32px 0 28px"}}>
          <div style={{fontSize:48,marginBottom:8}}>✈️</div>
          <div style={{fontFamily:"'Playfair Display',serif",fontSize:30,color:C.accent,fontWeight:700,letterSpacing:.5}}>SplitPay</div>
          <div style={{fontSize:13,color:C.muted,marginTop:4}}>旅行記帳・智能分帳</div>
        </div>

        <div style={{display:"flex",gap:10,marginBottom:28}}>
          <button onClick={()=>setShowNewTrip(true)} style={{...primaryBtn,flex:1}}>＋ 新增行程</button>
          <button onClick={()=>setShowJoin(true)} style={{...ghostBtn,flex:1}}>🔗 加入行程</button>
        </div>

        {data.trips.length===0 && (
          <div style={{textAlign:"center",color:C.muted,padding:"40px 0",fontSize:14}}>
            還沒有行程，點上方按鈕開始！
          </div>
        )}

        {[false,true].map(archived=>(
          <div key={String(archived)}>
            {data.trips.filter(t=>t.archived===archived).length>0 && (
              <div style={{fontSize:12,color:C.muted,marginBottom:8,marginTop:archived?20:0,letterSpacing:1,textTransform:"uppercase"}}>
                {archived?"📦 已封存":"📍 進行中"}
              </div>
            )}
            {data.trips.filter(t=>t.archived===archived).map(t=>(
              <div key={t.id} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:"16px",marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                  <div>
                    <div style={{fontSize:16,fontWeight:700,color:archived?C.muted:C.text}}>{t.name}</div>
                    <div style={{fontSize:11,color:C.muted,marginTop:2}}>{t.members.length}人・{t.expenses.length}筆帳目</div>
                  </div>
                  <div style={{display:"flex",gap:6}}>
                    {!archived && <button onClick={()=>{setTripId(t.id);setMyId(t.members[0].id);setScreen("trip");}} style={{...primaryBtn,padding:"7px 14px",fontSize:12}}>進入</button>}
                    {!archived && <button onClick={()=>archiveTrip(t.id)} style={{...ghostBtn,padding:"7px 12px",fontSize:12}}>封存</button>}
                    <button onClick={()=>deleteTrip(t.id)} style={{...dangerBtn,padding:"7px 12px",fontSize:12}}>刪除</button>
                  </div>
                </div>
                <div style={{fontSize:11,color:C.muted}}>
                  代碼：<span style={{fontFamily:"monospace",color:C.accent,letterSpacing:1}}>{t.id}</span>
                </div>
              </div>
            ))}
          </div>
        ))}
      </div>

      {showNewTrip && (
        <Modal title="新增行程" onClose={()=>setShowNewTrip(false)}>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>行程名稱</div>
            <input style={inp} placeholder="例：202607 台北、202608 北京" value={newTripName} onChange={e=>setNewTripName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&createTrip()} />
          </div>
          <div style={{marginBottom:20}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>你的名字</div>
            <input style={inp} placeholder="例：Alexis" value={myName} onChange={e=>setMyName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&createTrip()} />
          </div>
          <button onClick={createTrip} style={{...primaryBtn,width:"100%"}}>建立行程 →</button>
        </Modal>
      )}

      {showJoin && (
        <Modal title="加入行程" onClose={()=>setShowJoin(false)}>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>你的名字</div>
            <input style={inp} placeholder="例：Alexis" value={myName} onChange={e=>setMyName(e.target.value)} />
          </div>
          <div style={{marginBottom:20}}>
            <div style={{fontSize:12,color:C.muted,marginBottom:6}}>行程代碼（朋友分享的）</div>
            <input style={inp} placeholder="貼上代碼" value={joinCode} onChange={e=>setJoinCode(e.target.value)} />
          </div>
          <button onClick={joinTrip} style={{...primaryBtn,width:"100%"}}>加入行程 →</button>
        </Modal>
      )}
    </div>
  );

  // ── TRIP SCREEN ─────────────────────────────────────────────────────────
  if (!trip) return <div style={{color:C.text,padding:40,textAlign:"center"}}>找不到行程</div>;

  const tabs = [{key:"expenses",icon:"📋",label:"帳目"},{key:"balances",icon:"💰",label:"結算"},{key:"members",icon:"👥",label:"成員"},{key:"export",icon:"📤",label:"匯出"}];

  return (
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'Noto Sans TC',sans-serif",paddingBottom:80}}>
      {/* Header */}
      <div style={{position:"sticky",top:0,zIndex:50,background:"rgba(8,12,24,.94)",backdropFilter:"blur(12px)",borderBottom:`1px solid ${C.border}`}}>
        <div style={{maxWidth:480,margin:"0 auto",padding:"14px 16px 0"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <button onClick={()=>setScreen("home")} style={{background:"none",border:"none",color:C.muted,fontSize:20,cursor:"pointer",padding:0}}>←</button>
              <div>
                <div style={{fontFamily:"'Playfair Display',serif",fontSize:17,color:C.accent,fontWeight:700}}>{trip.name}</div>
                {trip.archived && <div style={{fontSize:10,color:C.gold,background:"rgba(251,191,36,.15)",borderRadius:8,padding:"1px 7px",display:"inline-block",marginTop:2}}>已封存</div>}
              </div>
            </div>
            <button onClick={()=>setShowRates(true)} style={{...ghostBtn,padding:"6px 11px",fontSize:12}}>💱 匯率</button>
          </div>

          {/* My balance */}
          <div style={{margin:"8px 0",padding:"9px 14px",borderRadius:12,background:myBalance>0.01?"rgba(52,211,153,.1)":myBalance<-0.01?"rgba(248,113,113,.1)":C.card,border:`1px solid ${myBalance>0.01?"rgba(52,211,153,.25)":myBalance<-0.01?"rgba(248,113,113,.25)":C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontSize:12,color:C.muted}}>我的結餘 ({memberMap[myId]?.name||"我"})</span>
            <span style={{fontSize:16,fontWeight:700,fontFamily:"monospace",color:myBalance>0.01?C.green:myBalance<-0.01?C.red:C.muted}}>
              {myBalance>0.01?`+HK$${fmt(myBalance)} 待收`:myBalance<-0.01?`-HK$${fmt(Math.abs(myBalance))} 待付`:"✅ 已結清"}
            </span>
          </div>

          {/* Tabs */}
          <div style={{display:"flex",gap:0}}>
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
          {!trip.archived && <button onClick={()=>setShowAddExp(true)} style={{...primaryBtn,width:"100%",marginBottom:16}}>＋ 新增帳目</button>}
          {expenses.length===0 && <div style={{textAlign:"center",color:C.muted,padding:"50px 0",fontSize:14}}>還沒有帳目</div>}
          {[...expenses].reverse().map(exp=>{
            const payer = memberMap[exp.paidBy];
            const curInfo = CURRENCIES.find(c=>c.code===exp.currency);
            const r = exp.currency===BASE?1:(rates[exp.currency]||1);
            const hkd = exp.amount*r;
            const perPerson = hkd/exp.splits.length;
            return (
              <div key={exp.id} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:"14px 16px",marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                  <div style={{flex:1}}>
                    <div style={{fontSize:15,fontWeight:600,color:C.text}}>{exp.desc}</div>
                    <div style={{fontSize:11,color:C.muted,marginTop:2}}>{exp.date || new Date(exp.createdAt).toLocaleDateString("zh-TW")}</div>
                  </div>
                  <div style={{textAlign:"right",flexShrink:0,marginLeft:10}}>
                    <div style={{fontSize:16,fontWeight:700,color:C.accent,fontFamily:"monospace"}}>{curInfo?.symbol}{exp.amount.toLocaleString()}</div>
                    {exp.currency!==BASE && <div style={{fontSize:11,color:C.muted}}>≈ HK${fmt(hkd)}</div>}
                    <div style={{fontSize:10,color:C.muted}}>每人 HK${fmt(perPerson)}</div>
                  </div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap",marginBottom:6}}>
                  <span style={{fontSize:11,color:C.muted}}>付款：</span>
                  {payer && <Tag color={payer.color} name={payer.name}/>}
                  <span style={{fontSize:11,color:C.muted}}>｜分攤 {exp.splits.length} 人</span>
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:8}}>
                  {exp.splits.map(sid=>memberMap[sid]?<Tag key={sid} color={memberMap[sid].color} name={memberMap[sid].name}/>:null)}
                </div>
                {exp.currency !== BASE && (
                  <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:exp.paidBy===myId&&!trip.archived?8:0}}>
                    <span style={{fontSize:11,color:C.muted}}>結算幣值：</span>
                    <div style={{display:"flex",gap:4,background:"rgba(255,255,255,0.05)",borderRadius:8,padding:2}}>
                      {[["hkd","HK$ 港幣"],["original",`${CURRENCIES.find(c=>c.code===exp.currency)?.symbol} 原幣`]].map(([m,l])=>(
                        <button key={m} onClick={()=>{ if(!trip.archived) updateTrip({...trip,expenses:expenses.map(e=>e.id===exp.id?{...e,settleMode:m}:e)}); }}
                          style={{border:"none",borderRadius:6,padding:"3px 9px",fontSize:11,cursor:trip.archived?"default":"pointer",fontFamily:"inherit",
                            background:(exp.settleMode||"hkd")===m?"rgba(56,189,248,.3)":"transparent",
                            color:(exp.settleMode||"hkd")===m?C.accent:C.muted,transition:"all .15s"}}>
                          {l}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {exp.paidBy===myId && !trip.archived && (
                  <button onClick={()=>deleteExpense(exp.id)} style={{...dangerBtn,fontSize:11,padding:"4px 10px"}}>刪除</button>
                )}
              </div>
            );
          })}
        </>}

        {/* ── BALANCES ── */}
        {tab==="balances" && <>
          <div style={{fontSize:12,color:C.muted,marginBottom:10,fontWeight:600,letterSpacing:.5}}>個人餘額</div>
          {members.map(m=>{
            const b=balances[m.id]||0;
            const spent = expenses.reduce((acc,e)=>{
              const r=e.currency===BASE?1:(rates[e.currency]||1);
              if(e.splits.includes(m.id)) acc+=e.amount*r/e.splits.length;
              return acc;
            },0);
            const paid = expenses.reduce((acc,e)=>{
              if(e.paidBy===m.id){const r=e.currency===BASE?1:(rates[e.currency]||1);acc+=e.amount*r;}
              return acc;
            },0);
            return (
              <div key={m.id} style={{background:C.card,border:`1px solid ${m.id===myId?m.color+"44":C.border}`,borderRadius:14,padding:"14px 16px",marginBottom:10}}>
                <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:10}}>
                  <Avatar member={m}/>
                  <div style={{flex:1}}>
                    <div style={{fontSize:15,fontWeight:600,color:C.text}}>{m.name}{m.id===myId&&" （我）"}</div>
                  </div>
                  <div style={{fontFamily:"monospace",fontWeight:700,fontSize:16,color:b>0.01?C.green:b<-0.01?C.red:C.muted}}>
                    {b>0.01?`+HK$${fmt(b)}`:b<-0.01?`-HK$${fmt(Math.abs(b))}`:"✅"}
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                  {[["💰 實際花費",`HK$${fmt(spent)}`],["💳 已付出",`HK$${fmt(paid)}`]].map(([l,v])=>(
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
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <div style={{fontSize:13,color:C.text,fontWeight:600}}>💡 結算清單</div>
              <div style={{fontSize:11,color:C.muted}}>{allTxns.length} 筆轉帳</div>
            </div>

              <div style={{fontSize:11,color:C.muted,marginBottom:12}}>
                HKD 帳目智能合併最少轉帳；原幣帳目按幣種列出（可在帳目頁逐筆切換）
              </div>
              {allTxns.length===0 && <div style={{textAlign:"center",color:C.green,padding:"20px 0",fontSize:14}}>🎉 大家已結清！</div>}
              {CURRENCIES.map(cur=>{
                const group = allTxns.filter(t=>t.currency===cur.code);
                if (!group.length) return null;
                const isHKD = cur.code === BASE;
                return (
                  <div key={cur.code} style={{marginBottom:16}}>
                    <div style={{fontSize:11,fontWeight:600,marginBottom:7,display:"flex",alignItems:"center",gap:6}}>
                      <span style={{background:isHKD?"rgba(52,211,153,.15)":"rgba(251,191,36,.15)",borderRadius:8,padding:"2px 9px",color:isHKD?C.green:C.gold}}>
                        {cur.symbol} {cur.code} {cur.name}{isHKD?" · 最少轉帳":""}
                      </span>
                    </div>
                    {group.map((t,i)=>{
                      const from=memberMap[t.from],to=memberMap[t.to];
                      const isMe=t.from===myId;
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
          {!trip.archived && (
            <div style={{display:"flex",gap:8,marginBottom:16}}>
              <input style={{...inp,flex:1}} placeholder="新增成員名字" value={newMemberName} onChange={e=>setNewMemberName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addMember()} />
              <button onClick={addMember} style={{...primaryBtn,padding:"9px 16px",whiteSpace:"nowrap"}}>加入</button>
            </div>
          )}
          {members.map(m=>(
            <div key={m.id} style={{display:"flex",alignItems:"center",gap:12,background:C.card,border:`1px solid ${m.id===myId?m.color+"44":C.border}`,borderRadius:12,padding:"12px 14px",marginBottom:8}}>
              <Avatar member={m}/>
              <div style={{flex:1}}>
                <div style={{fontSize:14,color:C.text,fontWeight:600}}>{m.name}{m.id===myId&&" （我）"}</div>
              </div>
            </div>
          ))}
          <div style={{marginTop:16,padding:14,background:"rgba(56,189,248,.07)",border:"1px solid rgba(56,189,248,.2)",borderRadius:12}}>
            <div style={{fontSize:13,color:C.accent,marginBottom:6,fontWeight:600}}>🔗 邀請朋友加入</div>
            <div style={{fontSize:12,color:C.muted,marginBottom:8}}>把以下行程代碼分享給朋友，他們在首頁選「加入行程」輸入代碼即可。</div>
            <div style={{fontFamily:"monospace",fontSize:13,color:C.text,background:"rgba(0,0,0,.3)",padding:"10px 12px",borderRadius:8,wordBreak:"break-all",letterSpacing:.5}}>{trip.id}</div>
            <button onClick={()=>navigator.clipboard?.writeText(trip.id)} style={{...ghostBtn,marginTop:10,width:"100%",fontSize:12}}>複製代碼</button>
          </div>
        </>}

        {/* ── EXPORT ── */}
        {tab==="export" && <>
          <div style={{fontSize:13,color:C.muted,marginBottom:16}}>匯出行程帳目，包含：帳目明細、個人餘額、結算清單</div>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <button onClick={()=>exportExcel(trip,members,expenses,rates)} style={{...primaryBtn,display:"flex",alignItems:"center",justifyContent:"center",gap:10,fontSize:15,padding:"16px"}}>
              <span style={{fontSize:24}}>📊</span>
              <div style={{textAlign:"left"}}>
                <div>匯出 Excel</div>
                <div style={{fontSize:11,opacity:.7,fontWeight:400}}>包含三個工作表</div>
              </div>
            </button>
            <button onClick={()=>exportPDF(trip,members,expenses,rates)} style={{...ghostBtn,display:"flex",alignItems:"center",justifyContent:"center",gap:10,fontSize:15,padding:"16px",borderRadius:12}}>
              <span style={{fontSize:24}}>📄</span>
              <div style={{textAlign:"left"}}>
                <div>匯出 / 列印 PDF</div>
                <div style={{fontSize:11,opacity:.6,fontWeight:400}}>開啟瀏覽器列印</div>
              </div>
            </button>
          </div>
          <div style={{marginTop:24,padding:14,background:"rgba(251,191,36,.07)",border:"1px solid rgba(251,191,36,.2)",borderRadius:12}}>
            <div style={{fontSize:13,color:C.gold,fontWeight:600,marginBottom:6}}>行程資訊</div>
            <div style={{fontSize:12,color:C.muted,lineHeight:2}}>
              行程：{trip.name}<br/>
              成員：{members.length} 人<br/>
              帳目：{expenses.length} 筆<br/>
              總金額：HK${fmt(expenses.reduce((s,e)=>s+e.amount*(e.currency===BASE?1:(rates[e.currency]||1)),0))}<br/>
              狀態：{trip.archived?"已封存":"進行中"}
            </div>
          </div>
        </>}
      </div>

      {/* ── Add Expense Modal ── */}
      {showAddExp && <AddExpenseModal trip={trip} members={members} rates={rates} myId={myId} onClose={()=>setShowAddExp(false)} onAdd={exp=>{ updateTrip({...trip,expenses:[...expenses,exp]}); setShowAddExp(false); }}/>}

      {/* ── Rates Modal ── */}
      {showRates && (
        <Modal title="匯率設定（1外幣 = ? HKD）" onClose={()=>setShowRates(false)} wide>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontSize:12,color:C.muted}}>可手動修改（如信用卡匯率）</div>
            <button onClick={refreshRates} style={{...ghostBtn,padding:"6px 12px",fontSize:12}}>
              {ratesLoading?"更新中…":ratesDone?"✅ 已更新":"🔄 即時匯率"}
            </button>
          </div>
          {CURRENCIES.filter(c=>c.code!==BASE).map(c=>(
            <div key={c.code} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
              <div style={{minWidth:70}}>
                <div style={{fontSize:14,fontWeight:600,color:C.text}}>{c.code}</div>
                <div style={{fontSize:11,color:C.muted}}>{c.name}</div>
              </div>
              <span style={{color:C.muted}}>=</span>
              <div style={{position:"relative",flex:1}}>
                <input type="number" step="0.0001" min="0" value={rates[c.code]||""} onChange={e=>updateTrip({...trip,rates:{...rates,[c.code]:parseFloat(e.target.value)||0}})} style={{...inp,paddingRight:50,textAlign:"right"}}/>
                <span style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",fontSize:11,color:C.muted,pointerEvents:"none"}}>HKD</span>
              </div>
            </div>
          ))}
        </Modal>
      )}
    </div>
  );
}

// ── Add Expense Modal ──────────────────────────────────────────────────────
function AddExpenseModal({trip,members,rates,myId,onClose,onAdd}){
  const [desc,setDesc]=useState("");
  const [amount,setAmount]=useState("");
  const [currency,setCurrency]=useState("HKD");
  const [paidBy,setPaidBy]=useState(myId);
  const [splits,setSplits]=useState(members.map(m=>m.id));
  const [date,setDate]=useState(new Date().toISOString().split('T')[0]);

  const toggle = id => setSplits(s=>s.includes(id)?s.filter(x=>x!==id):[...s,id]);
  const r = currency===BASE?1:(rates[currency]||1);
  const hkd = (parseFloat(amount)||0)*r;

  const submit = () => {
    if(!desc.trim()||!amount||splits.length===0) return;
    onAdd({id:uid(),desc:desc.trim(),amount:parseFloat(amount),currency,paidBy,splits,date,createdAt:Date.now()});
  };

  return (
    <Modal title="新增帳目" onClose={onClose}>
      <div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:10,marginBottom:14}}>
        <div>
          <div style={{fontSize:12,color:C.muted,marginBottom:6}}>描述</div>
          <input style={inp} placeholder="例：演唱會門票、晚餐、Uber…" value={desc} onChange={e=>setDesc(e.target.value)}/>
        </div>
        <div>
          <div style={{fontSize:12,color:C.muted,marginBottom:6}}>日期</div>
          <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{...inp,width:140}}/>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:currency!==BASE?4:14}}>
        <div>
          <div style={{fontSize:12,color:C.muted,marginBottom:6}}>金額</div>
          <input type="number" min="0" step="0.01" style={inp} placeholder="0.00" value={amount} onChange={e=>setAmount(e.target.value)}/>
        </div>
        <div>
          <div style={{fontSize:12,color:C.muted,marginBottom:6}}>貨幣</div>
          <select value={currency} onChange={e=>setCurrency(e.target.value)} style={{...inp,cursor:"pointer"}}>
            {CURRENCIES.map(c=><option key={c.code} value={c.code}>{c.code} {c.name}</option>)}
          </select>
        </div>
      </div>
      {currency!==BASE && hkd>0 && (
        <div style={{fontSize:12,color:C.muted,marginBottom:14,padding:"7px 12px",background:"rgba(56,189,248,.08)",borderRadius:8}}>
          ≈ HK${fmt(hkd)}（每人約 HK${splits.length>0?fmt(hkd/splits.length):"0"}）
        </div>
      )}
      <div style={{marginBottom:14}}>
        <div style={{fontSize:12,color:C.muted,marginBottom:8}}>由誰付款</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:7}}>
          {members.map(m=><button key={m.id} onClick={()=>setPaidBy(m.id)} style={pill(paidBy===m.id,m.color)}>{m.name}</button>)}
        </div>
      </div>
      <div style={{marginBottom:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <div style={{fontSize:12,color:C.muted}}>誰需要分攤（{splits.length} 人）</div>
          <div style={{display:"flex",gap:8}}>
            <span style={{fontSize:12,color:C.accent,cursor:"pointer"}} onClick={()=>setSplits(members.map(m=>m.id))}>全選</span>
            <span style={{fontSize:12,color:C.muted,cursor:"pointer"}} onClick={()=>setSplits([])}>清空</span>
          </div>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:7}}>
          {members.map(m=><button key={m.id} onClick={()=>toggle(m.id)} style={pill(splits.includes(m.id),m.color)}>{splits.includes(m.id)?"✓ ":""}{m.name}</button>)}
        </div>
      </div>
      <button onClick={submit} style={{...primaryBtn,width:"100%"}}>確認新增</button>
    </Modal>
  );
}
