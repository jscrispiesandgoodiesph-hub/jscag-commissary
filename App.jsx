// ============================================================
// JSCAG Food Solutions — Commissary Inventory System
// Production Version: Supabase + Excel Export + Auth + Customizable
//
// ENV VARS NEEDED (set in Vercel):
// VITE_SUPABASE_URL=https://xxxx.supabase.co
// VITE_SUPABASE_ANON_KEY=eyJ...
//
// This file is the complete src/App.jsx
// Deploy instructions in the Deployment Guide PDF.
// ============================================================
import { useState, useEffect, useMemo, createContext, useContext } from "react";
// ── SUPABASE CLIENT ───────────────────────────────────────────
// In production, install: npm install @supabase/supabase-js
// and uncomment below:
//
// import { createClient } from '@supabase/supabase-js';
// const supabase = createClient(
// import.meta.env.VITE_SUPABASE_URL,
// import.meta.env.VITE_SUPABASE_ANON_KEY
// );
//
// For this demo, we use in-memory data (same as before).
// The Supabase calls are shown as comments next to each function.
// ── EXCEL EXPORT HELPER ───────────────────────────────────────
// In production: npm install xlsx
// import * as XLSX from 'xlsx';
function exportToExcel(data, sheetName, filename) {
// Production version:
// const ws = XLSX.utils.json_to_sheet(data);
// const wb = XLSX.utils.book_new();
// XLSX.utils.book_append_sheet(wb, ws, sheetName);
// XLSX.writeFile(wb, filename + '.xlsx');
// Demo version — exports as CSV:
if (!data || data.length === 0) return;
const keys = Object.keys(data[0]);
const csv = [keys.join(","), ...data.map(r => keys.map(k => `"${r[k] ?? ""}"`).join(","))].
const blob = new Blob([csv], { type: "text/csv" });
const url = URL.createObjectURL(blob);
const a = document.createElement("a"); a.href = url; a.download = filename + ".csv"; a.clic
URL.revokeObjectURL(url);
}
// ── SEED DATA ─────────────────────────────────────────────────
const TODAY = new Date().toISOString().split("T")[0];
const YESTERDAY = new Date(Date.now() - 86400000).toISOString().split("T")[0];
const INITIAL_SETTINGS = {
companyName: "JSCAG Food Solutions",
tagline: "Pioneer ng Modern Flavored Crispy Chicken Isaw (since 2019)",
primaryColor: "#2563eb",
currency: "₱",
workingDaysPerMonth: 30,
alertEmail: "bryan@jscag.com",
lowStockAlerts: true,
rejectRateThreshold: 5, // % — alert if above this
timezone: "Asia/Manila",
};
const SEED_CATEGORIES = [
{ id: 1, name: "Kitchen", color: "#e74c3c", active: true, sort_order: 1 },
{ id: 2, name: "Packaging", color: "#3498db", active: true, sort_order: 2 },
{ id: 3, name: "Cleaning / Sanitation", color: "#2ecc71", active: true, sort_order: 3 },
{ id: 4, name: "E-commerce", color: "#9b59b6", active: true, sort_order: 4 },
{ id: 5, name: "Office", color: "#f39c12", active: true, sort_order: 5 },
{ id: 6, name: "Others", color: "#95a5a6", active: true, sort_order: 6 },
];
const SEED_ITEMS = [
{ id: 1, name: "Chicken Isaw", categoryId: 1, unit: "kg", purchaseUnit: "kg", purchasePrice
{ id: 2, name: "All-Purpose Flour", categoryId: 1, unit: "g", purchaseUnit: "1 kg", purchas
{ id: 3, name: "Cornstarch", categoryId: 1, unit: "g", purchaseUnit: "1 kg", purchasePrice:
{ id: 4, name: "Seasoning Mix (BBQ)", categoryId: 1, unit: "g", purchaseUnit: "200g pack",
{ id: 5, name: "Cooking Oil", categoryId: 1, unit: "ml", purchaseUnit: "1 liter", purchaseP
{ id: 6, name: "Plastic Packaging (50g)", categoryId: 2, unit: "pcs", purchaseUnit: "box 10
{ id: 7, name: "Sticker Label", categoryId: 2, unit: "pcs", purchaseUnit: "roll 500pcs", pu
{ id: 8, name: "Dishwashing Liquid", categoryId: 3, unit: "ml", purchaseUnit: "1 liter", pu
{ id: 9, name: "Alcohol 70%", categoryId: 3, unit: "ml", purchaseUnit: "1 liter", purchaseP
];
const SEED_PRODUCTS = [
{ id: 1, name: "J's Crispy Isaw - Classic", sku: "JCI-CL-50", targetPackSize: "50g", sellin
{ id: 2, name: "J's Crispy Isaw - BBQ", sku: "JCI-BBQ-50", targetPackSize: "50g", sellingPr
{ id: 3, name: "J's Crispy Isaw - Salted Egg", sku: "JCI-SE-50", targetPackSize: "50g", sel
{ id: 4, name: "J's Crispy Isaw - Plain w/ Suka", sku: "JCI-PL-50", targetPackSize: "50g",
{ id: 5, name: "J's Lettuce Crisps", sku: "JLC-50", targetPackSize: "50g", sellingPrice: 45
];
const SEED_INVENTORY = [
{ id: 1, date: YESTERDAY, itemId: 1, categoryId: 1, beginning: 15, ending: 8, unit: "kg", e
{ id: 2, date: YESTERDAY, itemId: 2, categoryId: 1, beginning: 5000, ending: 2800, unit: "g
{ id: 3, date: YESTERDAY, itemId: 3, categoryId: 1, beginning: 3000, ending: 1500, unit: "g
{ id: 4, date: YESTERDAY, itemId: 4, categoryId: 1, beginning: 600, ending: 200, unit: "g",
{ id: 5, date: YESTERDAY, itemId: 5, categoryId: 1, beginning: 5000, ending: 2000, unit: "m
{ id: 6, date: YESTERDAY, itemId: 6, categoryId: 2, beginning: 2000, ending: 800, unit: "pc
{ id: 7, date: YESTERDAY, itemId: 7, categoryId: 2, beginning: 1000, ending: 300, unit: "pc
{ id: 8, date: TODAY, itemId: 1, categoryId: 1, beginning: 8, ending: 3.5, unit: "kg", enco
{ id: 9, date: TODAY, itemId: 2, categoryId: 1, beginning: 2800, ending: 1200, unit: "g", e
{ id: 10, date: TODAY, itemId: 3, categoryId: 1, beginning: 1500, ending: 700, unit: "g", e
{ id: 11, date: TODAY, itemId: 4, categoryId: 1, beginning: 200, ending: 0, unit: "g", enco
{ id: 12, date: TODAY, itemId: 5, categoryId: 1, beginning: 2000, ending: 800, unit: "ml",
{ id: 13, date: TODAY, itemId: 6, categoryId: 2, beginning: 800, ending: 0, unit: "pcs", en
{ id: 14, date: TODAY, itemId: 7, categoryId: 2, beginning: 300, ending: 0, unit: "pcs", en
];
const SEED_PRODUCTION = [
{ id: 1, date: YESTERDAY, productId: 1, batchNo: "B-001", goodPacks: 820, rejects: 30, enco
{ id: 2, date: YESTERDAY, productId: 2, batchNo: "B-002", goodPacks: 350, rejects: 15, enco
{ id: 3, date: TODAY, productId: 1, batchNo: "B-003", goodPacks: 780, rejects: 20, encoder:
{ id: 4, date: TODAY, productId: 2, batchNo: "B-004", goodPacks: 220, rejects: 10, encoder:
];
const SEED_PAYROLL = [
{ id: 1, date: YESTERDAY, employee: "Jose Cruz", role: "Production Staff", dailyRate: 600,
{ id: 2, date: YESTERDAY, employee: "Maria Santos", role: "Inventory Auditor", dailyRate: 5
{ id: 3, date: YESTERDAY, employee: "Ana Reyes", role: "Packing Staff", dailyRate: 550, all
{ id: 4, date: TODAY, employee: "Jose Cruz", role: "Production Staff", dailyRate: 600, allo
{ id: 5, date: TODAY, employee: "Maria Santos", role: "Inventory Auditor", dailyRate: 550,
{ id: 6, date: TODAY, employee: "Ana Reyes", role: "Packing Staff", dailyRate: 550, allowan
];
const SEED_OVERHEAD = [
{ id: 1, date: YESTERDAY, costType: "Electricity", amount: 500, allocationType: "daily", no
{ id: 2, date: YESTERDAY, costType: "Gas / LPG", amount: 350, allocationType: "daily", note
{ id: 3, date: YESTERDAY, costType: "Rent", amount: 1000, allocationType: "daily", notes: "
{ id: 4, date: TODAY, costType: "Electricity", amount: 500, allocationType: "daily", notes:
{ id: 5, date: TODAY, costType: "Gas / LPG", amount: 350, allocationType: "daily", notes: "
{ id: 6, date: TODAY, costType: "Rent", amount: 1000, allocationType: "daily", notes: "" },
{ id: 7, date: TODAY, costType: "Water", amount: 100, allocationType: "daily", notes: "" },
{ id: 8, date: TODAY, costType: "Transportation", amount: 200, allocationType: "daily", not
];
// ── HELPERS ───────────────────────────────────────────────────
const peso = (n, currency = "₱") => `${currency}${Number(n || 0).toLocaleString("en-PH", { mi
const num = (n) => Number(n || 0).toLocaleString("en-PH", { maximumFractionDigits: 2 });
function getCostPerUnit(item) {
const { purchasePrice, purchaseUnit, unit } = item;
if (!purchasePrice || !purchaseUnit) return 0;
const pu = (purchaseUnit || "").toLowerCase();
if (unit === "g") {
if (pu.includes("kg") || pu.includes("kilo")) return purchasePrice / 1000;
if (pu.includes("500g")) return purchasePrice / 500;
if (pu.includes("200g")) return purchasePrice / 200;
const match = pu.match(/(\d+)g/); if (match) return purchasePrice / parseInt(match[1]);
return purchasePrice;
}
if (unit === "ml") {
if (pu.includes("liter") || pu.includes("litre") || pu.includes("1l")) return purchasePri
return purchasePrice;
}
if (unit === "pcs") {
const match = pu.match(/(\d+)\s*pcs/);
if (match) return purchasePrice / parseInt(match[1]);
return purchasePrice;
}
return purchasePrice;
}
function computeRecord(record, items) {
const item = items.find(i => i.id === record.itemId);
if (!item) return { ...record, used: 0, usedCost: 0, costPerUnit: 0 };
const used = Math.max(0, (record.beginning || 0) - (record.ending || 0));
const costPerUnit = getCostPerUnit(item);
return { ...record, used, usedCost: used * costPerUnit, costPerUnit };
}
function computePayrollRow(row) {
return { ...row, totalCost: (row.dailyRate||0)+(row.allowance||0)+(row.sss||0)+(row.philhea
}
// ── CONTEXT ───────────────────────────────────────────────────
const AppContext = createContext(null);
const useApp = () => useContext(AppContext);
// ── MOCK AUTH ─────────────────────────────────────────────────
// In production, replace with Supabase auth:
// const { data: { user } } = await supabase.auth.getUser();
const MOCK_USERS = [
{ id: 1, email: "bryan@jscag.com", password: "admin123", name: "Bryan", role: "admin" },
{ id: 2, email: "maria@jscag.com", password: "maria123", name: "Maria Santos", role: "audit
{ id: 3, email: "jose@jscag.com", password: "jose123", name: "Jose Cruz", role: "production
{ id: 4, email: "finance@jscag.com", password: "finance123", name: "Finance Team", role: "f
];
// ── MAIN APP ──────────────────────────────────────────────────
export default function App() {
const [user, setUser] = useState(null);
const [page, setPage] = useState("dashboard");
const [settings, setSettings] = useState(INITIAL_SETTINGS);
const [items, setItems] = useState(SEED_ITEMS);
const [categories, setCategories] = useState(SEED_CATEGORIES);
const [products, setProducts] = useState(SEED_PRODUCTS);
const [inventoryRecords, setInventoryRecords] = useState(SEED_INVENTORY);
const [productionRecords, setProductionRecords] = useState(SEED_PRODUCTION);
const [payrollRecords, setPayrollRecords] = useState(SEED_PAYROLL);
const [overheadRecords, setOverheadRecords] = useState(SEED_OVERHEAD);
const [toast, setToast] = useState(null);
const [sidebarOpen, setSidebarOpen] = useState(true);
// In production:
// useEffect(() => { supabase.auth.getSession().then(({data:{session}}) => setUser(session?
function showToast(msg, type = "success") {
setToast({ msg, type });
setTimeout(() => setToast(null), 3500);
}
const canAccess = (requiredRole) => {
if (!user) return false;
const hierarchy = { admin: 4, finance: 3, auditor: 2, production: 1 };
const roleAccess = { admin: ["admin"], finance: ["admin", "finance"], auditor: ["admin",
return roleAccess[requiredRole]?.includes(user.role) ?? false;
};
const NAV = [
{ id: "dashboard", label: "Dashboard", icon: "◼", roles: ["admin","auditor","production",
{ id: "inventory-list", label: "Item Master", icon: " ", roles: ["admin"] },
{ id: "daily-inventory", label: "Daily Inventory", icon: " ", roles: ["admin","auditor"]
{ id: "production", label: "Production Yield", icon: " ", roles: ["admin","production"]
{ id: "costing", label: "Product Costing", icon: " ", roles: ["admin","finance"] },
{ id: "payroll", label: "Payroll", icon: " ", roles: ["admin","finance"] },
{ id: "overhead", label: "Overhead", icon: " ", roles: ["admin","finance"] },
{ id: "reports", label: "Reports", icon: " ", roles: ["admin","finance","auditor"] },
{ id: "settings", label: "Settings", icon: " ", roles: ["admin"] },
].filter(n => !user || n.roles.includes(user.role));
const ctx = { user, settings, setSettings, items, setItems, categories, setCategories, prod
if (!user) return <LoginPage onLogin={(u) => setUser(u)} showToast={showToast} toast={toast
UI',sa
return (
<AppContext.Provider value={ctx}>
<div style={{ display: "flex", minHeight: "100vh", fontFamily: "'DM Sans','Segoe <style>{`
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;
*{box-sizing:border-box;margin:0;padding:0}
::-webkit-scrollbar{width:5px;height:5px} ::-webkit-scrollbar-track{background:#1a1
input,select,textarea{font-family:inherit}
.nav-item{display:flex;align-items:center;gap:10px;padding:9px 14px;border-radius:9
.nav-item:hover{background:#1e2235;color:#c8ccdc}
.nav-item.active{background:#1a2a4a;color:#4d9fff}
.card{background:#161924;border:1px solid #232840;border-radius:14px;padding:20px}
.btn{padding:9px 18px;border-radius:8px;border:none;cursor:pointer;font-size:13px;f
.btn-primary{background:#2563eb;color:#fff} .btn-primary:hover{background:#1d4ed8}
.btn-danger{background:#dc2626;color:#fff} .btn-danger:hover{background:#b91c1c}
.btn-ghost{background:#1e2235;color:#8b92a8} .btn-ghost:hover{background:#252a3d;co
.btn-green{background:#065f46;color:#6ee7b7} .btn-green:hover{background:#047857}
.btn-sm{padding:5px 12px;font-size:12px}
.inp{background:#1e2235;border:1px solid #2a3050;color:#e8eaf0;border-radius:8px;pa
.inp:focus{border-color:#2563eb} .inp::placeholder{color:#4a5168}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;padding:10px 14px;background:#1a1d27;color:#6b7390;font-weight:6
td{padding:10px 14px;border-bottom:1px solid #1a1d27;color:#c4c9dc;vertical-align:m
tr:last-child td{border-bottom:none} tr:hover td{background:#0d1020}
.badge{display:inline-block;padding:3px 9px;border-radius:20px;font-size:11px;font-
.badge-green{background:#0d2a1a;color:#34d399} .badge-red{background:#2a0d0d;color:
.badge-blue{background:#0d1a3a;color:#60a5fa} .badge-yellow{background:#2a1f0d;colo
label{font-size:11.5px;color:#6b7390;font-weight:600;text-transform:uppercase;lette
.stat-card{background:#161924;border:1px solid #232840;border-radius:14px;padding:2
.stat-val{font-size:24px;font-weight:700;color:#e8eaf0;font-family:'DM Mono',monosp
.stat-label{font-size:11.5px;color:#6b7390;margin-top:4px;font-weight:500}
.section-title{font-size:20px;font-weight:700;color:#e8eaf0;margin-bottom:4px}
.section-sub{font-size:13px;color:#6b7390;margin-bottom:20px}
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
.grid-4{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:16px}
@media(max-width:900px){.grid-4{grid-template-columns:1fr 1fr}.grid-3{grid-template
@media(max-width:600px){.grid-2,.grid-3,.grid-4{grid-template-columns:1fr}}
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,0.75);display:flex;alig
.modal{background:#161924;border:1px solid #2a3050;border-radius:16px;padding:28px;
.cost-row{display:flex;justify-content:space-between;align-items:center;padding:8px
.highlight-box{background:#0d1a3a;border:1px solid #1e3a6e;border-radius:10px;paddi
.alert-warn{background:#2a1f0d;border:1px solid #78350f;color:#fbbf24;padding:10px
.alert-info{background:#0d1a3a;border:1px solid #1e3a6e;color:#60a5fa;padding:10px
.tab-btn{padding:8px 18px;border-radius:8px;border:none;cursor:pointer;font-size:13
.tab-btn.active{background:#1a2a4a;color:#4d9fff}
`}</style>
v1.0</
{/* SIDEBAR */}
<div style={{ width: sidebarOpen ? 220 : 64, background: "#13161f", borderRight: "1px
<div style={{ display: "flex", alignItems: "center", justifyContent: "space-between
{sidebarOpen && <div>
<div style={{ fontSize: 13, fontWeight: 800, color: "#2563eb", letterSpacing: "
<div style={{ fontSize: 10, color: "#4a5168", marginTop: 1 }}>Commissary </div>}
<button style={{ background: "none", border: "none", color: "#4a5168", cursor: "p
</div>
<nav style={{ marginTop: 12, flex: 1 }}>
{NAV.map(n => (
<button key={n.id} className={`nav-item ${page === n.id ? "active" : ""}`} onCl
<span style={{ fontSize: 15, flexShrink: 0 }}>{n.icon}</span>
{sidebarOpen && <span>{n.label}</span>}
</button>
))}
</nav>
<div style={{ borderTop: "1px solid #1e2235", paddingTop: 12 }}>
{sidebarOpen && <div style={{ fontSize: 11, color: "#4a5168", marginBottom: 8, pa
<div style={{ color: "#8b92a8", fontWeight: 600 }}>{user.name}</div>
<div><span className="badge badge-blue" style={{ fontSize: 10, padding: "1px 6p
</div>}
<button className="nav-item" onClick={() => setUser(null)} title="Logout">
<span style={{ fontSize: 15 }}> </span>
{sidebarOpen && "Logout"}
</button>
</div>
</div>
{/* MAIN */}
<div style={{ flex: 1, padding: "28px 32px", overflowY: "auto", minWidth: 0 }}>
{/* BREADCRUMB */}
<div style={{ fontSize: 11, color: "#4a5168", marginBottom: 20, textTransform: "upp
{settings.companyName} / {NAV.find(n => n.id === page)?.label}
</div>
{page === "dashboard" && <Dashboard />}
{page === "inventory-list" && <ItemMaster />}
{page === "daily-inventory" && <DailyInventory />}
{page === "production" && <ProductionYield />}
{page === "costing" && <ProductCosting />}
{page === "payroll" && <Payroll />}
{page === "overhead" && <Overhead />}
{page === "reports" && <Reports />}
{page === "settings" && <Settings />}
</div>
{/* TOAST */}
{toast && <div style={{ position: "fixed", bottom: 24, right: 24, background: toast.t
{toast.msg}
</div>}
</div>
</AppContext.Provider>
);
}
// ── LOGIN PAGE ────────────────────────────────────────────────
function LoginPage({ onLogin, showToast, toast, settings }) {
const [email, setEmail] = useState("");
const [password, setPassword] = useState("");
const [loading, setLoading] = useState(false);
function handleLogin() {
setLoading(true);
// In production: const { data, error } = await supabase.auth.signInWithPassword({ setTimeout(() => {
const found = MOCK_USERS.find(u => u.email === email && u.password === password);
if (found) { onLogin(found); }
else { showToast("Mali ang email o password.", "error"); }
setLoading(false);
}, 600);
email,
}
return (
<div style={{ minHeight: "100vh", background: "#0f1117", display: "flex", alignItems: "ce
<style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700
<div style={{ width: "100%", maxWidth: 400 }}>
<div style={{ textAlign: "center", marginBottom: 40 }}>
<div style={{ fontSize: 32, fontWeight: 800, color: "#2563eb", letterSpacing: "0.04
<div style={{ fontSize: 13, color: "#6b7390", marginTop: 4 }}>Commissary Inventory
</div>
<div style={{ background: "#161924", border: "1px solid #232840", borderRadius: 16, p
<div style={{ fontWeight: 700, fontSize: 18, marginBottom: 24, color: "#e8eaf0" }}>
<div style={{ marginBottom: 16 }}>
<label>Email</label>
<input className="inp" type="email" placeholder="bryan@jscag.com" value={email} o
</div>
<div style={{ marginBottom: 24 }}>
<label>Password</label>
<input className="inp" type="password" placeholder="••••••••" value={password} on
</div>
<button onClick={handleLogin} disabled={loading} style={{ width: "100%", padding: "
{loading ? "Loading..." : "Pumasok"}
</button>
<div style={{ marginTop: 20, background: "#1a1d27", borderRadius: 8, padding: "12px
<div style={{ fontSize: 11, color: "#4a5168", marginBottom: 6, fontWeight: 600 }}
{[["bryan@jscag.com","admin123","Admin"],["maria@jscag.com","maria123","Auditor"]
<div key={e} style={{ fontSize: 11, color: "#6b7390", cursor: "pointer", paddin
<span style={{ color: "#4d9fff" }}>{r}</span>: {e} / {p}
</div>
))}
</div>
</div>
</div>
</div>
{toast && <div style={{ marginTop: 16, background: "#7f1d1d", border: "1px solid #991
);
}
produc
items)
// ── DASHBOARD ─────────────────────────────────────────────────
function Dashboard() {
const { inventoryRecords, productionRecords, payrollRecords, overheadRecords, items, const inv = inventoryRecords.filter(r => r.date === TODAY).map(r => computeRecord(r, const prod = productionRecords.filter(r => r.date === TODAY);
const pay = payrollRecords.filter(r => r.date === TODAY && r.present).map(computePayrollRow
const oh = overheadRecords.filter(r => r.date === TODAY);
const totalIng = inv.reduce((s, r) => s + r.usedCost, 0);
const totalPay = pay.reduce((s, r) => s + r.totalCost, 0);
const totalOH = oh.reduce((s, r) => s + r.amount, 0);
const totalProd = totalIng + totalPay + totalOH;
const goodPacks = prod.reduce((s, r) => s + r.goodPacks, 0);
const costPerPack = goodPacks > 0 ? totalProd / goodPacks : 0;
const lowStock = items.filter(item => {
const rec = [...inventoryRecords].filter(r => r.itemId === item.id).sort((a,b) => b.date.
return rec && rec.ending <= item.reorderLevel;
});
const topIng = [...inv].sort((a,b) => b.usedCost - a.usedCost).slice(0, 6);
const weekRejects = productionRecords.reduce((s,r) => s + r.rejects, 0);
const weekGood = productionRecords.reduce((s,r) => s + r.goodPacks, 0);
return (
<div>
<div className="section-title">Dashboard</div>
<div className="section-sub">{new Date(TODAY+"T00:00:00").toLocaleDateString("en-PH",{w
<div className="grid-4" style={{ marginBottom: 18 }}>
{[
" },
" },
" },
" " }
{ label: "Ingredient Cost", value: peso(totalIng), color: "#e74c3c", icon: " { label: "Payroll Cost", value: peso(totalPay), color: "#3498db", icon: " { label: "Overhead Cost", value: peso(totalOH), color: "#9b59b6", icon: " { label: "Total Prod. Cost", value: peso(totalProd), color: "#f39c12", icon: ].map(s => (
<div key={s.label} className="stat-card">
<div style={{ fontSize: 20, marginBottom: 8 }}>{s.icon}</div>
<div className="stat-val" style={{ color: s.color, fontSize: 20 }}>{s.value}</div
<div className="stat-label">{s.label}</div>
</div>
))}
</div>
<div className="grid-3" style={{ marginBottom: 18 }}>
<div className="highlight-box">
<div style={{ fontSize: 12, color: "#6b7390", marginBottom: 6 }}>Good Packs Today</
<div style={{ fontSize: 30, fontWeight: 700, color: "#34d399", fontFamily: "'DM Mon
<div style={{ fontSize: 11, color: "#4a5168", marginTop: 4 }}>packs produced</div>
</div>
<div className="highlight-box">
<div style={{ fontSize: 12, color: "#6b7390", marginBottom: 6 }}>Cost Per Pack</div
<div style={{ fontSize: 30, fontWeight: 700, color: "#60a5ff", fontFamily: "'DM Mon
<div style={{ fontSize: 11, color: "#4a5168", marginTop: 4 }}>total ÷ good packs</d
</div>
<div className="highlight-box">
<div style={{ fontSize: 12, color: "#6b7390", marginBottom: 6 }}>Reject Rate <div style={{ fontSize: 30, fontWeight: 700, color: weekGood+weekRejects > 0 {weekGood+weekRejects > 0 ? ((weekRejects/(weekGood+weekRejects))*100).toFixed(1)
</div>
<div style={{ fontSize: 11, color: "#4a5168", marginTop: 4 }}>threshold: {settings.
</div>
</div>
(All)<
&& (we
<div className="grid-2">
<div className="card">
<div style={{ fontWeight: 700, marginBottom: 14, color: lowStock.length > 0 ? "#fbb
{lowStock.length > 0 ? ` Low Stock (${lowStock.length})` : " Stock Levels OK"
</div>
{lowStock.length === 0
? <div style={{ color: "#4a5168", fontSize: 13 }}>Lahat ng items ay nasa safe lev
: lowStock.map(item => {
const rec = [...inventoryRecords].filter(r => r.itemId === item.id).sort((a,b)
return <div key={item.id} className="alert-warn" style={{ marginBottom: 8, bord
<strong>{item.name}</strong> — Stock: {rec?.ending} {item.unit} | Reorder sa:
</div>;
})
ngayon
}
</div>
<div className="card">
<div style={{ fontWeight: 700, marginBottom: 14, fontSize: 14 }}> Top Ingredients
{topIng.length === 0
? <div style={{ color: "#4a5168", fontSize: 13 }}>Walang inventory records : topIng.map(r => {
const item = items.find(i => i.id === r.itemId);
const pct = totalIng > 0 ? (r.usedCost / totalIng * 100).toFixed(0) : 0;
return <div key={r.id}>
<div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.
<span style={{ color: "#c4c9dc" }}>{item?.name}</span>
<span style={{ fontFamily: "'DM Mono',monospace", color: "#60a5ff", fontWei
</div>
<div style={{ height: 4, background: "#1a1d27", borderRadius: 2, marginBottom
<div style={{ height: "100%", width: pct + "%", background: "#2563eb", bord
</div>
</div>;
})
}
</div>
</div>
</div>
);
}
// ── ITEM MASTER ───────────────────────────────────────────────
function ItemMaster() {
const { items, setItems, categories, peso, showToast, getCostPerUnit, exportToExcel } = use
const [showForm, setShowForm] = useState(false);
const [editing, setEditing] = useState(null);
const [filter, setFilter] = useState("");
const [catFilter, setCatFilter] = useState("");
const blank = { name: "", categoryId: 1, unit: "g", purchaseUnit: "1 kg", purchasePrice: 0,
const [form, setForm] = useState(blank);
// Supabase: const {data} = await supabase.from('inventory_items').select('*');
const filtered = items.filter(i =>
i.name.toLowerCase().includes(filter.toLowerCase()) &&
(catFilter === "" || i.categoryId === Number(catFilter))
);
function save() {
if (!form.name.trim()) return;
if (editing) {
// supabase.from('inventory_items').update(form).eq('id', editing)
setItems(items.map(i => i.id === editing ? { ...form, id: editing } : i));
showToast("Item na-update.");
} else {
// supabase.from('inventory_items').insert(form)
setItems([...items, { ...form, id: Date.now() }]);
showToast("Item na-add.");
}
setShowForm(false);
}
function doExport() {
exportToExcel(filtered.map(i => ({
Name: i.name, Category: categories.find(c => c.id === i.categoryId)?.name,
Unit: i.unit, "Purchase Unit": i.purchaseUnit, "Purchase Price": i.purchasePrice,
"Cost/Unit": getCostPerUnit(i).toFixed(6), Supplier: i.supplier,
"Reorder Level": i.reorderLevel, Status: i.active ? "Active" : "Inactive"
})), "Item Master", "JSCAG-Item-Master");
}
return (
<div>
<div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start
<div><div className="section-title">Item Master List</div><div className="section-sub
<div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
<button className="btn btn-green btn-sm" onClick={doExport}> Export</button>
<button className="btn btn-primary" onClick={() => { setForm(blank); setEditing(nul
</div>
</div>
<div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
<input className="inp" style={{ maxWidth: 240 }} placeholder="Hanapin ang item..." va
<select className="inp" style={{ width: "auto" }} value={catFilter} onChange={e => se
<option value="">Lahat ng Categories</option>
{categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
</select>
</div>
<div className="card" style={{ padding: 0, overflow: "hidden" }}>
<table>
<thead><tr><th>Item</th><th>Category</th><th>Unit</th><th>Purchase Unit</th><th>Pur
<tbody>
{filtered.map(item => {
const cat = categories.find(c => c.id === item.categoryId);
return <tr key={item.id}>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{item.name}</td>
<td><span className="badge" style={{ background: cat?.color + "22", color: ca
<td style={{ fontFamily: "'DM Mono',monospace", color: "#6b7390" }}>{item.uni
<td style={{ color: "#6b7390" }}>{item.purchaseUnit}</td>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{peso(item.purchasePrice)}<
<td style={{ fontFamily: "'DM Mono',monospace", color: "#60a5ff", fontWeight:
<td style={{ fontFamily: "'DM Mono',monospace" }}>{item.reorderLevel} {item.u
<td style={{ color: "#6b7390", fontSize: 12 }}>{item.supplier}</td>
<td><span className={`badge ${item.active ? "badge-green" : "badge-red"}`}>{i
<td><button className="btn btn-ghost btn-sm" onClick={() => { setForm({...ite
</tr>;
})}
</tbody>
</table>
</div>
{filtered.length === 0 && <tr><td colSpan={10} style={{ textAlign: "center", colo
{showForm && <FormModal title={editing ? "I-Edit ang Item" : "Mag-add ng Item"} onClose
<div className="grid-2" style={{ gap: 12 }}>
<div style={{ gridColumn: "span 2" }}><label>Item Name *</label><input className="i
<div><label>Category</label>
<select className="inp" value={form.categoryId} onChange={e => setForm({...form,
{categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
</select>
</div>
<div><label>Unit</label>
<select className="inp" value={form.unit} onChange={e => setForm({...form, unit:
{["g","kg","ml","l","pcs","pack","box"].map(u => <option key={u}>{u}</option>)}
</select>
</div>
<div><label>Purchase Unit (paano binibili)</label><input className="inp" value={for
<div><label>Purchase Price (₱)</label><input className="inp" type="number" value={f
<div><label>Reorder Level</label><input className="inp" type="number" value={form.r
<div><label>Supplier</label><input className="inp" value={form.supplier} onChange={
<div><label>Status</label>
<select className="inp" value={form.active?"1":"0"} onChange={e => setForm({...fo
<option value="1">Active</option><option value="0">Inactive</option>
</select>
</div>
<div style={{ gridColumn: "span 2" }}><label>Notes</label><input className="inp" va
</div>
{form.purchasePrice > 0 && <div className="alert-info" style={{ marginTop: 12, Auto-compute: {peso(getCostPerUnit(form))} per {form.unit}
</div>}
</FormModal>}
</div>
border
);
}
// ── DAILY INVENTORY ───────────────────────────────────────────
function DailyInventory() {
const { items, categories, inventoryRecords, setInventoryRecords, showToast, computeRecord,
const [date, setDate] = useState(TODAY);
const [showForm, setShowForm] = useState(false);
const [form, setForm] = useState({ itemId: items[0]?.id, beginning: 0, ending: 0, encoder:
const records = inventoryRecords.filter(r => r.date === date).map(r => computeRecord(r, ite
const totalUsedCost = records.reduce((s,r) => s + r.usedCost, 0);
function autoCarryOver() {
const prev = new Date(date); prev.setDate(prev.getDate()-1);
const yStr = prev.toISOString().split("T")[0];
const prevRecs = inventoryRecords.filter(r => r.date === yStr);
let added = 0;
const newRecs = [...inventoryRecords];
prevRecs.forEach(pr => {
if (!inventoryRecords.find(r => r.date === date && r.itemId === pr.itemId)) {
newRecs.push({...pr, id: Date.now()+Math.random(), date, beginning: pr.ending, added++;
ending
}
});
setInventoryRecords(newRecs);
showToast(`${added} item(s) na-carry over mula ${yStr}.`);
}
function addRecord() {
const item = items.find(i => i.id === Number(form.itemId));
if (!item) return;
if (inventoryRecords.find(r => r.date === date && r.itemId === Number(form.itemId))) {
showToast("Naka-record na ang item na ito para sa date na ito.", "error"); return;
}
const cpu = getCostPerUnit(item);
setInventoryRecords([...inventoryRecords, { ...form, id: Date.now(), date, itemId: showToast("Record na-add.");
setShowForm(false);
Number
}
function doExport() {
exportToExcel(records.map(r => {
const item = items.find(i => i.id === r.itemId);
return { Date: date, Item: item?.name, Beginning: r.beginning, Ending: r.ending, }), "Inventory", `JSCAG-Inventory-${date}`);
Used:
}
return (
<div>
<div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start
<div><div className="section-title">Daily Inventory</div><div className="section-sub"
<div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
<input type="date" className="inp" style={{ width: "auto" }} value={date} onChange=
<button className="btn btn-ghost btn-sm" onClick={autoCarryOver}> Auto Carry-Over
<button className="btn btn-green btn-sm" onClick={doExport}> Export</button>
<button className="btn btn-primary" onClick={() => setShowForm(true)}>+ Add</button
</div>
</div>
<div className="grid-3" style={{ marginBottom: 16 }}>
<div className="stat-card"><div className="stat-val">{records.length}</div><div class
<div className="stat-card"><div className="stat-val" style={{ color: "#f87171" <div className="stat-card"><div className="stat-val" style={{ color: "#60a5ff" </div>
}}>{nu
}}>{pe
color:
<div className="card" style={{ padding: 0, overflow: "hidden" }}>
<table>
<thead><tr><th>Item</th><th>Category</th><th>Beginning</th><th>Ending </th><th>Us
<tbody>
{records.length === 0 && <tr><td colSpan={8} style={{ textAlign: "center", {records.map(r => {
const item = items.find(i => i.id === r.itemId);
const cat = categories.find(c => c.id === r.categoryId);
return <tr key={r.id}>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{item?.name}</td>
<td><span className="badge" style={{ background: cat?.color+"22", color: cat?
<td style={{ fontFamily: "'DM Mono',monospace" }}>{num(r.beginning)} {r.unit}
<td>
<div style={{ display: "flex", alignItems: "center", gap: 6 }}>
<input type="number" className="inp" style={{ width: 90, padding: "5px 8p
onChange={e => setInventoryRecords(inventoryRecords.map(ir => ir.id ===
<span style={{ color: "#6b7390", fontSize: 11 }}>{r.unit}</span>
</div>
</td>
<td style={{ fontFamily: "'DM Mono',monospace", color: r.used > 0 ? "#34d399"
<td style={{ fontFamily: "'DM Mono',monospace", color: "#6b7390" }}>{peso(r.c
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#60a
<td style={{ color: "#6b7390", fontSize: 12 }}>{r.encoder}</td>
</tr>;
})}
{records.length > 0 && <tr style={{ background: "#1a1d27" }}>
<td colSpan={6} style={{ textAlign: "right", fontWeight: 700, color: "#e8eaf0"
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#60a5f
<td></td>
</tr>}
</tbody>
</table>
</div>
{showForm && <FormModal title="Mag-add ng Inventory Record" onClose={() => setShowForm(
<div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
<div><label>Item</label>
<select className="inp" value={form.itemId} onChange={e => setForm({...form, item
{items.filter(i=>i.active).map(i => <option key={i.id} value={i.id}>{i.name} ({
</select>
</div>
<div className="grid-2" style={{ gap: 12 }}>
<div><label>Beginning</label><input className="inp" type="number" value={form.beg
<div><label>Ending</label><input className="inp" type="number" value={form.ending
</div>
<div><label>Encoder (Pangalan mo)</label><input className="inp" value={form.encoder
<div><label>Notes</label><input className="inp" value={form.notes} onChange={e => s
</div>
</FormModal>}
</div>
);
}
// ── PRODUCTION YIELD ──────────────────────────────────────────
function ProductionYield() {
const { products, productionRecords, setProductionRecords, showToast, num, exportToExcel, s
const [date, setDate] = useState(TODAY);
const [showForm, setShowForm] = useState(false);
const [form, setForm] = useState({ productId: products[0]?.id, batchNo: "", goodPacks: 0, r
const records = productionRecords.filter(r => r.date === date);
const totalGood = records.reduce((s,r) => s+r.goodPacks, 0);
const totalRejects = records.reduce((s,r) => s+r.rejects, 0);
const rejectRate = totalGood+totalRejects > 0 ? (totalRejects/(totalGood+totalRejects)*100)
function doExport() {
exportToExcel(records.map(r => {
const prod = products.find(p => p.id === r.productId);
return { Date: date, Product: prod?.name, "Batch No": r.batchNo, "Good Packs": r.goodPa
}), "Production", `JSCAG-Production-${date}`);
}
return (
<div>
<div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start
<div><div className="section-title">Production Yield</div><div className="section-sub
<div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
<input type="date" className="inp" style={{ width: "auto" }} value={date} onChange=
<button className="btn btn-green btn-sm" onClick={doExport}> Export</button>
<button className="btn btn-primary" onClick={() => setShowForm(true)}>+ Add Yield</
</div>
</div>
}}>{nu
}}>{nu
<div className="grid-3" style={{ marginBottom: 16 }}>
<div className="stat-card"><div className="stat-val" style={{ color: "#34d399" <div className="stat-card"><div className="stat-val" style={{ color: "#f87171" <div className="stat-card">
<div className="stat-val" style={{ color: rejectRate > settings.rejectRateThreshold
<div className="stat-label">Reject Rate {rejectRate > settings.rejectRateThreshold
</div>
</div>
color:
<div className="card" style={{ padding: 0, overflow: "hidden" }}>
<table>
<thead><tr><th>Product</th><th>Batch</th><th>Good Packs</th><th>Rejects</th><th>Tot
<tbody>
{records.length === 0 && <tr><td colSpan={7} style={{ textAlign: "center", {records.map(r => {
const prod = products.find(p => p.id === r.productId);
const total = r.goodPacks + r.rejects;
const rr = total > 0 ? (r.rejects/total*100).toFixed(1) : 0;
return <tr key={r.id}>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{prod?.name}</td>
<td style={{ fontFamily: "'DM Mono',monospace", color: "#6b7390" }}>{r.batchN
<td style={{ fontFamily: "'DM Mono',monospace", color: "#34d399", fontWeight:
<td style={{ fontFamily: "'DM Mono',monospace", color: "#f87171" }}>{num(r.re
<td style={{ fontFamily: "'DM Mono',monospace" }}>{num(total)}</td>
<td><span className={`badge ${rr > settings.rejectRateThreshold ? "badge-red"
<td style={{ color: "#6b7390", fontSize: 12 }}>{r.encoder}</td>
</tr>;
})}
</tbody>
</table>
</div>
{showForm && <FormModal title="Mag-add ng Production Yield" onClose={() => setShowForm(
setProductionRecords([...productionRecords, {...form, id: Date.now(), date, productId
showToast("Production record na-add."); setShowForm(false);
}}>
<div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
<div><label>Product</label>
<select className="inp" value={form.productId} onChange={e => setForm({...form, p
{products.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
</select>
</div>
<div className="grid-2" style={{ gap: 12 }}>
<div><label>Batch No.</label><input className="inp" value={form.batchNo} onChange
<div><label>Encoder</label><input className="inp" value={form.encoder} onChange={
<div><label>Good Packs</label><input className="inp" type="number" value={form.go
<div><label>Rejects</label><input className="inp" type="number" value={form.rejec
</div>
<div><label>Notes</label><input className="inp" value={form.notes} onChange={e => s
</div>
</FormModal>}
</div>
);
}
// ── PRODUCT COSTING ───────────────────────────────────────────
function ProductCosting() {
const { inventoryRecords, productionRecords, payrollRecords, overheadRecords, items, const [date, setDate] = useState(TODAY);
catego
const invRecs = inventoryRecords.filter(r => r.date === date).map(r => computeRecord(r, ite
const prodRecs = productionRecords.filter(r => r.date === date);
const payRecs = payrollRecords.filter(r => r.date === date && r.present).map(computePayroll
const ohRecs = overheadRecords.filter(r => r.date === date);
const getCat = (r) => { const item = items.find(i => i.id === r.itemId); return categories.
const totalKitchen = invRecs.filter(r => getCat(r) === "Kitchen").reduce((s,r) => s+r.usedC
const totalPkg = invRecs.filter(r => getCat(r) === "Packaging").reduce((s,r) => s+r.usedCos
const totalClean = invRecs.filter(r => getCat(r) === "Cleaning / Sanitation").reduce((s,r)
const totalOtherInv = invRecs.filter(r => !["Kitchen","Packaging","Cleaning / Sanitation"].
const totalPay = payRecs.reduce((s,r) => s+r.totalCost, 0);
const totalOH = ohRecs.reduce((s,r) => s+r.amount, 0);
const grandTotal = totalKitchen + totalPkg + totalClean + totalOtherInv + totalPay + const goodPacks = prodRecs.reduce((s,r) => s+r.goodPacks, 0);
const costPerPack = goodPacks > 0 ? grandTotal / goodPacks : 0;
totalO
const firstProd = products.find(p => prodRecs.find(r => r.productId === p.id));
const margin = firstProd ? firstProd.sellingPrice - costPerPack : 0;
const marginPct = firstProd?.sellingPrice > 0 ? (margin/firstProd.sellingPrice*100).toFixed
function doExport() {
exportToExcel([
{ Category: "Kitchen / Ingredients", Amount: totalKitchen },
{ Category: "Packaging", Amount: totalPkg },
{ Category: "Cleaning / Sanitation", Amount: totalClean },
{ Category: "Other Materials", Amount: totalOtherInv },
{ Category: "Payroll", Amount: totalPay },
{ Category: "Overhead", Amount: totalOH },
{ Category: "TOTAL", Amount: grandTotal },
{ Category: "Good Packs", Amount: goodPacks },
{ Category: "Cost Per Pack", Amount: costPerPack },
], "Costing", `JSCAG-Costing-${date}`);
}
return (
<div>
<div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start
<div><div className="section-title">Product Costing</div><div className="section-sub"
<div style={{ display: "flex", gap: 8 }}>
<input type="date" className="inp" style={{ width: "auto" }} value={date} onChange=
<button className="btn btn-green btn-sm" onClick={doExport}> Export</button>
</div>
</div>
<div className="grid-2">
<div className="card">
<div style={{ fontWeight: 700, fontSize: 14, marginBottom: 14 }}> Ingredient Brea
{invRecs.length === 0 ? <div style={{ color: "#4a5168", fontSize: 13 }}>Walang inve
: invRecs.map(r => {
const item = items.find(i => i.id === r.itemId);
const cat = categories.find(c => c.id === item?.categoryId);
return <div key={r.id} className="cost-row">
<span>
</span>
</div>;
<span style={{ color: "#e8eaf0" }}>{item?.name}</span>
<span style={{ display: "block", fontSize: 10.5, color: "#4a5168" }}>{r.use
<span style={{ fontFamily: "'DM Mono',monospace", color: "#c4c9dc" }}>{peso(r
})
}
</div>
<div className="card">
<div style={{ fontWeight: 700, fontSize: 14, marginBottom: 14 }}> {[
Cost Summary</d
{ label: "Kitchen / Ingredients", value: totalKitchen, color: "#e74c3c" },
{ label: "Packaging", value: totalPkg, color: "#3498db" },
{ label: "Cleaning / Sanitation", value: totalClean, color: "#2ecc71" },
{ label: "Other Materials", value: totalOtherInv, color: "#95a5a6" },
{ label: "Payroll", value: totalPay, color: "#9b59b6" },
{ label: "Overhead", value: totalOH, color: "#f39c12" },
].map(row => <div key={row.label} className="cost-row">
<span style={{ display: "flex", alignItems: "center", gap: 8 }}>
<span style={{ width: 8, height: 8, borderRadius: "50%", background: row.color,
<span style={{ color: "#c4c9dc" }}>{row.label}</span>
</span>
<span style={{ fontFamily: "'DM Mono',monospace", color: "#e8eaf0", fontWeight: 6
</div>)}
<div style={{ display: "flex", justifyContent: "space-between", padding: "14px 0 4p
<span>Total Daily Cost</span>
<span style={{ fontFamily: "'DM Mono',monospace", color: "#fbbf24" }}>{peso(grand
</div>
<div style={{ marginTop: 16, borderTop: "1px solid #1e2235", paddingTop: 16 }}>
<div style={{ fontSize: 12, color: "#6b7390", marginBottom: 6 }}>Good Packs Produ
<div style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, fontSize: 20, c
</div>
<div className="highlight-box" style={{ marginTop: 16 }}>
<div style={{ fontSize: 12, color: "#6b7390", marginBottom: 6 }}>COST PER PACK</d
<div style={{ fontSize: 36, fontWeight: 800, color: "#60a5ff", fontFamily: "'DM M
<div style={{ fontSize: 11, color: "#4a5168", marginTop: 4 }}>{peso(grandTotal)}
{firstProd && goodPacks > 0 && <div style={{ marginTop: 10, fontSize: 13, color:
Selling: {peso(firstProd.sellingPrice)} | Gross Margin: {peso(margin)} ({margin
</div>}
</div>
</div>
</div>
</div>
);
}
// ── PAYROLL ───────────────────────────────────────────────────
function Payroll() {
const { payrollRecords, setPayrollRecords, showToast, peso, computePayrollRow, exportToExce
const [date, setDate] = useState(TODAY);
const [showForm, setShowForm] = useState(false);
const [form, setForm] = useState({ employee: "", role: "Production Staff", dailyRate: 600,
const records = payrollRecords.filter(r => r.date === date).map(computePayrollRow);
const totalCost = records.filter(r => r.present).reduce((s,r) => s+r.totalCost, 0);
function doExport() {
exportToExcel(records.map(r => ({
Date: date, Employee: r.employee, Role: r.role, "Daily Rate": r.dailyRate,
Allowance: r.allowance, SSS: r.sss, PhilHealth: r.philhealth, "Pag-IBIG": r.pagibig,
"Other Benefits": r.otherBenefits, "Total Cost": r.totalCost, Present: r.present })), "Payroll", `JSCAG-Payroll-${date}`);
? "Yes
}
return (
<div>
<div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start
<div><div className="section-title">Payroll</div><div className="section-sub">Daily m
<div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
<input type="date" className="inp" style={{ width: "auto" }} value={date} onChange=
<button className="btn btn-green btn-sm" onClick={doExport}> Export</button>
<button className="btn btn-primary" onClick={() => setShowForm(true)}>+ Add Employe
</div>
</div>
<div className="grid-3" style={{ marginBottom: 16 }}>
<div className="stat-card"><div className="stat-val">{records.filter(r=>r.present).le
<div className="stat-card"><div className="stat-val" style={{ color: "#f87171" <div className="stat-card"><div className="stat-val" style={{ color: "#9b59b6" </div>
}}>{re
}}>{pe
color:
<div className="card" style={{ padding: 0, overflow: "hidden" }}>
<table>
<thead><tr><th>Employee</th><th>Role</th><th>Daily Rate</th><th>Allowance</th><th>C
<tbody>
{records.length === 0 && <tr><td colSpan={7} style={{ textAlign: "center", {records.map(r => <tr key={r.id} style={{ opacity: r.present ? 1 : 0.4 }}>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{r.employee}</td>
<td style={{ color: "#6b7390" }}>{r.role}</td>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{peso(r.dailyRate)}</td>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{peso(r.allowance)}</td>
<td style={{ fontFamily: "'DM Mono',monospace", color: "#6b7390" }}>{peso(r.sss
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#9b59b
<td>
<button className={`badge ${r.present ? "badge-green" : "badge-red"}`} onClick={() => setPayrollRecords(payrollRecords.map(pr => pr.id === r.id ?
{r.present ? "Present" : "Absent"}
</button>
</td>
</tr>)}
</tbody>
</table>
</div>
style=
{showForm && <FormModal title="Mag-add ng Employee" onClose={() => setShowForm(false)}
if (!form.employee) return;
setPayrollRecords([...payrollRecords, {...form, id: Date.now(), date, dailyRate: pars
showToast("Payroll record na-add."); setShowForm(false);
}}>
<div className="grid-2" style={{ gap: 12 }}>
<div style={{ gridColumn: "span 2" }}><label>Pangalan ng Employee *</label><input c
<div><label>Role</label>
<select className="inp" value={form.role} onChange={e => setForm({...form, role:
{["Production Staff","Packing Staff","Inventory Auditor","Finance","Driver","Su
</select>
</div>
<div><label>Daily Rate (₱)</label><input className="inp" type="number" value={form.
<div><label>Allowance (₱)</label><input className="inp" type="number" value={form.a
<div><label>SSS (₱)</label><input className="inp" type="number" value={form.sss} on
<div><label>PhilHealth (₱)</label><input className="inp" type="number" value={form.
<div><label>Pag-IBIG (₱)</label><input className="inp" type="number" value={form.pa
<div><label>Other Benefits (₱)</label><input className="inp" type="number" value={f
{form.dailyRate > 0 && <div style={{ gridColumn: "span 2" }} className="alert-info"
Total: {peso((parseFloat(form.dailyRate)||0)+(parseFloat(form.allowance)||0)+(par
</div>}
</div>
</FormModal>}
</div>
);
}
// ── OVERHEAD ──────────────────────────────────────────────────
function Overhead() {
const { overheadRecords, setOverheadRecords, showToast, peso, settings, exportToExcel } = u
const [date, setDate] = useState(TODAY);
const [showForm, setShowForm] = useState(false);
const [monthly, setMonthly] = useState(30000);
const [form, setForm] = useState({ costType: "Electricity", amount: 0, allocationType: "dai
const COST_TYPES = ["Electricity","Water","Gas / LPG","Rent","Maintenance","Transportation"
const records = overheadRecords.filter(r => r.date === date);
const total = records.reduce((s,r) => s+r.amount, 0);
function addRecord() {
setOverheadRecords([...overheadRecords, {...form, id: Date.now(), date, amount: parseFloa
showToast("Overhead record na-add."); setShowForm(false);
}
return (
<div>
<div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start
<div><div className="section-title">Overhead Costs</div><div className="section-sub">
<div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
<input type="date" className="inp" style={{ width: "auto" }} value={date} onChange=
<button className="btn btn-green btn-sm" onClick={() => exportToExcel(records.map(r
<button className="btn btn-primary" onClick={() => setShowForm(true)}>+ Add Overhea
</div>
</div>
<div className="grid-2" style={{ marginBottom: 16 }}>
<div className="stat-card"><div className="stat-val">{records.length}</div><div class
<div className="stat-card"><div className="stat-val" style={{ color: "#f39c12" }}>{pe
</div>
color:
<div className="card" style={{ padding: 0, overflow: "hidden", marginBottom: 16 }}>
<table>
<thead><tr><th>Cost Type</th><th>Amount</th><th>Allocation</th><th>Notes</th></tr><
<tbody>
{records.length === 0 && <tr><td colSpan={4} style={{ textAlign: "center", {records.map(r => <tr key={r.id}>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{r.costType}</td>
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#f39c1
<td><span className="badge badge-yellow">{r.allocationType}</span></td>
<td style={{ color: "#6b7390" }}>{r.notes || "—"}</td>
</tr>)}
{records.length > 0 && <tr style={{ background: "#1a1d27" }}>
<td style={{ fontWeight: 700, color: "#e8eaf0" }}>Total</td>
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#f39c1
<td colSpan={2}></td>
</tr>}
</tbody>
</table>
</div>
Monthly → Daily C
}}>
<div className="card">
<div style={{ fontWeight: 700, fontSize: 14, marginBottom: 14 }}> <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" <div style={{ flex: 1, minWidth: 180 }}>
<label>Monthly Amount (₱)</label>
<input className="inp" type="number" value={monthly} onChange={e => setMonthly(pa
</div>
<div style={{ fontSize: 20, color: "#4a5168", paddingTop: 16 }}>→</div>
<div style={{ flex: 1 }}>
<div style={{ fontSize: 12, color: "#6b7390", marginBottom: 4 }}>DAILY EQUIVALENT
<div style={{ fontFamily: "'DM Mono',monospace", fontSize: 24, fontWeight: <div style={{ fontSize: 11, color: "#4a5168" }}>÷ {settings.workingDaysPerMonth}
</div>
</div>
</div>
700, c
{showForm && <FormModal title="Mag-add ng Overhead Cost" onClose={() => setShowForm(fal
<div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
<div><label>Cost Type</label>
<select className="inp" value={form.costType} onChange={e => setForm({...form, co
{COST_TYPES.map(t => <option key={t}>{t}</option>)}
</select>
</div>
<div><label>Allocation Type</label>
<select className="inp" value={form.allocationType} onChange={e => setForm({...fo
<option value="daily">Daily</option>
<option value="weekly">Weekly</option>
<option value="monthly">Monthly</option>
</select>
</div>
<div>
<label>Amount (₱) {form.allocationType !== "daily" && `— Daily equiv: ${peso(form
<input className="inp" type="number" value={form.amount} onChange={e => setForm({
</div>
<div><label>Notes</label><input className="inp" value={form.notes} onChange={e => s
</div>
</FormModal>}
</div>
);
}
// ── REPORTS ───────────────────────────────────────────────────
function Reports() {
const { inventoryRecords, productionRecords, payrollRecords, overheadRecords, items, const [tab, setTab] = useState("costing");
const [dateFrom, setDateFrom] = useState(YESTERDAY);
const [dateTo, setDateTo] = useState(TODAY);
produc
const dates = [];
let cur = new Date(dateFrom+"T00:00:00"), endD = new Date(dateTo+"T00:00:00");
while (cur <= endD) { dates.push(cur.toISOString().split("T")[0]); cur.setDate(cur.getDate(
const costingRows = dates.map(date => {
const inv = inventoryRecords.filter(r => r.date === date).map(r => computeRecord(r, items
const prod = productionRecords.filter(r => r.date === date);
const pay = payrollRecords.filter(r => r.date === date && r.present).map(computePayrollRo
const oh = overheadRecords.filter(r => r.date === date);
const totalI = inv.reduce((s,r) => s+r.usedCost, 0);
const totalP = pay.reduce((s,r) => s+r.totalCost, 0);
const totalO = oh.reduce((s,r) => s+r.amount, 0);
const total = totalI+totalP+totalO;
const packs = prod.reduce((s,r) => s+r.goodPacks, 0);
const rejects = prod.reduce((s,r) => s+r.rejects, 0);
return { date, "Ingredient Cost": totalI, "Payroll": totalP, "Overhead": totalO, "Total C
});
const TABS = [
{ id: "costing", label: "Cost Per Pack History" },
{ id: "inventory", label: "Inventory Usage" },
{ id: "payroll", label: "Payroll Report" },
{ id: "overhead", label: "Overhead Report" },
];
return (
<div>
<div className="section-title">Reports</div>
<div className="section-sub">I-filter at i-export ang historical data.</div>
<div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
{TABS.map(t => <button key={t.id} className={`tab-btn ${tab === t.id ? "active" : ""}
</div>
<div style={{ display: "flex", gap: 12, marginBottom: 18, flexWrap: "wrap", alignItems:
<div><label>Mula sa</label><input type="date" className="inp" style={{ width: "auto"
<div><label>Hanggang</label><input type="date" className="inp" style={{ width: "auto"
<button className="btn btn-green" onClick={() => exportToExcel(
tab === "costing" ? costingRows :
tab === "inventory" ? inventoryRecords.filter(r=>r.date>=dateFrom&&r.date<=dateTo).
tab === "payroll" ? payrollRecords.filter(r=>r.date>=dateFrom&&r.date<=dateTo).map(
overheadRecords.filter(r=>r.date>=dateFrom&&r.date<=dateTo),
tab, `JSCAG-${tab}-${dateFrom}-to-${dateTo}`
)}> I-Export lahat</button>
</div>
{tab === "costing" && <div className="card" style={{ padding: 0, overflow: "hidden" }}>
<table>
<thead><tr><th>Date</th><th>Ingredient</th><th>Payroll</th><th>Overhead</th><th>Tot
<tbody>
{costingRows.map(r => <tr key={r.date}>
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 600 }}>{r.date}</td
<td style={{ fontFamily: "'DM Mono',monospace" }}>{peso(r["Ingredient Cost"])}<
<td style={{ fontFamily: "'DM Mono',monospace" }}>{peso(r["Payroll"])}</td>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{peso(r["Overhead"])}</td>
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#fbbf2
<td style={{ fontFamily: "'DM Mono',monospace", color: "#34d399" }}>{r["Good Pa
<td style={{ fontFamily: "'DM Mono',monospace", color: "#f87171" }}>{r["Rejects
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#60a5f
</tr>)}
</tbody>
</table>
</div>}
{tab === "inventory" && <div className="card" style={{ padding: 0, overflow: "hidden" }
<table>
<tbody>
<thead><tr><th>Date</th><th>Item</th><th>Beginning</th><th>Ending</th><th>Used</th>
{inventoryRecords.filter(r=>r.date>=dateFrom&&r.date<=dateTo).map(r => {
const c = computeRecord(r, items), item = items.find(i=>i.id===r.itemId);
return <tr key={r.id}>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{r.date}</td>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{item?.name}</td>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{r.beginning}</td>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{r.ending}</td>
<td style={{ fontFamily: "'DM Mono',monospace", color: "#34d399", fontWeight:
<td style={{ color: "#6b7390" }}>{r.unit}</td>
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#60a
</tr>;
})}
</tbody>
</table>
</div>}
{tab === "payroll" && <div className="card" style={{ padding: 0, overflow: "hidden" }}>
<table>
<thead><tr><th>Date</th><th>Employee</th><th>Role</th><th>Daily Rate</th><th>Allowa
<tbody>
{payrollRecords.filter(r=>r.date>=dateFrom&&r.date<=dateTo).map(r => {
const c = computePayrollRow(r);
return <tr key={r.id}>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{r.date}</td>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{r.employee}</td>
<td style={{ color: "#6b7390" }}>{r.role}</td>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{peso(r.dailyRate)}</td>
<td style={{ fontFamily: "'DM Mono',monospace" }}>{peso(r.allowance)}</td>
<td style={{ fontFamily: "'DM Mono',monospace", color: "#6b7390" }}>{peso(r.s
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#9b5
<td><span className={`badge ${r.present ? "badge-green" : "badge-red"}`}>{r.p
</tr>;
})}
</tbody>
</table>
</div>}
{tab === "overhead" && <div className="card" style={{ padding: 0, overflow: "hidden" }}
<table>
<thead><tr><th>Date</th><th>Cost Type</th><th>Amount</th><th>Allocation</th><th>Not
<tbody>
{overheadRecords.filter(r=>r.date>=dateFrom&&r.date<=dateTo).map(r => <tr key={r.
<td style={{ fontFamily: "'DM Mono',monospace" }}>{r.date}</td>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{r.costType}</td>
<td style={{ fontFamily: "'DM Mono',monospace", fontWeight: 700, color: "#f39c1
<td><span className="badge badge-yellow">{r.allocationType}</span></td>
<td style={{ color: "#6b7390" }}>{r.notes || "—"}</td>
</tr>)}
</tbody>
</table>
</div>}
</div>
);
}
// ── SETTINGS ──────────────────────────────────────────────────
function Settings() {
const { settings, setSettings, categories, setCategories, products, setProducts, showToast
const [tab, setTab] = useState("general");
const [form, setForm] = useState({...settings});
const [newCat, setNewCat] = useState({ name: "", color: "#2563eb" });
const [newProd, setNewProd] = useState({ name: "", sku: "", targetPackSize: "50g", sellingP
function saveSettings() {
setSettings(form);
showToast("Settings na-save.");
}
const TABS = [
{ id: "general", label: " General" },
{ id: "categories", label: " Categories" },
{ id: "products", label: " Products" },
{ id: "formulas", label: " Formulas" },
{ id: "integrations", label: " Integrations" },
];
return (
<div>
<div className="section-title">Settings</div>
<div className="section-sub">I-customize ang sistema ayon sa inyong business.</div>
<div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
{TABS.map(t => <button key={t.id} className={`tab-btn ${tab === t.id ? "active" : ""}
</div>
{tab === "general" && <div className="card" style={{ maxWidth: 560 }}>
<div style={{ fontWeight: 700, fontSize: 15, marginBottom: 18 }}>General Settings</di
<div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
<div><label>Company Name</label><input className="inp" value={form.companyName} onC
<div><label>Tagline / Description</label><input className="inp" value={form.tagline
<div><label>Currency Symbol</label><input className="inp" value={form.currency} onC
<div><label>Working Days per Month (para sa monthly-to-daily conversion)</label><in
<div><label>Alert Email (para sa low stock notifications)</label><input className="
<div><label>Reject Rate Threshold (%) — Mag-aalert kapag naabot na</label><input cl
<div style={{ display: "flex", alignItems: "center", gap: 10 }}>
<input type="checkbox" checked={form.lowStockAlerts} onChange={e => setForm({...f
<span style={{ fontSize: 13 }}>I-enable ang Low Stock Alerts</span>
</div>
</div>
</div>}
<button className="btn btn-primary" onClick={saveSettings}> I-Save ang Settings</
{tab === "categories" && <div>
<div className="card" style={{ marginBottom: 16 }}>
<div style={{ fontWeight: 700, fontSize: 15, marginBottom: 14 }}>Mag-add ng Categor
<div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
<div style={{ flex: 1, minWidth: 180 }}><label>Category Name</label><input classN
<div><label>Color</label><input type="color" value={newCat.color} onChange={e =>
<div style={{ paddingTop: 20 }}><button className="btn btn-primary" onClick={() =
if (!newCat.name.trim()) return;
setCategories([...categories, {...newCat, id: Date.now(), active: true}]);
setNewCat({ name: "", color: "#2563eb" }); showToast("Category na-add.");
}}>+ Add</button></div>
</div>
</div>
<div className="card" style={{ padding: 0, overflow: "hidden" }}>
<table>
<thead><tr><th>Category</th><th>Color</th><th>Status</th><th></th></tr></thead>
<tbody>
{categories.map(c => <tr key={c.id}>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{c.name}</td>
<td><div style={{ width: 24, height: 24, borderRadius: 6, background: c.color
<td><span className={`badge ${c.active ? "badge-green" : "badge-red"}`}>{c.ac
<td><button className="btn btn-ghost btn-sm" onClick={() => { setCategories(c
</tr>)}
</tbody>
</table>
</div>
</div>}
{tab === "products" && <div>
<div className="card" style={{ marginBottom: 16 }}>
<div style={{ fontWeight: 700, fontSize: 15, marginBottom: 14 }}>Mag-add ng Product
<div className="grid-2" style={{ gap: 12 }}>
<div><label>Product Name</label><input className="inp" value={newProd.name} onCha
<div><label>SKU</label><input className="inp" value={newProd.sku} onChange={e =>
<div><label>Target Pack Size</label><input className="inp" value={newProd.targetP
<div><label>Selling Price (₱)</label><input className="inp" type="number" value={
</div>
<button className="btn btn-primary" style={{ marginTop: 14 }} onClick={() => {
if (!newProd.name.trim()) return;
setProducts([...products, {...newProd, id: Date.now()}]);
setNewProd({ name: "", sku: "", targetPackSize: "50g", sellingPrice: 0, active: t
}}>+ Add Product</button>
</div>
<div className="card" style={{ padding: 0, overflow: "hidden" }}>
<table>
<thead><tr><th>Product</th><th>SKU</th><th>Pack Size</th><th>Selling Price</th><t
<tbody>
{products.map(p => <tr key={p.id}>
<td style={{ fontWeight: 600, color: "#e8eaf0" }}>{p.name}</td>
<td style={{ fontFamily: "'DM Mono',monospace", color: "#6b7390" }}>{p.sku}</
<td>{p.targetPackSize}</td>
<td style={{ fontFamily: "'DM Mono',monospace", color: "#34d399" }}>{peso(p.s
<td><span className={`badge ${p.active ? "badge-green" : "badge-red"}`}>{p.ac
</tr>)}
</tbody>
</table>
</div>
</div>}
{tab === "formulas" && <div className="card" style={{ maxWidth: 620 }}>
<div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}> Default Formulas<
<div style={{ fontSize: 13, color: "#8b92a8", marginBottom: 20 }}>Ito ang mga formula
{[
{ label: "Used Quantity", formula: "Beginning Inventory − Ending Inventory" },
{ label: "Used Cost (per item)", formula: "Used Quantity × Cost Per Unit" },
{ label: "Cost Per Unit (g)", formula: "Purchase Price ÷ 1000 (para sa kg to g)" },
{ label: "Cost Per Unit (ml)", formula: "Purchase Price ÷ 1000 (para sa liter to ml
{ label: "Daily Payroll Cost", formula: "Σ (Daily Rate + Allowance + SSS + PhilHeal
{ label: "Monthly → Daily Overhead", formula: "Monthly Amount ÷ Working Days per Mo
{ label: "Total Daily Production Cost", formula: "Total Ingredient Cost + Total Pay
{ label: "Cost Per Pack ", formula: "Total Daily Production Cost ÷ Total Good Pac
{ label: "Gross Margin", formula: "Selling Price − Cost Per Pack" },
{ label: "Gross Margin %", formula: "(Gross Margin ÷ Selling Price) × 100" },
].map(f => <div key={f.label} style={{ marginBottom: 14, padding: "12px 16px", <div style={{ fontSize: 11, color: "#6b7390", fontWeight: 700, marginBottom: <div style={{ fontFamily: "'DM Mono',monospace", fontSize: 13, color: "#60a5ff" }}>
</div>)}
</div>}
backgr
4, tex
{tab === "integrations" && <div className="card" style={{ maxWidth: 560 }}>
<div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}> Integrations</div
{[
{ name: "Supabase (Database)", status: "Kailangan i-connect", desc: "I-set ang VITE
{ name: "Resend (Email Alerts)", status: "Kailangan i-setup", desc: "I-set ang RESE
{ name: "GitHub + Vercel (Deployment)", status: "Kailangan i-connect", desc: "Sundi
{ name: "SMS Alerts (Semaphore PH)", status: "Future version", desc: "Available sa
{ name: "Excel Export (xlsx)", status: "Demo mode (CSV)", desc: "Sa production: npm
].map(i => <div key={i.name} style={{ marginBottom: 12, padding: "14px 16px", backgro
<div style={{ display: "flex", justifyContent: "space-between", alignItems: "center
<span style={{ fontWeight: 700, color: "#e8eaf0", fontSize: 13 }}>{i.name}</span>
<span className="badge" style={{ background: i.color+"22", color: i.color, fontSi
</div>
</div>)}
</div>}
</div>
<div style={{ fontSize: 12, color: "#6b7390" }}>{i.desc}</div>
);
}
// ── SHARED MODAL ──────────────────────────────────────────────
function FormModal({ title, children, onClose, onSave }) {
return (
<div className="modal-overlay" onClick={onClose}>
<div className="modal" onClick={e => e.stopPropagation()}>
<div style={{ fontWeight: 700, fontSize: 16, marginBottom: 20, color: "#e8eaf0" }}>{t
{children}
<div style={{ display: "flex", gap: 10, marginTop: 22 }}>
<button className="btn btn-primary" onClick={onSave}> I-Save</button>
<button className="btn btn-ghost" onClick={onClose}>Kanselahin</button>
</div>
</div>
</div>
);
}
