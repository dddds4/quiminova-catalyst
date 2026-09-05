import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  AreaChart, Area, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer,
} from 'recharts';
import * as XLSX from 'xlsx';
import { supabase } from './supabaseClient';
import {
  LayoutDashboard, ShoppingCart, ArrowDownCircle, ArrowUpCircle, Wallet,
  Boxes, FileSpreadsheet, Scale, TrendingUp, Plus, X, Search, Trash2,
  Download, Eye, Landmark, Check, Package, History, Upload, Pencil, Users, ChevronDown, ChevronUp, AlertTriangle,
} from 'lucide-react';

const LOGO_SRC = '/logo.webp';

/* ============================== THEME ============================== */
const C = {
  ink: '#102A43',
  inkSoft: '#51677E',
  teal: '#0D4A44',
  tealBright: '#1C9186',
  tealFaint: '#E6F1EF',
  bg: '#F4F6F5',
  surface: '#FFFFFF',
  line: '#E1E5E2',
  rust: '#A6472B',
  rustFaint: '#F7EAE7',
  amber: '#A6791F',
  amberFaint: '#F7EFDE',
};

/* ============================== HELPERS ============================== */
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const todayISO = () => new Date().toISOString().slice(0, 10);
const fmtCOP = (n) => 'COP $' + Math.round(n || 0).toLocaleString('es-CO');
const fmtMoney = (n) => '$' + Math.round(n || 0).toLocaleString('es-CO');
const fmtDate = (iso) => {
  if (!iso) return '—';
  const d = new Date(iso + 'T00:00:00');
  return d.toLocaleDateString('es-CO', { day: '2-digit', month: 'short', year: 'numeric' });
};
const addDays = (iso, days) => {
  const d = new Date(iso + 'T00:00:00');
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
};
const monthLabel = (key) => {
  const [y, m] = key.split('-');
  const d = new Date(Number(y), Number(m) - 1, 1);
  return d.toLocaleDateString('es-CO', { month: 'long', year: 'numeric' });
};

/* ============================== SEED DATA ============================== */
function seedData() {
  const clientId = uid();
  const productId = uid();
  const caja = uid();
  const banco = uid();
  const billetera = uid();
  return {
    company: {
      name: 'Quiminova',
      address: 'Cra. 4 Sur, Metropolitana, Barranquilla, Atlántico, Colombia',
      phone: '3003422909',
      email: 'sj5641067@gmail.com',
    },
    nextInvoiceNumber: 19,
    nextBillNumber: 1,
    clients: [{ id: clientId, name: 'Juan Ramirez', nit: '', phone: '', email: '', address: '' }],
    products: [{ id: productId, name: 'Alcohol Extraneutro 96%', unit: 'Litros', price: 4750 }],
    suppliers: [],
    accounts: [
      { id: caja, name: 'Caja general', type: 'Caja', initialBalance: 0 },
      { id: banco, name: 'Bancolombia', type: 'Banco', initialBalance: 0 },
      { id: billetera, name: 'Nequi', type: 'Billetera', initialBalance: 0 },
    ],
    invoices: [
      {
        id: uid(),
        number: 'INV0017',
        date: '2026-08-25',
        dueDate: '2026-09-24',
        clientId,
        items: [{ id: uid(), description: 'Alcohol Extraneutro 96%', lote: '0100826', rate: 4750, qty: 400, unit: 'Litros', total: 1900000 }],
        total: 1900000,
        payments: [],
      },
    ],
    payables: [],
    movements: [],
    inventoryMovements: [],
  };
}

/* ============================== DERIVED CALCS ============================== */
const paidOf = (doc) => (doc.payments || []).reduce((s, p) => s + p.amount, 0);
const balanceOf = (doc) => doc.total - paidOf(doc);
const statusOf = (doc) => {
  const bal = balanceOf(doc);
  if (bal <= 0.5) return 'Pagada';
  if (paidOf(doc) > 0) return 'Parcial';
  if (doc.dueDate && doc.dueDate < todayISO()) return 'Vencida';
  return 'Pendiente';
};
const accountBalance = (db, accId) => {
  const acc = db.accounts.find((a) => a.id === accId);
  if (!acc) return 0;
  const moved = db.movements.filter((m) => m.accountId === accId).reduce((s, m) => s + m.amount, 0);
  return acc.initialBalance + moved;
};
const totalCash = (db) => db.accounts.reduce((s, a) => s + accountBalance(db, a.id), 0);
const totalAR = (db) => db.invoices.reduce((s, i) => s + Math.max(balanceOf(i), 0), 0);
const totalAP = (db) => db.payables.reduce((s, p) => s + Math.max(balanceOf(p), 0), 0);

/* ---- Inventario y costeo (costo promedio ponderado de todas las compras) ---- */
const stockOf = (db, productId) => (db.inventoryMovements || []).filter((m) => m.productId === productId).reduce((s, m) => s + m.qty, 0);
const avgCostOf = (db, productId) => {
  const entradas = (db.inventoryMovements || []).filter((m) => m.productId === productId && m.qty > 0);
  const qty = entradas.reduce((s, m) => s + m.qty, 0);
  if (qty <= 0) return 0;
  const cost = entradas.reduce((s, m) => s + m.qty * m.unitCost, 0);
  return cost / qty;
};
const inventoryValue = (db) => db.products.reduce((s, p) => s + stockOf(db, p.id) * avgCostOf(db, p.id), 0);

/* ---- Rentabilidad: ingresos, costo de mercancía vendida (COGS) y gastos por periodo ---- */
function computePYG(db, monthKey) {
  const monthInvoices = db.invoices.filter((i) => i.date.slice(0, 7) === monthKey);
  const ingresos = monthInvoices.reduce((s, i) => s + i.total, 0);
  const cogs = monthInvoices.reduce((s, i) => s + i.items.reduce((a, it) => a + (it.cost || 0) * it.qty, 0), 0);
  const gastos = db.payables.filter((p) => p.type === 'Gasto' && p.date.slice(0, 7) === monthKey).reduce((s, p) => s + p.total, 0);
  const utilidadBruta = ingresos - cogs;
  const utilidadNeta = utilidadBruta - gastos;
  const margenBruto = ingresos > 0 ? (utilidadBruta / ingresos) * 100 : null;
  const margenNeto = ingresos > 0 ? (utilidadNeta / ingresos) * 100 : null;
  return { ingresos, cogs, gastos, utilidadBruta, utilidadNeta, margenBruto, margenNeto };
}

/* ---- Utilidades de fechas para históricos ---- */
const datesBetween = (startISO, endISO) => {
  const out = [];
  let d = new Date(startISO + 'T00:00:00');
  const end = new Date(endISO + 'T00:00:00');
  while (d <= end && out.length < 3660) {
    out.push(d.toISOString().slice(0, 10));
    d.setDate(d.getDate() + 1);
  }
  return out;
};
const allMonthsInData = (db) => {
  const set = new Set([todayISO().slice(0, 7)]);
  db.invoices.forEach((i) => set.add(i.date.slice(0, 7)));
  db.payables.forEach((p) => set.add(p.date.slice(0, 7)));
  return Array.from(set).sort();
};

/* ============================== STYLE BLOCK ============================== */
function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');

      .qn-root { font-family: 'Inter', sans-serif; color: ${C.ink}; background: ${C.bg}; }
      .qn-root * { box-sizing: border-box; }
      .qn-display { font-family: 'Space Grotesk', sans-serif; }
      .qn-mono { font-family: 'IBM Plex Mono', monospace; }

      .qn-shell { display: flex; min-height: 100vh; }
      .qn-sidebar { width: 232px; flex-shrink: 0; background: ${C.teal}; color: #EAF3F1; display: flex; flex-direction: column; padding: 22px 14px; }
      .qn-brand { display: flex; align-items: center; gap: 10px; padding: 0 8px 20px; border-bottom: 1px solid rgba(234,243,241,0.15); margin-bottom: 16px; }
      .qn-brand img { width: 34px; height: 34px; object-fit: contain; filter: brightness(0) invert(1); }
      .qn-brand-text { line-height: 1.1; }
      .qn-brand-text .name { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 15px; letter-spacing: 0.3px; }
      .qn-brand-text .sub { font-size: 10.5px; opacity: 0.65; letter-spacing: 0.4px; }
      .qn-nav { display: flex; flex-direction: column; gap: 2px; }
      .qn-nav-item { display: flex; align-items: center; gap: 10px; padding: 9px 10px; border-radius: 7px; font-size: 13.5px; font-weight: 500; color: rgba(234,243,241,0.75); cursor: pointer; border: none; background: transparent; text-align: left; width: 100%; }
      .qn-nav-item:hover { background: rgba(234,243,241,0.08); color: #EAF3F1; }
      .qn-nav-item.active { background: rgba(234,243,241,0.14); color: #fff; }
      .qn-nav-item svg { flex-shrink: 0; }
      .qn-nav-group-label { font-size: 10px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.45; padding: 14px 10px 4px; }

      .qn-main { flex: 1; padding: 30px 38px 60px; max-width: 1180px; }
      .qn-page-head { display: flex; align-items: baseline; justify-content: space-between; margin-bottom: 22px; flex-wrap: wrap; gap: 10px; }
      .qn-page-title { font-family: 'Space Grotesk', sans-serif; font-size: 22px; font-weight: 600; }
      .qn-page-sub { font-size: 13px; color: ${C.inkSoft}; margin-top: 2px; }

      .qn-btn { display: inline-flex; align-items: center; gap: 7px; font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 600; padding: 9px 15px; border-radius: 7px; border: 1px solid ${C.line}; background: ${C.surface}; color: ${C.ink}; cursor: pointer; white-space: nowrap; }
      .qn-btn:hover { border-color: ${C.tealBright}; }
      .qn-btn-primary { background: ${C.teal}; color: #fff; border-color: ${C.teal}; }
      .qn-btn-primary:hover { background: ${C.tealBright}; border-color: ${C.tealBright}; }
      .qn-btn-danger { color: ${C.rust}; }
      .qn-btn-sm { padding: 6px 10px; font-size: 12.5px; }
      .qn-btn-icon { padding: 7px; }

      .qn-card { background: ${C.surface}; border: 1px solid ${C.line}; border-radius: 10px; }
      .qn-stat { padding: 16px 18px; }
      .qn-stat .label { font-size: 12px; color: ${C.inkSoft}; margin-bottom: 8px; }
      .qn-stat .value { font-family: 'Space Grotesk', sans-serif; font-size: 24px; font-weight: 600; }
      .qn-stats-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 20px; }
      @media (max-width: 900px) { .qn-stats-row { grid-template-columns: repeat(2, 1fr); } }

      .qn-section { padding: 18px 20px; margin-bottom: 18px; }
      .qn-section-title { font-family: 'Space Grotesk', sans-serif; font-size: 14.5px; font-weight: 600; margin-bottom: 14px; display: flex; align-items: center; justify-content: space-between; }

      .qn-table { width: 100%; border-collapse: collapse; font-size: 13px; }
      .qn-table th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.6px; color: ${C.inkSoft}; font-weight: 600; padding: 0 10px 8px; border-bottom: 1px solid ${C.line}; }
      .qn-table td { padding: 11px 10px; border-bottom: 1px solid ${C.line}; vertical-align: middle; }
      .qn-table tr:last-child td { border-bottom: none; }
      .qn-table tr:hover td { background: ${C.bg}; }

      .qn-badge { display: inline-flex; align-items: center; font-size: 11px; font-weight: 600; padding: 3px 9px; border-radius: 20px; }
      .qn-badge.pagada { background: ${C.tealFaint}; color: ${C.teal}; }
      .qn-badge.parcial { background: ${C.amberFaint}; color: ${C.amber}; }
      .qn-badge.vencida { background: ${C.rustFaint}; color: ${C.rust}; }
      .qn-badge.pendiente { background: ${C.bg}; color: ${C.inkSoft}; border: 1px solid ${C.line}; }

      .qn-input, .qn-select { width: 100%; font-family: 'Inter', sans-serif; font-size: 13.5px; padding: 8px 10px; border-radius: 6px; border: 1px solid ${C.line}; background: #fff; color: ${C.ink}; }
      .qn-input:focus, .qn-select:focus { outline: 2px solid ${C.tealBright}; outline-offset: 0; border-color: ${C.tealBright}; }
      .qn-label { font-size: 11.5px; font-weight: 600; color: ${C.inkSoft}; margin-bottom: 5px; display: block; }
      .qn-field { margin-bottom: 13px; }
      .qn-row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
      .qn-row3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }

      .qn-modal-overlay { position: fixed; inset: 0; background: rgba(16,42,67,0.45); display: flex; align-items: flex-start; justify-content: center; padding: 40px 20px; overflow-y: auto; z-index: 50; }
      .qn-modal { background: #fff; border-radius: 12px; width: 100%; max-width: 560px; padding: 24px 26px 26px; }
      .qn-modal.wide { max-width: 720px; }
      .qn-modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 18px; }
      .qn-modal-title { font-family: 'Space Grotesk', sans-serif; font-size: 17px; font-weight: 600; }
      .qn-close { background: none; border: none; cursor: pointer; color: ${C.inkSoft}; padding: 4px; }

      .qn-empty { text-align: center; padding: 46px 20px; color: ${C.inkSoft}; }
      .qn-empty .t { font-family: 'Space Grotesk', sans-serif; font-size: 15px; color: ${C.ink}; margin-bottom: 5px; font-weight: 600; }
      .qn-empty .s { font-size: 13px; max-width: 340px; margin: 0 auto; }

      .qn-tabs { display: flex; gap: 4px; margin-bottom: 20px; border-bottom: 1px solid ${C.line}; }
      .qn-tab { padding: 9px 4px; margin-right: 22px; font-size: 13.5px; font-weight: 600; color: ${C.inkSoft}; background: none; border: none; border-bottom: 2px solid transparent; cursor: pointer; }
      .qn-tab.active { color: ${C.teal}; border-bottom-color: ${C.teal}; }

      .qn-item-row { display: grid; grid-template-columns: 2.4fr 1fr 0.9fr 0.9fr 1.1fr 30px; gap: 8px; align-items: end; margin-bottom: 10px; }
      .qn-linklike { background: none; border: none; color: ${C.tealBright}; font-weight: 600; font-size: 12.5px; cursor: pointer; padding: 2px 0; display: inline-flex; align-items: center; gap: 5px; }

      .qn-invoice-sheet { background: #fff; border: 1px solid ${C.line}; border-radius: 10px; padding: 30px 34px; }
      .qn-invoice-sheet .top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 22px; }
      .qn-invoice-sheet .co-name { font-family: 'Space Grotesk', sans-serif; font-size: 19px; font-weight: 700; }
      .qn-invoice-sheet .co-meta { font-size: 12px; color: ${C.inkSoft}; line-height: 1.7; margin-top: 4px; }
      .qn-invoice-sheet .meta-r { text-align: right; font-size: 11px; }
      .qn-invoice-sheet .meta-r .k { text-transform: uppercase; letter-spacing: 0.6px; color: ${C.inkSoft}; margin-top: 10px; }
      .qn-invoice-sheet .meta-r .v { font-size: 13px; font-weight: 600; margin-top: 2px; }
      .qn-invoice-sheet .meta-r .v.big { font-size: 16px; }
      .qn-invoice-sheet .logo-mark { width: 44px; height: 44px; object-fit: contain; margin-bottom: 8px; }
    `}</style>
  );
}

/* ============================== ATOMS ============================== */
function Badge({ status }) {
  const cls = status.toLowerCase();
  return <span className={`qn-badge ${cls}`}>{status}</span>;
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="qn-modal-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className={`qn-modal ${wide ? 'wide' : ''}`}>
        <div className="qn-modal-head">
          <div className="qn-modal-title">{title}</div>
          <button className="qn-close" onClick={onClose}><X size={18} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

function Empty({ title, sub }) {
  return (
    <div className="qn-empty">
      <div className="t">{title}</div>
      <div className="s">{sub}</div>
    </div>
  );
}

function ConfirmDelete({ onConfirm, label }) {
  const [arm, setArm] = useState(false);
  useEffect(() => {
    if (!arm) return;
    const t = setTimeout(() => setArm(false), 3000);
    return () => clearTimeout(t);
  }, [arm]);
  if (arm) {
    return (
      <button className="qn-btn qn-btn-sm qn-btn-danger" onClick={() => { onConfirm(); setArm(false); }}>
        <Check size={13} /> Confirmar
      </button>
    );
  }
  return (
    <button className="qn-btn qn-btn-sm qn-btn-icon" title={label || 'Eliminar'} onClick={() => setArm(true)}>
      <Trash2 size={14} />
    </button>
  );
}

function DueDateField({ label, baseDate, value, onChange }) {
  return (
    <div className="qn-field">
      <label className="qn-label">{label}</label>
      <input className="qn-input" type="date" value={value} onChange={(e) => onChange(e.target.value)} style={{ marginBottom: 6 }} />
      <div style={{ display: 'flex', gap: 12 }}>
        <button type="button" className="qn-linklike" onClick={() => onChange(addDays(baseDate, 15))}>+15 días</button>
        <button type="button" className="qn-linklike" onClick={() => onChange(addDays(baseDate, 30))}>+30 días</button>
      </div>
    </div>
  );
}

/* ============================== SIDEBAR ============================== */
function Sidebar({ tab, setTab, company }) {
  const items = [
    { group: 'General', links: [{ id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard }] },
    {
      group: 'Operación',
      links: [
        { id: 'ventas', label: 'Ventas', icon: ShoppingCart },
        { id: 'cxc', label: 'Cuentas por cobrar', icon: ArrowDownCircle },
        { id: 'cartera', label: 'Cartera de clientes', icon: Users },
        { id: 'cxp', label: 'Cuentas por pagar', icon: ArrowUpCircle },
        { id: 'inventario', label: 'Inventario', icon: Package },
        { id: 'caja', label: 'Caja y bancos', icon: Wallet },
      ],
    },
    {
      group: 'Finanzas',
      links: [
        { id: 'historicos', label: 'Histórico de ventas', icon: History },
        { id: 'balance', label: 'Balance general', icon: Scale },
        { id: 'pyg', label: 'Rentabilidad (P&G)', icon: TrendingUp },
      ],
    },
    { group: 'Configuración', links: [{ id: 'maestros', label: 'Clientes y productos', icon: Boxes }] },
  ];
  return (
    <div className="qn-sidebar">
      <div className="qn-brand">
        <img src={LOGO_SRC} alt="logo" />
        <div className="qn-brand-text">
          <div className="name">Catalyst</div>
          <div className="sub">{company.name.toUpperCase()}</div>
        </div>
      </div>
      <div className="qn-nav" style={{ flex: 1 }}>
        {items.map((g) => (
          <div key={g.group}>
            <div className="qn-nav-group-label">{g.group}</div>
            {g.links.map((l) => (
              <button key={l.id} className={`qn-nav-item ${tab === l.id ? 'active' : ''}`} onClick={() => setTab(l.id)}>
                <l.icon size={16} /> {l.label}
              </button>
            ))}
          </div>
        ))}
      </div>
      <button className="qn-nav-item" onClick={() => supabase.auth.signOut()} style={{ marginTop: 10 }}>
        <X size={16} /> Cerrar sesión
      </button>
    </div>
  );
}

/* ============================== EXCEL EXPORT ============================== */
function downloadInvoiceExcel(db, invoice) {
  const client = db.clients.find((c) => c.id === invoice.clientId);
  const rows = [];
  rows.push([db.company.name, '', '', '', 'FACTURA']);
  rows.push(['', '', '', '', invoice.number]);
  rows.push([db.company.address]);
  rows.push(['', '', '', '', 'FECHA']);
  rows.push(['', '', '', '', fmtDate(invoice.date)]);
  rows.push([db.company.phone]);
  rows.push(['', '', '', '', 'VENCIMIENTO']);
  rows.push(['', '', '', '', fmtDate(invoice.dueDate)]);
  rows.push([db.company.email]);
  rows.push(['', '', '', '', 'SALDO DEUDOR']);
  rows.push(['', '', '', '', balanceOf(invoice)]);
  rows.push([]);
  rows.push(['CLIENTE']);
  rows.push([client ? client.name : '(cliente eliminado)']);
  rows.push([]);
  rows.push(['ARTÍCULO', '', 'TARIFA', 'CANT.', 'TOTAL']);
  invoice.items.forEach((it) => {
    rows.push([it.lote ? `${it.description}  LOTE ${it.lote}` : it.description, '', it.rate, `${it.qty} ${it.unit}`, it.total]);
  });
  rows.push([]);
  rows.push(['', '', 'TOTAL', '', invoice.total]);
  rows.push([]);
  rows.push(['', '', 'SALDO DEUDOR', '', balanceOf(invoice)]);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 34 }, { wch: 10 }, { wch: 13 }, { wch: 13 }, { wch: 16 }];
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, invoice.number);
  XLSX.writeFile(wb, `${invoice.number}.xlsx`);
}

/* ============================== INVOICE PREVIEW ============================== */
function InvoicePreview({ db, invoice, onClose, onDeletePayment, onDeleteInvoice, onEditPaymentDate }) {
  const client = db.clients.find((c) => c.id === invoice.clientId);
  const bal = balanceOf(invoice);
  return (
    <Modal title={`Factura ${invoice.number}`} onClose={onClose} wide>
      <div className="qn-invoice-sheet">
        <div className="top">
          <div>
            <img className="logo-mark" src={LOGO_SRC} alt="logo" />
            <div className="co-name">{db.company.name}</div>
            <div className="co-meta">
              {db.company.address}<br />
              {db.company.phone}<br />
              {db.company.email}
            </div>
          </div>
          <div className="meta-r">
            <div style={{ fontSize: 11, color: C.inkSoft }}>FACTURA</div>
            <div className="v big qn-mono">{invoice.number}</div>
            <div className="k">Fecha</div>
            <div className="v">{fmtDate(invoice.date)}</div>
            <div className="k">Vencimiento</div>
            <div className="v">{fmtDate(invoice.dueDate)}</div>
            <div className="k">Saldo deudor</div>
            <div className="v big">{fmtCOP(bal)}</div>
          </div>
        </div>
        <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.6px', color: C.inkSoft, borderBottom: `1px solid ${C.line}`, paddingBottom: 6, marginBottom: 6 }}>Cliente</div>
        <div style={{ fontWeight: 600, marginBottom: 18 }}>{client ? client.name : '—'}</div>
        <table className="qn-table">
          <thead>
            <tr><th>Artículo</th><th>Tarifa</th><th>Cant.</th><th style={{ textAlign: 'right' }}>Total</th></tr>
          </thead>
          <tbody>
            {invoice.items.map((it) => (
              <tr key={it.id}>
                <td>{it.description}{it.lote ? <span className="qn-mono" style={{ color: C.inkSoft }}> · LOTE {it.lote}</span> : null}</td>
                <td>{fmtMoney(it.rate)}</td>
                <td>{it.qty} {it.unit}</td>
                <td style={{ textAlign: 'right' }}>{fmtMoney(it.total)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 40, marginTop: 16, paddingTop: 12, borderTop: `1px solid ${C.line}` }}>
          <div style={{ textAlign: 'right' }}>
            <div style={{ fontSize: 11, color: C.inkSoft }}>TOTAL</div>
            <div style={{ fontWeight: 700 }}>{fmtCOP(invoice.total)}</div>
          </div>
          <div style={{ textAlign: 'right' }}>
            <div style={{ fontSize: 11, color: C.inkSoft }}>SALDO DEUDOR</div>
            <div style={{ fontWeight: 700, color: bal > 0 ? C.rust : C.teal }}>{fmtCOP(bal)}</div>
          </div>
        </div>
      </div>

      <div style={{ marginTop: 18 }}>
        <div className="qn-section-title" style={{ marginBottom: 8 }}>Pagos registrados</div>
        {(invoice.payments || []).length === 0 ? (
          <div style={{ fontSize: 12.5, color: C.inkSoft }}>Aún no se ha registrado ningún pago.</div>
        ) : (
          <table className="qn-table">
            <thead><tr><th>Fecha</th><th>Cuenta</th><th>Método</th><th style={{ textAlign: 'right' }}>Monto</th><th></th></tr></thead>
            <tbody>
              {invoice.payments.map((p) => {
                const acc = db.accounts.find((a) => a.id === p.accountId);
                return (
                  <tr key={p.id}>
                    <td><input className="qn-input" type="date" value={p.date} onChange={(e) => onEditPaymentDate(p.id, e.target.value)} style={{ maxWidth: 150 }} /></td>
                    <td>{acc ? acc.name : '—'}</td>
                    <td>{p.method || '—'}</td>
                    <td style={{ textAlign: 'right', fontWeight: 600 }}>{fmtMoney(p.amount)}</td>
                    <td><ConfirmDelete onConfirm={() => onDeletePayment(p.id)} label="Eliminar pago" /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 18 }}>
        <ConfirmDelete onConfirm={onDeleteInvoice} label="Eliminar factura completa" />
        <div style={{ display: 'flex', gap: 10 }}>
          <button className="qn-btn" onClick={onClose}>Cerrar</button>
          <button className="qn-btn qn-btn-primary" onClick={() => downloadInvoiceExcel(db, invoice)}>
            <Download size={14} /> Descargar Excel
          </button>
        </div>
      </div>
    </Modal>
  );
}

/* ============================== PAYMENT MODAL (shared CXC / CXP) ============================== */
function PaymentModal({ db, doc, kind, onClose, onSave }) {
  const bal = balanceOf(doc);
  const [amount, setAmount] = useState(bal > 0 ? bal : 0);
  const [date, setDate] = useState(todayISO());
  const [accountId, setAccountId] = useState(db.accounts[0]?.id || '');
  const [method, setMethod] = useState('Transferencia');
  const label = kind === 'cxc' ? 'Cuenta destino (recibe el dinero)' : 'Cuenta origen (sale el dinero)';
  return (
    <Modal title={`Registrar pago · ${doc.number}`} onClose={onClose}>
      <div className="qn-field">
        <span className="qn-label">Saldo pendiente</span>
        <div className="qn-display" style={{ fontSize: 20, fontWeight: 600 }}>{fmtCOP(bal)}</div>
      </div>
      <div className="qn-row2">
        <div className="qn-field">
          <label className="qn-label">Monto a pagar</label>
          <input className="qn-input" type="number" value={amount} onChange={(e) => setAmount(Number(e.target.value))} />
        </div>
        <div className="qn-field">
          <label className="qn-label">Fecha</label>
          <input className="qn-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
      </div>
      <div className="qn-field">
        <label className="qn-label">{label}</label>
        <select className="qn-select" value={accountId} onChange={(e) => setAccountId(e.target.value)}>
          {db.accounts.map((a) => <option key={a.id} value={a.id}>{a.name} ({a.type})</option>)}
        </select>
      </div>
      <div className="qn-field">
        <label className="qn-label">Método</label>
        <select className="qn-select" value={method} onChange={(e) => setMethod(e.target.value)}>
          <option>Transferencia</option><option>Efectivo</option><option>Consignación</option><option>Otro</option>
        </select>
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 8 }}>
        <button className="qn-btn" onClick={onClose}>Cancelar</button>
        <button
          className="qn-btn qn-btn-primary"
          disabled={!amount || amount <= 0 || !accountId}
          onClick={() => { onSave({ amount, date, accountId, method }); onClose(); }}
        >
          <Check size={14} /> Registrar pago
        </button>
      </div>
    </Modal>
  );
}

/* ============================== VENTAS ============================== */
function NewSaleModal({ db, setDb, onClose }) {
  const [clientId, setClientId] = useState(db.clients[0]?.id || '');
  const [date, setDate] = useState(todayISO());
  const [dueDate, setDueDate] = useState(addDays(todayISO(), 30));
  const [items, setItems] = useState([{ id: uid(), productId: '', description: '', lote: '', rate: 0, qty: 1, unit: 'Unidades', total: 0 }]);
  const [showNewClient, setShowNewClient] = useState(false);
  const [ncName, setNcName] = useState('');
  const [ncPhone, setNcPhone] = useState('');

  const updateItem = (id, patch) => {
    setItems((prev) => prev.map((it) => {
      if (it.id !== id) return it;
      const next = { ...it, ...patch };
      next.total = Number(next.rate || 0) * Number(next.qty || 0);
      return next;
    }));
  };
  const pickProduct = (id, productId) => {
    const p = db.products.find((x) => x.id === productId);
    updateItem(id, { productId, description: p ? p.name : '', rate: p ? p.price : 0, unit: p ? p.unit : 'Unidades' });
  };
  const addRow = () => setItems((prev) => [...prev, { id: uid(), productId: '', description: '', lote: '', rate: 0, qty: 1, unit: 'Unidades', total: 0 }]);
  const removeRow = (id) => setItems((prev) => prev.filter((it) => it.id !== id));
  const total = items.reduce((s, it) => s + it.total, 0);

  const saveClient = () => {
    if (!ncName.trim()) return;
    const nc = { id: uid(), name: ncName.trim(), nit: '', phone: ncPhone.trim(), email: '', address: '' };
    setDb((prev) => ({ ...prev, clients: [...prev.clients, nc] }));
    setClientId(nc.id);
    setShowNewClient(false);
    setNcName(''); setNcPhone('');
  };

  const save = () => {
    if (!clientId || items.length === 0 || total <= 0) return;
    setDb((prev) => {
      const number = 'INV' + String(prev.nextInvoiceNumber).padStart(4, '0');
      const finalItems = items.map((it) => ({
        id: it.id, description: it.description, lote: it.lote, rate: Number(it.rate), qty: Number(it.qty), unit: it.unit, total: it.total,
        productId: it.productId || null, cost: it.productId ? avgCostOf(prev, it.productId) : 0,
      }));
      const invoice = { id: uid(), number, date, dueDate, clientId, items: finalItems, total, payments: [] };
      const invMoves = finalItems.filter((it) => it.productId).map((it) => ({
        id: uid(), productId: it.productId, date, qty: -it.qty, unitCost: it.cost, type: 'venta', refType: 'venta', refId: invoice.id,
      }));
      return {
        ...prev,
        invoices: [invoice, ...prev.invoices],
        inventoryMovements: [...(prev.inventoryMovements || []), ...invMoves],
        nextInvoiceNumber: prev.nextInvoiceNumber + 1,
      };
    });
    onClose();
  };

  return (
    <Modal title="Nueva venta" onClose={onClose} wide>
      <div className="qn-row3">
        <div className="qn-field">
          <label className="qn-label">Cliente</label>
          {!showNewClient ? (
            <select className="qn-select" value={clientId} onChange={(e) => setClientId(e.target.value)}>
              {db.clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          ) : (
            <div className="qn-row2">
              <input className="qn-input" placeholder="Nombre" value={ncName} onChange={(e) => setNcName(e.target.value)} />
              <input className="qn-input" placeholder="Teléfono" value={ncPhone} onChange={(e) => setNcPhone(e.target.value)} />
            </div>
          )}
          {!showNewClient ? (
            <button className="qn-linklike" style={{ marginTop: 6 }} onClick={() => setShowNewClient(true)}><Plus size={12} /> Nuevo cliente</button>
          ) : (
            <div style={{ display: 'flex', gap: 8, marginTop: 6 }}>
              <button className="qn-linklike" onClick={saveClient}><Check size={12} /> Guardar</button>
              <button className="qn-linklike" style={{ color: C.inkSoft }} onClick={() => setShowNewClient(false)}>Cancelar</button>
            </div>
          )}
        </div>
        <div className="qn-field">
          <label className="qn-label">Fecha de venta</label>
          <input className="qn-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
        <DueDateField label="Fecha de vencimiento" baseDate={date} value={dueDate} onChange={setDueDate} />
      </div>

      <div style={{ marginTop: 6 }}>
        <div className="qn-item-row" style={{ marginBottom: 6 }}>
          <span className="qn-label">Artículo</span><span className="qn-label">Tarifa</span><span className="qn-label">Cant.</span><span className="qn-label">Unidad</span><span className="qn-label">Lote</span><span />
        </div>
        {items.map((it) => (
          <div className="qn-item-row" key={it.id}>
            <div>
              <select className="qn-select" value={it.productId} onChange={(e) => pickProduct(it.id, e.target.value)} style={{ marginBottom: 5 }}>
                <option value="">— Producto libre —</option>
                {db.products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
              <input className="qn-input" placeholder="Descripción" value={it.description} onChange={(e) => updateItem(it.id, { description: e.target.value })} />
            </div>
            <input className="qn-input" type="number" value={it.rate} onChange={(e) => updateItem(it.id, { rate: e.target.value })} />
            <input className="qn-input" type="number" value={it.qty} onChange={(e) => updateItem(it.id, { qty: e.target.value })} />
            <input className="qn-input" value={it.unit} onChange={(e) => updateItem(it.id, { unit: e.target.value })} />
            <input className="qn-input" value={it.lote} onChange={(e) => updateItem(it.id, { lote: e.target.value })} />
            <button className="qn-close" onClick={() => removeRow(it.id)} title="Quitar"><X size={16} /></button>
          </div>
        ))}
        <button className="qn-linklike" onClick={addRow}><Plus size={13} /> Agregar artículo</button>
      </div>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 20, paddingTop: 14, borderTop: `1px solid ${C.line}` }}>
        <div>
          <div className="qn-label" style={{ marginBottom: 2 }}>Total factura</div>
          <div className="qn-display" style={{ fontSize: 22, fontWeight: 700 }}>{fmtCOP(total)}</div>
        </div>
        <div style={{ display: 'flex', gap: 10 }}>
          <button className="qn-btn" onClick={onClose}>Cancelar</button>
          <button className="qn-btn qn-btn-primary" disabled={!clientId || total <= 0} onClick={save}><Check size={14} /> Guardar venta</button>
        </div>
      </div>
    </Modal>
  );
}

/* ---- Editar una venta ya creada ---- */
function EditSaleModal({ db, setDb, invoice, onClose }) {
  const [clientId, setClientId] = useState(invoice.clientId);
  const [date, setDate] = useState(invoice.date);
  const [dueDate, setDueDate] = useState(invoice.dueDate);
  const [items, setItems] = useState(invoice.items.map((it) => ({ ...it })));
  const [showNewClient, setShowNewClient] = useState(false);
  const [ncName, setNcName] = useState('');
  const [ncPhone, setNcPhone] = useState('');

  const updateItem = (id, patch) => {
    setItems((prev) => prev.map((it) => {
      if (it.id !== id) return it;
      const next = { ...it, ...patch };
      next.total = Number(next.rate || 0) * Number(next.qty || 0);
      return next;
    }));
  };
  const pickProduct = (id, productId) => {
    const p = db.products.find((x) => x.id === productId);
    updateItem(id, { productId, description: p ? p.name : '', rate: p ? p.price : 0, unit: p ? p.unit : 'Unidades' });
  };
  const addRow = () => setItems((prev) => [...prev, { id: uid(), productId: '', description: '', lote: '', rate: 0, qty: 1, unit: 'Unidades', total: 0 }]);
  const removeRow = (id) => setItems((prev) => prev.filter((it) => it.id !== id));
  const total = items.reduce((s, it) => s + it.total, 0);

  const saveClient = () => {
    if (!ncName.trim()) return;
    const nc = { id: uid(), name: ncName.trim(), nit: '', phone: ncPhone.trim(), email: '', address: '' };
    setDb((prev) => ({ ...prev, clients: [...prev.clients, nc] }));
    setClientId(nc.id);
    setShowNewClient(false);
    setNcName(''); setNcPhone('');
  };

  const save = () => {
    if (!clientId || items.length === 0 || total <= 0) return;
    setDb((prev) => {
      const finalItems = items.map((it) => ({
        id: it.id || uid(), description: it.description, lote: it.lote, rate: Number(it.rate), qty: Number(it.qty), unit: it.unit, total: it.total,
        productId: it.productId || null, cost: it.productId ? avgCostOf(prev, it.productId) : 0,
      }));
      const invoices = prev.invoices.map((i) => i.id === invoice.id ? { ...i, date, dueDate, clientId, items: finalItems, total } : i);
      // Reconstruye los movimientos de inventario de esta factura para que coincidan con los artículos editados
      const otherInvMoves = (prev.inventoryMovements || []).filter((m) => !(m.refType === 'venta' && m.refId === invoice.id));
      const newInvMoves = finalItems.filter((it) => it.productId).map((it) => ({
        id: uid(), productId: it.productId, date, qty: -it.qty, unitCost: it.cost, type: 'venta', refType: 'venta', refId: invoice.id,
      }));
      return { ...prev, invoices, inventoryMovements: [...otherInvMoves, ...newInvMoves] };
    });
    onClose();
  };

  return (
    <Modal title={`Editar factura ${invoice.number}`} onClose={onClose} wide>
      <div className="qn-row3">
        <div className="qn-field">
          <label className="qn-label">Cliente</label>
          {!showNewClient ? (
            <select className="qn-select" value={clientId} onChange={(e) => setClientId(e.target.value)}>
              {db.clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          ) : (
            <div className="qn-row2">
              <input className="qn-input" placeholder="Nombre" value={ncName} onChange={(e) => setNcName(e.target.value)} />
              <input className="qn-input" placeholder="Teléfono" value={ncPhone} onChange={(e) => setNcPhone(e.target.value)} />
            </div>
          )}
          {!showNewClient ? (
            <button className="qn-linklike" style={{ marginTop: 6 }} onClick={() => setShowNewClient(true)}><Plus size={12} /> Nuevo cliente</button>
          ) : (
            <div style={{ display: 'flex', gap: 8, marginTop: 6 }}>
              <button className="qn-linklike" onClick={saveClient}><Check size={12} /> Guardar</button>
              <button className="qn-linklike" style={{ color: C.inkSoft }} onClick={() => setShowNewClient(false)}>Cancelar</button>
            </div>
          )}
        </div>
        <div className="qn-field">
          <label className="qn-label">Fecha de venta</label>
          <input className="qn-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
        <DueDateField label="Fecha de vencimiento" baseDate={date} value={dueDate} onChange={setDueDate} />
      </div>

      <div style={{ marginTop: 6 }}>
        <div className="qn-item-row" style={{ marginBottom: 6 }}>
          <span className="qn-label">Artículo</span><span className="qn-label">Tarifa</span><span className="qn-label">Cant.</span><span className="qn-label">Unidad</span><span className="qn-label">Lote</span><span />
        </div>
        {items.map((it) => (
          <div className="qn-item-row" key={it.id}>
            <div>
              <select className="qn-select" value={it.productId || ''} onChange={(e) => pickProduct(it.id, e.target.value)} style={{ marginBottom: 5 }}>
                <option value="">— Producto libre —</option>
                {db.products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
              <input className="qn-input" placeholder="Descripción" value={it.description} onChange={(e) => updateItem(it.id, { description: e.target.value })} />
            </div>
            <input className="qn-input" type="number" value={it.rate} onChange={(e) => updateItem(it.id, { rate: e.target.value })} />
            <input className="qn-input" type="number" value={it.qty} onChange={(e) => updateItem(it.id, { qty: e.target.value })} />
            <input className="qn-input" value={it.unit} onChange={(e) => updateItem(it.id, { unit: e.target.value })} />
            <input className="qn-input" value={it.lote} onChange={(e) => updateItem(it.id, { lote: e.target.value })} />
            <button className="qn-close" onClick={() => removeRow(it.id)} title="Quitar"><X size={16} /></button>
          </div>
        ))}
        <button className="qn-linklike" onClick={addRow}><Plus size={13} /> Agregar artículo</button>
      </div>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 20, paddingTop: 14, borderTop: `1px solid ${C.line}` }}>
        <div>
          <div className="qn-label" style={{ marginBottom: 2 }}>Total factura</div>
          <div className="qn-display" style={{ fontSize: 22, fontWeight: 700 }}>{fmtCOP(total)}</div>
        </div>
        <div style={{ display: 'flex', gap: 10 }}>
          <button className="qn-btn" onClick={onClose}>Cancelar</button>
          <button className="qn-btn qn-btn-primary" disabled={!clientId || total <= 0} onClick={save}><Check size={14} /> Guardar cambios</button>
        </div>
      </div>
    </Modal>
  );
}

/* ---- Importación masiva de ventas desde Excel ---- */
const SALES_HEADER_ALIASES = {
  fecha: 'date', 'fecha venta': 'date',
  'n factura': 'invoiceRef', 'no factura': 'invoiceRef', 'numero factura': 'invoiceRef', 'nfactura': 'invoiceRef', factura: 'invoiceRef',
  cliente: 'client',
  producto: 'product',
  descripcion: 'description',
  lote: 'lote',
  tarifa: 'rate', precio: 'rate', 'precio unitario': 'rate',
  cantidad: 'qty', cant: 'qty',
  unidad: 'unit',
  vencimiento: 'dueDate', 'fecha vencimiento': 'dueDate',
  'monto pagado': 'paidAmount', pago: 'paidAmount', abono: 'paidAmount',
  'cuenta de pago': 'accountName', cuenta: 'accountName',
};
const normalizeHeader = (h) => String(h || '')
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .toLowerCase().replace(/[°ºª.]/g, '').replace(/\s+/g, ' ').trim();

const normalizeDateValue = (v) => {
  if (v === undefined || v === null || v === '') return null;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const dt = new Date(s);
  if (!isNaN(dt.getTime())) return dt.toISOString().slice(0, 10);
  return null;
};

const parseNumberLike = (v) => {
  if (v === null || v === undefined || v === '') return NaN;
  if (typeof v === 'number') return v;
  let s = String(v).trim().replace(/[^0-9.,-]/g, '');
  if (!s) return NaN;
  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');
  if (lastComma > -1 && lastDot > -1) {
    s = lastComma > lastDot ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  } else if (lastComma > -1) {
    const parts = s.split(',');
    s = (parts.length === 2 && parts[1].length <= 2) ? parts[0] + '.' + parts[1] : s.replace(/,/g, '');
  } else if (lastDot > -1) {
    const parts = s.split('.');
    if (!(parts.length === 2 && parts[1].length <= 2)) s = s.replace(/\./g, '');
  }
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
};

function downloadSalesTemplate() {
  const headers = ['Fecha', 'N° Factura', 'Cliente', 'Producto', 'Lote', 'Tarifa', 'Cantidad', 'Unidad', 'Vencimiento', 'Monto pagado', 'Cuenta de pago'];
  const rows = [
    headers,
    ['2026-01-15', '', 'Juan Ramirez', 'Alcohol Extraneutro 96%', '0100826', 4750, 400, 'Litros', '2026-02-14', 1900000, 'Caja general'],
    ['2026-01-20', 'INV-HIST-01', 'María Torres', 'Alcohol Extraneutro 96%', '', 4750, 100, 'Litros', '', '', ''],
    ['2026-01-20', 'INV-HIST-01', 'María Torres', 'Esencia de vainilla', '', 12000, 5, 'Kg', '', '', ''],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 22 }, { wch: 26 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 14 }, { wch: 16 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Ventas');
  XLSX.writeFile(wb, 'Plantilla_ventas.xlsx');
}

function parseSalesRows(jsonRows) {
  const errors = [];
  const parsed = [];
  jsonRows.forEach((raw, idx) => {
    const rowNum = idx + 2;
    const row = {};
    Object.keys(raw).forEach((k) => {
      const canon = SALES_HEADER_ALIASES[normalizeHeader(k)];
      if (canon) row[canon] = raw[k];
    });
    const allEmpty = Object.values(raw).every((v) => String(v ?? '').trim() === '');
    if (allEmpty) return;
    const dateVal = normalizeDateValue(row.date);
    const client = String(row.client || '').trim();
    const rate = parseNumberLike(row.rate);
    const qty = parseNumberLike(row.qty);
    if (!dateVal) { errors.push(`Fila ${rowNum}: fecha inválida o vacía.`); return; }
    if (!client) { errors.push(`Fila ${rowNum}: falta el cliente.`); return; }
    if (!rate || rate <= 0 || isNaN(rate)) { errors.push(`Fila ${rowNum}: tarifa inválida.`); return; }
    if (!qty || qty <= 0 || isNaN(qty)) { errors.push(`Fila ${rowNum}: cantidad inválida.`); return; }
    parsed.push({
      rowNum, date: dateVal,
      invoiceRef: String(row.invoiceRef || '').trim(),
      client,
      product: String(row.product || '').trim(),
      description: String(row.description || row.product || 'Producto').trim(),
      lote: String(row.lote || '').trim(),
      rate, qty,
      unit: String(row.unit || '').trim() || 'Unidades',
      dueDate: normalizeDateValue(row.dueDate) || addDays(dateVal, 30),
      paidAmount: parseNumberLike(row.paidAmount) || 0,
      accountName: String(row.accountName || '').trim(),
    });
  });
  return { rows: parsed, errors };
}

function buildImportPlan(db, rows) {
  const groups = [];
  const byRef = {};
  rows.forEach((r) => {
    if (r.invoiceRef) {
      if (!byRef[r.invoiceRef]) { byRef[r.invoiceRef] = { rows: [] }; groups.push(byRef[r.invoiceRef]); }
      byRef[r.invoiceRef].rows.push(r);
    } else {
      groups.push({ rows: [r] });
    }
  });

  let clients = [...db.clients];
  let products = [...db.products];
  let nextInvoiceNumber = db.nextInvoiceNumber;
  const newInvoices = [];
  const newInvMoves = [];
  const newCashMoves = [];
  const errors = [];

  groups.forEach((g) => {
    const first = g.rows[0];
    let client = clients.find((c) => c.name.trim().toLowerCase() === first.client.toLowerCase());
    if (!client) {
      client = { id: uid(), name: first.client, nit: '', phone: '', email: '', address: '' };
      clients = [...clients, client];
    }
    const items = g.rows.map((r) => {
      let productId = null;
      if (r.product) {
        let product = products.find((p) => p.name.trim().toLowerCase() === r.product.toLowerCase());
        if (!product) {
          product = { id: uid(), name: r.product, unit: r.unit || 'Unidades', price: r.rate };
          products = [...products, product];
        }
        productId = product.id;
      }
      return {
        id: uid(), description: r.description, lote: r.lote, rate: r.rate, qty: r.qty, unit: r.unit,
        total: r.rate * r.qty, productId, cost: productId ? avgCostOf(db, productId) : 0,
      };
    });
    const total = items.reduce((s, it) => s + it.total, 0);
    const number = 'INV' + String(nextInvoiceNumber).padStart(4, '0');
    nextInvoiceNumber += 1;
    const invoiceId = uid();

    const paidAmount = g.rows.reduce((s, r) => s + (r.paidAmount || 0), 0);
    const accountName = g.rows.map((r) => r.accountName).find((a) => a);
    const payments = [];
    if (paidAmount > 0) {
      if (!accountName) {
        errors.push(`Factura ${number} (${first.client}): se indicó un monto pagado pero no una cuenta; se creó sin registrar el pago.`);
      } else {
        const account = db.accounts.find((a) => a.name.trim().toLowerCase() === accountName.toLowerCase());
        if (!account) {
          errors.push(`Factura ${number} (${first.client}): la cuenta "${accountName}" no existe; se creó sin registrar el pago.`);
        } else {
          const paymentId = uid();
          payments.push({ id: paymentId, date: first.date, amount: paidAmount, accountId: account.id, method: 'Importado' });
          newCashMoves.push({ id: uid(), paymentId, accountId: account.id, date: first.date, amount: paidAmount, concept: `Pago factura ${number} (importado)`, refType: 'venta', refId: invoiceId });
        }
      }
    }

    newInvoices.push({ id: invoiceId, number, date: first.date, dueDate: first.dueDate, clientId: client.id, items, total, payments });
    items.forEach((it) => {
      if (it.productId) newInvMoves.push({ id: uid(), productId: it.productId, date: first.date, qty: -it.qty, unitCost: it.cost, type: 'venta', refType: 'venta', refId: invoiceId });
    });
  });

  return {
    newDb: {
      ...db, clients, products,
      invoices: [...newInvoices, ...db.invoices],
      inventoryMovements: [...(db.inventoryMovements || []), ...newInvMoves],
      movements: [...db.movements, ...newCashMoves],
      nextInvoiceNumber,
    },
    createdCount: newInvoices.length,
    errors,
  };
}

function Ventas({ db, setDb }) {
  const [showNew, setShowNew] = useState(false);
  const [payFor, setPayFor] = useState(null);
  const [preview, setPreview] = useState(null);
  const [editing, setEditing] = useState(null);
  const [q, setQ] = useState('');
  const [importResult, setImportResult] = useState(null);
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);

  const handleFileChosen = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    setImporting(true);
    setImportResult(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const jsonRows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
        const { rows, errors: parseErrors } = parseSalesRows(jsonRows);
        if (rows.length === 0) {
          setImportResult({ created: 0, errors: parseErrors.length ? parseErrors : ['No se encontraron filas válidas en el archivo.'] });
        } else {
          const plan = buildImportPlan(db, rows);
          setDb(plan.newDb);
          setImportResult({ created: plan.createdCount, errors: [...parseErrors, ...plan.errors] });
        }
      } catch (err) {
        setImportResult({ created: 0, errors: ['No se pudo leer el archivo. Verifica que sea un .xlsx exportado desde la plantilla.'] });
      }
      setImporting(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const rows = useMemo(() => {
    return db.invoices.filter((inv) => {
      const client = db.clients.find((c) => c.id === inv.clientId);
      const text = `${inv.number} ${client ? client.name : ''}`.toLowerCase();
      return text.includes(q.toLowerCase());
    });
  }, [db.invoices, db.clients, q]);

  const addPayment = (invoiceId, payment) => {
    setDb((prev) => {
      const inv = prev.invoices.find((i) => i.id === invoiceId);
      const paymentId = uid();
      const invoices = prev.invoices.map((i) => i.id === invoiceId ? { ...i, payments: [...(i.payments || []), { ...payment, id: paymentId }] } : i);
      const movements = [...prev.movements, { id: uid(), paymentId, accountId: payment.accountId, date: payment.date, amount: payment.amount, concept: `Pago factura ${inv.number}`, refType: 'venta', refId: invoiceId }];
      return { ...prev, invoices, movements };
    });
  };

  const deleteInvoice = (invoiceId) => {
    setDb((prev) => ({
      ...prev,
      invoices: prev.invoices.filter((i) => i.id !== invoiceId),
      inventoryMovements: (prev.inventoryMovements || []).filter((m) => !(m.refType === 'venta' && m.refId === invoiceId)),
      movements: prev.movements.filter((m) => !(m.refType === 'venta' && m.refId === invoiceId)),
    }));
  };

  const deleteInvoicePayment = (invoiceId, paymentId) => {
    setDb((prev) => ({
      ...prev,
      invoices: prev.invoices.map((i) => i.id === invoiceId ? { ...i, payments: (i.payments || []).filter((p) => p.id !== paymentId) } : i),
      movements: prev.movements.filter((m) => m.paymentId !== paymentId),
    }));
  };

  const editInvoicePaymentDate = (invoiceId, paymentId, newDate) => {
    setDb((prev) => ({
      ...prev,
      invoices: prev.invoices.map((i) => i.id === invoiceId ? { ...i, payments: (i.payments || []).map((p) => p.id === paymentId ? { ...p, date: newDate } : p) } : i),
      movements: prev.movements.map((m) => m.paymentId === paymentId ? { ...m, date: newDate } : m),
    }));
  };

  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Ventas</div>
          <div className="qn-page-sub">Registra facturas de venta y descárgalas en Excel</div>
        </div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          <button className="qn-btn" onClick={downloadSalesTemplate}><Download size={14} /> Descargar plantilla</button>
          <button className="qn-btn" disabled={importing} onClick={() => fileInputRef.current && fileInputRef.current.click()}>
            <Upload size={14} /> {importing ? 'Importando…' : 'Subir ventas en Excel'}
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleFileChosen} />
          <button className="qn-btn qn-btn-primary" onClick={() => setShowNew(true)}><Plus size={15} /> Nueva venta</button>
        </div>
      </div>

      {importResult && (
        <div className="qn-card qn-section" style={{ borderColor: importResult.errors.length ? C.amber : C.teal, marginBottom: 18 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'start', gap: 12 }}>
            <div>
              <div style={{ fontWeight: 600, marginBottom: importResult.errors.length ? 8 : 0 }}>
                {importResult.created > 0
                  ? `Se importaron ${importResult.created} factura${importResult.created === 1 ? '' : 's'} correctamente.`
                  : 'No se importó ninguna factura.'}
              </div>
              {importResult.errors.length > 0 && (
                <ul style={{ margin: 0, paddingLeft: 18, fontSize: 12.5, color: C.inkSoft }}>
                  {importResult.errors.map((e, i) => <li key={i} style={{ marginBottom: 3 }}>{e}</li>)}
                </ul>
              )}
            </div>
            <button className="qn-close" onClick={() => setImportResult(null)}><X size={16} /></button>
          </div>
        </div>
      )}

      <div className="qn-card qn-section">
        <div style={{ display: 'flex', gap: 10, marginBottom: 14 }}>
          <div style={{ position: 'relative', flex: 1, maxWidth: 320 }}>
            <Search size={14} style={{ position: 'absolute', left: 10, top: 10, color: C.inkSoft }} />
            <input className="qn-input" style={{ paddingLeft: 30 }} placeholder="Buscar por factura o cliente" value={q} onChange={(e) => setQ(e.target.value)} />
          </div>
        </div>
        {rows.length === 0 ? (
          <Empty title="Aún no hay ventas" sub="Registra tu primera factura con el botón «Nueva venta»." />
        ) : (
          <table className="qn-table">
            <thead>
              <tr><th>Factura</th><th>Cliente</th><th>Fecha</th><th>Vence</th><th>Total</th><th>Saldo</th><th>Estado</th><th></th></tr>
            </thead>
            <tbody>
              {rows.map((inv) => {
                const client = db.clients.find((c) => c.id === inv.clientId);
                const bal = balanceOf(inv);
                return (
                  <tr key={inv.id}>
                    <td className="qn-mono">{inv.number}</td>
                    <td>{client ? client.name : '—'}</td>
                    <td>{fmtDate(inv.date)}</td>
                    <td>{fmtDate(inv.dueDate)}</td>
                    <td>{fmtMoney(inv.total)}</td>
                    <td style={{ fontWeight: 600, color: bal > 0 ? C.rust : C.teal }}>{fmtMoney(bal)}</td>
                    <td><Badge status={statusOf(inv)} /></td>
                    <td>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button className="qn-btn qn-btn-sm qn-btn-icon" title="Ver" onClick={() => setPreview(inv.id)}><Eye size={14} /></button>
                        <button className="qn-btn qn-btn-sm qn-btn-icon" title="Editar" onClick={() => setEditing(inv.id)}><Pencil size={14} /></button>
                        <button className="qn-btn qn-btn-sm qn-btn-icon" title="Excel" onClick={() => downloadInvoiceExcel(db, inv)}><FileSpreadsheet size={14} /></button>
                        {bal > 0 && <button className="qn-btn qn-btn-sm" onClick={() => setPayFor(inv)}>Registrar pago</button>}
                        <ConfirmDelete onConfirm={() => deleteInvoice(inv.id)} label="Eliminar factura" />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      {showNew && <NewSaleModal db={db} setDb={setDb} onClose={() => setShowNew(false)} />}
      {editing && db.invoices.find((i) => i.id === editing) && (
        <EditSaleModal db={db} setDb={setDb} invoice={db.invoices.find((i) => i.id === editing)} onClose={() => setEditing(null)} />
      )}
      {payFor && <PaymentModal db={db} doc={payFor} kind="cxc" onClose={() => setPayFor(null)} onSave={(p) => addPayment(payFor.id, p)} />}
      {preview && db.invoices.find((i) => i.id === preview) && (
        <InvoicePreview
          db={db}
          invoice={db.invoices.find((i) => i.id === preview)}
          onClose={() => setPreview(null)}
          onDeletePayment={(paymentId) => deleteInvoicePayment(preview, paymentId)}
          onEditPaymentDate={(paymentId, newDate) => editInvoicePaymentDate(preview, paymentId, newDate)}
          onDeleteInvoice={() => { deleteInvoice(preview); setPreview(null); }}
        />
      )}
    </div>
  );
}

/* ---- Importación masiva de pagos desde Excel (aplica a Ventas y Cuentas por pagar) ---- */
const PAYMENT_HEADER_ALIASES = {
  referencia: 'reference', 'no factura': 'reference', 'n factura': 'reference', factura: 'reference', ref: 'reference',
  fecha: 'date', 'fecha pago': 'date',
  monto: 'amount', valor: 'amount', pago: 'amount', abono: 'amount',
  cuenta: 'accountName', 'cuenta de pago': 'accountName',
  metodo: 'method',
};

function downloadPaymentsTemplate() {
  const headers = ['Referencia', 'Fecha', 'Monto', 'Cuenta', 'Método'];
  const rows = [
    headers,
    ['INV0019', '2026-09-01', 1900000, 'Caja general', 'Transferencia'],
    ['CXP0003', '2026-09-02', 500000, 'Bancolombia', 'Transferencia'],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 20 }, { wch: 16 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Pagos');
  XLSX.writeFile(wb, 'Plantilla_pagos.xlsx');
}

function parsePaymentsRows(jsonRows) {
  const errors = [];
  const parsed = [];
  jsonRows.forEach((raw, idx) => {
    const rowNum = idx + 2;
    const row = {};
    Object.keys(raw).forEach((k) => {
      const canon = PAYMENT_HEADER_ALIASES[normalizeHeader(k)];
      if (canon) row[canon] = raw[k];
    });
    const allEmpty = Object.values(raw).every((v) => String(v ?? '').trim() === '');
    if (allEmpty) return;
    const reference = String(row.reference || '').trim();
    const dateVal = normalizeDateValue(row.date);
    const amount = parseNumberLike(row.amount);
    const accountName = String(row.accountName || '').trim();
    if (!reference) { errors.push(`Fila ${rowNum}: falta la referencia (N° de factura o de cuenta por pagar).`); return; }
    if (!dateVal) { errors.push(`Fila ${rowNum}: fecha inválida o vacía.`); return; }
    if (!amount || amount <= 0 || isNaN(amount)) { errors.push(`Fila ${rowNum}: monto inválido.`); return; }
    if (!accountName) { errors.push(`Fila ${rowNum}: falta la cuenta.`); return; }
    parsed.push({ rowNum, reference, date: dateVal, amount, accountName, method: String(row.method || '').trim() });
  });
  return { rows: parsed, errors };
}

function buildPaymentsImportPlan(db, rows) {
  let invoices = [...db.invoices];
  let payables = [...db.payables];
  const newMovements = [];
  const errors = [];
  let appliedCount = 0;

  rows.forEach((r) => {
    const ref = r.reference.toUpperCase();
    const account = db.accounts.find((a) => a.name.trim().toLowerCase() === r.accountName.toLowerCase());
    if (!account) { errors.push(`Fila ${r.rowNum}: la cuenta "${r.accountName}" no existe.`); return; }

    if (ref.startsWith('INV')) {
      const idx = invoices.findIndex((i) => i.number.toUpperCase() === ref);
      if (idx === -1) { errors.push(`Fila ${r.rowNum}: no existe la factura ${r.reference}.`); return; }
      const inv = invoices[idx];
      const paymentId = uid();
      const payment = { id: paymentId, date: r.date, amount: r.amount, accountId: account.id, method: r.method || 'Importado' };
      invoices[idx] = { ...inv, payments: [...(inv.payments || []), payment] };
      newMovements.push({ id: uid(), paymentId, accountId: account.id, date: r.date, amount: r.amount, concept: `Pago factura ${inv.number} (importado)`, refType: 'venta', refId: inv.id });
      appliedCount += 1;
    } else if (ref.startsWith('CXP')) {
      const idx = payables.findIndex((p) => p.number.toUpperCase() === ref);
      if (idx === -1) { errors.push(`Fila ${r.rowNum}: no existe la cuenta por pagar ${r.reference}.`); return; }
      const pay = payables[idx];
      const paymentId = uid();
      const payment = { id: paymentId, date: r.date, amount: r.amount, accountId: account.id, method: r.method || 'Importado' };
      payables[idx] = { ...pay, payments: [...(pay.payments || []), payment] };
      newMovements.push({ id: uid(), paymentId, accountId: account.id, date: r.date, amount: -r.amount, concept: `Pago ${pay.number} (importado)`, refType: 'cxp', refId: pay.id });
      appliedCount += 1;
    } else {
      errors.push(`Fila ${r.rowNum}: la referencia "${r.reference}" no empieza con INV (factura) ni CXP (cuenta por pagar).`);
    }
  });

  return {
    newDb: { ...db, invoices, payables, movements: [...db.movements, ...newMovements] },
    appliedCount,
    errors,
  };
}

function ImportPaymentsControls({ db, setDb }) {
  const [importing, setImporting] = useState(false);
  const [result, setResult] = useState(null);
  const fileInputRef = useRef(null);

  const handleFileChosen = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    setImporting(true);
    setResult(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const jsonRows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
        const { rows, errors: parseErrors } = parsePaymentsRows(jsonRows);
        if (rows.length === 0) {
          setResult({ applied: 0, errors: parseErrors.length ? parseErrors : ['No se encontraron filas válidas en el archivo.'] });
        } else {
          const plan = buildPaymentsImportPlan(db, rows);
          setDb(plan.newDb);
          setResult({ applied: plan.appliedCount, errors: [...parseErrors, ...plan.errors] });
        }
      } catch (err) {
        setResult({ applied: 0, errors: ['No se pudo leer el archivo. Verifica que sea un .xlsx exportado desde la plantilla.'] });
      }
      setImporting(false);
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <>
      <button className="qn-btn" onClick={downloadPaymentsTemplate}><Download size={14} /> Plantilla de pagos</button>
      <button className="qn-btn" disabled={importing} onClick={() => fileInputRef.current && fileInputRef.current.click()}>
        <Upload size={14} /> {importing ? 'Importando…' : 'Subir pagos en Excel'}
      </button>
      <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleFileChosen} />
      {result && (
        <div className="qn-card qn-section" style={{ borderColor: result.errors.length ? C.amber : C.teal, width: '100%', order: 99 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'start', gap: 12 }}>
            <div>
              <div style={{ fontWeight: 600, marginBottom: result.errors.length ? 8 : 0 }}>
                {result.applied > 0 ? `Se registraron ${result.applied} pago${result.applied === 1 ? '' : 's'} correctamente.` : 'No se registró ningún pago.'}
              </div>
              {result.errors.length > 0 && (
                <ul style={{ margin: 0, paddingLeft: 18, fontSize: 12.5, color: C.inkSoft }}>
                  {result.errors.map((e, i) => <li key={i} style={{ marginBottom: 3 }}>{e}</li>)}
                </ul>
              )}
            </div>
            <button className="qn-close" onClick={() => setResult(null)}><X size={16} /></button>
          </div>
        </div>
      )}
    </>
  );
}

/* ============================== CUENTAS POR COBRAR ============================== */
function CuentasCobrar({ db, setDb }) {
  const [payFor, setPayFor] = useState(null);
  const [preview, setPreview] = useState(null);
  const [filter, setFilter] = useState('pendientes');
  const rows = db.invoices.filter((inv) => filter === 'todas' ? true : balanceOf(inv) > 0);
  const addPayment = (invoiceId, payment) => {
    setDb((prev) => {
      const inv = prev.invoices.find((i) => i.id === invoiceId);
      const paymentId = uid();
      const invoices = prev.invoices.map((i) => i.id === invoiceId ? { ...i, payments: [...(i.payments || []), { ...payment, id: paymentId }] } : i);
      const movements = [...prev.movements, { id: uid(), paymentId, accountId: payment.accountId, date: payment.date, amount: payment.amount, concept: `Pago factura ${inv.number}`, refType: 'venta', refId: invoiceId }];
      return { ...prev, invoices, movements };
    });
  };
  const deleteInvoice = (invoiceId) => {
    setDb((prev) => ({
      ...prev,
      invoices: prev.invoices.filter((i) => i.id !== invoiceId),
      inventoryMovements: (prev.inventoryMovements || []).filter((m) => !(m.refType === 'venta' && m.refId === invoiceId)),
      movements: prev.movements.filter((m) => !(m.refType === 'venta' && m.refId === invoiceId)),
    }));
  };
  const deleteInvoicePayment = (invoiceId, paymentId) => {
    setDb((prev) => ({
      ...prev,
      invoices: prev.invoices.map((i) => i.id === invoiceId ? { ...i, payments: (i.payments || []).filter((p) => p.id !== paymentId) } : i),
      movements: prev.movements.filter((m) => m.paymentId !== paymentId),
    }));
  };
  const editInvoicePaymentDate = (invoiceId, paymentId, newDate) => {
    setDb((prev) => ({
      ...prev,
      invoices: prev.invoices.map((i) => i.id === invoiceId ? { ...i, payments: (i.payments || []).map((p) => p.id === paymentId ? { ...p, date: newDate } : p) } : i),
      movements: prev.movements.map((m) => m.paymentId === paymentId ? { ...m, date: newDate } : m),
    }));
  };
  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Cuentas por cobrar</div>
          <div className="qn-page-sub">Saldo pendiente total: <strong>{fmtCOP(totalAR(db))}</strong></div>
        </div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'start', maxWidth: 520 }}>
          <select className="qn-select" style={{ width: 190 }} value={filter} onChange={(e) => setFilter(e.target.value)}>
            <option value="pendientes">Con saldo pendiente</option>
            <option value="todas">Todas las facturas</option>
          </select>
          <ImportPaymentsControls db={db} setDb={setDb} />
        </div>
      </div>
      <div className="qn-card qn-section">
        {rows.length === 0 ? (
          <Empty title="Todo al día" sub="No hay facturas con saldo pendiente en este filtro." />
        ) : (
          <table className="qn-table">
            <thead><tr><th>Factura</th><th>Cliente</th><th>Vence</th><th>Total</th><th>Saldo</th><th>Estado</th><th></th></tr></thead>
            <tbody>
              {rows.map((inv) => {
                const client = db.clients.find((c) => c.id === inv.clientId);
                const bal = balanceOf(inv);
                return (
                  <tr key={inv.id}>
                    <td className="qn-mono">{inv.number}</td>
                    <td>{client ? client.name : '—'}</td>
                    <td>{fmtDate(inv.dueDate)}</td>
                    <td>{fmtMoney(inv.total)}</td>
                    <td style={{ fontWeight: 600, color: bal > 0 ? C.rust : C.teal }}>{fmtMoney(bal)}</td>
                    <td><Badge status={statusOf(inv)} /></td>
                    <td>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button className="qn-btn qn-btn-sm qn-btn-icon" title="Ver" onClick={() => setPreview(inv.id)}><Eye size={14} /></button>
                        {bal > 0 && <button className="qn-btn qn-btn-sm" onClick={() => setPayFor(inv)}><ArrowDownCircle size={13} /> Registrar pago</button>}
                        <ConfirmDelete onConfirm={() => deleteInvoice(inv.id)} label="Eliminar factura" />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
      {payFor && <PaymentModal db={db} doc={payFor} kind="cxc" onClose={() => setPayFor(null)} onSave={(p) => addPayment(payFor.id, p)} />}
      {preview && db.invoices.find((i) => i.id === preview) && (
        <InvoicePreview
          db={db}
          invoice={db.invoices.find((i) => i.id === preview)}
          onClose={() => setPreview(null)}
          onDeletePayment={(paymentId) => deleteInvoicePayment(preview, paymentId)}
          onEditPaymentDate={(paymentId, newDate) => editInvoicePaymentDate(preview, paymentId, newDate)}
          onDeleteInvoice={() => { deleteInvoice(preview); setPreview(null); }}
        />
      )}
    </div>
  );
}

/* ============================== CARTERA DE CLIENTES ============================== */
function Cartera({ db }) {
  const [expanded, setExpanded] = useState(null);

  const rows = useMemo(() => {
    const map = {};
    db.invoices.forEach((inv) => {
      const bal = balanceOf(inv);
      if (bal <= 0.5) return;
      if (!map[inv.clientId]) map[inv.clientId] = { clientId: inv.clientId, invoices: [], facturado: 0, pagado: 0, saldo: 0 };
      map[inv.clientId].invoices.push(inv);
      map[inv.clientId].facturado += inv.total;
      map[inv.clientId].pagado += paidOf(inv);
      map[inv.clientId].saldo += bal;
    });
    return Object.values(map).sort((a, b) => b.saldo - a.saldo);
  }, [db.invoices]);

  const totalCartera = rows.reduce((s, r) => s + r.saldo, 0);

  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Cartera de clientes</div>
          <div className="qn-page-sub">Total en cartera: <strong>{fmtCOP(totalCartera)}</strong></div>
        </div>
      </div>
      <div className="qn-card qn-section">
        {rows.length === 0 ? (
          <Empty title="Cartera al día" sub="Ningún cliente tiene saldo pendiente en este momento." />
        ) : (
          <table className="qn-table">
            <thead>
              <tr><th>Cliente</th><th>Facturas pendientes</th><th style={{ textAlign: 'right' }}>Facturado</th><th style={{ textAlign: 'right' }}>Abonado</th><th style={{ textAlign: 'right' }}>Saldo pendiente</th><th></th></tr>
            </thead>
            <tbody>
              {rows.map((r) => {
                const client = db.clients.find((c) => c.id === r.clientId);
                const isOpen = expanded === r.clientId;
                return (
                  <React.Fragment key={r.clientId}>
                    <tr>
                      <td style={{ fontWeight: 600 }}>{client ? client.name : '(cliente eliminado)'}</td>
                      <td>{r.invoices.length}</td>
                      <td style={{ textAlign: 'right' }}>{fmtMoney(r.facturado)}</td>
                      <td style={{ textAlign: 'right' }}>{fmtMoney(r.pagado)}</td>
                      <td style={{ textAlign: 'right', fontWeight: 700, color: C.rust }}>{fmtMoney(r.saldo)}</td>
                      <td>
                        <button className="qn-btn qn-btn-sm qn-btn-icon" title="Ver facturas" onClick={() => setExpanded(isOpen ? null : r.clientId)}>
                          {isOpen ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                        </button>
                      </td>
                    </tr>
                    {isOpen && r.invoices
                      .sort((a, b) => a.dueDate.localeCompare(b.dueDate))
                      .map((inv) => (
                        <tr key={inv.id} style={{ background: C.bg }}>
                          <td colSpan={2} style={{ paddingLeft: 24 }}>
                            <span className="qn-mono">{inv.number}</span>
                            <span style={{ color: C.inkSoft, fontSize: 12 }}> · vence {fmtDate(inv.dueDate)}</span>
                          </td>
                          <td style={{ textAlign: 'right' }}>{fmtMoney(inv.total)}</td>
                          <td style={{ textAlign: 'right' }}>{fmtMoney(paidOf(inv))}</td>
                          <td style={{ textAlign: 'right', fontWeight: 600 }}>{fmtMoney(balanceOf(inv))}</td>
                          <td><Badge status={statusOf(inv)} /></td>
                        </tr>
                      ))}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

/* ============================== CUENTAS POR PAGAR ============================== */
function NewPayableModal({ db, setDb, onClose }) {
  const [type, setType] = useState('Proveedor');
  const [supplierId, setSupplierId] = useState(db.suppliers[0]?.id || '');
  const [newSupplier, setNewSupplier] = useState('');
  const [concept, setConcept] = useState('');
  const [total, setTotal] = useState(0);
  const [date, setDate] = useState(todayISO());
  const [dueDate, setDueDate] = useState(addDays(todayISO(), 15));

  const missingSupplier = type === 'Proveedor' && !supplierId && !newSupplier.trim();
  const save = () => {
    if (total <= 0 || missingSupplier) return;
    setDb((prev) => {
      let suppliers = prev.suppliers;
      let supId = supplierId;
      if (type === 'Proveedor' && !supId && newSupplier.trim()) {
        const s = { id: uid(), name: newSupplier.trim(), phone: '', email: '' };
        suppliers = [...suppliers, s];
        supId = s.id;
      }
      const number = 'CXP' + String(prev.nextBillNumber).padStart(4, '0');
      const payable = { id: uid(), number, type, supplierId: type === 'Proveedor' ? supId : null, concept, date, dueDate, total: Number(total), payments: [] };
      return { ...prev, suppliers, payables: [payable, ...prev.payables], nextBillNumber: prev.nextBillNumber + 1 };
    });
    onClose();
  };

  return (
    <Modal title="Nueva cuenta por pagar" onClose={onClose}>
      <div className="qn-field">
        <label className="qn-label">Tipo</label>
        <select className="qn-select" value={type} onChange={(e) => setType(e.target.value)}>
          <option value="Proveedor">Factura de proveedor</option>
          <option value="Gasto">Gasto general</option>
        </select>
      </div>
      {type === 'Proveedor' ? (
        <div className="qn-field">
          <label className="qn-label">Proveedor</label>
          <select className="qn-select" value={supplierId} onChange={(e) => setSupplierId(e.target.value)} style={{ marginBottom: 6 }}>
            <option value="">— Nuevo proveedor —</option>
            {db.suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          {!supplierId && <input className="qn-input" placeholder="Nombre del nuevo proveedor" value={newSupplier} onChange={(e) => setNewSupplier(e.target.value)} />}
        </div>
      ) : null}
      <div className="qn-field">
        <label className="qn-label">Concepto</label>
        <input className="qn-input" placeholder={type === 'Proveedor' ? 'Ej: Materia prima, insumos...' : 'Ej: Arriendo, servicios públicos...'} value={concept} onChange={(e) => setConcept(e.target.value)} />
      </div>
      <div className="qn-row3">
        <div className="qn-field">
          <label className="qn-label">Valor total</label>
          <input className="qn-input" type="number" value={total} onChange={(e) => setTotal(e.target.value)} />
        </div>
        <div className="qn-field">
          <label className="qn-label">Fecha</label>
          <input className="qn-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
        <div className="qn-field">
          <label className="qn-label">Vence</label>
          <input className="qn-input" type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} />
        </div>
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 8 }}>
        <button className="qn-btn" onClick={onClose}>Cancelar</button>
        <button className="qn-btn qn-btn-primary" disabled={total <= 0 || missingSupplier} onClick={save}><Check size={14} /> Guardar</button>
      </div>
    </Modal>
  );
}

function PayableDetail({ db, payable, onClose, onDeletePayment, onDeletePayable, onEditPaymentDate }) {
  const supplier = db.suppliers.find((s) => s.id === payable.supplierId);
  const bal = balanceOf(payable);
  return (
    <Modal title={`Cuenta por pagar ${payable.number}`} onClose={onClose} wide>
      <div className="qn-row3" style={{ marginBottom: 16 }}>
        <div>
          <div className="qn-label">Tipo</div>
          <div style={{ fontWeight: 600 }}>{payable.type}</div>
        </div>
        <div>
          <div className="qn-label">{payable.type === 'Proveedor' ? 'Proveedor' : 'Concepto'}</div>
          <div style={{ fontWeight: 600 }}>{payable.type === 'Proveedor' ? (supplier ? supplier.name : '—') : payable.concept}</div>
        </div>
        <div>
          <div className="qn-label">Vence</div>
          <div style={{ fontWeight: 600 }}>{fmtDate(payable.dueDate)}</div>
        </div>
      </div>
      {payable.type === 'Proveedor' && payable.concept && (
        <div style={{ fontSize: 13, color: C.inkSoft, marginBottom: 16 }}>{payable.concept}</div>
      )}
      <div style={{ display: 'flex', gap: 40, marginBottom: 20, paddingBottom: 14, borderBottom: `1px solid ${C.line}` }}>
        <div>
          <div className="qn-label">Total</div>
          <div className="qn-display" style={{ fontSize: 18, fontWeight: 700 }}>{fmtCOP(payable.total)}</div>
        </div>
        <div>
          <div className="qn-label">Saldo pendiente</div>
          <div className="qn-display" style={{ fontSize: 18, fontWeight: 700, color: bal > 0 ? C.rust : C.teal }}>{fmtCOP(bal)}</div>
        </div>
      </div>

      <div className="qn-section-title" style={{ marginBottom: 8 }}>Pagos registrados</div>
      {(payable.payments || []).length === 0 ? (
        <div style={{ fontSize: 12.5, color: C.inkSoft, marginBottom: 8 }}>Aún no se ha registrado ningún pago.</div>
      ) : (
        <table className="qn-table">
          <thead><tr><th>Fecha</th><th>Cuenta</th><th>Método</th><th style={{ textAlign: 'right' }}>Monto</th><th></th></tr></thead>
          <tbody>
            {payable.payments.map((p) => {
              const acc = db.accounts.find((a) => a.id === p.accountId);
              return (
                <tr key={p.id}>
                  <td><input className="qn-input" type="date" value={p.date} onChange={(e) => onEditPaymentDate(p.id, e.target.value)} style={{ maxWidth: 150 }} /></td>
                  <td>{acc ? acc.name : '—'}</td>
                  <td>{p.method || '—'}</td>
                  <td style={{ textAlign: 'right', fontWeight: 600 }}>{fmtMoney(p.amount)}</td>
                  <td><ConfirmDelete onConfirm={() => onDeletePayment(p.id)} label="Eliminar pago" /></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 18 }}>
        <ConfirmDelete onConfirm={onDeletePayable} label="Eliminar cuenta por pagar completa" />
        <button className="qn-btn" onClick={onClose}>Cerrar</button>
      </div>
    </Modal>
  );
}

function CuentasPagar({ db, setDb }) {
  const [showNew, setShowNew] = useState(false);
  const [payFor, setPayFor] = useState(null);
  const [preview, setPreview] = useState(null);
  const [filter, setFilter] = useState('pendientes');
  const rows = db.payables.filter((p) => filter === 'todas' ? true : balanceOf(p) > 0);

  const supplierTotals = useMemo(() => {
    const map = {};
    db.payables.filter((p) => p.type === 'Proveedor' && balanceOf(p) > 0).forEach((p) => {
      const sup = db.suppliers.find((s) => s.id === p.supplierId);
      const name = sup ? sup.name : '(proveedor eliminado)';
      if (!map[name]) map[name] = { total: 0, count: 0 };
      map[name].total += balanceOf(p);
      map[name].count += 1;
    });
    return Object.entries(map).sort((a, b) => b[1].total - a[1].total).map(([name, v]) => ({ name, ...v }));
  }, [db.payables, db.suppliers]);

  const addPayment = (payableId, payment) => {
    setDb((prev) => {
      const doc = prev.payables.find((p) => p.id === payableId);
      const paymentId = uid();
      const payables = prev.payables.map((p) => p.id === payableId ? { ...p, payments: [...(p.payments || []), { ...payment, id: paymentId }] } : p);
      const movements = [...prev.movements, { id: uid(), paymentId, accountId: payment.accountId, date: payment.date, amount: -Math.abs(payment.amount), concept: `Pago ${doc.number} · ${doc.concept || doc.type}`, refType: 'cxp', refId: payableId }];
      return { ...prev, payables, movements };
    });
  };

  const deletePayable = (payableId) => {
    setDb((prev) => ({
      ...prev,
      payables: prev.payables.filter((p) => p.id !== payableId),
      movements: prev.movements.filter((m) => !(m.refType === 'cxp' && m.refId === payableId)),
    }));
  };

  const deletePayablePayment = (payableId, paymentId) => {
    setDb((prev) => ({
      ...prev,
      payables: prev.payables.map((p) => p.id === payableId ? { ...p, payments: (p.payments || []).filter((x) => x.id !== paymentId) } : p),
      movements: prev.movements.filter((m) => m.paymentId !== paymentId),
    }));
  };

  const editPayablePaymentDate = (payableId, paymentId, newDate) => {
    setDb((prev) => ({
      ...prev,
      payables: prev.payables.map((p) => p.id === payableId ? { ...p, payments: (p.payments || []).map((x) => x.id === paymentId ? { ...x, date: newDate } : x) } : p),
      movements: prev.movements.map((m) => m.paymentId === paymentId ? { ...m, date: newDate } : m),
    }));
  };

  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Cuentas por pagar</div>
          <div className="qn-page-sub">Saldo pendiente total: <strong>{fmtCOP(totalAP(db))}</strong></div>
        </div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'flex-end', maxWidth: 620 }}>
          <select className="qn-select" style={{ width: 190 }} value={filter} onChange={(e) => setFilter(e.target.value)}>
            <option value="pendientes">Con saldo pendiente</option>
            <option value="todas">Todas</option>
          </select>
          <ImportPaymentsControls db={db} setDb={setDb} />
          <button className="qn-btn qn-btn-primary" onClick={() => setShowNew(true)}><Plus size={15} /> Nueva cuenta por pagar</button>
        </div>
      </div>

      {supplierTotals.length > 0 && (
        <div className="qn-card qn-section">
          <div className="qn-section-title">Total adeudado por proveedor</div>
          <table className="qn-table">
            <thead><tr><th>Proveedor</th><th>Facturas pendientes</th><th style={{ textAlign: 'right' }}>Total adeudado</th></tr></thead>
            <tbody>
              {supplierTotals.map((s) => (
                <tr key={s.name}>
                  <td>{s.name}</td>
                  <td>{s.count}</td>
                  <td style={{ textAlign: 'right', fontWeight: 600, color: C.rust }}>{fmtMoney(s.total)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="qn-card qn-section">
        {rows.length === 0 ? (
          <Empty title="Sin cuentas por pagar" sub="Registra facturas de proveedores o gastos generales." />
        ) : (
          <table className="qn-table">
            <thead><tr><th>Ref.</th><th>Tipo</th><th>Concepto</th><th>Vence</th><th>Total</th><th>Saldo</th><th>Estado</th><th></th></tr></thead>
            <tbody>
              {rows.map((p) => {
                const supplier = db.suppliers.find((s) => s.id === p.supplierId);
                const bal = balanceOf(p);
                return (
                  <tr key={p.id}>
                    <td className="qn-mono">{p.number}</td>
                    <td>{p.type}</td>
                    <td>{p.type === 'Proveedor' ? (supplier ? supplier.name : '—') : p.concept}</td>
                    <td>{fmtDate(p.dueDate)}</td>
                    <td>{fmtMoney(p.total)}</td>
                    <td style={{ fontWeight: 600, color: bal > 0 ? C.rust : C.teal }}>{fmtMoney(bal)}</td>
                    <td><Badge status={statusOf(p)} /></td>
                    <td>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button className="qn-btn qn-btn-sm qn-btn-icon" title="Ver" onClick={() => setPreview(p.id)}><Eye size={14} /></button>
                        {bal > 0 && <button className="qn-btn qn-btn-sm" onClick={() => setPayFor(p)}><ArrowUpCircle size={13} /> Registrar pago</button>}
                        <ConfirmDelete onConfirm={() => deletePayable(p.id)} label="Eliminar" />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
      {showNew && <NewPayableModal db={db} setDb={setDb} onClose={() => setShowNew(false)} />}
      {payFor && <PaymentModal db={db} doc={payFor} kind="cxp" onClose={() => setPayFor(null)} onSave={(p) => addPayment(payFor.id, p)} />}
      {preview && db.payables.find((p) => p.id === preview) && (
        <PayableDetail
          db={db}
          payable={db.payables.find((p) => p.id === preview)}
          onClose={() => setPreview(null)}
          onDeletePayment={(paymentId) => deletePayablePayment(preview, paymentId)}
          onEditPaymentDate={(paymentId, newDate) => editPayablePaymentDate(preview, paymentId, newDate)}
          onDeletePayable={() => { deletePayable(preview); setPreview(null); }}
        />
      )}
    </div>
  );
}

/* ============================== INVENTARIO ============================== */
function NewPurchaseModal({ db, setDb, onClose }) {
  const [productId, setProductId] = useState(db.products[0]?.id || '');
  const [newProductName, setNewProductName] = useState('');
  const [newProductUnit, setNewProductUnit] = useState('Unidades');
  const [qty, setQty] = useState(1);
  const [unitCost, setUnitCost] = useState(0);
  const [supplierId, setSupplierId] = useState(db.suppliers[0]?.id || '');
  const [newSupplier, setNewSupplier] = useState('');
  const [date, setDate] = useState(todayISO());
  const [dueDate, setDueDate] = useState(addDays(todayISO(), 15));
  const [paidNow, setPaidNow] = useState(false);
  const [accountId, setAccountId] = useState(db.accounts[0]?.id || '');

  const total = Number(qty || 0) * Number(unitCost || 0);
  const missingProduct = !productId && !newProductName.trim();
  const missingSupplier = !supplierId && !newSupplier.trim();

  const save = () => {
    if (total <= 0 || missingProduct || missingSupplier) return;
    setDb((prev) => {
      let products = prev.products;
      let prodId = productId;
      if (!prodId && newProductName.trim()) {
        const p = { id: uid(), name: newProductName.trim(), unit: newProductUnit, price: 0 };
        products = [...products, p];
        prodId = p.id;
      }
      let suppliers = prev.suppliers;
      let supId = supplierId;
      if (!supId && newSupplier.trim()) {
        const s = { id: uid(), name: newSupplier.trim(), phone: '', email: '' };
        suppliers = [...suppliers, s];
        supId = s.id;
      }
      const product = products.find((p) => p.id === prodId);
      const number = 'CXP' + String(prev.nextBillNumber).padStart(4, '0');
      const payable = {
        id: uid(), number, type: 'Proveedor', supplierId: supId,
        concept: `Compra ${qty} ${product?.unit || ''} · ${product?.name || ''}`,
        date, dueDate, total, payments: [],
      };
      const cashMoves = [];
      if (paidNow) {
        const paymentId = uid();
        payable.payments.push({ id: paymentId, date, amount: total, accountId, method: 'Efectivo/Transferencia' });
        cashMoves.push({ id: uid(), paymentId, accountId, date, amount: -total, concept: `Compra ${product?.name || ''} · ${payable.number}`, refType: 'cxp', refId: payable.id });
      }
      const invMove = { id: uid(), productId: prodId, date, qty: Number(qty), unitCost: Number(unitCost), type: 'compra', refType: 'compra', refId: payable.id };
      return {
        ...prev, products, suppliers,
        payables: [payable, ...prev.payables],
        nextBillNumber: prev.nextBillNumber + 1,
        inventoryMovements: [...(prev.inventoryMovements || []), invMove],
        movements: [...prev.movements, ...cashMoves],
      };
    });
    onClose();
  };

  return (
    <Modal title="Registrar compra de mercancía" onClose={onClose}>
      <div className="qn-field">
        <label className="qn-label">Producto</label>
        <select className="qn-select" value={productId} onChange={(e) => setProductId(e.target.value)} style={{ marginBottom: 6 }}>
          <option value="">— Producto nuevo —</option>
          {db.products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
        {!productId && (
          <div className="qn-row2">
            <input className="qn-input" placeholder="Nombre del producto" value={newProductName} onChange={(e) => setNewProductName(e.target.value)} />
            <input className="qn-input" placeholder="Unidad (Litros, Kg...)" value={newProductUnit} onChange={(e) => setNewProductUnit(e.target.value)} />
          </div>
        )}
      </div>
      <div className="qn-row2">
        <div className="qn-field"><label className="qn-label">Cantidad comprada</label><input className="qn-input" type="number" value={qty} onChange={(e) => setQty(e.target.value)} /></div>
        <div className="qn-field"><label className="qn-label">Costo por unidad</label><input className="qn-input" type="number" value={unitCost} onChange={(e) => setUnitCost(e.target.value)} /></div>
      </div>
      <div className="qn-field">
        <label className="qn-label">Proveedor</label>
        <select className="qn-select" value={supplierId} onChange={(e) => setSupplierId(e.target.value)} style={{ marginBottom: 6 }}>
          <option value="">— Nuevo proveedor —</option>
          {db.suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </select>
        {!supplierId && <input className="qn-input" placeholder="Nombre del nuevo proveedor" value={newSupplier} onChange={(e) => setNewSupplier(e.target.value)} />}
      </div>
      <div className="qn-row2">
        <div className="qn-field"><label className="qn-label">Fecha de compra</label><input className="qn-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} /></div>
        <div className="qn-field"><label className="qn-label">Vence (si es a crédito)</label><input className="qn-input" type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} /></div>
      </div>
      <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, marginBottom: 12 }}>
        <input type="checkbox" checked={paidNow} onChange={(e) => setPaidNow(e.target.checked)} /> Ya la pagué ahora
      </label>
      {paidNow && (
        <div className="qn-field">
          <label className="qn-label">Cuenta desde la que pagaste</label>
          <select className="qn-select" value={accountId} onChange={(e) => setAccountId(e.target.value)}>
            {db.accounts.map((a) => <option key={a.id} value={a.id}>{a.name} ({a.type})</option>)}
          </select>
        </div>
      )}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 8, paddingTop: 14, borderTop: `1px solid ${C.line}` }}>
        <div>
          <div className="qn-label" style={{ marginBottom: 2 }}>Total compra</div>
          <div className="qn-display" style={{ fontSize: 20, fontWeight: 700 }}>{fmtCOP(total)}</div>
        </div>
        <div style={{ display: 'flex', gap: 10 }}>
          <button className="qn-btn" onClick={onClose}>Cancelar</button>
          <button className="qn-btn qn-btn-primary" disabled={total <= 0 || missingProduct || missingSupplier} onClick={save}><Check size={14} /> Guardar compra</button>
        </div>
      </div>
    </Modal>
  );
}

function AdjustInventoryModal({ db, setDb, onClose }) {
  const [productId, setProductId] = useState(db.products[0]?.id || '');
  const [qty, setQty] = useState(0);
  const [reason, setReason] = useState('');
  const [date, setDate] = useState(todayISO());
  const save = () => {
    if (!productId || !qty) return;
    setDb((prev) => ({
      ...prev,
      inventoryMovements: [...(prev.inventoryMovements || []), { id: uid(), productId, date, qty: Number(qty), unitCost: 0, type: 'ajuste', refType: 'ajuste', refId: null, note: reason }],
    }));
    onClose();
  };
  return (
    <Modal title="Ajuste de inventario" onClose={onClose}>
      <div className="qn-field">
        <label className="qn-label">Producto</label>
        <select className="qn-select" value={productId} onChange={(e) => setProductId(e.target.value)}>
          {db.products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
      </div>
      <div className="qn-row2">
        <div className="qn-field">
          <label className="qn-label">Cantidad (negativo = merma o salida)</label>
          <input className="qn-input" type="number" value={qty} onChange={(e) => setQty(e.target.value)} />
        </div>
        <div className="qn-field"><label className="qn-label">Fecha</label><input className="qn-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} /></div>
      </div>
      <div className="qn-field"><label className="qn-label">Motivo</label><input className="qn-input" placeholder="Ej: conteo físico, merma, stock inicial" value={reason} onChange={(e) => setReason(e.target.value)} /></div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
        <button className="qn-btn" onClick={onClose}>Cancelar</button>
        <button className="qn-btn qn-btn-primary" disabled={!productId || !qty} onClick={save}><Check size={14} /> Guardar ajuste</button>
      </div>
    </Modal>
  );
}

function Inventario({ db, setDb }) {
  const [showPurchase, setShowPurchase] = useState(false);
  const [showAdjust, setShowAdjust] = useState(false);
  const totalValue = inventoryValue(db);
  const recent = [...(db.inventoryMovements || [])].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 12);
  const typeLabel = { compra: 'Compra', venta: 'Venta', ajuste: 'Ajuste' };

  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Inventario</div>
          <div className="qn-page-sub">Valor total en inventario: <strong>{fmtCOP(totalValue)}</strong></div>
        </div>
        <div style={{ display: 'flex', gap: 10 }}>
          <button className="qn-btn" onClick={() => setShowAdjust(true)}><Plus size={14} /> Ajuste</button>
          <button className="qn-btn qn-btn-primary" onClick={() => setShowPurchase(true)}><Plus size={15} /> Registrar compra</button>
        </div>
      </div>
      <div className="qn-card qn-section">
        <div className="qn-section-title">Existencias por producto</div>
        {db.products.length === 0 ? (
          <Empty title="Sin productos" sub="Crea uno desde «Clientes y productos» o al registrar tu primera compra." />
        ) : (
          <table className="qn-table">
            <thead><tr><th>Producto</th><th>Unidad</th><th>Stock actual</th><th>Costo promedio</th><th>Valor</th></tr></thead>
            <tbody>
              {db.products.map((p) => {
                const stock = stockOf(db, p.id);
                const cost = avgCostOf(db, p.id);
                return (
                  <tr key={p.id}>
                    <td>{p.name}</td>
                    <td>{p.unit}</td>
                    <td style={{ fontWeight: 600, color: stock < 0 ? C.rust : C.ink }}>{stock.toLocaleString('es-CO')}</td>
                    <td>{fmtMoney(cost)}</td>
                    <td>{fmtMoney(stock * cost)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
      <div className="qn-card qn-section">
        <div className="qn-section-title">Movimientos recientes</div>
        {recent.length === 0 ? (
          <Empty title="Sin movimientos" sub="Las compras y ventas de productos con inventario aparecerán aquí." />
        ) : (
          <table className="qn-table">
            <thead><tr><th>Fecha</th><th>Producto</th><th>Tipo</th><th style={{ textAlign: 'right' }}>Cantidad</th></tr></thead>
            <tbody>
              {recent.map((m) => {
                const p = db.products.find((x) => x.id === m.productId);
                return (
                  <tr key={m.id}>
                    <td>{fmtDate(m.date)}</td>
                    <td>{p ? p.name : '—'}</td>
                    <td>{typeLabel[m.type] || m.type}</td>
                    <td style={{ textAlign: 'right', fontWeight: 600, color: m.qty >= 0 ? C.teal : C.rust }}>{m.qty >= 0 ? '+' : ''}{m.qty.toLocaleString('es-CO')}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
      {showPurchase && <NewPurchaseModal db={db} setDb={setDb} onClose={() => setShowPurchase(false)} />}
      {showAdjust && <AdjustInventoryModal db={db} setDb={setDb} onClose={() => setShowAdjust(false)} />}
    </div>
  );
}

/* ============================== CAJA Y BANCOS ============================== */
function AccountModal({ onClose, onSave }) {
  const [name, setName] = useState('');
  const [type, setType] = useState('Banco');
  const [initialBalance, setInitialBalance] = useState(0);
  return (
    <Modal title="Nueva cuenta" onClose={onClose}>
      <div className="qn-field"><label className="qn-label">Nombre</label><input className="qn-input" placeholder="Ej: Cuenta de ahorros Davivienda" value={name} onChange={(e) => setName(e.target.value)} /></div>
      <div className="qn-row2">
        <div className="qn-field">
          <label className="qn-label">Tipo</label>
          <select className="qn-select" value={type} onChange={(e) => setType(e.target.value)}>
            <option>Caja</option><option>Banco</option><option>Billetera</option>
          </select>
        </div>
        <div className="qn-field"><label className="qn-label">Saldo inicial</label><input className="qn-input" type="number" value={initialBalance} onChange={(e) => setInitialBalance(Number(e.target.value))} /></div>
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
        <button className="qn-btn" onClick={onClose}>Cancelar</button>
        <button className="qn-btn qn-btn-primary" disabled={!name.trim()} onClick={() => { onSave({ id: uid(), name: name.trim(), type, initialBalance }); onClose(); }}><Check size={14} /> Crear cuenta</button>
      </div>
    </Modal>
  );
}

function CajaBancos({ db, setDb }) {
  const [showNew, setShowNew] = useState(false);
  const total = totalCash(db);
  const recent = [...db.movements].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 12);
  const iconFor = (type) => type === 'Caja' ? Wallet : type === 'Banco' ? Landmark : Boxes;
  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Caja y bancos</div>
          <div className="qn-page-sub">Saldo disponible total: <strong>{fmtCOP(total)}</strong></div>
        </div>
        <button className="qn-btn qn-btn-primary" onClick={() => setShowNew(true)}><Plus size={15} /> Nueva cuenta</button>
      </div>
      <div className="qn-stats-row" style={{ gridTemplateColumns: `repeat(${Math.min(db.accounts.length, 4) || 1}, 1fr)` }}>
        {db.accounts.map((a) => {
          const Icon = iconFor(a.type);
          return (
            <div className="qn-card qn-stat" key={a.id}>
              <div className="label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}><Icon size={13} /> {a.name} · {a.type}</div>
              <div className="value">{fmtMoney(accountBalance(db, a.id))}</div>
            </div>
          );
        })}
      </div>
      <div className="qn-card qn-section">
        <div className="qn-section-title">Movimientos recientes</div>
        {recent.length === 0 ? (
          <Empty title="Sin movimientos" sub="Los pagos de ventas y cuentas por pagar aparecerán aquí." />
        ) : (
          <table className="qn-table">
            <thead><tr><th>Fecha</th><th>Cuenta</th><th>Concepto</th><th style={{ textAlign: 'right' }}>Monto</th></tr></thead>
            <tbody>
              {recent.map((m) => {
                const acc = db.accounts.find((a) => a.id === m.accountId);
                return (
                  <tr key={m.id}>
                    <td>{fmtDate(m.date)}</td>
                    <td>{acc ? acc.name : '—'}</td>
                    <td>{m.concept}</td>
                    <td style={{ textAlign: 'right', fontWeight: 600, color: m.amount >= 0 ? C.teal : C.rust }}>{m.amount >= 0 ? '+' : ''}{fmtMoney(m.amount)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
      {showNew && <AccountModal onClose={() => setShowNew(false)} onSave={(acc) => setDb((prev) => ({ ...prev, accounts: [...prev.accounts, acc] }))} />}
    </div>
  );
}

/* ============================== MAESTROS ============================== */
function SimpleCrud({ title, items, columns, onAdd, onDelete, addFields }) {
  const [showAdd, setShowAdd] = useState(false);
  const blankForm = () => Object.fromEntries(addFields.map((f) => [f.key, f.default !== undefined ? f.default : '']));
  const [form, setForm] = useState(blankForm);
  const save = () => {
    if (addFields.some((f) => f.required && !String(form[f.key] || '').trim())) return;
    onAdd({ id: uid(), ...form });
    setForm(blankForm());
    setShowAdd(false);
  };
  return (
    <div className="qn-card qn-section">
      <div className="qn-section-title">
        {title}
        <button className="qn-btn qn-btn-sm qn-btn-primary" onClick={() => setShowAdd((s) => !s)}><Plus size={13} /> Agregar</button>
      </div>
      {showAdd && (
        <div style={{ display: 'grid', gridTemplateColumns: `repeat(${addFields.length}, 1fr) auto`, gap: 10, marginBottom: 16, alignItems: 'end' }}>
          {addFields.map((f) => (
            <div key={f.key}>
              <label className="qn-label">{f.label}</label>
              {f.type === 'select' ? (
                <select className="qn-select" value={form[f.key]} onChange={(e) => setForm({ ...form, [f.key]: e.target.value })}>
                  {f.options.map((o) => <option key={o} value={o}>{o}</option>)}
                </select>
              ) : (
                <input className="qn-input" type={f.type || 'text'} value={form[f.key]} onChange={(e) => setForm({ ...form, [f.key]: e.target.value })} />
              )}
            </div>
          ))}
          <button className="qn-btn qn-btn-primary qn-btn-sm" onClick={save}><Check size={13} /></button>
        </div>
      )}
      {items.length === 0 ? (
        <Empty title="Sin registros" sub="Agrega el primero con el botón de arriba." />
      ) : (
        <table className="qn-table">
          <thead><tr>{columns.map((c) => <th key={c.key}>{c.label}</th>)}<th></th></tr></thead>
          <tbody>
            {items.map((it) => (
              <tr key={it.id}>
                {columns.map((c) => <td key={c.key}>{c.render ? c.render(it) : (it[c.key] || '—')}</td>)}
                <td><ConfirmDelete onConfirm={() => onDelete(it.id)} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

function ResetAllModal({ onClose, onConfirm }) {
  const [text, setText] = useState('');
  const ok = text.trim().toUpperCase() === 'ELIMINAR TODO';
  return (
    <Modal title="Eliminar toda la información" onClose={onClose}>
      <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start', marginBottom: 16, padding: 12, background: C.rustFaint, borderRadius: 8 }}>
        <AlertTriangle size={18} color={C.rust} style={{ flexShrink: 0, marginTop: 1 }} />
        <div style={{ fontSize: 13, color: C.ink }}>
          Esto borra <strong>todas</strong> las ventas, cuentas por cobrar y por pagar, inventario, movimientos de caja, clientes, productos y proveedores. <strong>No se puede deshacer.</strong> Los datos de tu empresa (nombre, dirección, contacto) se mantienen.
        </div>
      </div>
      <div className="qn-field">
        <label className="qn-label">Escribe ELIMINAR TODO para confirmar</label>
        <input className="qn-input" value={text} onChange={(e) => setText(e.target.value)} placeholder="ELIMINAR TODO" autoFocus />
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
        <button className="qn-btn" onClick={onClose}>Cancelar</button>
        <button className="qn-btn qn-btn-danger" disabled={!ok} onClick={() => { onConfirm(); onClose(); }}>
          <Trash2 size={14} /> Eliminar todo definitivamente
        </button>
      </div>
    </Modal>
  );
}

function Maestros({ db, setDb }) {
  const [tab, setTab] = useState('clientes');
  const [showReset, setShowReset] = useState(false);

  const resetAllData = () => {
    setDb((prev) => ({
      company: prev.company,
      nextInvoiceNumber: 1,
      nextBillNumber: 1,
      clients: [],
      products: [],
      suppliers: [],
      accounts: [],
      invoices: [],
      payables: [],
      movements: [],
      inventoryMovements: [],
    }));
  };

  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Clientes y productos</div>
          <div className="qn-page-sub">Datos maestros usados en ventas y cuentas por pagar</div>
        </div>
      </div>
      <div className="qn-tabs">
        <button className={`qn-tab ${tab === 'clientes' ? 'active' : ''}`} onClick={() => setTab('clientes')}>Clientes</button>
        <button className={`qn-tab ${tab === 'productos' ? 'active' : ''}`} onClick={() => setTab('productos')}>Productos</button>
        <button className={`qn-tab ${tab === 'proveedores' ? 'active' : ''}`} onClick={() => setTab('proveedores')}>Proveedores</button>
      </div>
      {tab === 'clientes' && (
        <SimpleCrud
          title="Clientes"
          items={db.clients}
          columns={[{ key: 'name', label: 'Nombre' }, { key: 'phone', label: 'Teléfono' }, { key: 'email', label: 'Email' }, { key: 'address', label: 'Dirección' }]}
          addFields={[{ key: 'name', label: 'Nombre', required: true }, { key: 'phone', label: 'Teléfono' }, { key: 'email', label: 'Email' }, { key: 'address', label: 'Dirección' }]}
          onAdd={(c) => setDb((p) => ({ ...p, clients: [...p.clients, { nit: '', ...c }] }))}
          onDelete={(id) => setDb((p) => ({ ...p, clients: p.clients.filter((c) => c.id !== id) }))}
        />
      )}
      {tab === 'productos' && (
        <SimpleCrud
          title="Productos"
          items={db.products}
          columns={[{ key: 'name', label: 'Nombre' }, { key: 'unit', label: 'Unidad' }, { key: 'price', label: 'Precio ref.', render: (p) => fmtMoney(p.price) }]}
          addFields={[{ key: 'name', label: 'Nombre', required: true }, { key: 'unit', label: 'Unidad', default: 'Litros' }, { key: 'price', label: 'Precio', type: 'number', default: 0 }]}
          onAdd={(p) => setDb((prev) => ({ ...prev, products: [...prev.products, { ...p, price: Number(p.price) }] }))}
          onDelete={(id) => setDb((p) => ({ ...p, products: p.products.filter((x) => x.id !== id) }))}
        />
      )}
      {tab === 'proveedores' && (
        <SimpleCrud
          title="Proveedores"
          items={db.suppliers}
          columns={[{ key: 'name', label: 'Nombre' }, { key: 'phone', label: 'Teléfono' }, { key: 'email', label: 'Email' }]}
          addFields={[{ key: 'name', label: 'Nombre', required: true }, { key: 'phone', label: 'Teléfono' }, { key: 'email', label: 'Email' }]}
          onAdd={(s) => setDb((p) => ({ ...p, suppliers: [...p.suppliers, s] }))}
          onDelete={(id) => setDb((p) => ({ ...p, suppliers: p.suppliers.filter((s) => s.id !== id) }))}
        />
      )}

      <div className="qn-card qn-section" style={{ marginTop: 24, borderColor: C.rust }}>
        <div className="qn-section-title" style={{ color: C.rust }}>Zona de peligro</div>
        <div style={{ fontSize: 13, color: C.inkSoft, marginBottom: 14 }}>
          Elimina permanentemente toda la información del sistema (ventas, cuentas por cobrar y por pagar, inventario, caja y bancos, clientes, productos y proveedores). Los datos de tu empresa se conservan.
        </div>
        <button className="qn-btn qn-btn-danger" onClick={() => setShowReset(true)}>
          <Trash2 size={14} /> Eliminar toda la información del sistema
        </button>
      </div>

      {showReset && <ResetAllModal onClose={() => setShowReset(false)} onConfirm={resetAllData} />}
    </div>
  );
}

/* ============================== HISTÓRICOS ============================== */
function downloadSalesByDate(db, desde, hasta) {
  const days = datesBetween(desde, hasta);
  const invoicesInRange = [...db.invoices].filter((i) => i.date >= desde && i.date <= hasta).sort((a, b) => a.date.localeCompare(b.date));
  const byDay = {};
  invoicesInRange.forEach((i) => {
    byDay[i.date] = byDay[i.date] || { total: 0, count: 0 };
    byDay[i.date].total += i.total;
    byDay[i.date].count += 1;
  });
  const summaryRows = [['Fecha', 'Ventas totales', 'N° facturas']];
  let grandTotal = 0, grandCount = 0;
  days.forEach((d) => {
    const info = byDay[d] || { total: 0, count: 0 };
    grandTotal += info.total; grandCount += info.count;
    summaryRows.push([fmtDate(d), info.total, info.count]);
  });
  summaryRows.push(['TOTAL', grandTotal, grandCount]);

  const detailRows = [['Fecha', 'Factura', 'Cliente', 'Total', 'Saldo', 'Estado']];
  invoicesInRange.forEach((inv) => {
    const client = db.clients.find((c) => c.id === inv.clientId);
    detailRows.push([fmtDate(inv.date), inv.number, client ? client.name : '—', inv.total, balanceOf(inv), statusOf(inv)]);
  });

  const wb = XLSX.utils.book_new();
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  wsSummary['!cols'] = [{ wch: 16 }, { wch: 18 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Ventas por fecha');
  const wsDetail = XLSX.utils.aoa_to_sheet(detailRows);
  wsDetail['!cols'] = [{ wch: 14 }, { wch: 12 }, { wch: 28 }, { wch: 16 }, { wch: 16 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsDetail, 'Detalle facturas');
  XLSX.writeFile(wb, `Ventas_${desde}_a_${hasta}.xlsx`);
}

function Historicos({ db }) {
  const firstInvoiceDate = db.invoices.length ? [...db.invoices].sort((a, b) => a.date.localeCompare(b.date))[0].date : todayISO();
  const [desde, setDesde] = useState(firstInvoiceDate);
  const [hasta, setHasta] = useState(todayISO());

  const inRange = useMemo(() => db.invoices.filter((i) => i.date >= desde && i.date <= hasta), [db.invoices, desde, hasta]);
  const totalRango = inRange.reduce((s, i) => s + i.total, 0);

  const dailyRows = useMemo(() => {
    const byDay = {};
    inRange.forEach((i) => { byDay[i.date] = (byDay[i.date] || 0) + i.total; });
    return Object.entries(byDay).sort((a, b) => b[0].localeCompare(a[0])).map(([date, total]) => ({ date, total }));
  }, [inRange]);

  const months = allMonthsInData(db);
  const monthlyRows = months.map((m) => ({ month: m, ...computePYG(db, m) })).reverse();
  const chartData = months.map((m) => {
    const r = computePYG(db, m);
    return { month: `${m.slice(5, 7)}/${m.slice(2, 4)}`, ventas: r.ingresos, utilidad: r.utilidadNeta };
  });

  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Histórico de ventas</div>
          <div className="qn-page-sub">Consulta y descarga tus ventas de cualquier periodo</div>
        </div>
      </div>

      <div className="qn-card qn-section">
        <div style={{ display: 'flex', gap: 14, alignItems: 'end', flexWrap: 'wrap' }}>
          <div className="qn-field" style={{ marginBottom: 0 }}>
            <label className="qn-label">Desde</label>
            <input className="qn-input" type="date" value={desde} onChange={(e) => setDesde(e.target.value)} />
          </div>
          <div className="qn-field" style={{ marginBottom: 0 }}>
            <label className="qn-label">Hasta</label>
            <input className="qn-input" type="date" value={hasta} onChange={(e) => setHasta(e.target.value)} />
          </div>
          <div style={{ flex: 1 }} />
          <div style={{ textAlign: 'right' }}>
            <div className="qn-label" style={{ marginBottom: 2 }}>Total en el rango</div>
            <div className="qn-display" style={{ fontSize: 20, fontWeight: 700 }}>{fmtCOP(totalRango)}</div>
          </div>
          <button className="qn-btn qn-btn-primary" onClick={() => downloadSalesByDate(db, desde, hasta)}>
            <Download size={14} /> Descargar ventas por fecha
          </button>
        </div>
      </div>

      <div className="qn-card qn-section">
        <div className="qn-section-title">Ventas por día en el rango</div>
        {dailyRows.length === 0 ? (
          <Empty title="Sin ventas en este rango" sub="Ajusta las fechas para ver otro periodo." />
        ) : (
          <table className="qn-table">
            <thead><tr><th>Fecha</th><th style={{ textAlign: 'right' }}>Total vendido</th></tr></thead>
            <tbody>
              {dailyRows.map((r) => (
                <tr key={r.date}><td>{fmtDate(r.date)}</td><td style={{ textAlign: 'right', fontWeight: 600 }}>{fmtMoney(r.total)}</td></tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      <div className="qn-card qn-section">
        <div className="qn-section-title">Histórico mensual</div>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={chartData} margin={{ left: -20, right: 10, top: 5 }}>
            <CartesianGrid stroke={C.line} vertical={false} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={false} tickLine={false} tickFormatter={(v) => v >= 1000000 ? `${v / 1000000}M` : v} />
            <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ borderRadius: 8, border: `1px solid ${C.line}`, fontSize: 12 }} />
            <Bar dataKey="ventas" fill={C.tealBright} radius={[3, 3, 0, 0]} barSize={14} />
            <Bar dataKey="utilidad" fill={C.ink} radius={[3, 3, 0, 0]} barSize={14} />
          </BarChart>
        </ResponsiveContainer>
        <div style={{ display: 'flex', gap: 18, fontSize: 12, color: C.inkSoft, marginTop: 4, marginBottom: 14 }}>
          <span><span style={{ display: 'inline-block', width: 9, height: 9, background: C.tealBright, borderRadius: 2, marginRight: 5 }} />Ventas</span>
          <span><span style={{ display: 'inline-block', width: 9, height: 9, background: C.ink, borderRadius: 2, marginRight: 5 }} />Utilidad neta</span>
        </div>
        <table className="qn-table">
          <thead><tr><th>Mes</th><th style={{ textAlign: 'right' }}>Ventas</th><th style={{ textAlign: 'right' }}>Utilidad neta</th><th style={{ textAlign: 'right' }}>Margen</th></tr></thead>
          <tbody>
            {monthlyRows.map((r) => (
              <tr key={r.month}>
                <td>{monthLabel(r.month)}</td>
                <td style={{ textAlign: 'right' }}>{fmtMoney(r.ingresos)}</td>
                <td style={{ textAlign: 'right', fontWeight: 600, color: r.utilidadNeta >= 0 ? C.teal : C.rust }}>{fmtMoney(r.utilidadNeta)}</td>
                <td style={{ textAlign: 'right' }}>{r.margenNeto === null ? '—' : `${r.margenNeto.toFixed(1)}%`}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== BALANCE GENERAL ============================== */
function BalanceGeneral({ db }) {
  const cash = totalCash(db);
  const ar = totalAR(db);
  const ap = totalAP(db);
  const assets = cash + ar;
  const equity = assets - ap;
  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Balance general</div>
          <div className="qn-page-sub">Corte al {fmtDate(todayISO())}</div>
        </div>
      </div>
      <div className="qn-row2" style={{ alignItems: 'start' }}>
        <div className="qn-card qn-section">
          <div className="qn-section-title">Activos</div>
          <div style={{ fontSize: 12, color: C.inkSoft, marginBottom: 6, marginTop: 4 }}>Efectivo y bancos</div>
          {db.accounts.map((a) => (
            <div key={a.id} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '5px 0' }}>
              <span>{a.name}</span><span className="qn-mono">{fmtMoney(accountBalance(db, a.id))}</span>
            </div>
          ))}
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '6px 0', fontWeight: 600, borderTop: `1px solid ${C.line}`, marginTop: 4 }}>
            <span>Subtotal caja y bancos</span><span>{fmtMoney(cash)}</span>
          </div>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '10px 0 5px' }}>
            <span>Cuentas por cobrar</span><span className="qn-mono">{fmtMoney(ar)}</span>
          </div>
          <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 12, paddingTop: 10, borderTop: `1px solid ${C.ink}`, fontWeight: 700 }}>
            <span>Total activos</span><span>{fmtMoney(assets)}</span>
          </div>
        </div>
        <div>
          <div className="qn-card qn-section" style={{ marginBottom: 14 }}>
            <div className="qn-section-title">Pasivos</div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '5px 0' }}>
              <span>Cuentas por pagar</span><span className="qn-mono">{fmtMoney(ap)}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 12, paddingTop: 10, borderTop: `1px solid ${C.ink}`, fontWeight: 700 }}>
              <span>Total pasivos</span><span>{fmtMoney(ap)}</span>
            </div>
          </div>
          <div className="qn-card qn-section">
            <div className="qn-section-title">Patrimonio</div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '5px 0' }}>
              <span>Utilidades acumuladas (activos − pasivos)</span><span className="qn-mono">{fmtMoney(equity)}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 12, paddingTop: 10, borderTop: `1px solid ${C.ink}`, fontWeight: 700 }}>
              <span>Total pasivo + patrimonio</span><span>{fmtMoney(ap + equity)}</span>
            </div>
          </div>
        </div>
      </div>
      <div style={{ fontSize: 11.5, color: C.inkSoft, marginTop: 14 }}>
        Balance simplificado: se calcula automáticamente a partir de tus saldos de caja/bancos, cuentas por cobrar y cuentas por pagar. No sustituye un balance contable formal.
      </div>
    </div>
  );
}

/* ============================== ESTADO DE RESULTADOS ============================== */
function EstadoResultados({ db }) {
  const months = useMemo(() => {
    const set = new Set([todayISO().slice(0, 7)]);
    db.invoices.forEach((i) => set.add(i.date.slice(0, 7)));
    db.payables.forEach((p) => set.add(p.date.slice(0, 7)));
    return Array.from(set).sort().reverse();
  }, [db]);
  const [month, setMonth] = useState(months[0]);
  const r = computePYG(db, month);
  const pct = (v) => v === null ? '—' : `${v.toFixed(1)}%`;
  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Rentabilidad · Estado de resultados (P&amp;G)</div>
          <div className="qn-page-sub">Periodo: {monthLabel(month)}</div>
        </div>
        <select className="qn-select" style={{ width: 200 }} value={month} onChange={(e) => setMonth(e.target.value)}>
          {months.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
        </select>
      </div>

      <div className="qn-stats-row" style={{ gridTemplateColumns: 'repeat(3, 1fr)' }}>
        <div className="qn-card qn-stat">
          <div className="label">Utilidad bruta</div>
          <div className="value" style={{ color: r.utilidadBruta >= 0 ? C.teal : C.rust }}>{fmtMoney(r.utilidadBruta)}</div>
        </div>
        <div className="qn-card qn-stat">
          <div className="label">Utilidad neta</div>
          <div className="value" style={{ color: r.utilidadNeta >= 0 ? C.teal : C.rust }}>{fmtMoney(r.utilidadNeta)}</div>
        </div>
        <div className="qn-card qn-stat">
          <div className="label">Margen neto</div>
          <div className="value">{pct(r.margenNeto)}</div>
        </div>
      </div>

      <div className="qn-card qn-section">
        <div style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 0', fontWeight: 600 }}>
          <span>Ingresos por ventas</span><span className="qn-mono">{fmtMoney(r.ingresos)}</span>
        </div>
        <div style={{ height: 1, background: C.line, margin: '6px 0 14px' }} />
        <div style={{ display: 'flex', justifyContent: 'space-between', padding: '5px 0', fontSize: 13, color: C.inkSoft }}>
          <span>Costo de mercancía vendida (COGS)</span><span className="qn-mono">−{fmtMoney(r.cogs)}</span>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 0', fontWeight: 600, borderTop: `1px solid ${C.line}`, marginTop: 6 }}>
          <span>Utilidad bruta ({pct(r.margenBruto)})</span><span>{fmtMoney(r.utilidadBruta)}</span>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', padding: '12px 0 5px', fontSize: 13, color: C.inkSoft }}>
          <span>Gastos generales</span><span className="qn-mono">−{fmtMoney(r.gastos)}</span>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 14, paddingTop: 14, borderTop: `2px solid ${C.ink}` }}>
          <span className="qn-display" style={{ fontWeight: 700, fontSize: 16 }}>Utilidad neta del periodo</span>
          <span className="qn-display" style={{ fontWeight: 700, fontSize: 18, color: r.utilidadNeta >= 0 ? C.teal : C.rust }}>{fmtMoney(r.utilidadNeta)}</span>
        </div>
      </div>
      <div style={{ fontSize: 11.5, color: C.inkSoft, marginTop: 14 }}>
        El costo de mercancía vendida usa el costo promedio del producto en el momento de cada venta (registrado en Inventario). Los artículos «producto libre» sin costo cargado cuentan como costo cero. Cálculo por causación: se toma la fecha de la factura y de la cuenta por pagar, sin importar si ya se cobró o pagó.
      </div>
    </div>
  );
}

/* ============================== DASHBOARD ============================== */
function Dashboard({ db }) {
  const currentMonth = todayISO().slice(0, 7);
  const monthInvoices = db.invoices.filter((i) => i.date.slice(0, 7) === currentMonth);
  const ventasMes = monthInvoices.reduce((s, i) => s + i.total, 0);
  const pyg = computePYG(db, currentMonth);
  const cobradoMes = db.invoices.reduce((s, i) => s + (i.payments || []).filter((p) => p.date.slice(0, 7) === currentMonth).reduce((a, p) => a + p.amount, 0), 0);

  const daysInMonth = new Date(Number(currentMonth.slice(0, 4)), Number(currentMonth.slice(5, 7)), 0).getDate();
  const dailySeries = Array.from({ length: daysInMonth }, (_, idx) => {
    const day = idx + 1;
    const dayStr = `${currentMonth}-${String(day).padStart(2, '0')}`;
    const total = monthInvoices.filter((i) => i.date === dayStr).reduce((s, i) => s + i.total, 0);
    return { day, total };
  });

  const productTotals = {};
  db.invoices.forEach((inv) => inv.items.forEach((it) => { productTotals[it.description] = (productTotals[it.description] || 0) + it.total; }));
  const topProducts = Object.entries(productTotals).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([name, total]) => ({ name: name.length > 16 ? name.slice(0, 16) + '…' : name, total }));

  const upcomingAR = [...db.invoices].filter((i) => balanceOf(i) > 0).sort((a, b) => a.dueDate.localeCompare(b.dueDate)).slice(0, 5);
  const upcomingAP = [...db.payables].filter((p) => balanceOf(p) > 0).sort((a, b) => a.dueDate.localeCompare(b.dueDate)).slice(0, 5);

  return (
    <div>
      <div className="qn-page-head">
        <div>
          <div className="qn-page-title">Dashboard</div>
          <div className="qn-page-sub">{monthLabel(currentMonth)} · actualizado en tiempo real</div>
        </div>
      </div>

      <div className="qn-stats-row">
        <div className="qn-card qn-stat">
          <div className="label">Ventas del mes</div>
          <div className="value" style={{ color: C.teal }}>{fmtMoney(ventasMes)}</div>
        </div>
        <div className="qn-card qn-stat">
          <div className="label">Cobrado en el mes</div>
          <div className="value">{fmtMoney(cobradoMes)}</div>
        </div>
        <div className="qn-card qn-stat">
          <div className="label">Por cobrar (total)</div>
          <div className="value" style={{ color: C.rust }}>{fmtMoney(totalAR(db))}</div>
        </div>
        <div className="qn-card qn-stat">
          <div className="label">Saldo en caja y bancos</div>
          <div className="value">{fmtMoney(totalCash(db))}</div>
        </div>
      </div>

      <div className="qn-stats-row" style={{ gridTemplateColumns: 'repeat(3, 1fr)' }}>
        <div className="qn-card qn-stat">
          <div className="label">Utilidad neta del mes</div>
          <div className="value" style={{ color: pyg.utilidadNeta >= 0 ? C.teal : C.rust }}>{fmtMoney(pyg.utilidadNeta)}</div>
        </div>
        <div className="qn-card qn-stat">
          <div className="label">Margen neto del mes</div>
          <div className="value">{pyg.margenNeto === null ? '—' : `${pyg.margenNeto.toFixed(1)}%`}</div>
        </div>
        <div className="qn-card qn-stat">
          <div className="label">Valor en inventario</div>
          <div className="value">{fmtMoney(inventoryValue(db))}</div>
        </div>
      </div>

      <div className="qn-row2" style={{ alignItems: 'start', marginBottom: 4 }}>
        <div className="qn-card qn-section">
          <div className="qn-section-title">Ventas por día · {monthLabel(currentMonth)}</div>
          <ResponsiveContainer width="100%" height={220}>
            <AreaChart data={dailySeries} margin={{ left: -20, right: 10, top: 5 }}>
              <defs>
                <linearGradient id="qnArea" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={C.tealBright} stopOpacity={0.35} />
                  <stop offset="100%" stopColor={C.tealBright} stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={C.line} vertical={false} />
              <XAxis dataKey="day" tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={false} tickLine={false} tickFormatter={(v) => v >= 1000000 ? `${v / 1000000}M` : v} />
              <Tooltip formatter={(v) => fmtMoney(v)} labelFormatter={(d) => `Día ${d}`} contentStyle={{ borderRadius: 8, border: `1px solid ${C.line}`, fontSize: 12 }} />
              <Area type="monotone" dataKey="total" stroke={C.tealBright} strokeWidth={2} fill="url(#qnArea)" />
            </AreaChart>
          </ResponsiveContainer>
        </div>
        <div className="qn-card qn-section">
          <div className="qn-section-title">Top productos (histórico)</div>
          {topProducts.length === 0 ? <Empty title="Sin datos aún" sub="Aparecerá cuando registres ventas." /> : (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={topProducts} layout="vertical" margin={{ left: 10, right: 20, top: 5 }}>
                <CartesianGrid stroke={C.line} horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={false} tickLine={false} tickFormatter={(v) => v >= 1000000 ? `${v / 1000000}M` : v} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 11.5, fill: C.ink }} axisLine={false} tickLine={false} width={110} />
                <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ borderRadius: 8, border: `1px solid ${C.line}`, fontSize: 12 }} />
                <Bar dataKey="total" fill={C.teal} radius={[0, 4, 4, 0]} barSize={16} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>

      <div className="qn-row2" style={{ alignItems: 'start' }}>
        <div className="qn-card qn-section">
          <div className="qn-section-title">Próximas a vencer · por cobrar</div>
          {upcomingAR.length === 0 ? <Empty title="Nada pendiente" sub="No hay facturas con saldo por cobrar." /> : (
            <table className="qn-table">
              <tbody>
                {upcomingAR.map((inv) => {
                  const client = db.clients.find((c) => c.id === inv.clientId);
                  return (
                    <tr key={inv.id}>
                      <td className="qn-mono" style={{ width: 80 }}>{inv.number}</td>
                      <td>{client ? client.name : '—'}</td>
                      <td>{fmtDate(inv.dueDate)}</td>
                      <td style={{ textAlign: 'right', fontWeight: 600 }}>{fmtMoney(balanceOf(inv))}</td>
                      <td><Badge status={statusOf(inv)} /></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>
        <div className="qn-card qn-section">
          <div className="qn-section-title">Próximas a vencer · por pagar</div>
          {upcomingAP.length === 0 ? <Empty title="Nada pendiente" sub="No hay cuentas por pagar activas." /> : (
            <table className="qn-table">
              <tbody>
                {upcomingAP.map((p) => {
                  const s = db.suppliers.find((x) => x.id === p.supplierId);
                  return (
                    <tr key={p.id}>
                      <td className="qn-mono" style={{ width: 80 }}>{p.number}</td>
                      <td>{p.type === 'Proveedor' ? (s ? s.name : '—') : p.concept}</td>
                      <td>{fmtDate(p.dueDate)}</td>
                      <td style={{ textAlign: 'right', fontWeight: 600 }}>{fmtMoney(balanceOf(p))}</td>
                      <td><Badge status={statusOf(p)} /></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================== APP (Supabase + auth) ============================== */
function LoadingScreen({ text }) {
  return (
    <div className="qn-root" style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <GlobalStyle />
      <div style={{ color: C.inkSoft, fontSize: 13 }}>{text || 'Cargando Catalyst…'}</div>
    </div>
  );
}

function LoginScreen() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setLoading(true);
    setError('');
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (error) setError('Correo o contraseña incorrectos.');
  };

  return (
    <div className="qn-root" style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <GlobalStyle />
      <form onSubmit={submit} className="qn-card qn-section" style={{ width: 340 }}>
        <img src={LOGO_SRC} alt="logo" style={{ width: 40, marginBottom: 14 }} />
        <div className="qn-page-title" style={{ marginBottom: 16 }}>Entrar a Catalyst</div>
        <div className="qn-field">
          <label className="qn-label">Correo</label>
          <input className="qn-input" type="email" value={email} onChange={(e) => setEmail(e.target.value)} required />
        </div>
        <div className="qn-field">
          <label className="qn-label">Contraseña</label>
          <input className="qn-input" type="password" value={password} onChange={(e) => setPassword(e.target.value)} required />
        </div>
        {error && <div style={{ color: C.rust, fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
        <button className="qn-btn qn-btn-primary" style={{ width: '100%', justifyContent: 'center' }} disabled={loading}>
          {loading ? 'Entrando…' : 'Entrar'}
        </button>
      </form>
    </div>
  );
}

function App({ userId }) {
  const [db, setDb] = useState(null);
  const [loadError, setLoadError] = useState(null);
  const [reloadKey, setReloadKey] = useState(0);
  const [tab, setTab] = useState('dashboard');
  const saveTimer = useRef(null);
  const firstLoad = useRef(true);

  useEffect(() => {
    let cancelled = false;
    async function load(attempt) {
      const { data, error } = await supabase.from('app_data').select('data').eq('user_id', userId).maybeSingle();
      if (cancelled) return;
      if (error) {
        if (attempt < 2) { setTimeout(() => load(attempt + 1), 600); return; }
        setLoadError('No se pudo conectar con tu base de datos: ' + error.message + '. Tu información NO se ha borrado — intenta de nuevo.');
        return;
      }
      if (data && data.data) {
        setDb(data.data);
        return;
      }
      const seed = seedData();
      const { error: upsertError } = await supabase.from('app_data').upsert({ user_id: userId, data: seed, updated_at: new Date().toISOString() });
      if (cancelled) return;
      if (upsertError) {
        setLoadError('No se pudo inicializar tu información: ' + upsertError.message);
        return;
      }
      setDb(seed);
    }
    setLoadError(null);
    load(0);
    return () => { cancelled = true; };
  }, [userId, reloadKey]);

  useEffect(() => {
    if (!db) return;
    if (firstLoad.current) { firstLoad.current = false; return; }
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      supabase.from('app_data').upsert({ user_id: userId, data: db, updated_at: new Date().toISOString() }).then(({ error }) => {
        if (error) console.error('Error guardando en Supabase:', error);
      });
    }, 500);
    return () => clearTimeout(saveTimer.current);
  }, [db, userId]);

  if (loadError) {
    return (
      <div className="qn-root" style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <GlobalStyle />
        <div style={{ textAlign: 'center', maxWidth: 380, padding: 20 }}>
          <div style={{ fontWeight: 700, marginBottom: 8, fontFamily: "'Space Grotesk', sans-serif" }}>No se pudo cargar tu información</div>
          <div style={{ fontSize: 13, color: C.inkSoft, marginBottom: 18 }}>{loadError}</div>
          <button className="qn-btn qn-btn-primary" onClick={() => setReloadKey((k) => k + 1)}>Reintentar</button>
        </div>
      </div>
    );
  }

  if (!db) return <LoadingScreen />;

  return (
    <div className="qn-root">
      <GlobalStyle />
      <div className="qn-shell">
        <Sidebar tab={tab} setTab={setTab} company={db.company} />
        <div className="qn-main">
          {tab === 'dashboard' && <Dashboard db={db} />}
          {tab === 'ventas' && <Ventas db={db} setDb={setDb} />}
          {tab === 'cxc' && <CuentasCobrar db={db} setDb={setDb} />}
          {tab === 'cartera' && <Cartera db={db} />}
          {tab === 'cxp' && <CuentasPagar db={db} setDb={setDb} />}
          {tab === 'inventario' && <Inventario db={db} setDb={setDb} />}
          {tab === 'caja' && <CajaBancos db={db} setDb={setDb} />}
          {tab === 'historicos' && <Historicos db={db} />}
          {tab === 'balance' && <BalanceGeneral db={db} />}
          {tab === 'pyg' && <EstadoResultados db={db} />}
          {tab === 'maestros' && <Maestros db={db} setDb={setDb} />}
        </div>
      </div>
    </div>
  );
}

export default function Root() {
  const [session, setSession] = useState(undefined);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, sess) => setSession(sess));
    return () => sub.subscription.unsubscribe();
  }, []);

  if (session === undefined) return <LoadingScreen text="Verificando sesión…" />;
  if (!session) return <LoginScreen />;
  return <App userId={session.user.id} />;
}
