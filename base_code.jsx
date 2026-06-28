import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";


const STORAGE_KEY = "portafolio-v1";

// ── Diseño: tokens ───────────────────────────────────────────────────────────
const C = {
  bg: "#E8ECEF",
  card: "#FFFFFF",
  line: "#D3DAE1",
  ink: "#0F1B26",
  muted: "#5D6C7B",
  blue: "#1F5EFF",
  blueSoft: "#EBF1FF",
  green: "#0B8A4B",
  greenSoft: "#E8F6EE",
  red: "#C7392F",
  redSoft: "#FBEDEB",
};
const MONO = "'SF Mono', 'Cascadia Code', Consolas, 'Roboto Mono', monospace";

// ── Helpers ──────────────────────────────────────────────────────────────────
const num = (v) => {
  if (typeof v === "number") return v;
  let s = String(v ?? "").trim().replace(/[$\s"]/g, "");
  if (s.startsWith("(") && s.endsWith(")")) s = "-" + s.slice(1, -1); // (123) → -123
  if (s.includes(",") && s.includes(".")) s = s.replace(/,/g, ""); // 1,234.56
  else if (/^-?\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, ""); // 1,234
  else if (s.includes(",")) s = s.replace(",", "."); // 12,34 → 12.34
  const n = parseFloat(s);
  return isFinite(n) ? n : NaN;
};
const fmt = (v, d = 2) =>
  !isFinite(v)
    ? "—"
    : v.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtSigned = (v, d = 2) => (!isFinite(v) ? "—" : (v > 0 ? "+" : "") + fmt(v, d));

// ── Lectura del archivo ──────────────────────────────────────────────────────
// Soporta tres formatos:
// 1) IBKR Transaction History (CSV en español con sección "Transaction History")
// 2) IBKR Open Positions (Activity Statement clásico)
// 3) Hoja propia del usuario (columnas Ticker / Acciones / Precio Promedio)
function parseWorkbook(wb) {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null, raw: false });
    const txn = parseIBKRTransactions(rows);
    if (txn.length) return txn;
    const ibkr = parseIBKROpenPositions(rows);
    if (ibkr.length) return ibkr;
  }
  let best = [];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null });
    const own = parseSheetRows(rows);
    const gen = parseGenericPositions(rows);
    const found = own.length >= gen.length ? own : gen;
    if (found.length > best.length) best = found;
  }
  return best;
}

// ── Parser 1: IBKR Transaction History ──────────────────────────────────────
// Columnas: Date | Account | Description | Transaction Type | Symbol |
//           Quantity | Price | Price Currency | Gross Amount | Commission | Net Amount
// Parsea fecha en formato M/D/YYYY o YYYY-MM-DD → número comparable
function parseDate(s) {
  if (!s) return 0;
  s = s.trim();
  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return new Date(s).getTime();
  // M/D/YYYY
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(`${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`).getTime();
  return 0;
}

function parseIBKRTransactions(rows) {
  const low = (c) => String(c ?? "").trim().toLowerCase();
  let cols = null;
  const txns = []; // recopilar todas antes de procesar

  for (const row of rows) {
    if (!row || low(row[0]) !== "transaction history") continue;
    if (low(row[1]) === "header") { cols = row.map((c) => low(c)); continue; }
    if (low(row[1]) !== "data" || !cols) continue;
    const get = (name) => { const i = cols.indexOf(name); return i > -1 ? String(row[i] ?? "").trim() : ""; };
    const type = get("transaction type");
    if (type !== "Buy" && type !== "Sell") continue;
    const sym = get("symbol").toUpperCase();
    if (!sym || sym === "-") continue;
    if (get("price currency") !== "USD") continue;
    txns.push({
      date:  parseDate(get("date")),
      type,  sym,
      qty:   num(get("quantity")),
      gross: num(get("gross amount")),
      comm:  num(get("commission")),
      desc:  get("description"),
    });
  }

  // Ordenar de más antiguo a más reciente para calcular correctamente el cost basis
  txns.sort((a, b) => a.date - b.date);

  const map = new Map();
  for (const tx of txns) {
    if (!map.has(tx.sym)) map.set(tx.sym, {
      sh: 0, cost: 0,   // weighted avg (para Promedio y calculadoras)
      lots: [],          // FIFO (para Invertido Actual)
      hist: 0, proceeds: 0, totalBought: 0,
      n: tx.desc,
    });
    const pos = map.get(tx.sym);

    if (tx.type === "Buy") {
      const paid = Math.abs(tx.gross) + Math.abs(isFinite(tx.comm) ? tx.comm : 0);
      const bqty = isFinite(tx.qty) ? tx.qty : 0;
      pos.sh          += bqty;
      pos.cost        += paid;
      pos.lots.push({ qty: bqty, cost: paid }); // FIFO lot
      pos.hist        += paid;
      pos.totalBought += bqty;
    } else {
      const soldQty  = Math.abs(isFinite(tx.qty)   ? tx.qty   : 0);
      const received = Math.abs(isFinite(tx.gross)  ? tx.gross : 0)
                     - Math.abs(isFinite(tx.comm)   ? tx.comm  : 0);
      // Weighted avg
      const avgCost = pos.sh > 0 ? pos.cost / pos.sh : 0;
      pos.sh       -= soldQty;
      pos.cost     -= avgCost * soldQty;
      // FIFO: descuenta desde el lote más antiguo
      let toSell = soldQty;
      while (toSell > 0.00001 && pos.lots.length > 0) {
        const lot = pos.lots[0];
        if (lot.qty <= toSell + 0.00001) { toSell -= lot.qty; pos.lots.shift(); }
        else { const f = toSell / lot.qty; lot.cost -= lot.cost * f; lot.qty -= toSell; toSell = 0; }
      }
      pos.proceeds += received;
    }
    if (!pos.n) pos.n = tx.desc;
  }

  const result = [];
  for (const [t, p] of map) {
    if (!(p.sh > 0.0001)) continue;
    const avg     = p.cost / p.sh;
    const invFifo = p.lots.reduce((a, l) => a + l.cost, 0);
    const netCost = p.hist - p.proceeds;
    const prom    = p.totalBought > 0 ? p.hist / p.totalBought : avg;
    result.push({
      t, n: p.n,
      sh:          Math.round(p.sh          * 10000) / 10000,
      avg:         Math.round(avg           * 10000) / 10000,  // weighted avg (calculadoras)
      prom:        Math.round(prom          * 10000) / 10000,  // hist / totalBought
      inv:         Math.round(invFifo       * 100)   / 100,    // Invertido Actual (FIFO)
      hist:        Math.round(p.hist        * 100)   / 100,    // Invertido Histórico
      proceeds:    Math.round(p.proceeds    * 100)   / 100,
      totalBought: Math.round(p.totalBought * 10000) / 10000,
      cef:         Math.round((netCost / p.sh) * 10000) / 10000,
    });
  }
  return result.sort((a, b) => a.t.localeCompare(b.t));
}

// ── Parser 2: IBKR Activity Statement ("Open Positions") ────────────────────
function parseIBKROpenPositions(rows) {
  const low = (c) => String(c ?? "").trim().toLowerCase();
  const names = {};
  let fiHeader = null;
  for (const r of rows) {
    if (!r || low(r[0]) !== "financial instrument information") continue;
    if (low(r[1]) === "header") { fiHeader = r.map(low); continue; }
    if (low(r[1]) === "data" && fiHeader) {
      const si = fiHeader.indexOf("symbol"), di = fiHeader.indexOf("description");
      if (si > -1 && di > -1 && r[si]) names[String(r[si]).trim().toUpperCase()] = String(r[di] ?? "").trim();
    }
  }
  const isPosSection = (s) => s === "open positions" || s === "positions";
  let header = null;
  const raw = [];
  for (const r of rows) {
    if (!r || !isPosSection(low(r[0]))) continue;
    if (low(r[1]) === "header") { header = r.map(low); continue; }
    if (low(r[1]) !== "data" || !header) continue;
    const col = (n) => header.indexOf(n);
    const dd = col("datadiscriminator");
    if (dd > -1 && r[dd] && low(r[dd]) !== "summary") continue;
    const cat = col("asset category");
    if (cat > -1 && r[cat] && !low(r[cat]).includes("stock")) continue;
    const t = String(r[col("symbol")] ?? "").trim().toUpperCase();
    const sh = num(r[col("quantity")]);
    let avg = col("cost price") > -1 ? num(r[col("cost price")]) : NaN;
    let inv = col("cost basis") > -1 ? num(r[col("cost basis")]) : NaN;
    if (!t || !(sh > 0)) continue;
    if (!isFinite(avg) && isFinite(inv)) avg = inv / sh;
    if (!isFinite(inv) && isFinite(avg)) inv = sh * avg;
    if (!(avg > 0)) continue;
    raw.push({ t, n: names[t] || "", sh, avg, inv, hist: inv });
  }
  const map = new Map();
  for (const p of raw) {
    const ex = map.get(p.t);
    if (ex) { const sh = ex.sh + p.sh, inv = ex.inv + p.inv; map.set(p.t, { ...ex, sh, inv, avg: inv / sh, hist: inv }); }
    else map.set(p.t, p);
  }
  return [...map.values()].sort((a, b) => a.t.localeCompare(b.t));
}

// ── Parser 3: tabla genérica ─────────────────────────────────────────────────
function parseGenericPositions(rows) {
  const low = (c) => String(c ?? "").trim().toLowerCase();
  let hIdx = -1, cols = {};
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const r = rows[i] || [];
    const c = { t: -1, n: -1, sh: -1, avg: -1, inv: -1 };
    r.forEach((cell, j) => {
      const s = low(cell);
      if (!s) return;
      if (c.t === -1 && (s === "symbol" || s.includes("instrument") || s === "ticker")) c.t = j;
      else if (c.n === -1 && (s.includes("description") || s.includes("nombre") || s === "name")) c.n = j;
      else if (c.sh === -1 && (s === "position" || s === "quantity" || s.includes("accion") || s === "shares" || s === "qty")) c.sh = j;
      else if (c.avg === -1 && (s.includes("average price") || s.includes("avg price") || s.includes("cost price") || s.includes("avg cost") || s.includes("promedio"))) c.avg = j;
      else if (c.inv === -1 && s.includes("cost basis")) c.inv = j;
    });
    if (c.t > -1 && c.sh > -1 && c.avg > -1) { hIdx = i; cols = c; break; }
  }
  if (hIdx === -1) return [];
  const out = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const t = String(r[cols.t] ?? "").trim();
    if (!t) continue;
    if (low(t) === "total" || t.startsWith("💡")) break;
    const sh = num(r[cols.sh]), avg = num(r[cols.avg]);
    if (!(sh > 0) || !(avg > 0)) continue;
    const inv = cols.inv > -1 && isFinite(num(r[cols.inv])) ? num(r[cols.inv]) : sh * avg;
    out.push({ t: t.toUpperCase(), n: cols.n > -1 ? String(r[cols.n] ?? "").trim() : "", sh, avg, inv, hist: inv });
  }
  return out;
}

function parseSheetRows(rows) {
  const low = (c) => String(c ?? "").trim().toLowerCase();
  // 1) localizar la fila de encabezados
  let hIdx = -1, cols = {};
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const r = rows[i] || [];
    const ti = r.findIndex((c) => low(c) === "ticker");
    if (ti === -1) continue;
    cols = { t: ti, n: -1, sh: -1, avg: -1, inv: -1, hist: -1 };
    r.forEach((c, j) => {
      const s = low(c);
      if (s.includes("nombre") || s.includes("name")) cols.n = j;
      else if (s.includes("accion")) cols.sh = j;
      else if (s.includes("promedio") || s.includes("avg")) cols.avg = j;
      else if (s.includes("hist")) cols.hist = j;
      else if (s.includes("invertido") || s.includes("invested")) cols.inv = j;
    });
    if (cols.sh !== -1 && cols.avg !== -1) { hIdx = i; break; }
  }
  if (hIdx === -1) return [];

  // 2) leer filas de datos hasta TOTAL o fila vacía
  const out = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const t = String(r[cols.t] ?? "").trim();
    if (!t) break;
    if (low(t) === "total" || t.startsWith("💡")) break;
    const sh = num(r[cols.sh]);
    const avg = num(r[cols.avg]);
    if (!(sh > 0) || !(avg > 0)) continue;
    const inv = cols.inv !== -1 && isFinite(num(r[cols.inv])) ? num(r[cols.inv]) : sh * avg;
    const hist = cols.hist !== -1 && isFinite(num(r[cols.hist])) ? num(r[cols.hist]) : inv;
    out.push({
      t: t.toUpperCase(),
      n: cols.n !== -1 ? String(r[cols.n] ?? "").trim() : "",
      sh, avg, inv, hist,
    });
  }
  return out;
}

// Merge de dos cuentas: suma shares, inv y hist de tickers duplicados.
// Tickers nuevos se agregan. El avg resultante es el promedio ponderado real.
function mergePortfolios(current, incoming) {
  const map = new Map(current.map((p) => [p.t, { ...p }]));
  for (const np of incoming) {
    const ex = map.get(np.t);
    if (ex) {
      const sh          = ex.sh          + np.sh;
      const inv         = ex.inv         + np.inv;
      const hist        = ex.hist        + np.hist;
      const proceeds    = (ex.proceeds    ?? 0) + (np.proceeds    ?? 0);
      const totalBought = (ex.totalBought ?? 0) + (np.totalBought ?? 0);
      const netCost     = hist - proceeds;
      const prom        = totalBought > 0 ? hist / totalBought : (sh > 0 ? inv / sh : 0);
      map.set(np.t, {
        t: np.t, n: np.n || ex.n,
        sh, inv, hist, proceeds, totalBought, prom,
        avg:  sh > 0 ? inv / sh : 0,
        cef:  proceeds > 0 ? netCost / sh : (sh > 0 ? inv / sh : 0),
      });
    } else {
      map.set(np.t, { ...np });
    }
  }
  return [...map.values()].sort((a, b) => a.t.localeCompare(b.t));
}

// ── Componentes base ─────────────────────────────────────────────────────────
function Label({ children }) {
  return (
    <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: C.muted, marginBottom: 4 }}>
      {children}
    </div>
  );
}

function Input({ label, value, onChange, unit, placeholder }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <Label>{label}</Label>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <input
          type="text"
          inputMode="decimal"
          value={value}
          placeholder={placeholder || "0.00"}
          onChange={(e) => onChange(e.target.value)}
          style={{
            flex: 1, minWidth: 0, fontFamily: MONO, fontSize: 18, fontWeight: 600,
            color: C.ink, background: C.blueSoft, border: `2px solid ${C.blue}`,
            borderRadius: 8, padding: "10px 12px", outline: "none",
          }}
        />
        {unit && <span style={{ fontSize: 12, color: C.muted, fontWeight: 600, width: 58 }}>{unit}</span>}
      </div>
    </div>
  );
}

function Row({ label, value, unit, strong, color }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 12, padding: "9px 0", borderBottom: `1px dashed ${C.line}` }}>
      <span style={{ fontSize: 13, color: C.muted }}>{label}</span>
      <span style={{ fontFamily: MONO, fontSize: strong ? 20 : 15, fontWeight: strong ? 700 : 600, color: color || C.ink, whiteSpace: "nowrap" }}>
        {value} {unit && <span style={{ fontSize: 11, color: C.muted, fontWeight: 500 }}>{unit}</span>}
      </span>
    </div>
  );
}

function Card({ title, children, right }) {
  return (
    <div style={{ background: C.card, border: `1px solid ${C.line}`, borderRadius: 14, padding: 18, marginBottom: 16 }}>
      {title && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", paddingBottom: 10, marginBottom: 6, borderBottom: `2px solid ${C.ink}` }}>
          <span style={{ fontSize: 12, fontWeight: 800, letterSpacing: "0.1em", textTransform: "uppercase", color: C.ink }}>{title}</span>
          {right}
        </div>
      )}
      {children}
    </div>
  );
}

function Btn({ children, onClick, kind = "primary", small }) {
  const styles = {
    primary: { background: C.ink, color: "#fff", border: `1px solid ${C.ink}` },
    blue: { background: C.blue, color: "#fff", border: `1px solid ${C.blue}` },
    green: { background: C.green, color: "#fff", border: `1px solid ${C.green}` },
    ghost: { background: "transparent", color: C.muted, border: `1px solid ${C.line}` },
  }[kind];
  return (
    <button
      onClick={onClick}
      style={{
        ...styles, borderRadius: 8, cursor: "pointer", fontWeight: 700,
        fontSize: small ? 12 : 14, padding: small ? "7px 12px" : "12px 16px", width: small ? "auto" : "100%",
      }}
    >
      {children}
    </button>
  );
}

// ── Selector de posición ─────────────────────────────────────────────────────
function PositionPicker({ portfolio, ticker, setTicker, shares, setShares, avg, setAvg, setCef }) {
  const onPick = (t) => {
    setTicker(t);
    const p = portfolio.find((x) => x.t === t);
    if (p) {
      setShares(String(p.sh));
      setAvg(String(p.avg));
      setCef(p.cef ?? null); // cef del portafolio, null si no tiene ventas previas
    } else {
      setShares(""); setAvg(""); setCef(null);
    }
  };
  return (
    <Card title="Tu posición actual">
      <div style={{ marginBottom: 14 }}>
        <Label>Ticker</Label>
        <select
          value={ticker}
          onChange={(e) => onPick(e.target.value)}
          style={{
            width: "100%", fontFamily: MONO, fontSize: 16, fontWeight: 700, color: C.ink,
            background: C.blueSoft, border: `2px solid ${C.blue}`, borderRadius: 8,
            padding: "10px 12px", outline: "none", appearance: "none",
          }}
        >
          <option value="">— selecciona —</option>
          {portfolio.map((p) => (
            <option key={p.t} value={p.t}>{p.t}{p.n ? ` · ${p.n}` : ""}</option>
          ))}
          <option value="OTRO">OTRO (manual)</option>
        </select>
      </div>
      <Input label="Acciones que tienes" value={shares} onChange={setShares} unit="acciones" />
      <Input label="Precio promedio de compra" value={avg} onChange={setAvg} unit="USD" />
      <Row label="Total invertido" value={fmt(num(shares) * num(avg))} unit="USD" />
    </Card>
  );
}

// ── Calculadora de promedio ──────────────────────────────────────────────────
function Promedio({ pos }) {
  const [price, setPrice] = useState("");
  const [amount, setAmount] = useState("");
  const sh = num(pos.shares), avg = num(pos.avg), p = num(price), m = num(amount);
  const currentCef = (pos.cef != null && isFinite(pos.cef)) ? pos.cef : avg;
  const hasCef = pos.cef != null && isFinite(pos.cef) && Math.abs(pos.cef - avg) > 0.001;

  const invested   = sh * avg;
  const newShares  = m / p;
  const totShares  = sh + newShares;
  const totCap     = invested + m;
  const newAvg     = totCap / totShares;
  const delta      = newAvg - avg;
  const pct        = (delta / avg) * 100;
  const ready      = isFinite(newAvg) && p > 0 && m > 0 && sh > 0 && avg > 0;
  const lower      = ready && delta < 0;

  // Nuevo CEF tras la recompra:
  // netCost actual = cef × sh (lo que tienes "realmente" invertido neto)
  // + monto de la nueva compra
  const netCostActual = currentCef * sh;
  const newNetCost    = netCostActual + m;
  const newCef        = newNetCost / totShares;
  const cefDelta      = newCef - currentCef;
  const cefPct        = (cefDelta / currentCef) * 100;
  const cefLower      = ready && cefDelta < 0;

  return (
    <>
      <Card title="Nueva compra">
        <Input label="Precio actual de la acción" value={price} onChange={setPrice} unit="USD" />
        <Input label="Monto que quieres invertir" value={amount} onChange={setAmount} unit="USD" />
        <Row label="Acciones que comprarías" value={fmt(newShares, 4)} unit="acciones" />
      </Card>

      <Card title="Después de la compra">
        <Row label="Total acciones acumuladas" value={fmt(totShares, 4)} unit="acciones" />
        <Row label="Capital total invertido" value={fmt(totCap)} unit="USD" />
        <Row label="Nuevo precio promedio" value={fmt(newAvg, 4)} unit="USD" strong color={ready ? (lower ? C.green : C.red) : C.ink} />
        {ready && hasCef && (
          <Row
            label="Nuevo costo efectivo promedio"
            value={fmt(newCef, 4)}
            unit="USD"
            strong
            color={cefLower ? C.green : C.red}
          />
        )}
        {ready && (
          <div style={{ marginTop: 14, padding: 14, borderRadius: 10, background: lower ? C.greenSoft : C.redSoft, border: `1px solid ${lower ? C.green : C.red}` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color: lower ? C.green : C.red, marginBottom: 6 }}>
              {lower ? "🟢 Compraste más barato — bajas tu promedio" : "🔴 Compraste más caro — subes tu promedio"}
            </div>
            <div style={{ fontFamily: MONO, fontSize: 15, color: C.ink }}>
              Promedio: {fmtSigned(delta, 4)} USD/acción · {fmtSigned(pct, 2)}%
            </div>
            {hasCef && (
              <div style={{ fontFamily: MONO, fontSize: 15, color: C.ink, marginTop: 4 }}>
                Costo ef.: {fmtSigned(cefDelta, 4)} USD/acción · {fmtSigned(cefPct, 2)}%
              </div>
            )}
          </div>
        )}
      </Card>
    </>
  );
}

// ── Calculadora de venta ─────────────────────────────────────────────────────
function Venta({ pos }) {
  const [current, setCurrent] = useState("");
  const [target, setTarget] = useState("");
  const sh = num(pos.shares), avg = num(pos.avg), cur = num(current), tgt = num(target);
  const invested = sh * avg;
  const gainVsCurrent = ((tgt - cur) / cur) * 100;
  const gainVsAvg = ((tgt - avg) / avg) * 100;
  const posValue = sh * tgt;
  const totalGain = posValue - invested;
  const ready = isFinite(totalGain) && sh > 0 && avg > 0 && tgt > 0;

  return (
    <>
      <Card title="Precio objetivo de venta">
        <Input label="Precio actual de la acción" value={current} onChange={setCurrent} unit="USD" />
        <Input label="Precio al que deseas vender" value={target} onChange={setTarget} unit="USD" />
      </Card>

      <Card title="Resultados esperados">
        <Row label="📈 Ganancia vs precio actual" value={fmtSigned(gainVsCurrent)} unit="%" color={gainVsCurrent >= 0 ? C.green : C.red} />
        <Row label="💵 Ganancia sobre tu inversión" value={fmtSigned(gainVsAvg)} unit="%" color={gainVsAvg >= 0 ? C.green : C.red} />
        <Row label="🏦 Valor esperado de tu posición" value={fmt(posValue)} unit="USD" />
        <Row label="💰 Ganancia total esperada" value={fmtSigned(totalGain)} unit="USD" strong color={ready ? (totalGain >= 0 ? C.green : C.red) : C.ink} />
      </Card>

      {sh > 0 && avg > 0 && (
        <MetaGanancia sh={sh} avg={avg} invested={invested} />
      )}

    </>
  );
}

function MetaGanancia({ sh, avg, invested }) {
  const [meta, setMeta] = useState("");
  const pct  = num(meta);
  const sell  = avg * (1 + pct / 100);
  const gain  = invested * (pct / 100);
  const toSell = gain / sell;
  const basis  = (sh - toSell) * avg;
  const ready  = isFinite(pct) && pct > 0 && sh > 0 && avg > 0;

  return (
    <Card title="🎯 Meta de ganancia objetivo">
      <div style={{ marginBottom: 14 }}>
        <Label>Meta en %</Label>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <input
            type="text"
            inputMode="decimal"
            value={meta}
            placeholder="ej. 30"
            onChange={(e) => setMeta(e.target.value)}
            style={{
              flex: 1, minWidth: 0, fontFamily: MONO, fontSize: 18, fontWeight: 600,
              color: C.ink, background: C.blueSoft, border: `2px solid ${C.blue}`,
              borderRadius: 8, padding: "10px 12px", outline: "none",
            }}
          />
          <span style={{ fontSize: 12, color: C.muted, fontWeight: 600, width: 28 }}>%</span>
        </div>
      </div>

      {ready && (
        <>
          <Row label="Precio de venta objetivo"   value={fmt(sell, 4)}   unit="USD" strong color={C.blue} />
          <Row label="Ganancia en USD"             value={fmt(gain)}      unit="USD" color={C.green} />
          <Row label="Acciones a vender"           value={fmt(toSell, 4)} unit="acciones" />
          <Row label="Nuevo cost basis"            value={fmt(basis)}     unit="USD" />
          <div style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>
            💡 Acciones a vender = las necesarias para retirar solo la ganancia, conservando el resto de tu posición.
          </div>
        </>
      )}
    </Card>
  );
}

function CicloVentaRecompra({ pos }) {
  const [pct, setPct]         = useState("");
  const [sellPrice, setSell]  = useState("");
  const [buyPrice, setBuy]    = useState("");

  const sh   = num(pos.shares);
  const avg  = num(pos.avg);
  // Promedio histórico: hist / totalBought (todas las compras alguna vez)
  const baseProm = (pos.prom != null && isFinite(pos.prom)) ? pos.prom : avg;
  // CEF histórico: ya incorpora ventas previas
  const baseCef  = (pos.cef  != null && isFinite(pos.cef))  ? pos.cef  : avg;
  // Para recalcular prom: necesitamos hist y totalBought originales
  const histOrig  = (pos.hist        != null && isFinite(pos.hist))        ? pos.hist        : sh * avg;
  const tbOrig    = (pos.totalBought != null && isFinite(pos.totalBought)) ? pos.totalBought : sh;
  const procOrig  = (pos.proceeds    != null && isFinite(pos.proceeds))    ? pos.proceeds    : 0;

  const p   = num(pct), sp = num(sellPrice), bp = num(buyPrice);

  const sharesToSell   = sh * (p / 100);
  const simProceeds    = sharesToSell * sp;
  const profit         = simProceeds - (sharesToSell * avg);
  const repurchaseCost = sharesToSell * bp;

  // Nuevo Promedio: (hist total + recompra) / (totalBought + shares recompradas)
  // = promedio de TODAS las compras que habrás hecho alguna vez
  const newHistTotal = histOrig + repurchaseCost;
  const newTbTotal   = tbOrig  + sharesToSell;
  const newProm      = newHistTotal / newTbTotal;
  const promDelta    = newProm - baseProm;
  const promPct      = (promDelta / baseProm) * 100;

  // Nuevo CEF: parte del CEF histórico ya reducido + esta operación
  // = (cef_hist × sh - simProceeds + repurchaseCost) / sh
  const newCef     = (baseCef * sh - simProceeds + repurchaseCost) / sh;
  const cefDelta   = newCef - baseCef;
  const cefPct     = (cefDelta / baseCef) * 100;

  const ready  = sh > 0 && avg > 0 && p > 0 && p <= 100 && sp > 0 && bp > 0;
  const better = ready && newCef < baseCef;

  return (
    <Card title="🔄 Simulador: Venta + Recompra">
      <div style={{ fontSize: 12, color: C.muted, marginBottom: 14, lineHeight: 1.5 }}>
        Simula vender un % de tu posición a cierto precio y recomprar esas mismas acciones más barato.
        Ve en cuánto quedan tu <b style={{ color: C.ink }}>Promedio</b> y tu <b style={{ color: C.ink }}>Costo Efectivo Promedio</b> después del ciclo completo.
      </div>

      <Input label="% de la posición a vender" value={pct}       onChange={setPct}   unit="%" placeholder="ej. 30" />
      <Input label="Precio de venta"            value={sellPrice} onChange={setSell}  unit="USD" />
      <Input label="Precio de recompra"         value={buyPrice}  onChange={setBuy}   unit="USD" />

      {ready && (
        <>
          <div style={{ borderTop: `1px solid ${C.line}`, margin: "14px 0 10px" }} />
          <Row label="Acciones a vender y recomprar" value={fmt(sharesToSell, 4)}  unit="acciones" />
          <Row label="Recibido de la venta"           value={fmt(simProceeds)}       unit="USD" />
          <Row label="Ganancia realizada en venta"    value={fmtSigned(profit)}      unit="USD" color={profit >= 0 ? C.green : C.red} />
          <Row label="Costo de la recompra"           value={fmt(repurchaseCost)}    unit="USD" />

          <div style={{ borderTop: `1px solid ${C.line}`, margin: "14px 0 10px" }} />

          {/* Nuevo Promedio: hist total / totalBought total */}
          <Row
            label={`Nuevo promedio histórico (antes: $${fmt(baseProm, 2)})`}
            value={fmt(newProm, 4)}
            unit="USD"
            strong
            color={promDelta < 0 ? C.green : promDelta > 0 ? C.red : C.muted}
          />

          {/* Nuevo CEF */}
          <div style={{
            marginTop: 14, padding: 16, borderRadius: 10,
            background: better ? C.greenSoft : C.redSoft,
            border: `1px solid ${better ? C.green : C.red}`,
          }}>
            <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.07em", textTransform: "uppercase", color: better ? C.green : C.red, marginBottom: 6 }}>
              Costo Efectivo Promedio resultante
            </div>
            <div style={{ fontFamily: MONO, fontSize: 26, fontWeight: 800, color: better ? C.green : C.red }}>
              ${fmt(newCef, 4)}
            </div>
            <div style={{ fontFamily: MONO, fontSize: 13, color: C.ink, marginTop: 6 }}>
              {better ? "▼" : "▲"} {fmtSigned(cefDelta, 4)} USD · {fmtSigned(cefPct, 2)}% vs CEF actual (${fmt(baseCef, 2)})
            </div>
            {!better && bp >= sp && (
              <div style={{ fontSize: 12, color: C.red, marginTop: 8 }}>
                ⚠️ Recompras más caro que vendes — el CEF sube. Intenta un precio de recompra menor al de venta.
              </div>
            )}
          </div>
        </>
      )}
    </Card>
  );
}

// ── Portafolio + carga de Excel ──────────────────────────────────────────────
function Portafolio({ portfolio, onLoad, onFileParsed, onClear, status }) {
  const fileRef = useRef(null);
  const totInv = portfolio.reduce((a, p) => a + p.inv, 0);
  const totHist = portfolio.reduce((a, p) => a + p.hist, 0);

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = ""; // permite re-subir el mismo archivo
    if (!file) return;
    try {
      let wb;
      if (/\.csv$/i.test(file.name)) {
        const text = await file.text();
        wb = XLSX.read(text, { type: "string", raw: true });
      } else {
        const buf = await file.arrayBuffer();
        wb = XLSX.read(buf, { type: "array" });
      }
      const parsed = parseWorkbook(wb);
      onFileParsed(parsed, file.name);
    } catch (err) {
      onFileParsed(null, file.name);
    }
  };

  return (
    <>
      <Card
        title="📥 Actualizar desde Excel"
      >
        <div style={{ fontSize: 13, color: C.muted, marginBottom: 12 }}>
          Sube el <b style={{ color: C.ink }}>CSV de IBKR</b> (Activity Statement o reporte de posiciones).
          La app lee la sección <i>Open Positions</i>: símbolo, cantidad, Cost Price y Cost Basis.
          También acepta Excel con columnas Ticker / Acciones / Precio Promedio.
          Después de leerlo te preguntaré si quieres <b style={{ color: C.ink }}>reemplazar</b> los datos
          actuales o <b style={{ color: C.ink }}>añadirlos</b> a los que ya están.
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.xlsm,.csv" onChange={handleFile} style={{ display: "none" }} />
        <Btn kind="blue" onClick={() => fileRef.current?.click()}>📂 Subir CSV de IBKR o Excel</Btn>
        {portfolio.length > 0 && (
          <div style={{ marginTop: 8 }}>
            <Btn kind="ghost" onClick={onClear}>🗑 Limpiar portafolio</Btn>
          </div>
        )}
        {status && (
          <div style={{ marginTop: 10, fontSize: 12, fontFamily: MONO, color: status.ok ? C.green : C.red }}>
            {status.msg}
          </div>
        )}
      </Card>

      <Card title="📊 Portafolio unificado">
        {portfolio.length === 0 ? (
          <div style={{ textAlign: "center", padding: "24px 0", color: C.muted }}>
            <div style={{ fontSize: 32, marginBottom: 8 }}>📂</div>
            <div style={{ fontSize: 14, fontWeight: 600, color: C.ink, marginBottom: 4 }}>Sin posiciones cargadas</div>
            <div style={{ fontSize: 12 }}>Sube tu CSV de IBKR arriba para ver tu portafolio aquí.</div>
          </div>
        ) : (
          <>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: MONO, fontSize: 13 }}>
                <thead>
                  <tr>
                    {["Ticker", "Acciones", "Prom.", "Costo Ef.", "Inv. Actual", "Inv. Histórico"].map((h, i) => (
                      <th key={h} style={{ textAlign: i === 0 ? "left" : "right", fontSize: 10, textTransform: "uppercase", letterSpacing: "0.05em", color: C.muted, padding: "6px 8px", borderBottom: `2px solid ${C.ink}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {portfolio.map((p) => (
                    <tr key={p.t} onClick={() => onLoad(p.t)} style={{ cursor: "pointer" }} title={`${p.n} — toca para cargar en las calculadoras`}>
                      <td style={{ padding: "8px", fontWeight: 700, color: C.blue, borderBottom: `1px dashed ${C.line}` }}>{p.t}</td>
                      <td style={{ padding: "8px", textAlign: "right", borderBottom: `1px dashed ${C.line}` }}>{fmt(p.sh, 2)}</td>
                      <td style={{ padding: "8px", textAlign: "right", borderBottom: `1px dashed ${C.line}` }}>{fmt(p.prom ?? p.avg, 2)}</td>
                      <td style={{ padding: "8px", textAlign: "right", borderBottom: `1px dashed ${C.line}`,
                        color: p.cef < p.avg ? C.green : p.cef > p.avg ? C.red : C.muted,
                        fontWeight: p.cef !== p.avg ? 700 : 400,
                      }}>{fmt(p.cef ?? p.avg, 2)}</td>
                      <td style={{ padding: "8px", textAlign: "right", borderBottom: `1px dashed ${C.line}` }}>{fmt(p.inv)}</td>
                      <td style={{ padding: "8px", textAlign: "right", color: C.muted, borderBottom: `1px dashed ${C.line}` }}>{fmt(p.hist)}</td>
                    </tr>
                  ))}
                  <tr>
                    <td style={{ padding: "10px 8px", fontWeight: 800 }}>TOTAL</td>
                    <td style={{ padding: "10px 8px", textAlign: "right", color: C.muted }}>{portfolio.length} pos.</td>
                    <td /><td />
                    <td style={{ padding: "10px 8px", textAlign: "right", fontWeight: 800 }}>{fmt(totInv)}</td>
                    <td style={{ padding: "10px 8px", textAlign: "right", fontWeight: 800, color: C.muted }}>{fmt(totHist)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>
              💡 Toca cualquier fila y elige si quieres simular una compra (Promedio) o una venta.
            </div>
          </>
        )}
      </Card>
    </>
  );
}

// ── Diálogo: ¿reemplazar o añadir? ───────────────────────────────────────────
function ImportDialog({ pending, currentCount, onReplace, onMerge, onCancel }) {
  const dupes = pending.dupes, news = pending.rows.length - dupes;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,27,38,0.55)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 50 }}>
      <div style={{ background: C.card, borderRadius: 16, padding: 22, maxWidth: 420, width: "100%", border: `1px solid ${C.line}` }}>
        <div style={{ fontSize: 16, fontWeight: 800, marginBottom: 6 }}>📥 {pending.fileName}</div>
        <div style={{ fontSize: 13, color: C.muted, marginBottom: 16, lineHeight: 1.5 }}>
          Encontré <b style={{ color: C.ink }}>{pending.rows.length} posiciones</b> en el archivo
          ({news} nuevas, {dupes} que ya existen). Tu portafolio actual tiene {currentCount} posiciones.
          <br />¿Qué quieres hacer?
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <Btn kind="blue" onClick={onReplace}>🔄 Reemplazar todo con el archivo</Btn>
          <Btn kind="green" onClick={onMerge}>➕ Añadir posiciones nuevas</Btn>
          <Btn kind="ghost" onClick={onCancel}>Cancelar</Btn>
        </div>
        <div style={{ fontSize: 11, color: C.muted, marginTop: 12, lineHeight: 1.5 }}>
          <b>Reemplazar:</b> borra todo lo actual y usa solo este archivo.<br />
          <b>Añadir (Merge):</b> suma las posiciones de tickers que ya existen (para combinar dos cuentas) y agrega los nuevos.
        </div>
      </div>
    </div>
  );
}

// ── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("promedio");
  const [portfolio, setPortfolio] = useState([]);
  const [ticker, setTicker] = useState("");
  const [shares, setShares] = useState("");
  const [avg, setAvg] = useState("");
  const [cef, setCef] = useState(null); // Costo Efectivo Promedio del ticker cargado
  const [pending, setPending] = useState(null); // { rows, fileName, dupes }
  const [status, setStatus] = useState(null);
  const [loaded, setLoaded] = useState(false);

  // Cargar portafolio guardado (persistente entre sesiones)
  useEffect(() => {
    (async () => {
      try {
        if (window.storage) {
          const r = await window.storage.get(STORAGE_KEY);
          if (r?.value) {
            const data = JSON.parse(r.value);
            if (Array.isArray(data) && data.length) setPortfolio(data);
          }
        }
      } catch {
        // sin datos guardados todavía — se usa el portafolio del Excel original
      }
      setLoaded(true);
    })();
  }, []);

  const persist = async (data) => {
    setPortfolio(data);
    try {
      if (window.storage) await window.storage.set(STORAGE_KEY, JSON.stringify(data));
    } catch {
      // si falla el guardado, la app sigue funcionando en memoria
    }
  };

  const onFileParsed = (rows, fileName) => {
    if (!rows) {
      setStatus({ ok: false, msg: `✕ No pude leer "${fileName}". ¿Es un Excel válido?` });
      return;
    }
    if (!rows.length) {
      setStatus({ ok: false, msg: `✕ No encontré posiciones en "${fileName}". Verifica que tenga columnas Ticker / Acciones / Precio Promedio.` });
      return;
    }
    const existing = new Set(portfolio.map((p) => p.t));
    const dupes = rows.filter((r) => existing.has(r.t)).length;
    setStatus(null);
    setPending({ rows, fileName, dupes });
  };

  const applyReplace = async () => {
    await persist([...pending.rows].sort((a, b) => a.t.localeCompare(b.t)));
    setStatus({ ok: true, msg: `✓ Portafolio reemplazado: ${pending.rows.length} posiciones de "${pending.fileName}".` });
    setPending(null);
    setTicker(""); setShares(""); setAvg("");
  };

  const applyMerge = async () => {
    const merged = mergePortfolios(portfolio, pending.rows);
    await persist(merged);
    const nuevas = pending.rows.length - pending.dupes;
    setStatus({ ok: true, msg: `✓ Portafolio actualizado: ${merged.length} posiciones (${pending.dupes} actualizadas, ${nuevas} nuevas).` });
    setPending(null);
    setTicker(""); setShares(""); setAvg("");
  };

  const [destPick, setDestPick] = useState(null);

  const onClear = async () => {
    await persist([]);
    setTicker(""); setShares(""); setAvg(""); setCef(null);
    setStatus({ ok: true, msg: "✓ Portafolio limpio." });
  }; // ticker pendiente de elegir destino

  const loadFromPortfolio = (t) => {
    const p = portfolio.find((x) => x.t === t);
    if (p) setDestPick(t);
  };

  const goTo = (t, destino) => {
    const p = portfolio.find((x) => x.t === t);
    if (p) {
      setTicker(t);
      setShares(String(p.sh));
      setAvg(String(p.avg));
      setCef(p.cef ?? null);
      setTab(destino);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
    setDestPick(null);
  };

  const tabs = [
    { id: "promedio",   label: "📈 Promedio" },
    { id: "venta",      label: "💰 Venta" },
    { id: "simulador",  label: "🔄 Simulador" },
    { id: "portafolio", label: "📊 Portafolio" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.ink, fontFamily: "system-ui, -apple-system, 'Segoe UI', sans-serif" }}>
      <div style={{ maxWidth: 640, margin: "0 auto", padding: "20px 14px 48px" }}>
        <div style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: MONO, fontSize: 11, fontWeight: 700, letterSpacing: "0.18em", color: C.muted }}>
            COMPRA · VENTA · PORTAFOLIO
          </div>
          <h1 style={{ margin: "2px 0 0", fontSize: 26, fontWeight: 800, letterSpacing: "-0.02em" }}>
            Calculadora de inversiones
          </h1>
          {ticker && (
            <div style={{ marginTop: 6, display: "inline-block", fontFamily: MONO, fontSize: 13, fontWeight: 700, color: C.blue, background: C.blueSoft, border: `1px solid ${C.blue}`, borderRadius: 6, padding: "3px 10px" }}>
              {ticker} cargado
            </div>
          )}
        </div>

        <div style={{ display: "flex", gap: 6, marginBottom: 18, background: C.card, border: `1px solid ${C.line}`, borderRadius: 12, padding: 5 }}>
          {tabs.map((t) => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              style={{
                flex: 1, padding: "10px 4px", fontSize: 13, fontWeight: 700, border: "none",
                borderRadius: 8, cursor: "pointer",
                background: tab === t.id ? C.ink : "transparent",
                color: tab === t.id ? "#fff" : C.muted,
                transition: "background .15s",
              }}
            >
              {t.label}
            </button>
          ))}
        </div>

        {!loaded ? (
          <Card><div style={{ fontSize: 13, color: C.muted, textAlign: "center", padding: 10 }}>Cargando tu portafolio…</div></Card>
        ) : (
          <>
            {tab !== "portafolio" && (
              <PositionPicker
                portfolio={portfolio}
                ticker={ticker} setTicker={setTicker}
                shares={shares} setShares={setShares}
                avg={avg} setAvg={setAvg}
                setCef={setCef}
              />
            )}
            {tab === "promedio" && <Promedio pos={{ shares, avg, cef }} />}
            {tab === "venta" && <Venta pos={{ shares, avg, cef, prom: portfolio.find(p=>p.t===ticker)?.prom, hist: portfolio.find(p=>p.t===ticker)?.hist, totalBought: portfolio.find(p=>p.t===ticker)?.totalBought, proceeds: portfolio.find(p=>p.t===ticker)?.proceeds }} />}
            {tab === "simulador" && <CicloVentaRecompra pos={{ shares, avg, cef, prom: portfolio.find(p=>p.t===ticker)?.prom, hist: portfolio.find(p=>p.t===ticker)?.hist, totalBought: portfolio.find(p=>p.t===ticker)?.totalBought, proceeds: portfolio.find(p=>p.t===ticker)?.proceeds }} />}
            {tab === "portafolio" && (
              <Portafolio
                portfolio={portfolio}
                onLoad={loadFromPortfolio}
                onFileParsed={onFileParsed}
                onClear={onClear}
                status={status}
              />
            )}
          </>
        )}

        <div style={{ fontSize: 11, color: C.muted, textAlign: "center", marginTop: 8 }}>
          💡 Escribe en los campos con borde azul · todo lo demás se calcula solo.
          <br />
          Herramienta de cálculo, no asesoría financiera.
        </div>
      </div>

      {pending && (
        <ImportDialog
          pending={pending}
          currentCount={portfolio.length}
          onReplace={applyReplace}
          onMerge={applyMerge}
          onCancel={() => setPending(null)}
        />
      )}

      {destPick && (
        <div
          onClick={() => setDestPick(null)}
          style={{ position: "fixed", inset: 0, background: "rgba(15,27,38,0.55)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 50 }}
        >
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.card, borderRadius: 16, padding: 22, maxWidth: 360, width: "100%", border: `1px solid ${C.line}` }}>
            <div style={{ fontFamily: MONO, fontSize: 18, fontWeight: 800, color: C.blue, marginBottom: 4 }}>
              {destPick}
            </div>
            <div style={{ fontSize: 13, color: C.muted, marginBottom: 16 }}>
              {portfolio.find((x) => x.t === destPick)?.n || ""}
              <br />¿Qué quieres calcular con esta posición?
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              <Btn kind="blue" onClick={() => goTo(destPick, "promedio")}>📈 Promedio — Simular una compra</Btn>
              <Btn kind="green" onClick={() => goTo(destPick, "venta")}>💰 Venta — Simular una venta</Btn>
              <Btn kind="ghost" onClick={() => setDestPick(null)}>Cancelar</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

