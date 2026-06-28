import * as XLSX from "xlsx";
import { num } from "./utils.ts";
import type { Position } from "./types.ts";

type Row = unknown[];

const low = (c: unknown): string => String(c ?? "").trim().toLowerCase();

export function parseWorkbook(wb: XLSX.WorkBook): Position[] {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json<Row>(wb.Sheets[name], { header: 1, defval: null, raw: false });
    const txn = parseIBKRTransactions(rows);
    if (txn.length) return txn;
    const ibkr = parseIBKROpenPositions(rows);
    if (ibkr.length) return ibkr;
  }
  let best: Position[] = [];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json<Row>(wb.Sheets[name], { header: 1, defval: null });
    const own = parseSheetRows(rows);
    const gen = parseGenericPositions(rows);
    const found = own.length >= gen.length ? own : gen;
    if (found.length > best.length) best = found;
  }
  return best;
}

function parseDate(s: unknown): number {
  if (!s) return 0;
  const str = String(s).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return new Date(str).getTime();
  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(`${m[3]}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`).getTime();
  return 0;
}

interface TxLot { qty: number; cost: number }
interface TxPos { sh: number; cost: number; lots: TxLot[]; hist: number; proceeds: number; totalBought: number; n: string }

function parseIBKRTransactions(rows: Row[]): Position[] {
  let cols: string[] | null = null;
  const txns: Array<{ date: number; type: string; sym: string; qty: number; gross: number; comm: number; desc: string }> = [];

  for (const row of rows) {
    if (!row || low(row[0]) !== "transaction history") continue;
    if (low(row[1]) === "header") { cols = row.map(low); continue; }
    if (low(row[1]) !== "data" || !cols) continue;
    const c = cols;
    const get = (name: string): string => {
      const i = c.indexOf(name);
      return i > -1 ? String(row[i] ?? "").trim() : "";
    };
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

  txns.sort((a, b) => a.date - b.date);

  const map = new Map<string, TxPos>();
  for (const tx of txns) {
    if (!map.has(tx.sym)) map.set(tx.sym, { sh: 0, cost: 0, lots: [], hist: 0, proceeds: 0, totalBought: 0, n: tx.desc });
    const pos = map.get(tx.sym)!;

    if (tx.type === "Buy") {
      const paid = Math.abs(tx.gross) + Math.abs(isFinite(tx.comm) ? tx.comm : 0);
      const bqty = isFinite(tx.qty) ? tx.qty : 0;
      pos.sh          += bqty;
      pos.cost        += paid;
      pos.lots.push({ qty: bqty, cost: paid });
      pos.hist        += paid;
      pos.totalBought += bqty;
    } else {
      const soldQty  = Math.abs(isFinite(tx.qty)  ? tx.qty  : 0);
      const received = Math.abs(isFinite(tx.gross) ? tx.gross : 0)
                     - Math.abs(isFinite(tx.comm)  ? tx.comm  : 0);
      const avgCost  = pos.sh > 0 ? pos.cost / pos.sh : 0;
      pos.sh   -= soldQty;
      pos.cost -= avgCost * soldQty;
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

  const result: Position[] = [];
  for (const [t, p] of map) {
    if (!(p.sh > 0.0001)) continue;
    const avg     = p.cost / p.sh;
    const invFifo = p.lots.reduce((a, l) => a + l.cost, 0);
    const netCost = p.hist - p.proceeds;
    const prom    = p.totalBought > 0 ? p.hist / p.totalBought : avg;
    result.push({
      t, n: p.n,
      sh:          Math.round(p.sh          * 10000) / 10000,
      avg:         Math.round(avg           * 10000) / 10000,
      prom:        Math.round(prom          * 10000) / 10000,
      inv:         Math.round(invFifo       * 100)   / 100,
      hist:        Math.round(p.hist        * 100)   / 100,
      proceeds:    Math.round(p.proceeds    * 100)   / 100,
      totalBought: Math.round(p.totalBought * 10000) / 10000,
      cef:         Math.round((netCost / p.sh) * 10000) / 10000,
    });
  }
  return result.sort((a, b) => a.t.localeCompare(b.t));
}

function parseIBKROpenPositions(rows: Row[]): Position[] {
  const names: Record<string, string> = {};
  let fiHeader: string[] | null = null;
  for (const r of rows) {
    if (!r || low(r[0]) !== "financial instrument information") continue;
    if (low(r[1]) === "header") { fiHeader = r.map(low); continue; }
    if (low(r[1]) === "data" && fiHeader) {
      const si = fiHeader.indexOf("symbol"), di = fiHeader.indexOf("description");
      if (si > -1 && di > -1 && r[si]) names[String(r[si]).trim().toUpperCase()] = String(r[di] ?? "").trim();
    }
  }
  const isPosSection = (s: string) => s === "open positions" || s === "positions";
  let header: string[] | null = null;
  const raw: Position[] = [];
  for (const r of rows) {
    if (!r || !isPosSection(low(r[0]))) continue;
    if (low(r[1]) === "header") { header = r.map(low); continue; }
    if (low(r[1]) !== "data" || !header) continue;
    const h = header;
    const col = (n: string) => h.indexOf(n);
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
  const map = new Map<string, Position>();
  for (const p of raw) {
    const ex = map.get(p.t);
    if (ex) {
      const sh = ex.sh + p.sh, inv = ex.inv + p.inv;
      map.set(p.t, { ...ex, sh, inv, avg: inv / sh, hist: inv });
    } else {
      map.set(p.t, p);
    }
  }
  return [...map.values()].sort((a, b) => a.t.localeCompare(b.t));
}

interface ColMap { t: number; n: number; sh: number; avg: number; inv: number }

function parseGenericPositions(rows: Row[]): Position[] {
  let hIdx = -1;
  let cols: ColMap = { t: -1, n: -1, sh: -1, avg: -1, inv: -1 };
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const r = rows[i];
    const c: ColMap = { t: -1, n: -1, sh: -1, avg: -1, inv: -1 };
    r.forEach((cell, j) => {
      const s = low(cell);
      if (!s) return;
      if      (c.t   === -1 && (s === "symbol" || s.includes("instrument") || s === "ticker")) c.t = j;
      else if (c.n   === -1 && (s.includes("description") || s.includes("nombre") || s === "name")) c.n = j;
      else if (c.sh  === -1 && (s === "position" || s === "quantity" || s.includes("accion") || s === "shares" || s === "qty")) c.sh = j;
      else if (c.avg === -1 && (s.includes("average price") || s.includes("avg price") || s.includes("cost price") || s.includes("avg cost") || s.includes("promedio"))) c.avg = j;
      else if (c.inv === -1 && s.includes("cost basis")) c.inv = j;
    });
    if (c.t > -1 && c.sh > -1 && c.avg > -1) { hIdx = i; cols = c; break; }
  }
  if (hIdx === -1) return [];
  const out: Position[] = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i];
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

interface SheetColMap { t: number; n: number; sh: number; avg: number; inv: number; hist: number }

function parseSheetRows(rows: Row[]): Position[] {
  let hIdx = -1;
  let cols: SheetColMap = { t: -1, n: -1, sh: -1, avg: -1, inv: -1, hist: -1 };
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const r = rows[i];
    const ti = r.findIndex((c) => low(c) === "ticker");
    if (ti === -1) continue;
    cols = { t: ti, n: -1, sh: -1, avg: -1, inv: -1, hist: -1 };
    r.forEach((c, j) => {
      const s = low(c);
      if      (s.includes("nombre") || s.includes("name")) cols.n = j;
      else if (s.includes("accion")) cols.sh = j;
      else if (s.includes("promedio") || s.includes("avg")) cols.avg = j;
      else if (s.includes("hist")) cols.hist = j;
      else if (s.includes("invertido") || s.includes("invested")) cols.inv = j;
    });
    if (cols.sh !== -1 && cols.avg !== -1) { hIdx = i; break; }
  }
  if (hIdx === -1) return [];
  const out: Position[] = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const t = String(r[cols.t] ?? "").trim();
    if (!t) break;
    if (low(t) === "total" || t.startsWith("💡")) break;
    const sh = num(r[cols.sh]), avg = num(r[cols.avg]);
    if (!(sh > 0) || !(avg > 0)) continue;
    const inv  = cols.inv  !== -1 && isFinite(num(r[cols.inv]))  ? num(r[cols.inv])  : sh * avg;
    const hist = cols.hist !== -1 && isFinite(num(r[cols.hist])) ? num(r[cols.hist]) : inv;
    out.push({ t: t.toUpperCase(), n: cols.n !== -1 ? String(r[cols.n] ?? "").trim() : "", sh, avg, inv, hist });
  }
  return out;
}

export function mergePortfolios(current: Position[], incoming: Position[]): Position[] {
  const key = (t: string) => String(t ?? "").trim().toUpperCase();
  const map = new Map<string, Position>(current.map((p) => [key(p.t), { ...p, t: key(p.t) }]));
  for (const np of incoming) {
    const k  = key(np.t);
    const ex = map.get(k);
    if (ex) {
      const sh          = Number(ex.sh)          + Number(np.sh);
      const inv         = Number(ex.inv)         + Number(np.inv);
      const hist        = Number(ex.hist)        + Number(np.hist);
      const proceeds    = (ex.proceeds    ?? 0)  + (np.proceeds    ?? 0);
      const totalBought = (ex.totalBought ?? 0)  + (np.totalBought ?? 0);
      const netCost     = hist - proceeds;
      const prom        = totalBought > 0 ? hist / totalBought : (sh > 0 ? inv / sh : 0);
      map.set(k, {
        t: k, n: np.n || ex.n,
        sh, inv, hist, proceeds, totalBought, prom,
        avg: sh > 0 ? inv / sh : 0,
        cef: proceeds > 0 ? netCost / sh : (sh > 0 ? inv / sh : 0),
      });
    } else {
      map.set(k, { ...np, t: k });
    }
  }
  return [...map.values()].sort((a, b) => a.t.localeCompare(b.t));
}
