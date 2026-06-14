import { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

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
	let s = String(v ?? "")
		.trim()
		.replace(/[$\s"]/g, "");
	if (s.startsWith("(") && s.endsWith(")")) s = "-" + s.slice(1, -1); // (123) → -123
	if (s.includes(",") && s.includes("."))
		s = s.replace(/,/g, ""); // 1,234.56
	else if (/^-?\d{1,3}(,\d{3})+$/.test(s))
		s = s.replace(/,/g, ""); // 1,234
	else if (s.includes(",")) s = s.replace(",", "."); // 12,34 → 12.34
	const n = parseFloat(s);
	return isFinite(n) ? n : NaN;
};
const fmt = (v, d = 2) =>
	!isFinite(v)
		? "—"
		: v.toLocaleString("en-US", {
				minimumFractionDigits: d,
				maximumFractionDigits: d,
			});
const fmtSigned = (v, d = 2) =>
	!isFinite(v) ? "—" : (v > 0 ? "+" : "") + fmt(v, d);

// ── Lectura del archivo ──────────────────────────────────────────────────────
// Orden de intento: 1) IBKR Open Positions, 2) IBKR Transaction History,
// 3) hoja propia con "Ticker", 4) tabla genérica.
function parseWorkbook(wb) {
	let best = [];
	for (const name of wb.SheetNames) {
		const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], {
			header: 1,
			defval: null,
		});
		const ibkr = parseIBKR(rows);
		if (ibkr.length) return ibkr;
		const txns = parseIBKRTransactions(rows);
		if (txns.length) return txns;
		const own = parseSheetRows(rows);
		const gen = parseGenericPositions(rows);
		const found = own.length >= gen.length ? own : gen;
		if (found.length > best.length) best = found;
	}
	return best;
}

// IBKR Transaction History CSV — calcula posiciones acumulando compras y
// reduciendo cost basis en ventas (promedio ponderado).
// Formato: Transaction History,Header,Date,Account,Description,Transaction Type,
//          Symbol,Quantity,Price,Price Currency,Gross Amount,Commission,Net Amount
function parseIBKRTransactions(rows) {
	const low = (c) => String(c ?? "").trim().toLowerCase();
	let header = null;
	const txns = [];

	for (const r of rows) {
		if (!r || low(r[0]) !== "transaction history") continue;
		if (low(r[1]) === "header") { header = r.map(low); continue; }
		if (low(r[1]) !== "data" || !header) continue;

		const col = (n) => header.indexOf(n);
		const txType = String(r[col("transaction type")] ?? "").trim();
		const currency = String(r[col("price currency")] ?? "").trim().toUpperCase();
		if (!["Buy", "Sell"].includes(txType) || currency !== "USD") continue;

		const sym = String(r[col("symbol")] ?? "").trim().toUpperCase();
		const qty = num(r[col("quantity")]);
		const gross = num(r[col("gross amount")]);
		const comm = num(r[col("commission")]);
		const desc = String(r[col("description")] ?? "").trim();
		if (!sym || !isFinite(qty)) continue;
		txns.push({ sym, txType, qty, gross, comm, desc });
	}

	if (!txns.length) return [];

	const posMap = new Map();
	for (const tx of txns) {
		const pos = posMap.get(tx.sym) || { sh: 0, cost: 0, n: "" };
		if (!pos.n && tx.desc) pos.n = tx.desc;
		if (tx.txType === "Buy") {
			pos.sh += tx.qty;
			pos.cost += Math.abs(tx.gross) + Math.abs(tx.comm);
		} else {
			const sold = Math.abs(tx.qty);
			if (pos.sh > 0) {
				pos.cost = Math.max(0, pos.cost - (pos.cost / pos.sh) * sold);
				pos.sh = Math.max(0, pos.sh - sold);
			}
		}
		posMap.set(tx.sym, pos);
	}

	return [...posMap.entries()]
		.filter(([, p]) => p.sh > 0.001 && p.cost > 0)
		.map(([t, p]) => {
			const avg = p.cost / p.sh;
			return { t, n: p.n, sh: p.sh, avg, inv: p.cost, hist: p.cost };
		})
		.sort((a, b) => a.t.localeCompare(b.t));
}

// Formato IBKR (Activity Statement / Flex): filas tipo
// "Open Positions,Header,DataDiscriminator,Asset Category,Currency,Symbol,Quantity,Mult,Cost Price,Cost Basis,..."
// "Open Positions,Data,Summary,Stocks,USD,AAPL,10,1,150.25,1502.50,..."
function parseIBKR(rows) {
	const low = (c) =>
		String(c ?? "")
			.trim()
			.toLowerCase();

	// Nombres de empresas desde "Financial Instrument Information"
	const names = {};
	let fiHeader = null;
	for (const r of rows) {
		if (!r || low(r[0]) !== "financial instrument information") continue;
		if (low(r[1]) === "header") {
			fiHeader = r.map(low);
			continue;
		}
		if (low(r[1]) === "data" && fiHeader) {
			const si = fiHeader.indexOf("symbol"),
				di = fiHeader.indexOf("description");
			if (si > -1 && di > -1 && r[si]) {
				names[String(r[si]).trim().toUpperCase()] = String(r[di] ?? "").trim();
			}
		}
	}

	// Posiciones desde "Open Positions" (también acepta "Positions")
	const isPosSection = (s) => s === "open positions" || s === "positions";
	let header = null;
	const raw = [];
	for (const r of rows) {
		if (!r || !isPosSection(low(r[0]))) continue;
		if (low(r[1]) === "header") {
			header = r.map(low);
			continue;
		}
		if (low(r[1]) !== "data" || !header) continue;

		const col = (n) => header.indexOf(n);
		const dd = col("datadiscriminator");
		if (dd > -1 && r[dd] && low(r[dd]) !== "summary") continue; // omite lotes, toma el resumen
		const cat = col("asset category");
		if (cat > -1 && r[cat] && !low(r[cat]).includes("stock")) continue; // solo acciones/ETF

		const t = String(r[col("symbol")] ?? "")
			.trim()
			.toUpperCase();
		const sh = num(r[col("quantity")]);
		let avg = col("cost price") > -1 ? num(r[col("cost price")]) : NaN;
		let inv = col("cost basis") > -1 ? num(r[col("cost basis")]) : NaN;
		if (!t || !(sh > 0)) continue;
		if (!isFinite(avg) && isFinite(inv)) avg = inv / sh;
		if (!isFinite(inv) && isFinite(avg)) inv = sh * avg;
		if (!(avg > 0)) continue;
		raw.push({ t, n: names[t] || "", sh, avg, inv, hist: inv });
	}

	// Combina duplicados (varias cuentas en un mismo reporte)
	const map = new Map();
	for (const p of raw) {
		const ex = map.get(p.t);
		if (ex) {
			const sh = ex.sh + p.sh,
				inv = ex.inv + p.inv;
			map.set(p.t, { ...ex, sh, inv, avg: inv / sh, hist: inv });
		} else map.set(p.t, p);
	}
	return [...map.values()].sort((a, b) => a.t.localeCompare(b.t));
}

// Tabla genérica (p. ej. exportación del portal de IBKR u otro broker):
// encabezados con Symbol/Instrument + Position/Quantity + Average/Cost Price.
function parseGenericPositions(rows) {
	const low = (c) =>
		String(c ?? "")
			.trim()
			.toLowerCase();
	let hIdx = -1,
		cols = {};
	for (let i = 0; i < Math.min(rows.length, 40); i++) {
		const r = rows[i] || [];
		const c = { t: -1, n: -1, sh: -1, avg: -1, inv: -1 };
		r.forEach((cell, j) => {
			const s = low(cell);
			if (!s) return;
			if (
				c.t === -1 &&
				(s === "symbol" || s.includes("instrument") || s === "ticker")
			)
				c.t = j;
			else if (
				c.n === -1 &&
				(s.includes("description") || s.includes("nombre") || s === "name")
			)
				c.n = j;
			else if (
				c.sh === -1 &&
				(s === "position" ||
					s === "quantity" ||
					s.includes("accion") ||
					s === "shares" ||
					s === "qty")
			)
				c.sh = j;
			else if (
				c.avg === -1 &&
				(s.includes("average price") ||
					s.includes("avg price") ||
					s.includes("cost price") ||
					s.includes("avg cost") ||
					s.includes("promedio"))
			)
				c.avg = j;
			else if (c.inv === -1 && s.includes("cost basis")) c.inv = j;
		});
		if (c.t > -1 && c.sh > -1 && c.avg > -1) {
			hIdx = i;
			cols = c;
			break;
		}
	}
	if (hIdx === -1) return [];

	const out = [];
	for (let i = hIdx + 1; i < rows.length; i++) {
		const r = rows[i] || [];
		const t = String(r[cols.t] ?? "").trim();
		if (!t) continue;
		if (low(t) === "total" || t.startsWith("💡")) break;
		const sh = num(r[cols.sh]);
		const avg = num(r[cols.avg]);
		if (!(sh > 0) || !(avg > 0)) continue;
		const inv =
			cols.inv > -1 && isFinite(num(r[cols.inv])) ? num(r[cols.inv]) : sh * avg;
		out.push({
			t: t.toUpperCase(),
			n: cols.n > -1 ? String(r[cols.n] ?? "").trim() : "",
			sh,
			avg,
			inv,
			hist: inv,
		});
	}
	return out;
}

function parseSheetRows(rows) {
	const low = (c) =>
		String(c ?? "")
			.trim()
			.toLowerCase();
	// 1) localizar la fila de encabezados
	let hIdx = -1,
		cols = {};
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
		if (cols.sh !== -1 && cols.avg !== -1) {
			hIdx = i;
			break;
		}
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
		const inv =
			cols.inv !== -1 && isFinite(num(r[cols.inv]))
				? num(r[cols.inv])
				: sh * avg;
		const hist =
			cols.hist !== -1 && isFinite(num(r[cols.hist])) ? num(r[cols.hist]) : inv;
		out.push({
			t: t.toUpperCase(),
			n: cols.n !== -1 ? String(r[cols.n] ?? "").trim() : "",
			sh,
			avg,
			inv,
			hist,
		});
	}
	return out;
}

// Combinar: tickers repetidos se suman (promedio ponderado), nuevos se agregan.
function mergePortfolios(current, incoming) {
	const map = new Map(current.map((p) => [p.t, { ...p }]));
	for (const np of incoming) {
		const ex = map.get(np.t);
		if (ex) {
			const sh = ex.sh + np.sh;
			const inv = ex.inv + np.inv;
			map.set(np.t, {
				t: np.t,
				n: np.n || ex.n,
				sh,
				inv,
				avg: sh > 0 ? inv / sh : 0,
				hist: ex.hist + np.hist,
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
		<div
			style={{
				fontSize: 11,
				fontWeight: 700,
				letterSpacing: "0.08em",
				textTransform: "uppercase",
				color: C.muted,
				marginBottom: 4,
			}}
		>
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
						flex: 1,
						minWidth: 0,
						fontFamily: MONO,
						fontSize: 18,
						fontWeight: 600,
						color: C.ink,
						background: C.blueSoft,
						border: `2px solid ${C.blue}`,
						borderRadius: 8,
						padding: "10px 12px",
						outline: "none",
					}}
				/>
				{unit && (
					<span
						style={{ fontSize: 12, color: C.muted, fontWeight: 600, width: 58 }}
					>
						{unit}
					</span>
				)}
			</div>
		</div>
	);
}

function Row({ label, value, unit, strong, color }) {
	return (
		<div
			style={{
				display: "flex",
				justifyContent: "space-between",
				alignItems: "baseline",
				gap: 12,
				padding: "9px 0",
				borderBottom: `1px dashed ${C.line}`,
			}}
		>
			<span style={{ fontSize: 13, color: C.muted }}>{label}</span>
			<span
				style={{
					fontFamily: MONO,
					fontSize: strong ? 20 : 15,
					fontWeight: strong ? 700 : 600,
					color: color || C.ink,
					whiteSpace: "nowrap",
				}}
			>
				{value}{" "}
				{unit && (
					<span style={{ fontSize: 11, color: C.muted, fontWeight: 500 }}>
						{unit}
					</span>
				)}
			</span>
		</div>
	);
}

function Card({ title, children, right }) {
	return (
		<div
			style={{
				background: C.card,
				border: `1px solid ${C.line}`,
				borderRadius: 14,
				padding: 18,
				marginBottom: 16,
			}}
		>
			{title && (
				<div
					style={{
						display: "flex",
						justifyContent: "space-between",
						alignItems: "center",
						paddingBottom: 10,
						marginBottom: 6,
						borderBottom: `2px solid ${C.ink}`,
					}}
				>
					<span
						style={{
							fontSize: 12,
							fontWeight: 800,
							letterSpacing: "0.1em",
							textTransform: "uppercase",
							color: C.ink,
						}}
					>
						{title}
					</span>
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
		green: {
			background: C.green,
			color: "#fff",
			border: `1px solid ${C.green}`,
		},
		ghost: {
			background: "transparent",
			color: C.muted,
			border: `1px solid ${C.line}`,
		},
	}[kind];
	return (
		<button
			onClick={onClick}
			style={{
				...styles,
				borderRadius: 8,
				cursor: "pointer",
				fontWeight: 700,
				fontSize: small ? 12 : 14,
				padding: small ? "7px 12px" : "12px 16px",
				width: small ? "auto" : "100%",
			}}
		>
			{children}
		</button>
	);
}

// ── Selector de posición ─────────────────────────────────────────────────────
function PositionPicker({
	portfolio,
	ticker,
	setTicker,
	shares,
	setShares,
	avg,
	setAvg,
}) {
	const onPick = (t) => {
		setTicker(t);
		const p = portfolio.find((x) => x.t === t);
		if (p) {
			setShares(String(p.sh));
			setAvg(String(p.avg));
		} else {
			setShares("");
			setAvg("");
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
						width: "100%",
						fontFamily: MONO,
						fontSize: 16,
						fontWeight: 700,
						color: C.ink,
						background: C.blueSoft,
						border: `2px solid ${C.blue}`,
						borderRadius: 8,
						padding: "10px 12px",
						outline: "none",
						appearance: "none",
					}}
				>
					<option value="">— selecciona —</option>
					{portfolio.map((p) => (
						<option key={p.t} value={p.t}>
							{p.t}
							{p.n ? ` · ${p.n}` : ""}
						</option>
					))}
					<option value="OTRO">OTRO (manual)</option>
				</select>
			</div>
			<Input
				label="Acciones que tienes"
				value={shares}
				onChange={setShares}
				unit="acciones"
			/>
			<Input
				label="Precio promedio de compra"
				value={avg}
				onChange={setAvg}
				unit="USD"
			/>
			<Row
				label="Total invertido"
				value={fmt(num(shares) * num(avg))}
				unit="USD"
			/>
		</Card>
	);
}

// ── Calculadora de promedio ──────────────────────────────────────────────────
function Promedio({ pos }) {
	const [price, setPrice] = useState("");
	const [amount, setAmount] = useState("");
	const sh = num(pos.shares),
		avg = num(pos.avg),
		p = num(price),
		m = num(amount);
	const invested = sh * avg;
	const newShares = m / p;
	const totShares = sh + newShares;
	const totCap = invested + m;
	const newAvg = totCap / totShares;
	const delta = newAvg - avg;
	const pct = (delta / avg) * 100;
	const ready = isFinite(newAvg) && p > 0 && m > 0 && sh > 0 && avg > 0;
	const lower = ready && delta < 0;

	return (
		<>
			<Card title="Nueva compra">
				<Input
					label="Precio actual de la acción"
					value={price}
					onChange={setPrice}
					unit="USD"
				/>
				<Input
					label="Monto que quieres invertir"
					value={amount}
					onChange={setAmount}
					unit="USD"
				/>
				<Row
					label="Acciones que comprarías"
					value={fmt(newShares, 4)}
					unit="acciones"
				/>
			</Card>

			<Card title="Después de la compra">
				<Row
					label="Total acciones acumuladas"
					value={fmt(totShares, 4)}
					unit="acciones"
				/>
				<Row label="Capital total invertido" value={fmt(totCap)} unit="USD" />
				<Row
					label="Nuevo precio promedio"
					value={fmt(newAvg, 4)}
					unit="USD"
					strong
					color={ready ? (lower ? C.green : C.red) : C.ink}
				/>
				{ready && (
					<div
						style={{
							marginTop: 14,
							padding: 14,
							borderRadius: 10,
							background: lower ? C.greenSoft : C.redSoft,
							border: `1px solid ${lower ? C.green : C.red}`,
						}}
					>
						<div
							style={{
								fontSize: 13,
								fontWeight: 800,
								color: lower ? C.green : C.red,
								marginBottom: 6,
							}}
						>
							{lower
								? "🟢 Compraste más barato — bajas tu promedio"
								: "🔴 Compraste más caro — subes tu promedio"}
						</div>
						<div style={{ fontFamily: MONO, fontSize: 15, color: C.ink }}>
							{fmtSigned(delta, 4)} USD/acción · {fmtSigned(pct, 2)}%
						</div>
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
	const sh = num(pos.shares),
		avg = num(pos.avg),
		cur = num(current),
		tgt = num(target);
	const invested = sh * avg;
	const gainVsCurrent = ((tgt - cur) / cur) * 100;
	const gainVsAvg = ((tgt - avg) / avg) * 100;
	const posValue = sh * tgt;
	const totalGain = posValue - invested;
	const ready = isFinite(totalGain) && sh > 0 && avg > 0 && tgt > 0;

	const levels = useMemo(() => {
		if (!(sh > 0 && avg > 0)) return [];
		return [20, 30, 40, 50, 60, 70, 80, 90, 100].map((t) => {
			const sell = avg * (1 + t / 100);
			const gain = invested * (t / 100);
			const toSell = gain / sell;
			const basis = (sh - toSell) * avg;
			return { t, sell, gain, toSell, basis };
		});
	}, [sh, avg, invested]);

	const th = (h) => (
		<th
			key={h}
			style={{
				textAlign: "right",
				fontSize: 10,
				textTransform: "uppercase",
				letterSpacing: "0.05em",
				color: C.muted,
				padding: "6px 8px",
				borderBottom: `2px solid ${C.ink}`,
				whiteSpace: "nowrap",
			}}
		>
			{h}
		</th>
	);
	const td = (v, extra = {}) => (
		<td
			style={{
				padding: "7px 8px",
				textAlign: "right",
				borderBottom: `1px dashed ${C.line}`,
				...extra,
			}}
		>
			{v}
		</td>
	);

	return (
		<>
			<Card title="Precio objetivo de venta">
				<Input
					label="Precio actual de la acción"
					value={current}
					onChange={setCurrent}
					unit="USD"
				/>
				<Input
					label="Precio al que deseas vender"
					value={target}
					onChange={setTarget}
					unit="USD"
				/>
			</Card>

			<Card title="Resultados esperados">
				<Row
					label="📈 Ganancia vs precio actual"
					value={fmtSigned(gainVsCurrent)}
					unit="%"
					color={gainVsCurrent >= 0 ? C.green : C.red}
				/>
				<Row
					label="💵 Ganancia sobre tu inversión"
					value={fmtSigned(gainVsAvg)}
					unit="%"
					color={gainVsAvg >= 0 ? C.green : C.red}
				/>
				<Row
					label="🏦 Valor esperado de tu posición"
					value={fmt(posValue)}
					unit="USD"
				/>
				<Row
					label="💰 Ganancia total esperada"
					value={fmtSigned(totalGain)}
					unit="USD"
					strong
					color={ready ? (totalGain >= 0 ? C.green : C.red) : C.ink}
				/>
			</Card>

			{levels.length > 0 && (
				<Card title="🎯 Niveles de ganancia objetivo">
					<div style={{ overflowX: "auto" }}>
						<table
							style={{
								width: "100%",
								borderCollapse: "collapse",
								fontFamily: MONO,
								fontSize: 13,
							}}
						>
							<thead>
								<tr>
									{[
										"Meta",
										"Precio venta",
										"Ganancia USD",
										"Acciones a vender",
										"Nuevo cost basis",
									].map(th)}
								</tr>
							</thead>
							<tbody>
								{levels.map((l) => (
									<tr key={l.t}>
										{td(`${l.t}%`, { fontWeight: 700, color: C.blue })}
										{td(fmt(l.sell, 4))}
										{td(fmt(l.gain), { color: C.green, fontWeight: 600 })}
										{td(fmt(l.toSell, 4))}
										{td(fmt(l.basis))}
									</tr>
								))}
							</tbody>
						</table>
					</div>
					<div style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>
						💡 "Acciones a vender" = las necesarias para retirar solo la
						ganancia, conservando el resto de tu posición. Los precios se
						calculan sobre tu promedio de compra.
					</div>
				</Card>
			)}
		</>
	);
}

// ── Portafolio + carga de Excel ──────────────────────────────────────────────
function Portafolio({ portfolio, onLoad, onFileParsed, status }) {
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
			<Card title="📥 Actualizar desde Excel">
				<div style={{ fontSize: 13, color: C.muted, marginBottom: 12 }}>
					Sube el <b style={{ color: C.ink }}>CSV de IBKR</b> (Activity
					Statement o reporte de posiciones). La app lee la sección{" "}
					<i>Open Positions</i>: símbolo, cantidad, Cost Price y Cost Basis.
					También acepta Excel con columnas Ticker / Acciones / Precio Promedio.
					Después de leerlo te preguntaré si quieres{" "}
					<b style={{ color: C.ink }}>reemplazar</b> los datos actuales o{" "}
					<b style={{ color: C.ink }}>añadirlos</b> a los que ya están.
				</div>
				<input
					ref={fileRef}
					type="file"
					accept=".xlsx,.xls,.xlsm,.csv"
					onChange={handleFile}
					style={{ display: "none" }}
				/>
				<Btn kind="blue" onClick={() => fileRef.current?.click()}>
					📂 Subir CSV de IBKR o Excel
				</Btn>
				{status && (
					<div
						style={{
							marginTop: 10,
							fontSize: 12,
							fontFamily: MONO,
							color: status.ok ? C.green : C.red,
						}}
					>
						{status.msg}
					</div>
				)}
			</Card>

			<Card title="📊 Portafolio unificado">
				{portfolio.length === 0 ? (
					<div
						style={{
							padding: "28px 0",
							textAlign: "center",
							fontSize: 13,
							color: C.muted,
						}}
					>
						Sin posiciones cargadas — sube tu CSV de IBKR arriba.
					</div>
				) : (
				<>
				<div style={{ overflowX: "auto" }}>
					<table
						style={{
							width: "100%",
							borderCollapse: "collapse",
							fontFamily: MONO,
							fontSize: 13,
						}}
					>
						<thead>
							<tr>
								{["Ticker", "Acciones", "Prom.", "Invertido", "Histórico"].map(
									(h, i) => (
										<th
											key={h}
											style={{
												textAlign: i === 0 ? "left" : "right",
												fontSize: 10,
												textTransform: "uppercase",
												letterSpacing: "0.05em",
												color: C.muted,
												padding: "6px 8px",
												borderBottom: `2px solid ${C.ink}`,
											}}
										>
											{h}
										</th>
									),
								)}
							</tr>
						</thead>
						<tbody>
							{portfolio.map((p) => (
								<tr
									key={p.t}
									onClick={() => onLoad(p.t)}
									style={{ cursor: "pointer" }}
									title={`${p.n} — toca para cargar en las calculadoras`}
								>
									<td
										style={{
											padding: "8px",
											fontWeight: 700,
											color: C.blue,
											borderBottom: `1px dashed ${C.line}`,
										}}
									>
										{p.t}
									</td>
									<td
										style={{
											padding: "8px",
											textAlign: "right",
											borderBottom: `1px dashed ${C.line}`,
										}}
									>
										{fmt(p.sh, 2)}
									</td>
									<td
										style={{
											padding: "8px",
											textAlign: "right",
											borderBottom: `1px dashed ${C.line}`,
										}}
									>
										{fmt(p.avg, 2)}
									</td>
									<td
										style={{
											padding: "8px",
											textAlign: "right",
											borderBottom: `1px dashed ${C.line}`,
										}}
									>
										{fmt(p.inv)}
									</td>
									<td
										style={{
											padding: "8px",
											textAlign: "right",
											color: C.muted,
											borderBottom: `1px dashed ${C.line}`,
										}}
									>
										{fmt(p.hist)}
									</td>
								</tr>
							))}
							<tr>
								<td style={{ padding: "10px 8px", fontWeight: 800 }}>TOTAL</td>
								<td
									style={{
										padding: "10px 8px",
										textAlign: "right",
										color: C.muted,
									}}
								>
									{portfolio.length} pos.
								</td>
								<td />
								<td
									style={{
										padding: "10px 8px",
										textAlign: "right",
										fontWeight: 800,
									}}
								>
									{fmt(totInv)}
								</td>
								<td
									style={{
										padding: "10px 8px",
										textAlign: "right",
										fontWeight: 800,
										color: C.muted,
									}}
								>
									{fmt(totHist)}
								</td>
							</tr>
						</tbody>
					</table>
				</div>
				<div style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>
					💡 Toca cualquier fila y elige si quieres simular una compra
					(Promedio) o una venta.
				</div>
				</>
				)}
			</Card>
		</>
	);
}

// ── Diálogo: ¿reemplazar o añadir? ───────────────────────────────────────────
function ImportDialog({ pending, currentCount, onReplace, onMerge, onCancel }) {
	const dupes = pending.dupes,
		news = pending.rows.length - dupes;
	return (
		<div
			style={{
				position: "fixed",
				inset: 0,
				background: "rgba(15,27,38,0.55)",
				display: "flex",
				alignItems: "center",
				justifyContent: "center",
				padding: 16,
				zIndex: 50,
			}}
		>
			<div
				style={{
					background: C.card,
					borderRadius: 16,
					padding: 22,
					maxWidth: 420,
					width: "100%",
					border: `1px solid ${C.line}`,
				}}
			>
				<div style={{ fontSize: 16, fontWeight: 800, marginBottom: 6 }}>
					📥 {pending.fileName}
				</div>
				<div
					style={{
						fontSize: 13,
						color: C.muted,
						marginBottom: 16,
						lineHeight: 1.5,
					}}
				>
					Encontré{" "}
					<b style={{ color: C.ink }}>{pending.rows.length} posiciones</b> en el
					archivo ({news} nuevas, {dupes} que ya existen). Tu portafolio actual
					tiene {currentCount} posiciones.
					<br />
					¿Qué quieres hacer?
				</div>
				<div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
					<Btn kind="blue" onClick={onReplace}>
						🔄 Reemplazar todo con el archivo
					</Btn>
					<Btn kind="green" onClick={onMerge}>
						➕ Añadir a lo que ya está
					</Btn>
					<Btn kind="ghost" onClick={onCancel}>
						Cancelar
					</Btn>
				</div>
				<div
					style={{
						fontSize: 11,
						color: C.muted,
						marginTop: 12,
						lineHeight: 1.5,
					}}
				>
					<b>Reemplazar:</b> borra la lista actual y usa solo el archivo.
					<br />
					<b>Añadir:</b> los tickers repetidos se combinan (acciones sumadas y
					promedio ponderado); los nuevos se agregan.
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
	const [pending, setPending] = useState(null); // { rows, fileName, dupes }
	const [status, setStatus] = useState(null);

	const onFileParsed = (rows, fileName) => {
		if (!rows) {
			setStatus({
				ok: false,
				msg: `✕ No pude leer "${fileName}". ¿Es un Excel válido?`,
			});
			return;
		}
		if (!rows.length) {
			setStatus({
				ok: false,
				msg: `✕ No encontré posiciones en "${fileName}". Verifica que tenga columnas Ticker / Acciones / Precio Promedio.`,
			});
			return;
		}
		const existing = new Set(portfolio.map((p) => p.t));
		const dupes = rows.filter((r) => existing.has(r.t)).length;
		setStatus(null);
		setPending({ rows, fileName, dupes });
	};

	const applyReplace = () => {
		setPortfolio([...pending.rows].sort((a, b) => a.t.localeCompare(b.t)));
		setStatus({
			ok: true,
			msg: `✓ Portafolio reemplazado: ${pending.rows.length} posiciones de "${pending.fileName}".`,
		});
		setPending(null);
		setTicker("");
		setShares("");
		setAvg("");
	};

	const applyMerge = () => {
		const merged = mergePortfolios(portfolio, pending.rows);
		setPortfolio(merged);
		setStatus({
			ok: true,
			msg: `✓ Datos añadidos: ahora tienes ${merged.length} posiciones (${pending.dupes} combinadas con promedio ponderado).`,
		});
		setPending(null);
		setTicker("");
		setShares("");
		setAvg("");
	};

	const [destPick, setDestPick] = useState(null); // ticker pendiente de elegir destino

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
			setTab(destino);
			window.scrollTo({ top: 0, behavior: "smooth" });
		}
		setDestPick(null);
	};

	const tabs = [
		{ id: "promedio", label: "📈 Promedio" },
		{ id: "venta", label: "💰 Venta" },
		{ id: "portafolio", label: "📊 Portafolio" },
	];

	return (
		<div
			style={{
				minHeight: "100vh",
				background: C.bg,
				color: C.ink,
				fontFamily: "system-ui, -apple-system, 'Segoe UI', sans-serif",
			}}
		>
			<div
				style={{ maxWidth: 640, margin: "0 auto", padding: "20px 14px 48px" }}
			>
				<div style={{ marginBottom: 18 }}>
					<div
						style={{
							fontFamily: MONO,
							fontSize: 11,
							fontWeight: 700,
							letterSpacing: "0.18em",
							color: C.muted,
						}}
					>
						COMPRA · VENTA · PORTAFOLIO
					</div>
					<h1
						style={{
							margin: "2px 0 0",
							fontSize: 26,
							fontWeight: 800,
							letterSpacing: "-0.02em",
						}}
					>
						Calculadora de inversiones
					</h1>
					{ticker && (
						<div
							style={{
								marginTop: 6,
								display: "inline-block",
								fontFamily: MONO,
								fontSize: 13,
								fontWeight: 700,
								color: C.blue,
								background: C.blueSoft,
								border: `1px solid ${C.blue}`,
								borderRadius: 6,
								padding: "3px 10px",
							}}
						>
							{ticker} cargado
						</div>
					)}
				</div>

				<div
					style={{
						display: "flex",
						gap: 6,
						marginBottom: 18,
						background: C.card,
						border: `1px solid ${C.line}`,
						borderRadius: 12,
						padding: 5,
					}}
				>
					{tabs.map((t) => (
						<button
							key={t.id}
							onClick={() => setTab(t.id)}
							style={{
								flex: 1,
								padding: "10px 4px",
								fontSize: 13,
								fontWeight: 700,
								border: "none",
								borderRadius: 8,
								cursor: "pointer",
								background: tab === t.id ? C.ink : "transparent",
								color: tab === t.id ? "#fff" : C.muted,
								transition: "background .15s",
							}}
						>
							{t.label}
						</button>
					))}
				</div>

				<>
					{tab !== "portafolio" && (
						<PositionPicker
							portfolio={portfolio}
							ticker={ticker}
							setTicker={setTicker}
							shares={shares}
							setShares={setShares}
							avg={avg}
							setAvg={setAvg}
						/>
					)}
					{tab === "promedio" && <Promedio pos={{ shares, avg }} />}
					{tab === "venta" && <Venta pos={{ shares, avg }} />}
					{tab === "portafolio" && (
						<Portafolio
							portfolio={portfolio}
							onLoad={loadFromPortfolio}
							onFileParsed={onFileParsed}
							status={status}
						/>
					)}
				</>

				<div
					style={{
						fontSize: 11,
						color: C.muted,
						textAlign: "center",
						marginTop: 8,
					}}
				>
					💡 Escribe en los campos con borde azul · todo lo demás se calcula
					solo.
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
					style={{
						position: "fixed",
						inset: 0,
						background: "rgba(15,27,38,0.55)",
						display: "flex",
						alignItems: "center",
						justifyContent: "center",
						padding: 16,
						zIndex: 50,
					}}
				>
					<div
						onClick={(e) => e.stopPropagation()}
						style={{
							background: C.card,
							borderRadius: 16,
							padding: 22,
							maxWidth: 360,
							width: "100%",
							border: `1px solid ${C.line}`,
						}}
					>
						<div
							style={{
								fontFamily: MONO,
								fontSize: 18,
								fontWeight: 800,
								color: C.blue,
								marginBottom: 4,
							}}
						>
							{destPick}
						</div>
						<div style={{ fontSize: 13, color: C.muted, marginBottom: 16 }}>
							{portfolio.find((x) => x.t === destPick)?.n || ""}
							<br />
							¿Qué quieres calcular con esta posición?
						</div>
						<div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
							<Btn kind="blue" onClick={() => goTo(destPick, "promedio")}>
								📈 Promedio — Simular una compra
							</Btn>
							<Btn kind="green" onClick={() => goTo(destPick, "venta")}>
								💰 Venta — <Samp></Samp>imular una venta
							</Btn>
							<Btn kind="ghost" onClick={() => setDestPick(null)}>
								Cancelar
							</Btn>
						</div>
					</div>
				</div>
			)}
		</div>
	);
}
