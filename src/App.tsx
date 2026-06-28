import { useState, useEffect, useRef } from "react";
import { STORAGE_KEY, C, MONO } from "./lib/constants.ts";
import { mergePortfolios } from "./lib/parsers.ts";
import { Card, Btn } from "./components/ui.tsx";
import PositionPicker from "./components/PositionPicker.tsx";
import AverageCalculator from "./components/AverageCalculator.tsx";
import SaleCalculator from "./components/SaleCalculator.tsx";
import SellBuySimulator from "./components/SellBuySimulator.tsx";
import Portfolio from "./components/Portfolio.tsx";
import ImportDialog from "./components/ImportDialog.tsx";
import type { Position, PendingImport, StatusMsg } from "./lib/types.ts";

type TabId = "average" | "sale" | "simulator" | "portfolio";

const TABS: { id: TabId; label: string }[] = [
  { id: "average",   label: "📈 Promedio" },
  { id: "sale",      label: "💰 Venta" },
  { id: "simulator", label: "🔄 Simulador" },
  { id: "portfolio", label: "📊 Portafolio" },
];

export default function App() {
  const [tab, setTab]             = useState<TabId>("average");
  const [portfolio, setPortfolio] = useState<Position[]>([]);
  const portfolioRef              = useRef<Position[]>([]);
  portfolioRef.current            = portfolio;
  const [ticker, setTicker]       = useState("");
  const [shares, setShares]       = useState("");
  const [avg, setAvg]             = useState("");
  const [cef, setCef]             = useState<number | null>(null);
  const [pending, setPending]     = useState<PendingImport | null>(null);
  const [status, setStatus]       = useState<StatusMsg | null>(null);
  const [loaded, setLoaded]       = useState(false);
  const [destPick, setDestPick]   = useState<string | null>(null);

  useEffect(() => {
    (async () => {
      try {
        if (window.storage) {
          const r = await window.storage.get(STORAGE_KEY);
          if (r?.value) {
            const data: unknown = JSON.parse(r.value);
            if (Array.isArray(data) && data.length) setPortfolio(data as Position[]);
          }
        }
      } catch { /* no saved data yet */ }
      setLoaded(true);
    })();
  }, []);

  const persist = async (data: Position[]) => {
    setPortfolio(data);
    try {
      if (window.storage) await window.storage.set(STORAGE_KEY, JSON.stringify(data));
    } catch { /* storage unavailable */ }
  };

  const onFileParsed = (rows: Position[] | null, fileName: string) => {
    if (!rows) {
      setStatus({ ok: false, msg: `✕ No pude leer "${fileName}". ¿Es un Excel válido?` });
      return;
    }
    if (!rows.length) {
      setStatus({ ok: false, msg: `✕ No encontré posiciones en "${fileName}". Verifica que tenga columnas Ticker / Acciones / Precio Promedio.` });
      return;
    }
    const existing = new Set(portfolioRef.current.map((p) => p.t));
    setStatus(null);
    setPending({ rows, fileName, dupes: rows.filter((r) => existing.has(r.t)).length });
  };

  const applyReplace = async () => {
    if (!pending) return;
    await persist([...pending.rows].sort((a, b) => a.t.localeCompare(b.t)));
    setStatus({ ok: true, msg: `✓ Portafolio reemplazado: ${pending.rows.length} posiciones de "${pending.fileName}".` });
    setPending(null); setTicker(""); setShares(""); setAvg("");
  };

  const applyMerge = async () => {
    if (!pending) return;
    const merged = mergePortfolios(portfolio, pending.rows);
    await persist(merged);
    const nuevas = pending.rows.length - pending.dupes;
    setStatus({ ok: true, msg: `✓ Portafolio actualizado: ${merged.length} posiciones (${pending.dupes} actualizadas, ${nuevas} nuevas).` });
    setPending(null); setTicker(""); setShares(""); setAvg("");
  };

  const onClear = async () => {
    await persist([]);
    setTicker(""); setShares(""); setAvg(""); setCef(null);
    setStatus({ ok: true, msg: "✓ Portafolio limpio." });
  };

  const goTo = (t: string, dest: TabId) => {
    const p = portfolio.find((x) => x.t === t);
    if (p) {
      setTicker(t); setShares(String(p.sh)); setAvg(String(p.avg)); setCef(p.cef ?? null);
      setTab(dest);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
    setDestPick(null);
  };

  const posFull = portfolio.find((p) => p.t === ticker);
  const posExtra = { prom: posFull?.prom, hist: posFull?.hist, totalBought: posFull?.totalBought, proceeds: posFull?.proceeds };

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
          {TABS.map((t) => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              flex: 1, padding: "10px 4px", fontSize: 13, fontWeight: 700, border: "none",
              borderRadius: 8, cursor: "pointer", transition: "background .15s",
              background: tab === t.id ? C.ink : "transparent",
              color:      tab === t.id ? "#fff" : C.muted,
            }}>
              {t.label}
            </button>
          ))}
        </div>

        {!loaded ? (
          <Card><div style={{ fontSize: 13, color: C.muted, textAlign: "center", padding: 10 }}>Cargando tu portafolio…</div></Card>
        ) : (
          <>
            {tab !== "portfolio" && (
              <PositionPicker
                portfolio={portfolio}
                ticker={ticker} setTicker={setTicker}
                shares={shares} setShares={setShares}
                avg={avg}       setAvg={setAvg}
                setCef={setCef}
              />
            )}
            {tab === "average"   && <AverageCalculator pos={{ shares, avg, cef }} />}
            {tab === "sale"      && <SaleCalculator     pos={{ shares, avg, cef, ...posExtra }} />}
            {tab === "simulator" && <SellBuySimulator   pos={{ shares, avg, cef, ...posExtra }} />}
            {tab === "portfolio" && (
              <Portfolio
                portfolio={portfolio}
                onLoad={(t) => { if (portfolio.find((x) => x.t === t)) setDestPick(t); }}
                onFileParsed={onFileParsed}
                onClear={onClear}
                status={status}
              />
            )}
          </>
        )}

        <div style={{ fontSize: 11, color: C.muted, textAlign: "center", marginTop: 8 }}>
          💡 Escribe en los campos con borde azul · todo lo demás se calcula solo.
          <br />Herramienta de cálculo, no asesoría financiera.
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
        <div onClick={() => setDestPick(null)} style={{ position: "fixed", inset: 0, background: "rgba(15,27,38,0.55)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 50 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.card, borderRadius: 16, padding: 22, maxWidth: 360, width: "100%", border: `1px solid ${C.line}` }}>
            <div style={{ fontFamily: MONO, fontSize: 18, fontWeight: 800, color: C.blue, marginBottom: 4 }}>{destPick}</div>
            <div style={{ fontSize: 13, color: C.muted, marginBottom: 16 }}>
              {portfolio.find((x) => x.t === destPick)?.n || ""}
              <br />¿Qué quieres calcular con esta posición?
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              <Btn kind="blue"  onClick={() => goTo(destPick, "average")}>📈 Promedio — Simular una compra</Btn>
              <Btn kind="green" onClick={() => goTo(destPick, "sale")}>💰 Venta — Simular una venta</Btn>
              <Btn kind="ghost" onClick={() => setDestPick(null)}>Cancelar</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
