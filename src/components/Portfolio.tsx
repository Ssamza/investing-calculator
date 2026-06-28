import { useRef } from "react";
import * as XLSX from "xlsx";
import { C, MONO } from "../lib/constants.ts";
import { fmt } from "../lib/utils.ts";
import { Card, Btn } from "./ui.tsx";
import { parseWorkbook } from "../lib/parsers.ts";
import type { Position, StatusMsg } from "../lib/types.ts";

interface PortfolioProps {
  portfolio: Position[];
  onLoad: (ticker: string) => void;
  onFileParsed: (rows: Position[] | null, fileName: string) => void;
  onClear: () => void;
  status: StatusMsg | null;
}

export default function Portfolio({ portfolio, onLoad, onFileParsed, onClear, status }: PortfolioProps) {
  const fileRef = useRef<HTMLInputElement>(null);
  const totInv  = portfolio.reduce((a, p) => a + p.inv,  0);
  const totHist = portfolio.reduce((a, p) => a + p.hist, 0);

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      let wb: XLSX.WorkBook;
      if (/\.csv$/i.test(file.name)) {
        const text = await file.text();
        wb = XLSX.read(text, { type: "string", raw: true });
      } else {
        const buf = await file.arrayBuffer();
        wb = XLSX.read(buf, { type: "array" });
      }
      onFileParsed(parseWorkbook(wb), file.name);
    } catch {
      onFileParsed(null, file.name);
    }
  };

  return (
    <>
      <Card title="📥 Actualizar desde Excel">
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
                        color: (p.cef ?? p.avg) < p.avg ? C.green : (p.cef ?? p.avg) > p.avg ? C.red : C.muted,
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
