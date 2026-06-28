import { useState } from "react";
import { C, MONO } from "../lib/constants.ts";
import { num, fmt, fmtSigned } from "../lib/utils.ts";
import { Input, Row, Card } from "./ui.tsx";
import type { PosInput } from "../lib/types.ts";

export default function AverageCalculator({ pos }: { pos: PosInput }) {
  const [price, setPrice]   = useState("");
  const [amount, setAmount] = useState("");
  const sh = num(pos.shares), avg = num(pos.avg), p = num(price), m = num(amount);
  const currentCef = (pos.cef != null && isFinite(pos.cef)) ? pos.cef : avg;
  const hasCef = pos.cef != null && isFinite(pos.cef) && Math.abs(pos.cef - avg) > 0.001;

  const invested  = sh * avg;
  const newShares = m / p;
  const totShares = sh + newShares;
  const totCap    = invested + m;
  const newAvg    = totCap / totShares;
  const delta     = newAvg - avg;
  const pct       = (delta / avg) * 100;
  const ready     = isFinite(newAvg) && p > 0 && m > 0 && sh > 0 && avg > 0;
  const lower     = ready && delta < 0;

  const newCef   = (currentCef * sh + m) / totShares;
  const cefDelta = newCef - currentCef;
  const cefPct   = (cefDelta / currentCef) * 100;
  const cefLower = ready && cefDelta < 0;

  return (
    <>
      <Card title="Nueva compra">
        <Input label="Precio actual de la acción" value={price}  onChange={setPrice}  unit="USD" />
        <Input label="Monto que quieres invertir"  value={amount} onChange={setAmount} unit="USD" />
        <Row label="Acciones que comprarías" value={fmt(newShares, 4)} unit="acciones" />
      </Card>

      <Card title="Después de la compra">
        <Row label="Total acciones acumuladas" value={fmt(totShares, 4)} unit="acciones" />
        <Row label="Capital total invertido"   value={fmt(totCap)}       unit="USD" />
        <Row label="Nuevo precio promedio"     value={fmt(newAvg, 4)}    unit="USD" strong color={ready ? (lower ? C.green : C.red) : C.ink} />
        {ready && hasCef && (
          <Row label="Nuevo costo efectivo promedio" value={fmt(newCef, 4)} unit="USD" strong color={cefLower ? C.green : C.red} />
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
