import { useState } from "react";
import { C, MONO } from "../lib/constants.ts";
import { num, fmt, fmtSigned } from "../lib/utils.ts";
import { Label, Input, Row, Card } from "./ui.tsx";
import type { PosInput } from "../lib/types.ts";

function ProfitTarget({ sh, avg, invested }: { sh: number; avg: number; invested: number }) {
  const [meta, setMeta] = useState("");
  const pct    = num(meta);
  const sell   = avg * (1 + pct / 100);
  const gain   = invested * (pct / 100);
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
          <Row label="Precio de venta objetivo" value={fmt(sell, 4)}   unit="USD" strong color={C.blue} />
          <Row label="Ganancia en USD"           value={fmt(gain)}      unit="USD" color={C.green} />
          <Row label="Acciones a vender"         value={fmt(toSell, 4)} unit="acciones" />
          <Row label="Nuevo cost basis"          value={fmt(basis)}     unit="USD" />
          <div style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>
            💡 Acciones a vender = las necesarias para retirar solo la ganancia, conservando el resto de tu posición.
          </div>
        </>
      )}
    </Card>
  );
}

export default function SaleCalculator({ pos }: { pos: PosInput }) {
  const [current, setCurrent] = useState("");
  const [target, setTarget]   = useState("");
  const sh = num(pos.shares), avg = num(pos.avg), cur = num(current), tgt = num(target);
  const invested      = sh * avg;
  const gainVsCurrent = ((tgt - cur) / cur) * 100;
  const gainVsAvg     = ((tgt - avg) / avg) * 100;
  const posValue      = sh * tgt;
  const totalGain     = posValue - invested;
  const ready         = isFinite(totalGain) && sh > 0 && avg > 0 && tgt > 0;

  return (
    <>
      <Card title="Precio objetivo de venta">
        <Input label="Precio actual de la acción"   value={current} onChange={setCurrent} unit="USD" />
        <Input label="Precio al que deseas vender"  value={target}  onChange={setTarget}  unit="USD" />
      </Card>
      <Card title="Resultados esperados">
        <Row label="📈 Ganancia vs precio actual"     value={fmtSigned(gainVsCurrent)} unit="%" color={gainVsCurrent >= 0 ? C.green : C.red} />
        <Row label="💵 Ganancia sobre tu inversión"   value={fmtSigned(gainVsAvg)}     unit="%" color={gainVsAvg >= 0 ? C.green : C.red} />
        <Row label="🏦 Valor esperado de tu posición" value={fmt(posValue)}             unit="USD" />
        <Row label="💰 Ganancia total esperada"       value={fmtSigned(totalGain)}      unit="USD" strong color={ready ? (totalGain >= 0 ? C.green : C.red) : C.ink} />
      </Card>
      {sh > 0 && avg > 0 && <ProfitTarget sh={sh} avg={avg} invested={invested} />}
    </>
  );
}
