import { useState } from "react";
import { C, MONO } from "../lib/constants.ts";
import { num, fmt, fmtSigned } from "../lib/utils.ts";
import { Input, Row, Card } from "./ui.tsx";
import type { PosInput } from "../lib/types.ts";

export default function SellBuySimulator({ pos }: { pos: PosInput }) {
  const [pct, setPct]        = useState("");
  const [sellPrice, setSell] = useState("");
  const [buyPrice, setBuy]   = useState("");

  const sh       = num(pos.shares);
  const avg      = num(pos.avg);
  const baseProm = (pos.prom != null && isFinite(pos.prom)) ? pos.prom : avg;
  const baseCef  = (pos.cef  != null && isFinite(pos.cef))  ? pos.cef  : avg;
  const histOrig = (pos.hist        != null && isFinite(pos.hist))        ? pos.hist        : sh * avg;
  const tbOrig   = (pos.totalBought != null && isFinite(pos.totalBought)) ? pos.totalBought : sh;

  const p = num(pct), sp = num(sellPrice), bp = num(buyPrice);

  const sharesToSell   = sh * (p / 100);
  const simProceeds    = sharesToSell * sp;
  const profit         = simProceeds - (sharesToSell * avg);
  const repurchaseCost = sharesToSell * bp;

  const newProm   = (histOrig + repurchaseCost) / (tbOrig + sharesToSell);
  const promDelta = newProm - baseProm;

  const newCef   = (baseCef * sh - simProceeds + repurchaseCost) / sh;
  const cefDelta = newCef - baseCef;
  const cefPct   = (cefDelta / baseCef) * 100;

  const ready  = sh > 0 && avg > 0 && p > 0 && p <= 100 && sp > 0 && bp > 0;
  const better = ready && newCef < baseCef;

  return (
    <Card title="🔄 Simulador: Venta + Recompra">
      <div style={{ fontSize: 12, color: C.muted, marginBottom: 14, lineHeight: 1.5 }}>
        Simula vender un % de tu posición a cierto precio y recomprar esas mismas acciones más barato.
        Ve en cuánto quedan tu <b style={{ color: C.ink }}>Promedio</b> y tu{" "}
        <b style={{ color: C.ink }}>Costo Efectivo Promedio</b> después del ciclo completo.
      </div>

      <Input label="% de la posición a vender" value={pct}       onChange={setPct}   unit="%" placeholder="ej. 30" />
      <Input label="Precio de venta"            value={sellPrice} onChange={setSell}  unit="USD" />
      <Input label="Precio de recompra"         value={buyPrice}  onChange={setBuy}   unit="USD" />

      {ready && (
        <>
          <div style={{ borderTop: `1px solid ${C.line}`, margin: "14px 0 10px" }} />
          <Row label="Acciones a vender y recomprar" value={fmt(sharesToSell, 4)} unit="acciones" />
          <Row label="Recibido de la venta"           value={fmt(simProceeds)}     unit="USD" />
          <Row label="Ganancia realizada en venta"    value={fmtSigned(profit)}    unit="USD" color={profit >= 0 ? C.green : C.red} />
          <Row label="Costo de la recompra"           value={fmt(repurchaseCost)}  unit="USD" />

          <div style={{ borderTop: `1px solid ${C.line}`, margin: "14px 0 10px" }} />

          <Row
            label={`Nuevo promedio histórico (antes: $${fmt(baseProm, 2)})`}
            value={fmt(newProm, 4)}
            unit="USD"
            strong
            color={promDelta < 0 ? C.green : promDelta > 0 ? C.red : C.muted}
          />

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
