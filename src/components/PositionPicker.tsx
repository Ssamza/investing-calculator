import { C, MONO } from "../lib/constants.ts";
import { num, fmt } from "../lib/utils.ts";
import { Label, Input, Row, Card } from "./ui.tsx";
import type { Position } from "../lib/types.ts";

interface PositionPickerProps {
  portfolio: Position[];
  ticker: string;
  setTicker: (t: string) => void;
  shares: string;
  setShares: (s: string) => void;
  avg: string;
  setAvg: (a: string) => void;
  setCef: (c: number | null) => void;
}

export default function PositionPicker({ portfolio, ticker, setTicker, shares, setShares, avg, setAvg, setCef }: PositionPickerProps) {
  const onPick = (t: string) => {
    setTicker(t);
    const p = portfolio.find((x) => x.t === t);
    if (p) {
      setShares(String(p.sh));
      setAvg(String(p.avg));
      setCef(p.cef ?? null);
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
