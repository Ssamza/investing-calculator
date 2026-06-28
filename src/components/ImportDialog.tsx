import { C } from "../lib/constants.ts";
import { Btn } from "./ui.tsx";
import type { PendingImport } from "../lib/types.ts";

interface ImportDialogProps {
  pending: PendingImport;
  currentCount: number;
  onReplace: () => void;
  onMerge: () => void;
  onCancel: () => void;
}

export default function ImportDialog({ pending, currentCount, onReplace, onMerge, onCancel }: ImportDialogProps) {
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
          <Btn kind="blue"  onClick={onReplace}>🔄 Reemplazar todo con el archivo</Btn>
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
