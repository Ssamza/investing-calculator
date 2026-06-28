import type { ReactNode, CSSProperties } from "react";
import { C, MONO } from "../lib/constants.ts";

export function Label({ children }: { children: ReactNode }) {
  return (
    <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: C.muted, marginBottom: 4 }}>
      {children}
    </div>
  );
}

interface InputProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  unit?: string;
  placeholder?: string;
}

export function Input({ label, value, onChange, unit, placeholder }: InputProps) {
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

interface RowProps {
  label: string;
  value: string;
  unit?: string;
  strong?: boolean;
  color?: string;
}

export function Row({ label, value, unit, strong, color }: RowProps) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 12, padding: "9px 0", borderBottom: `1px dashed ${C.line}` }}>
      <span style={{ fontSize: 13, color: C.muted }}>{label}</span>
      <span style={{ fontFamily: MONO, fontSize: strong ? 20 : 15, fontWeight: strong ? 700 : 600, color: color || C.ink, whiteSpace: "nowrap" } as CSSProperties}>
        {value} {unit && <span style={{ fontSize: 11, color: C.muted, fontWeight: 500 }}>{unit}</span>}
      </span>
    </div>
  );
}

interface CardProps {
  title?: string;
  children: ReactNode;
  right?: ReactNode;
}

export function Card({ title, children, right }: CardProps) {
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

type BtnKind = "primary" | "blue" | "green" | "ghost";

interface BtnProps {
  children: ReactNode;
  onClick?: () => void;
  kind?: BtnKind;
  small?: boolean;
}

export function Btn({ children, onClick, kind = "primary", small }: BtnProps) {
  const styles: Record<BtnKind, CSSProperties> = {
    primary: { background: C.ink,          color: "#fff", border: `1px solid ${C.ink}` },
    blue:    { background: C.blue,         color: "#fff", border: `1px solid ${C.blue}` },
    green:   { background: C.green,        color: "#fff", border: `1px solid ${C.green}` },
    ghost:   { background: "transparent",  color: C.muted, border: `1px solid ${C.line}` },
  };
  return (
    <button
      onClick={onClick}
      style={{
        ...styles[kind], borderRadius: 8, cursor: "pointer", fontWeight: 700,
        fontSize: small ? 12 : 14, padding: small ? "7px 12px" : "12px 16px", width: small ? "auto" : "100%",
      }}
    >
      {children}
    </button>
  );
}
