export const num = (v: unknown): number => {
  if (typeof v === "number") return v;
  let s = String(v ?? "").trim().replace(/[$\s"]/g, "");
  if (s.startsWith("(") && s.endsWith(")")) s = "-" + s.slice(1, -1);
  if (s.includes(",") && s.includes(".")) s = s.replace(/,/g, "");
  else if (/^-?\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, "");
  else if (s.includes(",")) s = s.replace(",", ".");
  const n = parseFloat(s);
  return isFinite(n) ? n : NaN;
};

export const fmt = (v: number, d = 2): string =>
  !isFinite(v)
    ? "—"
    : v.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });

export const fmtSigned = (v: number, d = 2): string =>
  !isFinite(v) ? "—" : (v > 0 ? "+" : "") + fmt(v, d);
