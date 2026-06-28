export interface Position {
  t: string;
  n: string;
  sh: number;
  avg: number;
  prom?: number;
  inv: number;
  hist: number;
  proceeds?: number;
  totalBought?: number;
  cef?: number;
}

export interface PosInput {
  shares: string;
  avg: string;
  cef?: number | null;
  prom?: number;
  hist?: number;
  totalBought?: number;
  proceeds?: number;
}

export interface PendingImport {
  rows: Position[];
  fileName: string;
  dupes: number;
}

export interface StatusMsg {
  ok: boolean;
  msg: string;
}
