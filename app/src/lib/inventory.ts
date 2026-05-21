import type { StockItem, StockStatus } from "./types";

export const saldoOf = (i: StockItem) => i.entrada - i.saida;

export function statusOf(i: StockItem): StockStatus {
  const s = saldoOf(i);
  if (s <= i.ES) return "crit";
  if (s <= i.PP) return "pedir";
  return "ok";
}

export const statusLabel = (s: StockStatus) =>
  s === "ok" ? "OK" : s === "pedir" ? "PEDIR AGORA" : "CRÍTICO";

export const statusTone = (s: StockStatus) =>
  s === "ok"
    ? "success"
    : s === "pedir"
      ? "warning"
      : ("destructive" as const);

export function summarize(items: StockItem[]) {
  const total = items.length;
  const ok = items.filter((i) => statusOf(i) === "ok").length;
  const pedir = items.filter((i) => statusOf(i) === "pedir").length;
  const crit = items.filter((i) => statusOf(i) === "crit").length;
  const totalSaldo = items.reduce((acc, i) => acc + saldoOf(i), 0);
  return { total, ok, pedir, crit, totalSaldo };
}
