import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export const nf = (n: number) =>
  Number.isFinite(n)
    ? n.toLocaleString("pt-BR", { maximumFractionDigits: 2 })
    : "0";

export const pct = (n: number) =>
  `${(Number.isFinite(n) ? n : 0).toLocaleString("pt-BR", {
    maximumFractionDigits: 0,
  })}%`;

export const uid = () =>
  `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
