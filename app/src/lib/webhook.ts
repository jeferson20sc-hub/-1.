import type { Movement } from "./types";

const WEBHOOK_KEY = "senai-cba::webhook-url";

export function getWebhookUrl(): string {
  if (typeof window === "undefined") return "";
  return window.localStorage.getItem(WEBHOOK_KEY) ?? "";
}

export function setWebhookUrl(url: string) {
  window.localStorage.setItem(WEBHOOK_KEY, url.trim());
}

export function buildPayload(m: Movement) {
  return {
    Hora: m.hora,
    Data: m.data,
    Operador: m.operador,
    Codigo: m.codigo,
    Descricao: m.descricao,
    Tipo: m.tipo,
    Quantidade: m.qtd,
    Saldo: m.saldo,
    EstoqueSeg: m.ES,
    PontoPedido: m.PP,
    Status: m.status,
    Origem: m.origem,
  };
}

export async function postMovement(
  url: string,
  movement: Movement,
): Promise<{ ok: true } | { ok: false; status?: number; error: string }> {
  try {
    const res = await fetch(url, {
      method: "POST",
      mode: "cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(buildPayload(movement)),
    });
    if (res.ok || res.status === 202) return { ok: true };
    return { ok: false, status: res.status, error: `HTTP ${res.status}` };
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return { ok: false, error: msg };
  }
}
