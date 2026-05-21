export type StockStatus = "ok" | "pedir" | "crit";

export type StockItem = {
  cod: string;
  desc: string;
  cen: string;
  c: string;
  tr: string;
  kes: string;
  C: number;
  TR: number;
  ES: number;
  PP: number;
  entrada: number;
  saida: number;
  mem: string;
};

export type MovementType = "Entrada" | "Saída";

export type SyncStatus = "local" | "queued" | "sending" | "ok" | "error";

export type Movement = {
  id: string;
  data: string;
  hora: string;
  isoTimestamp: string;
  operador: string;
  codigo: string;
  descricao: string;
  tipo: MovementType;
  qtd: number;
  saldo: number;
  ES: number;
  PP: number;
  status: "OK" | "PEDIR AGORA" | "CRÍTICO";
  origem: string;
  _sync: SyncStatus;
  _attempts: number;
  _error?: string;
};
