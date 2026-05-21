import * as XLSX from "xlsx";

import { saldoOf, statusLabel, statusOf } from "./inventory";
import type { Movement, StockItem } from "./types";

export function exportToExcel(items: StockItem[], movements: Movement[]) {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Lançamentos
  const lancRows = movements.map((m) => ({
    Data: m.data,
    Hora: m.hora,
    Operador: m.operador,
    Código: m.codigo,
    Descrição: m.descricao,
    Tipo: m.tipo,
    Quantidade: m.qtd,
    Saldo: m.saldo,
    "ES (Estoque Seg.)": m.ES,
    "PP (Ponto Pedido)": m.PP,
    Status: m.status,
    Origem: m.origem,
    "Sync Excel": m._sync === "ok" ? "Sim" : "Não",
  }));
  const wsLanc = XLSX.utils.json_to_sheet(lancRows);
  XLSX.utils.book_append_sheet(wb, wsLanc, "Lançamentos");

  // Sheet 2: Controle de Estoque
  const estoqueRows = items.map((i) => ({
    Código: i.cod,
    Descrição: i.desc,
    Cenário: i.cen,
    Entrada: i.entrada,
    Saída: i.saida,
    Saldo: saldoOf(i),
    "ES (Estoque Seg.)": i.ES,
    "PP (Ponto Pedido)": i.PP,
    Status: statusLabel(statusOf(i)),
  }));
  const wsEstoque = XLSX.utils.json_to_sheet(estoqueRows);
  XLSX.utils.book_append_sheet(wb, wsEstoque, "Controle de Estoque");

  // Sheet 3: Resumo / Dashboard
  const okCount = items.filter((i) => statusOf(i) === "ok").length;
  const pedirCount = items.filter((i) => statusOf(i) === "pedir").length;
  const critCount = items.filter((i) => statusOf(i) === "crit").length;
  const totalSaldo = items.reduce((s, i) => s + saldoOf(i), 0);
  const resumoRows = [
    { Indicador: "Total de itens monitorados", Valor: items.length },
    { Indicador: "Itens em estoque saudável (OK)", Valor: okCount },
    { Indicador: "Itens a pedir (abaixo do PP)", Valor: pedirCount },
    { Indicador: "Itens críticos (abaixo do ES)", Valor: critCount },
    { Indicador: "Saldo total em estoque", Valor: totalSaldo },
    { Indicador: "Total de lançamentos", Valor: movements.length },
    {
      Indicador: "Exportado em",
      Valor: new Date().toLocaleString("pt-BR"),
    },
  ];
  const wsResumo = XLSX.utils.json_to_sheet(resumoRows);
  XLSX.utils.book_append_sheet(wb, wsResumo, "Resumo");

  const filename = `Gestao_Estoque_LIS_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, filename);
}

export function printAsPdf() {
  window.print();
}
