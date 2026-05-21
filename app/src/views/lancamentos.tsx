import {
  type Column,
  type ColumnDef,
  flexRender,
  getCoreRowModel,
  getPaginationRowModel,
  getSortedRowModel,
  type SortingState,
  useReactTable,
} from "@tanstack/react-table";
import { format } from "date-fns";
import { ptBR } from "date-fns/locale";
import { motion } from "framer-motion";
import {
  ArrowUpDown,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  CircleAlert,
  Cloud,
  CloudUpload,
  Loader2,
  Package,
  QrCode,
  ScanLine,
} from "lucide-react";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { saldoOf, statusOf } from "@/lib/inventory";
import type { Movement, MovementType, StockItem, SyncStatus } from "@/lib/types";
import { cn, nf, uid } from "@/lib/utils";

interface Props {
  items: StockItem[];
  movements: Movement[];
  operator: string;
  onMovement: (m: Movement, applyTo: StockItem) => void;
  onOpenScanner: () => void;
  pendingCode?: string | null;
  onPendingCodeConsumed?: () => void;
}

const SYNC_TONE: Record<
  SyncStatus,
  {
    label: string;
    variant: "success" | "warning" | "destructive" | "muted";
    icon: typeof Cloud;
  }
> = {
  ok: { label: "Excel", variant: "success", icon: Cloud },
  sending: { label: "Enviando", variant: "warning", icon: CloudUpload },
  queued: { label: "Fila", variant: "warning", icon: Loader2 },
  error: { label: "Erro", variant: "destructive", icon: CircleAlert },
  local: { label: "Local", variant: "muted", icon: CheckCircle2 },
};

const PAGE_SIZE = 20;

export function LancamentosView({
  items,
  movements,
  operator,
  onMovement,
  onOpenScanner,
  pendingCode,
  onPendingCodeConsumed,
}: Props) {
  const [code, setCode] = useState("");
  const [type, setType] = useState<MovementType | "">("");
  const [qty, setQty] = useState("");
  const [sorting, setSorting] = useState<SortingState>([
    { id: "isoTimestamp", desc: true },
  ]);

  const selected = useMemo(
    () => items.find((i) => i.cod === code.trim()),
    [items, code],
  );

  // Apply scanned code from QrScanner modal (passed as prop to avoid race condition)
  useEffect(() => {
    if (pendingCode) {
      setCode(pendingCode);
      onPendingCodeConsumed?.();
    }
  }, [pendingCode, onPendingCodeConsumed]);

  const columns = useMemo<ColumnDef<Movement>[]>(
    () => [
      {
        accessorKey: "isoTimestamp",
        header: ({ column }) => <SortHeader label="Quando" column={column} />,
        cell: ({ row }) => (
          <div className="text-xs">
            <div className="font-semibold">{row.original.data}</div>
            <div className="text-muted-foreground">{row.original.hora}</div>
          </div>
        ),
      },
      {
        accessorKey: "operador",
        header: ({ column }) => (
          <SortHeader label="Operador" column={column} />
        ),
        cell: ({ getValue }) => (
          <span className="text-xs">{getValue<string>()}</span>
        ),
      },
      {
        accessorKey: "codigo",
        header: ({ column }) => <SortHeader label="Item" column={column} />,
        cell: ({ row }) => (
          <div>
            <div className="font-mono text-xs font-semibold">
              {row.original.codigo}
            </div>
            <div className="text-xs text-muted-foreground truncate max-w-[24ch]">
              {row.original.descricao}
            </div>
          </div>
        ),
      },
      {
        accessorKey: "tipo",
        header: "Tipo",
        cell: ({ getValue }) => (
          <Badge variant={getValue<string>() === "Entrada" ? "success" : "destructive"}>
            {getValue<string>()}
          </Badge>
        ),
      },
      {
        accessorKey: "qtd",
        header: ({ column }) => (
          <SortHeader label="Qtd" column={column} right />
        ),
        cell: ({ getValue }) => (
          <span className="font-mono text-xs">{nf(getValue<number>())}</span>
        ),
        meta: { right: true },
      },
      {
        accessorKey: "saldo",
        header: ({ column }) => (
          <SortHeader label="Saldo" column={column} right />
        ),
        cell: ({ getValue }) => (
          <span className="font-mono font-semibold text-xs">
            {nf(getValue<number>())}
          </span>
        ),
        meta: { right: true },
      },
      {
        accessorKey: "_sync",
        header: "Excel",
        cell: ({ row }) => {
          const tone = SYNC_TONE[row.original._sync];
          const Icon = tone.icon;
          return (
            <Badge variant={tone.variant} title={row.original._error}>
              <Icon
                className={cn(
                  "size-3",
                  (row.original._sync === "sending" ||
                    row.original._sync === "queued") &&
                    "animate-spin",
                )}
              />
              {tone.label}
            </Badge>
          );
        },
      },
    ],
    [],
  );

  const table = useReactTable({
    data: movements,
    columns,
    state: { sorting },
    onSortingChange: setSorting,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
    initialState: { pagination: { pageSize: PAGE_SIZE } },
  });

  function submit() {
    if (!selected) {
      toast.error("Código não encontrado", {
        description: "Confira o código do item ou use o scanner.",
      });
      return;
    }
    if (!type) {
      toast.error("Selecione Entrada ou Saída");
      return;
    }
    const q = Number(qty);
    if (!Number.isFinite(q) || q <= 0) {
      toast.error("Quantidade inválida");
      return;
    }

    const next: StockItem = {
      ...selected,
      entrada: selected.entrada + (type === "Entrada" ? q : 0),
      saida: selected.saida + (type === "Saída" ? q : 0),
    };
    const saldo = saldoOf(next);
    const st = statusOf(next);
    const now = new Date();
    const m: Movement = {
      id: uid(),
      data: format(now, "dd/MM/yyyy", { locale: ptBR }),
      hora: format(now, "HH:mm:ss"),
      isoTimestamp: now.toISOString(),
      operador: operator,
      codigo: next.cod,
      descricao: next.desc,
      tipo: type,
      qtd: q,
      saldo,
      ES: next.ES,
      PP: next.PP,
      status: st === "ok" ? "OK" : st === "pedir" ? "PEDIR AGORA" : "CRÍTICO",
      origem: "Web",
      _sync: "queued",
      _attempts: 0,
    };

    onMovement(m, next);
    toast.success(`${type} de ${nf(q)} em ${next.cod}`, {
      description: next.desc,
    });
    setQty("");
    setType("");
  }

  return (
    <div className="grid grid-cols-1 xl:grid-cols-[420px_1fr] gap-6">
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Package className="size-4 text-primary" /> Novo lançamento
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="space-y-2">
            <Label htmlFor="codigo">Código do item</Label>
            <div className="flex gap-2">
              <Input
                id="codigo"
                placeholder="Ex.: 001"
                value={code}
                onChange={(e) => setCode(e.target.value)}
                className="font-mono"
              />
              <Button
                type="button"
                variant="outline"
                size="icon"
                onClick={onOpenScanner}
                aria-label="Abrir scanner"
              >
                <ScanLine className="size-4" />
              </Button>
            </div>
            {selected && (
              <div className="rounded-xl border border-border bg-muted/30 px-3 py-2 text-xs">
                <div className="font-semibold">{selected.desc}</div>
                <div className="text-muted-foreground">
                  Saldo atual: <strong>{nf(saldoOf(selected))}</strong> · ES{" "}
                  {nf(selected.ES)} · PP {nf(selected.PP)}
                </div>
              </div>
            )}
            {!selected && code.length > 0 && (
              <p className="text-xs text-destructive">
                Código não cadastrado nos itens.
              </p>
            )}
          </div>

          <div className="space-y-2">
            <Label>Tipo de movimentação</Label>
            <div className="grid grid-cols-2 gap-2">
              {(["Entrada", "Saída"] as const).map((t) => (
                <button
                  key={t}
                  type="button"
                  onClick={() => setType(t)}
                  className={cn(
                    "h-11 rounded-xl border text-sm font-semibold transition-all",
                    type === t
                      ? t === "Entrada"
                        ? "border-success bg-success/10 text-success"
                        : "border-destructive bg-destructive/10 text-destructive"
                      : "border-border hover:bg-accent",
                  )}
                >
                  {t}
                </button>
              ))}
            </div>
          </div>

          <div className="space-y-2">
            <Label htmlFor="qty">Quantidade</Label>
            <Input
              id="qty"
              type="number"
              inputMode="numeric"
              min={1}
              placeholder="0"
              value={qty}
              onChange={(e) => setQty(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && submit()}
              className="text-base"
            />
          </div>

          <Button onClick={submit} size="lg" className="w-full">
            Registrar lançamento
          </Button>
        </CardContent>
      </Card>

      <Card className="overflow-hidden">
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <QrCode className="size-4 text-primary" /> Histórico de lançamentos
            </span>
            <Badge variant="outline">{movements.length}</Badge>
          </CardTitle>
        </CardHeader>
        <CardContent className="p-0">
          <div className="overflow-x-auto scrollbar-thin">
            <table className="w-full text-sm">
              <thead>
                {table.getHeaderGroups().map((hg) => (
                  <tr
                    key={hg.id}
                    className="text-left text-xs text-muted-foreground uppercase tracking-wider border-b border-border bg-muted/20"
                  >
                    {hg.headers.map((h) => (
                      <th
                        key={h.id}
                        className={cn(
                          "px-4 py-3",
                          (h.column.columnDef.meta as { right?: boolean })
                            ?.right && "text-right",
                        )}
                      >
                        {h.isPlaceholder
                          ? null
                          : flexRender(
                              h.column.columnDef.header,
                              h.getContext(),
                            )}
                      </th>
                    ))}
                  </tr>
                ))}
              </thead>
              <tbody>
                {table.getRowModel().rows.length === 0 ? (
                  <tr>
                    <td
                      colSpan={columns.length}
                      className="px-4 py-12 text-center text-muted-foreground text-sm"
                    >
                      Nenhum lançamento ainda. Use o formulário ao lado.
                    </td>
                  </tr>
                ) : (
                  table.getRowModel().rows.map((row) => (
                    <motion.tr
                      key={row.id}
                      layout
                      initial={{ opacity: 0, x: -8 }}
                      animate={{ opacity: 1, x: 0 }}
                      className="border-b border-border last:border-0"
                    >
                      {row.getVisibleCells().map((cell) => (
                        <td
                          key={cell.id}
                          className={cn(
                            "px-4 py-3",
                            (cell.column.columnDef.meta as { right?: boolean })
                              ?.right && "text-right",
                          )}
                        >
                          {flexRender(
                            cell.column.columnDef.cell,
                            cell.getContext(),
                          )}
                        </td>
                      ))}
                    </motion.tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination */}
          {table.getPageCount() > 1 && (
            <div className="flex items-center justify-between px-4 py-3 border-t border-border">
              <span className="text-xs text-muted-foreground">
                {movements.length} lançamento(s) · Pág.{" "}
                {table.getState().pagination.pageIndex + 1} /{" "}
                {table.getPageCount()}
              </span>
              <div className="flex items-center gap-1">
                <Button
                  variant="outline"
                  size="icon"
                  className="size-7"
                  onClick={() => table.previousPage()}
                  disabled={!table.getCanPreviousPage()}
                >
                  <ChevronLeft className="size-3.5" />
                </Button>
                <Button
                  variant="outline"
                  size="icon"
                  className="size-7"
                  onClick={() => table.nextPage()}
                  disabled={!table.getCanNextPage()}
                >
                  <ChevronRight className="size-3.5" />
                </Button>
              </div>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

function SortHeader({
  label,
  column,
  right,
}: {
  label: string;
  column: Column<Movement, unknown>;
  right?: boolean;
}) {
  return (
    <button
      onClick={column.getToggleSortingHandler()}
      className={cn(
        "flex items-center gap-1 hover:text-foreground transition-colors",
        right && "ml-auto",
      )}
    >
      {label}
      <ArrowUpDown className="size-3 opacity-50" />
      {column.getIsSorted() === "asc" && " ↑"}
      {column.getIsSorted() === "desc" && " ↓"}
    </button>
  );
}
