import {
  type Column,
  type ColumnDef,
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getPaginationRowModel,
  getSortedRowModel,
  type SortingState,
  useReactTable,
} from "@tanstack/react-table";
import { ArrowUpDown, ChevronLeft, ChevronRight } from "lucide-react";
import { useMemo, useState } from "react";
import { Search } from "lucide-react";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { saldoOf, statusLabel, statusOf, statusTone } from "@/lib/inventory";
import type { StockItem } from "@/lib/types";
import { cn, nf } from "@/lib/utils";

interface Props {
  items: StockItem[];
}

type Row = StockItem & { saldo: number; st: ReturnType<typeof statusOf> };

const PAGE_SIZE = 15;

export function EstoqueView({ items }: Props) {
  const [sorting, setSorting] = useState<SortingState>([]);
  const [globalFilter, setGlobalFilter] = useState("");

  const data = useMemo<Row[]>(
    () =>
      items.map((i) => ({
        ...i,
        saldo: saldoOf(i),
        st: statusOf(i),
      })),
    [items],
  );

  const columns = useMemo<ColumnDef<Row>[]>(
    () => [
      {
        accessorKey: "cod",
        header: ({ column }) => (
          <SortHeader label="Código" column={column} />
        ),
        cell: ({ getValue }) => (
          <span className="font-mono font-semibold text-xs">{getValue<string>()}</span>
        ),
      },
      {
        accessorKey: "desc",
        header: ({ column }) => (
          <SortHeader label="Descrição" column={column} />
        ),
        cell: ({ row }) => (
          <div>
            <div className="font-medium text-sm">{row.original.desc}</div>
            <div className="text-xs text-muted-foreground">{row.original.cen}</div>
          </div>
        ),
      },
      {
        accessorKey: "entrada",
        header: ({ column }) => (
          <SortHeader label="Entrada" column={column} right />
        ),
        cell: ({ getValue }) => (
          <span className="font-mono text-success text-xs">{nf(getValue<number>())}</span>
        ),
        meta: { right: true },
      },
      {
        accessorKey: "saida",
        header: ({ column }) => (
          <SortHeader label="Saída" column={column} right />
        ),
        cell: ({ getValue }) => (
          <span className="font-mono text-destructive text-xs">{nf(getValue<number>())}</span>
        ),
        meta: { right: true },
      },
      {
        accessorKey: "saldo",
        header: ({ column }) => (
          <SortHeader label="Saldo" column={column} right />
        ),
        cell: ({ getValue }) => (
          <span className="font-mono font-bold text-xs">{nf(getValue<number>())}</span>
        ),
        meta: { right: true },
      },
      {
        accessorKey: "ES",
        header: ({ column }) => (
          <SortHeader label="ES" column={column} right />
        ),
        cell: ({ getValue }) => (
          <span className="font-mono text-muted-foreground text-xs">{nf(getValue<number>())}</span>
        ),
        meta: { right: true },
      },
      {
        accessorKey: "PP",
        header: ({ column }) => (
          <SortHeader label="PP" column={column} right />
        ),
        cell: ({ getValue }) => (
          <span className="font-mono text-muted-foreground text-xs">{nf(getValue<number>())}</span>
        ),
        meta: { right: true },
      },
      {
        accessorKey: "st",
        header: "Status",
        cell: ({ getValue }) => {
          const st = getValue<ReturnType<typeof statusOf>>();
          return <Badge variant={statusTone(st)}>{statusLabel(st)}</Badge>;
        },
        filterFn: "equals",
      },
    ],
    [],
  );

  const table = useReactTable({
    data,
    columns,
    state: { sorting, globalFilter },
    onSortingChange: setSorting,
    onGlobalFilterChange: setGlobalFilter,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
    initialState: { pagination: { pageSize: PAGE_SIZE } },
  });

  return (
    <Card>
      <CardHeader>
        <div className="flex flex-wrap items-center justify-between gap-3">
          <CardTitle>Controle de Estoque</CardTitle>
          <div className="relative w-full sm:w-72">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 size-4 text-muted-foreground" />
            <Input
              placeholder="Buscar código, descrição ou cenário…"
              value={globalFilter}
              onChange={(e) => setGlobalFilter(e.target.value)}
              className="pl-9"
            />
          </div>
        </div>
      </CardHeader>
      <CardContent className="p-0">
        <div className="overflow-x-auto scrollbar-thin">
          <table className="w-full text-sm">
            <thead>
              {table.getHeaderGroups().map((hg) => (
                <tr
                  key={hg.id}
                  className="text-left text-xs uppercase tracking-wider text-muted-foreground border-b border-border bg-muted/20"
                >
                  {hg.headers.map((h) => (
                    <th
                      key={h.id}
                      className={cn(
                        "px-4 py-3",
                        (h.column.columnDef.meta as { right?: boolean })?.right &&
                          "text-right",
                      )}
                    >
                      {h.isPlaceholder
                        ? null
                        : flexRender(h.column.columnDef.header, h.getContext())}
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
                    Nenhum item encontrado.
                  </td>
                </tr>
              ) : (
                table.getRowModel().rows.map((row) => (
                  <tr
                    key={row.id}
                    className={cn(
                      "border-b border-border last:border-0",
                      row.original.st === "crit" && "bg-destructive/[0.03]",
                      row.original.st === "pedir" && "bg-warning/[0.03]",
                    )}
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
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        {table.getPageCount() > 1 && (
          <div className="flex items-center justify-between px-4 py-3 border-t border-border">
            <span className="text-xs text-muted-foreground">
              {table.getFilteredRowModel().rows.length} item(s) ·{" "}
              Página {table.getState().pagination.pageIndex + 1} de{" "}
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
  );
}

function SortHeader({
  label,
  column,
  right,
}: {
  label: string;
  column: Column<Row, unknown>;
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
