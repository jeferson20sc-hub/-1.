import { useMemo, useState } from "react";
import { Search } from "lucide-react";

import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { saldoOf, statusLabel, statusOf, statusTone } from "@/lib/inventory";
import type { StockItem } from "@/lib/types";
import { cn, nf } from "@/lib/utils";

interface Props {
  items: StockItem[];
}

export function EstoqueView({ items }: Props) {
  const [query, setQuery] = useState("");
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return items;
    return items.filter(
      (i) =>
        i.cod.toLowerCase().includes(q) ||
        i.desc.toLowerCase().includes(q) ||
        i.cen.toLowerCase().includes(q),
    );
  }, [items, query]);

  return (
    <Card>
      <CardHeader>
        <div className="flex flex-wrap items-center justify-between gap-3">
          <CardTitle>Controle de Estoque</CardTitle>
          <div className="relative w-full sm:w-72">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 size-4 text-muted-foreground" />
            <Input
              placeholder="Buscar por código, descrição ou cenário"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              className="pl-9"
            />
          </div>
        </div>
      </CardHeader>
      <CardContent className="p-0">
        <div className="overflow-x-auto scrollbar-thin">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-xs uppercase tracking-wider text-muted-foreground border-b border-border bg-muted/20">
                <th className="px-4 py-3">Código</th>
                <th className="px-4 py-3">Descrição</th>
                <th className="px-4 py-3 text-right">Entrada</th>
                <th className="px-4 py-3 text-right">Saída</th>
                <th className="px-4 py-3 text-right">Saldo</th>
                <th className="px-4 py-3 text-right">ES</th>
                <th className="px-4 py-3 text-right">PP</th>
                <th className="px-4 py-3">Status</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((it) => {
                const st = statusOf(it);
                const tone = statusTone(st);
                return (
                  <tr
                    key={it.cod}
                    className={cn(
                      "border-b border-border last:border-0",
                      st === "crit" && "bg-destructive/[0.03]",
                      st === "pedir" && "bg-warning/[0.03]",
                    )}
                  >
                    <td className="px-4 py-3 font-mono font-semibold">{it.cod}</td>
                    <td className="px-4 py-3">
                      <div className="font-medium">{it.desc}</div>
                      <div className="text-xs text-muted-foreground">{it.cen}</div>
                    </td>
                    <td className="px-4 py-3 text-right font-mono text-success">
                      {nf(it.entrada)}
                    </td>
                    <td className="px-4 py-3 text-right font-mono text-destructive">
                      {nf(it.saida)}
                    </td>
                    <td className="px-4 py-3 text-right font-mono font-bold">
                      {nf(saldoOf(it))}
                    </td>
                    <td className="px-4 py-3 text-right font-mono text-muted-foreground">
                      {nf(it.ES)}
                    </td>
                    <td className="px-4 py-3 text-right font-mono text-muted-foreground">
                      {nf(it.PP)}
                    </td>
                    <td className="px-4 py-3">
                      <Badge variant={tone}>{statusLabel(st)}</Badge>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </CardContent>
    </Card>
  );
}
