import { motion } from "framer-motion";
import {
  AlertTriangle,
  Boxes,
  CheckCircle2,
  ShoppingCart,
  TrendingUp,
} from "lucide-react";

import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { saldoOf, statusOf, statusLabel, summarize } from "@/lib/inventory";
import type { Movement, StockItem } from "@/lib/types";
import { cn, nf } from "@/lib/utils";

interface Props {
  items: StockItem[];
  movements: Movement[];
}

const ICONS = {
  total: Boxes,
  ok: CheckCircle2,
  pedir: ShoppingCart,
  crit: AlertTriangle,
};

export function DashboardView({ items, movements }: Props) {
  const { total, ok, pedir, crit, totalSaldo } = summarize(items);
  const last24h = movements.filter((m) => {
    const ts = Date.parse(m.isoTimestamp);
    return Number.isFinite(ts) && Date.now() - ts < 24 * 60 * 60 * 1000;
  }).length;

  const kpis = [
    {
      key: "total",
      label: "Itens monitorados",
      value: total,
      hint: `${nf(totalSaldo)} unidades em estoque`,
      tone: "primary" as const,
    },
    {
      key: "ok",
      label: "Estoque saudável",
      value: ok,
      hint: total > 0 ? `${Math.round((ok / total) * 100)}% dos itens` : "—",
      tone: "success" as const,
    },
    {
      key: "pedir",
      label: "Pedir agora",
      value: pedir,
      hint: "Abaixo do Ponto de Pedido",
      tone: "warning" as const,
    },
    {
      key: "crit",
      label: "Críticos",
      value: crit,
      hint: "Abaixo do Estoque de Segurança",
      tone: "destructive" as const,
    },
  ];

  const criticos = items
    .filter((i) => statusOf(i) === "crit")
    .sort((a, b) => saldoOf(a) - saldoOf(b))
    .slice(0, 5);

  const aPedir = items
    .filter((i) => statusOf(i) === "pedir")
    .sort((a, b) => saldoOf(a) - saldoOf(b))
    .slice(0, 5);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {kpis.map((k, i) => {
          const Icon = ICONS[k.key as keyof typeof ICONS];
          return (
            <motion.div
              key={k.key}
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: i * 0.05 }}
            >
              <Card className="overflow-hidden relative">
                <div
                  className={cn(
                    "absolute inset-0 opacity-[0.04]",
                    k.tone === "primary" && "bg-primary",
                    k.tone === "success" && "bg-success",
                    k.tone === "warning" && "bg-warning",
                    k.tone === "destructive" && "bg-destructive",
                  )}
                />
                <CardContent className="p-5 relative">
                  <div className="flex items-start justify-between">
                    <div>
                      <div className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                        {k.label}
                      </div>
                      <div className="text-4xl font-bold mt-2 tracking-tight">
                        {nf(k.value)}
                      </div>
                      <div className="text-xs text-muted-foreground mt-2">
                        {k.hint}
                      </div>
                    </div>
                    <div
                      className={cn(
                        "size-10 rounded-xl grid place-items-center",
                        k.tone === "primary" && "bg-primary/10 text-primary",
                        k.tone === "success" && "bg-success/10 text-success",
                        k.tone === "warning" && "bg-warning/10 text-warning",
                        k.tone === "destructive" &&
                          "bg-destructive/10 text-destructive",
                      )}
                    >
                      <Icon className="size-5" />
                    </div>
                  </div>
                </CardContent>
              </Card>
            </motion.div>
          );
        })}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-destructive">
              <AlertTriangle className="size-4" />
              Itens críticos
            </CardTitle>
          </CardHeader>
          <CardContent>
            {criticos.length === 0 ? (
              <p className="text-sm text-muted-foreground">
                Nenhum item crítico no momento.
              </p>
            ) : (
              <ul className="divide-y divide-border">
                {criticos.map((it) => (
                  <li
                    key={it.cod}
                    className="py-3 flex items-center justify-between gap-3"
                  >
                    <div className="min-w-0">
                      <div className="text-sm font-semibold truncate">
                        {it.cod} · {it.desc}
                      </div>
                      <div className="text-xs text-muted-foreground">
                        Saldo {nf(saldoOf(it))} · ES {nf(it.ES)} · PP {nf(it.PP)}
                      </div>
                    </div>
                    <Badge variant="destructive">{statusLabel("crit")}</Badge>
                  </li>
                ))}
              </ul>
            )}
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-warning">
              <ShoppingCart className="size-4" />
              Pedir agora
            </CardTitle>
          </CardHeader>
          <CardContent>
            {aPedir.length === 0 ? (
              <p className="text-sm text-muted-foreground">
                Todos os itens acima do Ponto de Pedido.
              </p>
            ) : (
              <ul className="divide-y divide-border">
                {aPedir.map((it) => (
                  <li
                    key={it.cod}
                    className="py-3 flex items-center justify-between gap-3"
                  >
                    <div className="min-w-0">
                      <div className="text-sm font-semibold truncate">
                        {it.cod} · {it.desc}
                      </div>
                      <div className="text-xs text-muted-foreground">
                        Saldo {nf(saldoOf(it))} · ES {nf(it.ES)} · PP {nf(it.PP)}
                      </div>
                    </div>
                    <Badge variant="warning">{statusLabel("pedir")}</Badge>
                  </li>
                ))}
              </ul>
            )}
          </CardContent>
        </Card>
      </div>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <TrendingUp className="size-4 text-primary" />
            Atividade nas últimas 24h
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-4xl font-bold">{nf(last24h)}</div>
          <p className="text-sm text-muted-foreground mt-1">
            movimentações registradas
          </p>
        </CardContent>
      </Card>
    </div>
  );
}
