import { useMemo, useState } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  ComposedChart,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { saldoOf, statusLabel, summarize } from "@/lib/inventory";
import type { Movement, StockItem } from "@/lib/types";
import { nf } from "@/lib/utils";

interface Props {
  items: StockItem[];
  movements: Movement[];
}

// Static palette that reads well on both themes
const C = {
  primary: "#3b82f6",
  success: "#22c55e",
  warning: "#f59e0b",
  destructive: "#ef4444",
  muted: "#94a3b8",
  purple: "#a855f7",
};

// ---------------------------------------------------------------------------
// Sub-charts
// ---------------------------------------------------------------------------

function MovsByDay({ movements }: { movements: Movement[] }) {
  const data = useMemo(() => {
    const map = new Map<string, { data: string; entradas: number; saidas: number }>();
    const cutoff = Date.now() - 30 * 24 * 60 * 60 * 1000;
    for (const m of movements) {
      if (Date.parse(m.isoTimestamp) < cutoff) continue;
      const key = m.data;
      const entry = map.get(key) ?? { data: key, entradas: 0, saidas: 0 };
      if (m.tipo === "Entrada") entry.entradas += m.qtd;
      else entry.saidas += m.qtd;
      map.set(key, entry);
    }
    return [...map.values()].sort((a, b) => a.data.localeCompare(b.data));
  }, [movements]);

  if (data.length === 0) {
    return (
      <p className="text-sm text-muted-foreground py-8 text-center">
        Nenhuma movimentação nos últimos 30 dias.
      </p>
    );
  }

  return (
    <ResponsiveContainer width="100%" height={220}>
      <BarChart data={data} margin={{ top: 4, right: 8, bottom: 4, left: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="rgba(128,128,128,0.15)" />
        <XAxis dataKey="data" tick={{ fontSize: 10 }} tickLine={false} />
        <YAxis tick={{ fontSize: 10 }} tickLine={false} axisLine={false} />
        <Tooltip
          contentStyle={{ borderRadius: 10, border: "1px solid rgba(128,128,128,0.2)", fontSize: 12 }}
          formatter={(v: number, name: string) => [nf(v), name === "entradas" ? "Entradas" : "Saídas"]}
        />
        <Bar dataKey="entradas" fill={C.success} radius={[4, 4, 0, 0]} maxBarSize={28} />
        <Bar dataKey="saidas" fill={C.destructive} radius={[4, 4, 0, 0]} maxBarSize={28} />
      </BarChart>
    </ResponsiveContainer>
  );
}

function ParetoChart({ items }: { items: StockItem[] }) {
  const data = useMemo(() => {
    const sorted = [...items]
      .map((i) => ({ cod: i.cod, desc: i.desc, saldo: saldoOf(i) }))
      .filter((i) => i.saldo > 0)
      .sort((a, b) => b.saldo - a.saldo);
    const total = sorted.reduce((s, i) => s + i.saldo, 0);
    let cum = 0;
    return sorted.map((i) => {
      cum += i.saldo;
      return { ...i, cumPct: total > 0 ? Math.round((cum / total) * 100) : 0 };
    });
  }, [items]);

  if (data.length === 0) {
    return (
      <p className="text-sm text-muted-foreground py-8 text-center">
        Nenhum item com saldo positivo.
      </p>
    );
  }

  return (
    <ResponsiveContainer width="100%" height={220}>
      <ComposedChart data={data} margin={{ top: 4, right: 32, bottom: 4, left: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="rgba(128,128,128,0.15)" />
        <XAxis dataKey="cod" tick={{ fontSize: 10 }} tickLine={false} />
        <YAxis yAxisId="left" tick={{ fontSize: 10 }} tickLine={false} axisLine={false} />
        <YAxis
          yAxisId="right"
          orientation="right"
          tick={{ fontSize: 10 }}
          tickLine={false}
          axisLine={false}
          tickFormatter={(v) => `${v}%`}
          domain={[0, 100]}
        />
        <Tooltip
          contentStyle={{ borderRadius: 10, border: "1px solid rgba(128,128,128,0.2)", fontSize: 12 }}
          formatter={(v: number, name: string) =>
            name === "cumPct" ? [`${v}%`, "Acumulado"] : [nf(v), "Saldo"]
          }
        />
        <Bar yAxisId="left" dataKey="saldo" fill={C.primary} radius={[4, 4, 0, 0]} maxBarSize={32} />
        <Line
          yAxisId="right"
          dataKey="cumPct"
          stroke={C.warning}
          strokeWidth={2}
          dot={{ r: 3, fill: C.warning }}
          type="monotone"
        />
      </ComposedChart>
    </ResponsiveContainer>
  );
}

function StatusPie({ items }: { items: StockItem[] }) {
  const { ok, pedir, crit } = summarize(items);
  const data = [
    { name: statusLabel("ok"), value: ok, color: C.success },
    { name: statusLabel("pedir"), value: pedir, color: C.warning },
    { name: statusLabel("crit"), value: crit, color: C.destructive },
  ].filter((d) => d.value > 0);

  if (data.length === 0) {
    return (
      <p className="text-sm text-muted-foreground py-8 text-center">
        Nenhum item cadastrado.
      </p>
    );
  }

  return (
    <ResponsiveContainer width="100%" height={220}>
      <PieChart>
        <Pie
          data={data}
          cx="50%"
          cy="50%"
          innerRadius={55}
          outerRadius={85}
          paddingAngle={3}
          dataKey="value"
          label={({ name, value }) => `${name}: ${value}`}
          labelLine={false}
        >
          {data.map((d) => (
            <Cell key={d.name} fill={d.color} />
          ))}
        </Pie>
        <Tooltip
          contentStyle={{ borderRadius: 10, border: "1px solid rgba(128,128,128,0.2)", fontSize: 12 }}
          formatter={(v: number) => [v, "Itens"]}
        />
      </PieChart>
    </ResponsiveContainer>
  );
}

function DenteDeSerra({
  items,
  movements,
}: {
  items: StockItem[];
  movements: Movement[];
}) {
  const [cod, setCod] = useState(items[0]?.cod ?? "");

  const data = useMemo(() => {
    const pts: { label: string; saldo: number }[] = [];
    const item = items.find((i) => i.cod === cod);
    if (!item) return pts;
    const filtered = movements
      .filter((m) => m.codigo === cod)
      .sort((a, b) => a.isoTimestamp.localeCompare(b.isoTimestamp));
    for (const m of filtered) {
      pts.push({ label: `${m.data} ${m.hora}`, saldo: m.saldo });
    }
    return pts;
  }, [cod, items, movements]);

  return (
    <div className="space-y-3">
      <div className="flex items-center gap-3 flex-wrap">
        <span className="text-xs text-muted-foreground font-medium">Item:</span>
        <Select value={cod} onValueChange={setCod}>
          <SelectTrigger className="w-56 h-8 text-xs">
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            {items.map((i) => (
              <SelectItem key={i.cod} value={i.cod} className="text-xs">
                {i.cod} · {i.desc}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>

      {data.length === 0 ? (
        <p className="text-sm text-muted-foreground py-6 text-center">
          Nenhuma movimentação registrada para este item.
        </p>
      ) : (
        <ResponsiveContainer width="100%" height={200}>
          <LineChart data={data} margin={{ top: 4, right: 8, bottom: 4, left: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="rgba(128,128,128,0.15)" />
            <XAxis dataKey="label" tick={{ fontSize: 9 }} tickLine={false} interval="preserveStartEnd" />
            <YAxis tick={{ fontSize: 10 }} tickLine={false} axisLine={false} />
            <Tooltip
              contentStyle={{ borderRadius: 10, border: "1px solid rgba(128,128,128,0.2)", fontSize: 12 }}
              formatter={(v: number) => [nf(v), "Saldo"]}
            />
            {/* ES reference line rendered as a dashed area via second line */}
            <Line dataKey="saldo" stroke={C.primary} strokeWidth={2} dot={{ r: 3 }} type="stepAfter" />
          </LineChart>
        </ResponsiveContainer>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Main view
// ---------------------------------------------------------------------------

export function GraficosView({ items, movements }: Props) {
  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <Card>
          <CardHeader>
            <CardTitle className="text-sm">Movimentações por dia (últimos 30 dias)</CardTitle>
          </CardHeader>
          <CardContent className="pt-0">
            <div className="flex gap-4 mb-3 text-xs">
              <span className="flex items-center gap-1.5">
                <span className="size-2.5 rounded-sm inline-block" style={{ background: C.success }} />
                Entradas
              </span>
              <span className="flex items-center gap-1.5">
                <span className="size-2.5 rounded-sm inline-block" style={{ background: C.destructive }} />
                Saídas
              </span>
            </div>
            <MovsByDay movements={movements} />
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle className="text-sm">Distribuição de status</CardTitle>
          </CardHeader>
          <CardContent className="pt-0">
            <StatusPie items={items} />
          </CardContent>
        </Card>
      </div>

      <Card>
        <CardHeader>
          <CardTitle className="text-sm">Pareto 80/20 — saldo por item</CardTitle>
        </CardHeader>
        <CardContent className="pt-0">
          <div className="flex gap-4 mb-3 text-xs">
            <span className="flex items-center gap-1.5">
              <span className="size-2.5 rounded-sm inline-block" style={{ background: C.primary }} />
              Saldo (eixo esq.)
            </span>
            <span className="flex items-center gap-1.5">
              <span className="size-2.5 rounded-full inline-block" style={{ background: C.warning }} />
              % acumulado (eixo dir.)
            </span>
          </div>
          <ParetoChart items={items} />
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="text-sm">Dente de serra — saldo ao longo do tempo</CardTitle>
        </CardHeader>
        <CardContent className="pt-0">
          <DenteDeSerra items={items} movements={movements} />
        </CardContent>
      </Card>
    </div>
  );
}
