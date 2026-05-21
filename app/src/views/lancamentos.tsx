import { format } from "date-fns";
import { ptBR } from "date-fns/locale";
import { motion } from "framer-motion";
import {
  CheckCircle2,
  CircleAlert,
  Cloud,
  CloudUpload,
  Loader2,
  Package,
  QrCode,
  ScanLine,
} from "lucide-react";
import { useMemo, useState } from "react";
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
}

const SYNC_TONE: Record<
  SyncStatus,
  { label: string; variant: "success" | "warning" | "destructive" | "muted"; icon: typeof Cloud }
> = {
  ok: { label: "Excel", variant: "success", icon: Cloud },
  sending: { label: "Enviando", variant: "warning", icon: CloudUpload },
  queued: { label: "Fila", variant: "warning", icon: Loader2 },
  error: { label: "Erro", variant: "destructive", icon: CircleAlert },
  local: { label: "Local", variant: "muted", icon: CheckCircle2 },
};

export function LancamentosView({
  items,
  movements,
  operator,
  onMovement,
  onOpenScanner,
}: Props) {
  const [code, setCode] = useState("");
  const [type, setType] = useState<MovementType | "">("");
  const [qty, setQty] = useState("");

  const selected = useMemo(
    () => items.find((i) => i.cod === code.trim()),
    [items, code],
  );

  const recent = movements.slice(0, 25);

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
              <QrCode className="size-4 text-primary" /> Últimos lançamentos
            </span>
            <Badge variant="outline">{movements.length}</Badge>
          </CardTitle>
        </CardHeader>
        <CardContent className="p-0">
          <div className="overflow-x-auto scrollbar-thin">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-muted-foreground uppercase tracking-wider border-b border-border bg-muted/20">
                  <th className="px-4 py-3">Quando</th>
                  <th className="px-4 py-3">Operador</th>
                  <th className="px-4 py-3">Item</th>
                  <th className="px-4 py-3">Tipo</th>
                  <th className="px-4 py-3 text-right">Qtd</th>
                  <th className="px-4 py-3 text-right">Saldo</th>
                  <th className="px-4 py-3">Excel</th>
                </tr>
              </thead>
              <tbody>
                {recent.length === 0 ? (
                  <tr>
                    <td
                      colSpan={7}
                      className="px-4 py-12 text-center text-muted-foreground text-sm"
                    >
                      Nenhum lançamento ainda. Use o formulário ao lado.
                    </td>
                  </tr>
                ) : (
                  recent.map((m) => {
                    const tone = SYNC_TONE[m._sync];
                    const Icon = tone.icon;
                    return (
                      <motion.tr
                        key={m.id}
                        layout
                        initial={{ opacity: 0, x: -8 }}
                        animate={{ opacity: 1, x: 0 }}
                        className="border-b border-border last:border-0"
                      >
                        <td className="px-4 py-3 text-xs">
                          <div className="font-semibold">{m.data}</div>
                          <div className="text-muted-foreground">{m.hora}</div>
                        </td>
                        <td className="px-4 py-3 text-xs">{m.operador}</td>
                        <td className="px-4 py-3">
                          <div className="font-mono text-xs font-semibold">
                            {m.codigo}
                          </div>
                          <div className="text-xs text-muted-foreground truncate max-w-[28ch]">
                            {m.descricao}
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <Badge
                            variant={
                              m.tipo === "Entrada" ? "success" : "destructive"
                            }
                          >
                            {m.tipo}
                          </Badge>
                        </td>
                        <td className="px-4 py-3 text-right font-mono">
                          {nf(m.qtd)}
                        </td>
                        <td className="px-4 py-3 text-right font-mono font-semibold">
                          {nf(m.saldo)}
                        </td>
                        <td className="px-4 py-3">
                          <Badge variant={tone.variant} title={m._error}>
                            <Icon
                              className={cn(
                                "size-3",
                                m._sync === "sending" && "animate-spin",
                                m._sync === "queued" && "animate-spin",
                              )}
                            />
                            {tone.label}
                          </Badge>
                        </td>
                      </motion.tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
