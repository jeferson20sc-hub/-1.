import {
  ArrowLeftRight,
  BarChart3,
  Boxes,
  LayoutDashboard,
  QrCode,
  Settings,
  Sparkles,
} from "lucide-react";
import { motion } from "framer-motion";

import { cn } from "@/lib/utils";

export type View =
  | "dashboard"
  | "lancamentos"
  | "estoque"
  | "graficos"
  | "scanner"
  | "config";

interface Item {
  id: View;
  label: string;
  icon: typeof LayoutDashboard;
  hint?: string;
}

const ITEMS: Item[] = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "lancamentos", label: "Lançamentos", icon: ArrowLeftRight },
  { id: "estoque", label: "Estoque", icon: Boxes },
  { id: "graficos", label: "Gráficos", icon: BarChart3, hint: "Pareto, dente de serra" },
  { id: "scanner", label: "Scanner", icon: QrCode, hint: "QR / Código de barras" },
  { id: "config", label: "Configurações", icon: Settings },
];

interface Props {
  current: View;
  onChange: (v: View) => void;
  operator: string;
}

export function Sidebar({ current, onChange, operator }: Props) {
  return (
    <aside className="hidden lg:flex w-64 shrink-0 flex-col border-r border-border/60 bg-card/40 backdrop-blur-xl">
      <div className="flex items-center gap-3 px-5 h-16 border-b border-border/60">
        <div className="size-10 rounded-xl bg-gradient-to-br from-primary to-primary/60 grid place-items-center shadow-md shadow-primary/30">
          <Sparkles className="size-5 text-primary-foreground" strokeWidth={2.4} />
        </div>
        <div className="leading-tight">
          <div className="font-bold text-sm">SENAI · Estoque LIS</div>
          <div className="text-[11px] text-muted-foreground">
            Gestão Profissional
          </div>
        </div>
      </div>

      <nav className="flex-1 p-3 space-y-1 scrollbar-thin overflow-y-auto">
        {ITEMS.map((it) => {
          const active = current === it.id;
          const Icon = it.icon;
          return (
            <button
              key={it.id}
              onClick={() => onChange(it.id)}
              className={cn(
                "group relative w-full flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-semibold transition-colors",
                active
                  ? "text-primary-foreground"
                  : "text-muted-foreground hover:bg-accent hover:text-foreground",
              )}
            >
              {active && (
                <motion.div
                  layoutId="sidebar-active"
                  className="absolute inset-0 rounded-xl bg-gradient-to-r from-primary to-primary/80 shadow-lg shadow-primary/30"
                  transition={{ type: "spring", stiffness: 280, damping: 28 }}
                />
              )}
              <Icon className="size-4 relative z-10" strokeWidth={2.2} />
              <span className="relative z-10">{it.label}</span>
              {it.hint && (
                <span
                  className={cn(
                    "ml-auto relative z-10 text-[10px] uppercase tracking-wider",
                    active ? "text-primary-foreground/70" : "text-muted-foreground/70",
                  )}
                >
                  {it.hint}
                </span>
              )}
            </button>
          );
        })}
      </nav>

      <div className="p-3 border-t border-border/60">
        <div className="rounded-xl bg-muted/50 p-3 flex items-center gap-3">
          <div className="size-9 rounded-full bg-gradient-to-br from-success to-success/60 grid place-items-center text-success-foreground font-bold text-sm shadow">
            {operator.charAt(0).toUpperCase() || "?"}
          </div>
          <div className="min-w-0">
            <div className="text-xs font-semibold truncate">
              {operator || "Sem operador"}
            </div>
            <div className="text-[11px] text-muted-foreground">Operador ativo</div>
          </div>
        </div>
      </div>
    </aside>
  );
}
