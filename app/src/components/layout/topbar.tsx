import { Cloud, CloudOff, LogOut, Moon, RefreshCw, Sun } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { useTheme } from "@/components/theme-provider";

interface Props {
  title: string;
  subtitle?: string;
  syncSummary: { pending: number; failed: number; ok: number };
  webhookConfigured: boolean;
  onRetryAll: () => void;
  onLogout: () => void;
}

export function Topbar({
  title,
  subtitle,
  syncSummary,
  webhookConfigured,
  onRetryAll,
  onLogout,
}: Props) {
  const { resolved, toggle } = useTheme();
  const Icon = resolved === "dark" ? Sun : Moon;
  const SyncIcon = webhookConfigured ? Cloud : CloudOff;

  return (
    <header className="h-16 shrink-0 flex items-center gap-4 border-b border-border/60 bg-background/60 backdrop-blur-xl px-4 lg:px-6">
      <div className="min-w-0 flex-1">
        <h1 className="text-base font-bold leading-tight truncate">{title}</h1>
        {subtitle && (
          <p className="text-xs text-muted-foreground truncate">{subtitle}</p>
        )}
      </div>

      <div className="flex items-center gap-2">
        <Badge variant={webhookConfigured ? "success" : "muted"}>
          <SyncIcon className="size-3" />
          {webhookConfigured ? "Excel conectado" : "Modo local"}
        </Badge>

        {syncSummary.pending > 0 && (
          <Badge variant="warning">{syncSummary.pending} sincronizando</Badge>
        )}
        {syncSummary.failed > 0 && (
          <Button variant="outline" size="sm" onClick={onRetryAll}>
            <RefreshCw className="size-3.5" />
            Reenviar {syncSummary.failed}
          </Button>
        )}

        <Button
          variant="ghost"
          size="icon"
          onClick={toggle}
          aria-label="Alternar tema"
        >
          <Icon className="size-4" />
        </Button>
        <Button
          variant="ghost"
          size="icon"
          onClick={onLogout}
          aria-label="Trocar operador"
        >
          <LogOut className="size-4" />
        </Button>
      </div>
    </header>
  );
}
