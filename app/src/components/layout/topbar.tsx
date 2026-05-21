import {
  Cloud,
  CloudOff,
  FileDown,
  LogOut,
  Menu,
  Moon,
  Printer,
  RefreshCw,
  Sun,
} from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { useTheme } from "@/components/theme-provider";

interface Props {
  title: string;
  subtitle?: string;
  syncSummary: { pending: number; failed: number; ok: number };
  webhookConfigured: boolean;
  onRetryAll: () => void;
  onLogout: () => void;
  onMenuOpen: () => void;
  onExportExcel: () => void;
  onExportPdf: () => void;
}

export function Topbar({
  title,
  subtitle,
  syncSummary,
  webhookConfigured,
  onRetryAll,
  onLogout,
  onMenuOpen,
  onExportExcel,
  onExportPdf,
}: Props) {
  const { resolved, toggle } = useTheme();
  const Icon = resolved === "dark" ? Sun : Moon;
  const SyncIcon = webhookConfigured ? Cloud : CloudOff;

  return (
    <header className="h-16 shrink-0 flex items-center gap-2 lg:gap-4 border-b border-border/60 bg-background/60 backdrop-blur-xl px-3 lg:px-6 print:hidden">
      {/* Mobile hamburger */}
      <Button
        variant="ghost"
        size="icon"
        onClick={onMenuOpen}
        className="lg:hidden shrink-0"
        aria-label="Abrir menu"
      >
        <Menu className="size-5" />
      </Button>

      <div className="min-w-0 flex-1">
        <h1 className="text-base font-bold leading-tight truncate">{title}</h1>
        {subtitle && (
          <p className="hidden sm:block text-xs text-muted-foreground truncate">
            {subtitle}
          </p>
        )}
      </div>

      <div className="flex items-center gap-1 lg:gap-2">
        <Badge
          variant={webhookConfigured ? "success" : "muted"}
          className="hidden sm:inline-flex"
        >
          <SyncIcon className="size-3" />
          {webhookConfigured ? "Excel" : "Local"}
        </Badge>

        {syncSummary.pending > 0 && (
          <Badge variant="warning" className="hidden sm:inline-flex">
            {syncSummary.pending} enviando
          </Badge>
        )}
        {syncSummary.failed > 0 && (
          <Button variant="outline" size="sm" onClick={onRetryAll}>
            <RefreshCw className="size-3.5" />
            <span className="hidden sm:inline ml-1">Reenviar {syncSummary.failed}</span>
          </Button>
        )}

        <Button
          variant="ghost"
          size="icon"
          onClick={onExportExcel}
          aria-label="Exportar Excel"
          title="Exportar Excel (.xlsx)"
        >
          <FileDown className="size-4" />
        </Button>

        <Button
          variant="ghost"
          size="icon"
          onClick={onExportPdf}
          aria-label="Imprimir / PDF"
          title="Imprimir / Salvar como PDF"
        >
          <Printer className="size-4" />
        </Button>

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
