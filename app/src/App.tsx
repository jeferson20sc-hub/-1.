import { AnimatePresence, motion } from "framer-motion";
import { useEffect, useMemo, useState } from "react";
import { Toaster, toast } from "sonner";

import { OperatorGate } from "@/components/operator-gate";
import { Sidebar, type View } from "@/components/layout/sidebar";
import { Topbar } from "@/components/layout/topbar";
import { ComingSoonView } from "@/views/coming-soon";
import { ConfigView } from "@/views/config";
import { DashboardView } from "@/views/dashboard";
import { EstoqueView } from "@/views/estoque";
import { LancamentosView } from "@/views/lancamentos";
import { DEFAULT_ITEMS } from "@/lib/items";
import { useLocalState } from "@/lib/storage";
import { syncEngine } from "@/lib/sync";
import type { Movement, StockItem } from "@/lib/types";
import { getWebhookUrl } from "@/lib/webhook";

const VIEW_TITLES: Record<View, { title: string; subtitle: string }> = {
  dashboard: {
    title: "Dashboard",
    subtitle: "Visão geral do estoque em tempo real",
  },
  lancamentos: {
    title: "Lançamentos",
    subtitle: "Registre entradas e saídas — sincronia automática com Excel",
  },
  estoque: {
    title: "Controle de Estoque",
    subtitle: "Status de cada item monitorado",
  },
  graficos: { title: "Gráficos", subtitle: "Análises e tendências" },
  scanner: { title: "Scanner", subtitle: "Leitura de QR Code e código de barras" },
  config: { title: "Configurações", subtitle: "Integrações e preferências" },
};

export default function App() {
  const [operator, setOperator] = useLocalState<string>(
    "senai-cba::operator",
    "",
  );
  const [recentOps, setRecentOps] = useLocalState<string[]>(
    "senai-cba::operator-recents",
    [],
  );
  const [items, setItems] = useLocalState<StockItem[]>(
    "senai-cba::items",
    DEFAULT_ITEMS,
  );
  const [movements, setMovements] = useLocalState<Movement[]>(
    "senai-cba::movements",
    [],
  );
  const [view, setView] = useState<View>("dashboard");
  const [webhookUrl, setWebhookUrlState] = useState<string>(() =>
    getWebhookUrl(),
  );

  // Seed the sync engine with persisted state on mount, then mirror its
  // emissions back into React state. We intentionally don't depend on
  // `movements` here — the engine is the source of truth for live updates.
  useEffect(() => {
    syncEngine.setMovements(movements);
    const unsub = syncEngine.subscribe(setMovements);
    return () => {
      unsub();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const syncSummary = useMemo(() => {
    let pending = 0;
    let failed = 0;
    let ok = 0;
    for (const m of movements) {
      if (m._sync === "ok") ok++;
      else if (m._sync === "error") failed++;
      else if (m._sync === "queued" || m._sync === "sending") pending++;
    }
    return { pending, failed, ok };
  }, [movements]);

  if (!operator) {
    return (
      <>
        <OperatorGate
          recents={recentOps}
          onEnter={(name) => {
            setOperator(name);
            setRecentOps((prev) =>
              [name, ...prev.filter((p) => p !== name)].slice(0, 8),
            );
            toast.success(`Bem-vindo, ${name}`);
          }}
        />
        <Toaster richColors position="top-right" />
      </>
    );
  }

  const meta = VIEW_TITLES[view];

  function handleMovement(m: Movement, applied: StockItem) {
    setItems((prev) => prev.map((i) => (i.cod === applied.cod ? applied : i)));
    // enqueue() prepends to the engine's list and emits, which feeds setMovements
    syncEngine.enqueue(m);
  }

  return (
    <div className="min-h-screen flex bg-background">
      <Sidebar current={view} onChange={setView} operator={operator} />
      <main className="flex-1 flex flex-col min-w-0 gradient-mesh">
        <Topbar
          title={meta.title}
          subtitle={meta.subtitle}
          syncSummary={syncSummary}
          webhookConfigured={!!webhookUrl}
          onRetryAll={() => {
            syncEngine.retryAll();
            toast.info("Reenviando movimentações pendentes…");
          }}
          onLogout={() => {
            setOperator("");
          }}
        />

        <div className="flex-1 overflow-y-auto scrollbar-thin">
          <div className="p-4 lg:p-6 max-w-[1500px] mx-auto w-full">
            <AnimatePresence mode="wait">
              <motion.div
                key={view}
                initial={{ opacity: 0, y: 12 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -8 }}
                transition={{ duration: 0.2 }}
              >
                {view === "dashboard" && (
                  <DashboardView items={items} movements={movements} />
                )}
                {view === "lancamentos" && (
                  <LancamentosView
                    items={items}
                    movements={movements}
                    operator={operator}
                    onMovement={handleMovement}
                    onOpenScanner={() => setView("scanner")}
                  />
                )}
                {view === "estoque" && <EstoqueView items={items} />}
                {view === "graficos" && (
                  <ComingSoonView
                    title="Gráficos avançados"
                    description="Pareto 80/20, dente de serra por item e movimentações por dia chegam na próxima entrega."
                  />
                )}
                {view === "scanner" && (
                  <ComingSoonView
                    title="Scanner QR / Código de Barras"
                    description="Leitor html5-qrcode com câmera em tela cheia, troca frontal/traseira e upload de imagem chega na próxima entrega."
                  />
                )}
                {view === "config" && (
                  <ConfigView
                    initialUrl={webhookUrl}
                    onSaved={(u) => {
                      setWebhookUrlState(u);
                      // Triggers any queued/error items to flush to the new URL
                      syncEngine.run();
                    }}
                    onClearMovements={() => {
                      syncEngine.setMovements([]);
                      setMovements([]);
                      toast.success("Histórico local limpo");
                    }}
                  />
                )}
              </motion.div>
            </AnimatePresence>
          </div>
        </div>
      </main>
      <Toaster richColors position="top-right" />
    </div>
  );
}
