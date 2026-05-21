import { motion } from "framer-motion";
import { ArrowRight, UserRound } from "lucide-react";
import { useState, type FormEvent } from "react";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";

interface Props {
  recents: string[];
  onEnter: (name: string) => void;
}

export function OperatorGate({ recents, onEnter }: Props) {
  const [value, setValue] = useState("");

  function submit(e: FormEvent) {
    e.preventDefault();
    const v = value.trim();
    if (!v) return;
    onEnter(v);
  }

  return (
    <div className="min-h-screen gradient-mesh grid place-items-center px-4">
      <motion.div
        initial={{ opacity: 0, y: 24 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.4, ease: "easeOut" }}
        className="glass w-full max-w-md rounded-3xl p-8 shadow-2xl shadow-primary/10"
      >
        <div className="flex items-center justify-center mb-6">
          <div className="size-16 rounded-2xl bg-gradient-to-br from-primary to-primary/60 grid place-items-center shadow-lg shadow-primary/30">
            <UserRound className="size-8 text-primary-foreground" strokeWidth={2.4} />
          </div>
        </div>
        <h1 className="text-2xl font-bold text-center tracking-tight">
          Identifique-se
        </h1>
        <p className="text-sm text-muted-foreground text-center mt-2 mb-7">
          Digite seu nome para começar a registrar lançamentos no estoque.
        </p>

        <form onSubmit={submit} className="space-y-3">
          <div className="space-y-2">
            <Label htmlFor="op-name">Nome do operador</Label>
            <Input
              id="op-name"
              autoFocus
              autoComplete="off"
              placeholder="Seu nome"
              value={value}
              onChange={(e) => setValue(e.target.value)}
              className="h-12 text-base"
            />
          </div>
          <Button type="submit" size="lg" className="w-full">
            Entrar <ArrowRight className="size-4" />
          </Button>
        </form>

        {recents.length > 0 && (
          <div className="mt-6">
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground mb-2">
              Operadores recentes
            </p>
            <div className="flex flex-wrap gap-2">
              {recents.slice(0, 6).map((n) => (
                <button
                  key={n}
                  type="button"
                  onClick={() => onEnter(n)}
                  className="px-3 py-1.5 rounded-full text-xs font-semibold bg-secondary text-secondary-foreground hover:bg-secondary/70 transition-colors"
                >
                  {n}
                </button>
              ))}
            </div>
          </div>
        )}
      </motion.div>
    </div>
  );
}
