import { Check, Cloud, ExternalLink, Info } from "lucide-react";
import { useState } from "react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { setWebhookUrl } from "@/lib/webhook";

interface Props {
  initialUrl: string;
  onSaved: (url: string) => void;
  onClearMovements: () => void;
}

export function ConfigView({ initialUrl, onSaved, onClearMovements }: Props) {
  const [url, setUrl] = useState(initialUrl);

  function save() {
    const v = url.trim();
    setWebhookUrl(v);
    onSaved(v);
    toast.success(v ? "URL do Power Automate salva" : "Integração desativada");
  }

  return (
    <div className="grid grid-cols-1 lg:grid-cols-[1fr_360px] gap-6 max-w-5xl">
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Cloud className="size-4 text-primary" />
            Integração com Excel Online
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="space-y-2">
            <Label htmlFor="hook">URL do Power Automate (gatilho HTTP)</Label>
            <Input
              id="hook"
              placeholder="https://prod-XX.logic.azure.com/workflows/.../invoke?...sig=..."
              value={url}
              onChange={(e) => setUrl(e.target.value)}
              className="font-mono text-xs"
            />
            <p className="text-xs text-muted-foreground">
              A URL fica armazenada apenas no seu navegador (localStorage). Nada é
              comitado no repositório.
            </p>
          </div>
          <div className="flex gap-2">
            <Button onClick={save}>
              <Check className="size-4" /> Salvar
            </Button>
            <Button variant="outline" asChild>
              <a
                href="https://make.powerautomate.com"
                target="_blank"
                rel="noreferrer"
              >
                <ExternalLink className="size-4" /> Abrir Power Automate
              </a>
            </Button>
          </div>

          <div className="rounded-xl border border-border bg-muted/30 p-4 text-xs space-y-2">
            <div className="flex items-center gap-2 font-semibold">
              <Info className="size-4 text-primary" /> Estrutura esperada na
              tabela do Excel
            </div>
            <p className="text-muted-foreground">
              Aba <strong>Base de lancamentos</strong>, tabela{" "}
              <strong>Tabela</strong>, colunas:
            </p>
            <code className="block font-mono bg-background rounded-lg p-2 leading-relaxed">
              ID · Hora · Data · Operador · Codigo · Descricao · Tipo · Quantidade
              · Saldo · EstoqueSeg · PontoPedido · Status · Origem
            </code>
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="text-destructive">Zona de perigo</CardTitle>
        </CardHeader>
        <CardContent className="space-y-3">
          <p className="text-sm text-muted-foreground">
            Limpa o histórico local de movimentações. Os dados que já chegaram no
            Excel permanecem lá.
          </p>
          <Button
            variant="destructive"
            onClick={() => {
              if (confirm("Apagar histórico local de movimentações?"))
                onClearMovements();
            }}
          >
            Limpar histórico local
          </Button>
        </CardContent>
      </Card>
    </div>
  );
}
