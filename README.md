# SENAI · Gestão de Estoque LIS

Sistema de controle de estoque com sincronização para Excel Online via Power Automate.

- **Site público:** https://jeferson20sc-hub.github.io/senai-cba/
- **Stack do app:** Vite + React 18 + TypeScript + Tailwind + shadcn/ui + Framer Motion + Sonner
- **Offline-first:** funciona 100% local (localStorage) e sincroniza com Excel quando online
- **PWA:** instalável no celular e desktop

---

## Estrutura do repositório

```
/
├── index.html              # versão legada (HTML único) — preservada
├── app/                    # NOVO app moderno (Vite + React + TS)
│   ├── src/
│   ├── package.json
│   └── vite.config.ts
└── .github/workflows/deploy.yml   # CI/CD: build automático e deploy
```

## Como ativar o app novo (2 cliques)

1. **Faça merge do PR.**
2. No GitHub vá em **Settings → Pages** e mude o **Source** para **"GitHub Actions"**.

Pronto. A cada push em `app/**` o GitHub Actions vai buildar e publicar automaticamente.

## Conectar ao Excel Online (Power Automate)

1. Acesse https://make.powerautomate.com e crie um fluxo automatizado
2. Gatilho: **"Quando uma solicitação HTTP é recebida"**
3. Cole este schema JSON:
   ```json
   {
     "type": "object",
     "properties": {
       "Hora":         {"type": "string"},
       "Data":         {"type": "string"},
       "Operador":     {"type": "string"},
       "Codigo":       {"type": "string"},
       "Descricao":    {"type": "string"},
       "Tipo":         {"type": "string"},
       "Quantidade":   {"type": "number"},
       "Saldo":        {"type": "number"},
       "EstoqueSeg":   {"type": "number"},
       "PontoPedido":  {"type": "number"},
       "Status":       {"type": "string"},
       "Origem":       {"type": "string"}
     }
   }
   ```
4. Ação: **Excel Online (Business) → Adicionar uma linha em uma tabela**
   - Arquivo: `Gestao_Estoque_LIS.xlsx`
   - Tabela: `Tabela` (na aba `Base de lancamentos`)
   - Mapeie cada coluna com o conteúdo dinâmico correspondente do gatilho
5. **Salve** o fluxo e **copie a URL HTTP POST** gerada pelo gatilho
6. No app, vá em **Configurações → Integração com Excel Online**, cole a URL e salve

A URL fica armazenada **apenas no seu navegador** (localStorage). Nada é comitado.

## Desenvolvimento local (opcional)

Se algum dia quiser rodar localmente — precisa de Node 20+:

```bash
cd app
npm install
npm run dev
```

## Recursos do app

- **Sem login fixo** — o operador digita o próprio nome (ou escolhe da lista de recentes salva)
- **Dashboard** com KPIs, itens críticos e a pedir
- **Lançamentos** com formulário rápido, tabela em tempo real e status de sincronização Excel por linha (✔ Excel / ⏳ Enviando / ❌ Erro / ✔ Local)
- **Controle de estoque** completo com busca
- **Fila de sincronização offline** — registros feitos sem internet ficam em fila e sincronizam quando voltar online, com retry e backoff exponencial
- **Tema claro/escuro** automático ou manual
- **PWA** instalável (manifest + service worker)

### Em breve (próxima entrega)

- Scanner QR / código de barras com `html5-qrcode` (câmera em tela cheia, troca frontal/traseira, upload)
- Gráficos: Pareto 80/20, dente de serra, movimentações por dia (Recharts)
- Exportação Excel + PDF
