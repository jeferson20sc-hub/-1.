# Como abrir a página e (opcional) conectar ao Excel online

## 1. Abrir sem precisar baixar (resolve o problema do Teams)

O Teams **sempre baixa** um arquivo `.html` anexado — isso é uma limitação do
próprio Teams, não tem conserto dentro do arquivo. A solução é mandar um **link**.

Passos (uma vez só):

1. No GitHub, abra o repositório → **Settings** → **Pages**.
2. Em **Build and deployment → Source**, escolha **Deploy from a branch**.
3. Em **Branch**, selecione a branch onde está o `index.html` (ou `main` após o
   merge do Pull Request) e a pasta **/(root)**. Clique **Save**.
4. Aguarde ~1 minuto. O GitHub mostra o link, no formato:
   `https://jeferson20sc-hub.github.io/zyth/`
5. **Mande esse link no Teams.** O professor clica e abre na hora, no navegador
   do PC ou do celular, **sem baixar nada**.

A página funciona 100% sozinha: login, animações, modo DEMO, planilha de
controle com 🟢🟡🔴, dashboard, gráficos por exercício e botão **Exportar para
Excel (.xlsx)**. Os lançamentos ficam salvos no aparelho de quem usa.

## 2. (Opcional) Cair direto numa planilha Excel online compartilhada

Por padrão **não precisa** disso — o botão "Exportar para Excel" já gera um
`.xlsx` no formato do professor. Mas se quiser que **todo lançamento de qualquer
aluno** caia automaticamente numa mesma planilha online, crie um webhook grátis
e cole a URL no `index.html`.

### Opção A — Google Apps Script (100% grátis, recomendado)

1. Crie uma planilha no Google Sheets (abas: `Lancamentos`).
2. Menu **Extensões → Apps Script** e cole:

   ```js
   function doPost(e){
     var d = JSON.parse(e.postData.contents);
     var sh = SpreadsheetApp.getActive().getSheetByName('Lancamentos');
     sh.appendRow([d.data, d.operador, d.codigo, d.descricao, d.tipo, d.qtd]);
     return ContentService.createTextOutput('ok');
   }
   ```
3. **Implantar → Nova implantação → App da Web**, acesso "Qualquer pessoa".
   Copie a URL gerada.
4. No `index.html`, troque a linha:
   `const WEBHOOK_URL = "";` → `const WEBHOOK_URL = "SUA_URL_AQUI";`
5. Pronto: cada "Salvar movimentação" envia a linha para a planilha online.

### Opção B — Power Automate (para Excel no OneDrive/SharePoint)

Use o gatilho **"When a HTTP request is received"** → ação **"Add a row into a
table"** apontando para a tabela do Excel online. Copie a URL do gatilho e cole
em `WEBHOOK_URL` igual ao passo 4 acima. (Esse gatilho HTTP exige um plano
Power Automate pago — por isso a Opção A é a recomendada e gratuita.)

> Importante: **nunca** coloque senha de planilha dentro do `index.html` —
> o código é público no navegador. A proteção do arquivo Excel do professor
> (senha) continua só no arquivo dele, fora desta página.
