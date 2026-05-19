# Conectar o site ao Excel online (OneDrive) via Power Automate

O site **https://jeferson20sc-hub.github.io/senai-cba/** já envia cada
lançamento para o Power Automate. Falta criar o fluxo que recebe esse
envio e grava na planilha **Gestao_Estoque_LIS.xlsx**.

## 1. Subir a planilha

1. Baixe o arquivo `Gestao_Estoque_LIS.xlsx` deste repositório.
2. Envie para a sua pasta no OneDrive:
   `Documents/Lis/Sistema_Estoque_LIS`.
3. A planilha já vem com **1 aba** (`Lancamentos`) e uma **Tabela**
   chamada `Lancamentos` com filtros ativados. Não renomeie a aba nem a
   tabela. A linha de exemplo (`LIS-EXEMPLO-001`) pode ser apagada depois
   do primeiro teste.

## 2. Criar o fluxo no Power Automate

1. Acesse **make.powerautomate.com** > **Criar** > **Fluxo de nuvem
   instantâneo**.
2. Gatilho: **Quando uma solicitação HTTP é recebida**.
3. No campo **Esquema JSON do corpo da solicitação**, cole:

```json
{
  "type": "object",
  "properties": {
    "id": { "type": "string" },
    "data": { "type": "string" },
    "dataBR": { "type": "string" },
    "operador": { "type": "string" },
    "codigo": { "type": "string" },
    "descricao": { "type": "string" },
    "tipo": { "type": "string" },
    "qtd": { "type": "number" },
    "saldo": { "type": "number" },
    "estoqueSeg": { "type": "number" },
    "pontoPedido": { "type": "number" },
    "status": { "type": "string" },
    "origem": { "type": "string" }
  }
}
```

> O site envia como `text/plain` (modo `no-cors`, para funcionar no
> GitHub Pages sem erro de CORS). Se algum campo vier vazio, adicione
> antes da próxima ação um passo **Analisar JSON** (Parse JSON) usando
> `triggerBody()` e o mesmo esquema acima.

## 3. Gravar na planilha

1. **Nova etapa** > Excel Online (Business) > **Adicionar uma linha a
   uma tabela**.
2. Preencha:
   - **Localização**: OneDrive for Business
   - **Biblioteca**: OneDrive
   - **Arquivo**: `.../Lis/Sistema_Estoque_LIS/Gestao_Estoque_LIS.xlsx`
   - **Tabela**: `Lancamentos`
3. Mapeie as colunas com os campos do gatilho:

| Coluna da tabela | Campo dinâmico |
|------------------|----------------|
| ID               | `id`           |
| DataHora         | `data`         |
| Data             | `dataBR`       |
| Operador         | `operador`     |
| Codigo           | `codigo`       |
| Descricao        | `descricao`    |
| Tipo             | `tipo`         |
| Quantidade       | `qtd`          |
| Saldo            | `saldo`        |
| EstoqueSeg       | `estoqueSeg`   |
| PontoPedido      | `pontoPedido`  |
| Status           | `status`       |
| Origem           | `origem`       |

## 4. Pegar a URL e conferir

1. **Salve** o fluxo. Volte ao gatilho HTTP e copie a **URL HTTP POST**
   gerada.
2. Ela já está configurada em `index.html` na constante `WEBHOOK_URL`.
   Se o Power Automate gerar uma URL diferente, substitua o valor dessa
   constante e publique de novo no GitHub Pages.
3. No site, abra o painel > aba **Lançamentos** > botão
   **🔌 Testar conexão com o Power Automate**. Deve aparecer uma linha
   com código `000` na aba `Lancamentos` da planilha.

Pronto: cada Entrada/Saída registrada no site cai automaticamente na
tabela do Excel online, com filtros prontos para análise.
