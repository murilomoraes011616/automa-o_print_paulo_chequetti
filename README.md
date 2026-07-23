# 📊 Automação Mapa de Faturamento → WhatsApp

Script em Python que automatiza a geração e o envio do relatório **Mapa de Faturamento** (KPI), eliminando o trabalho manual de atualizar a planilha, exportar o PDF e enviar para a gestora responsável via WhatsApp.

## ✅ O que o script faz

1. **Abre o Excel em segundo plano** (`xlwings`), sem exibir a interface e sem pop-ups de alerta.
2. **Atualiza a data de referência** na célula `X41` da aba `tabelas-auxiliares` com a data do dia.
3. **Atualiza todos os dados** da pasta de trabalho (`RefreshAll`), puxando as informações mais recentes do SAP HANA, e aguarda o tempo necessário para a atualização terminar.
4. **Configura a área de impressão** da aba `Indicador Faturamento-julho` (intervalo, orientação, zoom e ajuste de página) e **exporta a tabela como PDF**, salvando o arquivo com a data do dia no nome.
5. **Abre o WhatsApp Web** com uma sessão salva (via `Playwright`), localiza o contato configurado, anexa o PDF gerado e **envia a mensagem automaticamente**.
6. Em caso de erro, tira um print de tela para facilitar o diagnóstico e mantém o navegador aberto para investigação.

## 🛠️ Tecnologias utilizadas

- [`xlwings`](https://www.xlwings.org/) — manipulação do Excel (via COM/VBA por trás dos panos)
- [`playwright`](https://playwright.dev/python/) — automação do navegador para envio via WhatsApp Web
- `datetime` — geração da data formatada para o nome do arquivo
- `time` — controle de espera entre etapas

## 📂 Estrutura do fluxo

```
Excel (.xlsm ligado ao SAP HANA)
   └── xlwings
         ├── Atualiza data (X41)
         ├── RefreshAll()
         └── Exporta aba como PDF
               └── Playwright
                     └── WhatsApp Web
                           └── Envia PDF para o contato configurado
```

## ⚙️ Configuração necessária

Antes de rodar, ajuste no código:

| Item | Onde configurar |
|---|---|
| Caminho do arquivo `.xlsm` | `wb = app.books.open(r'...')` |
| Nome da aba com a tabela dinâmica | `wb.sheets('Indicador Faturamento-julho')` |
| Área de impressão | `aba.api.PageSetup.PrintArea` |
| Pasta de destino do PDF | `aba.api.ExportAsFixedFormat(0, r'...')` |
| Caminho da sessão do WhatsApp Web | `caminho_da_sessao` |
| Nome do contato no WhatsApp | `nome_do_contato` |

> A sessão do WhatsApp Web precisa estar previamente autenticada (QR Code escaneado ao menos uma vez) para que o Playwright reutilize o login salvo em `caminho_da_sessao`.

## ▶️ Como executar

```bash
pip install xlwings playwright
playwright install chromium
python nome_do_script.py
```

## 📌 Observações

- O script foi desenhado para rodar sem interação manual (headless no navegador, invisível no Excel), sendo ideal para agendamento automático (ex: Agendador de Tarefas do Windows).
- Todos os passos possuem tratamento de erro básico com print de tela de debug para facilitar a manutenção.
- Próximas melhorias possíveis: logging estruturado, arquivo de configuração externo (paths, contato, horários), espera ativa em vez de `time.sleep()` fixo, e modularização em arquivos separados (`excel.py`, `whatsapp.py`, `config.py`, etc).

## ✍️ Autor

Desenvolvido por **Murilo Moraes** como parte de um processo de aprendizado prático de Python, com foco em automações reais aplicadas ao dia a dia de trabalho.