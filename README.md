# 📊 Automacao SLA Mensal

Automacao para coletar SLA mensal do Zabbix, gerar emails de SLA e criar rascunhos para revisao quando o SLA nao atinge a meta.

## 🎯 Objetivo

1. Buscar o SLA do mes anterior no Zabbix.
2. Enviar email automatico para regionais com SLA >= 99.
3. Criar rascunho no Outlook para regionais com SLA < 98.
4. Gerar resumo em XLSX por execucao.

## 📂 Estrutura

- clients/                 clientes de integracao (Graph e Zabbix)
- services/                templates e regras de negocio
- scripts/                 scripts de automacao (login de cache e agendamento)
- data/                    planilhas de contatos
- exports/                 resumo mensal e arquivos gerados
- image/                   imagens e assinatura

## ⚙️ Variaveis do .env

Obrigatorias:
- M365_TENANT_ID
- M365_CLIENT_ID
- M365_CLIENT_SECRET
- M365_SENDER_UPN
- REPLY_TO_GROUP_EMAIL
- REGIONAIS_CONTATOS_PATH
- REGIONAIS_CONTATOS_SHEET
- ZABBIX_URL
- ZABBIX_TOKEN
- USE_ZABBIX
- DRY_RUN

Teste:
- SAFE_TEST_TO=email@dominio.com.br (quando preenchido, todos os envios vao para esse email)
- SUMMARY_EMAIL_RECIPIENTS=email@dominio.com.br[,outro@dominio.com.br] (destinatarios do sumario executivo por email)
- SUMMARY_SAFE_TEST_TO=email@dominio.com.br (sobrescreve apenas os destinatarios do sumario executivo por email)
- ENABLE_TEAMS_SUMMARY=True|False (habilita envio do sumario para Teams)
- TEAMS_SUMMARY_RECIPIENTS=email@dominio.com.br[,outro@dominio.com.br] (destinatarios do sumario executivo no Teams)
- TEAMS_SUMMARY_SAFE_TEST_TO=email@dominio.com.br (sobrescreve apenas os destinatarios do sumario executivo no Teams)

Comportamento em teste:
- `SAFE_TEST_TO` sobrescreve apenas os destinatarios `To`.
- Os emails enviados continuam saindo da conta configurada em `M365_SENDER_UPN`.
- Os rascunhos sao criados na caixa da conta autenticada no cache delegado. No fluxo atual, isso normalmente coincide com `M365_SENDER_UPN`.
- Para teste controlado, use `DRY_RUN=False` junto com `SAFE_TEST_TO` preenchido.

Cache delegado (rascunho):
- GRAPH_USE_AUTH_CACHE_FOR_DRAFT=True
- GRAPH_AUTH_CACHE_PATH=.auth_cache/sla_token_cache.bin
- GRAPH_DELEGATED_SCOPES=Mail.ReadWrite

Cache delegado (Teams sumario):
- reutiliza `GRAPH_AUTH_CACHE_PATH`
- inclua tambem os escopos `Chat.ReadWrite`, `ChatMessage.Send` e `User.Read` em `GRAPH_DELEGATED_SCOPES`

## ▶️ Como rodar

PowerShell:

```powershell
& "C:\automacao_sla\.venv\Scripts\python.exe" .\main.py
```

## 🔐 Gerar cache delegado (rascunho)

```powershell
& "C:\automacao_sla\.venv\Scripts\python.exe" .\scripts\graph_login_cache.py
```

Esse comando gera o cache em .auth_cache/sla_token_cache.bin.

## 🗓️ Agendar no Windows (dia 3 as 10:00)

1) Rodar o script de agendamento:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_task.ps1
```

2) Rodar manualmente (teste):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_main.ps1
```

3) Reenviar apenas o sumario (sem reenviar os emails das regionais):

```powershell
& "C:\automacao_sla\.venv\Scripts\python.exe" .\scripts\send_summary_test.py
```

## 📄 Saidas

- Resumo XLSX:
  exports/AAAA/mes_abrev/DD/envio_sla_mes.xlsx
- Sumario executivo por email:
  tabela HTML com resumo por regional e tabela separada de problemas locais detectados na execucao

Colunas do resumo XLSX:
- `emails_originais`: destinatarios vindos da planilha.
- `emails_utilizados`: destinatarios efetivamente usados na execucao.
- `safe_test_to_aplicado`: valor aplicado via `SAFE_TEST_TO`, quando houver.
- `anexos_pdf`: nome dos PDFs anexados.
- `anexos_pdf_paths`: caminho completo dos PDFs anexados.
- `anexos_pdf_tids`: TIDs FortiAnalyzer usados para cada PDF.
- `anexos_pdf_reports`: nome do relatorio Forti associado ao PDF.
- `resultado`: resultado final da linha (`enviado`, `rascunho_criado`, `dry_run_send`, `dry_run_draft`).
- `draft_id`: id do rascunho no Graph quando a linha gera rascunho real.
- `problemas`: detalhe de falha ou alerta local detectado para a regional.

## 🧭 Observacoes

- SLA >= 99: envia email.
- SLA < 98: cria rascunho no Outlook.
- SLA entre 98 e 99: ignora.
- O PDF anexado e selecionado por correspondencia da regional Forti (`NOME_REG_FORTI`) com deduplicacao por identidade do relatorio.
- O sumario executivo por email usa `SUMMARY_EMAIL_RECIPIENTS`; se essa variavel nao estiver preenchida, o sumario nao e enviado.
- `SUMMARY_SAFE_TEST_TO` e `TEAMS_SUMMARY_SAFE_TEST_TO` permitem testar apenas o sumario sem alterar a lista oficial de destinatarios.
- O sumario do Teams usa `TEAMS_SUMMARY_RECIPIENTS` e depende de cache delegado com escopos de chat.
- O sumario executivo cobre falhas detectadas localmente na execucao. NDRs que chegam depois no Outlook nao sao confirmados no momento do envio pelo Graph.
