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

Cache delegado (rascunho):
- GRAPH_USE_AUTH_CACHE_FOR_DRAFT=True
- GRAPH_AUTH_CACHE_PATH=.auth_cache/sla_token_cache.bin
- GRAPH_DELEGATED_SCOPES=Mail.ReadWrite

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

## 🗓️ Agendar no Windows (dia 5 as 10:00)

1) Rodar o script de agendamento:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install_task.ps1
```

2) Rodar manualmente (teste):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_main.ps1
```

## 📄 Saidas

- Resumo XLSX:
  exports/AAAA/mes_abrev/DD/envio_sla_mes.xlsx

## 🧭 Observacoes

- SLA >= 99: envia email.
- SLA < 98: cria rascunho no Outlook.
- SLA entre 98 e 99: ignora.
