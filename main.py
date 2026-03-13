import os
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import warnings
import urllib3

from dotenv import load_dotenv

from clients.graph_client import GraphMailClient
from clients.zabbix_client import ZabbixClient
from services.email_templates import build_email_acima_99, build_email_abaixo_99
from services.email_signature import build_signature_inline_attachments
from services.recipients_service import RecipientsService
from services.sla_service import SlaService


# ======================================================
# LOG
# ======================================================
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger("sla_mensal")

# Carrega .env (override garante que o .env prevaleça)
load_dotenv(override=True)

DRY_RUN = os.getenv("DRY_RUN", "True").strip().lower() in ("1", "true", "yes", "y", "on")
USE_ZABBIX = os.getenv("USE_ZABBIX", "False").strip().lower() in ("1", "true", "yes", "y", "on")
SAFE_TEST_TO = os.getenv("SAFE_TEST_TO", "").strip()
ZABBIX_VERIFY_SSL = os.getenv("ZABBIX_VERIFY_SSL", "True").strip().lower() in (
    "1", "true", "yes", "y", "on"
)

if not ZABBIX_VERIFY_SSL:
    # Evita poluir o log com warning de SSL autoassinado
    warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)


def main():
    base_dir = Path(__file__).resolve().parent

    # Referencia do mes anterior
    hoje = datetime.now()
    first_day_current = hoje.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_day_prev = first_day_current - timedelta(seconds=1)
    first_day_prev = last_day_prev.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    periodo_inicio = int(first_day_prev.timestamp())
    periodo_fim = int(last_day_prev.timestamp())
    meses = [
        "Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
    ]
    mes_referencia = meses[first_day_prev.month - 1]
    ano_referencia = str(first_day_prev.year)

    # =========================
    # FONTE DE DADOS (Zabbix ou mock)
    # =========================
    if USE_ZABBIX:
        sla_service = SlaService(ZabbixClient())
        regionais_ok = sla_service.get_regionals_sla(
            time_from=periodo_inicio,
            time_to=periodo_fim,
        )
    else:
        regionais_ok = [
            {"regional": "Integrada Rio de Janeiro", "sla": 100.0},
            # {"regional": "Integrada X", "sla": 99.2},
        ]

    # Planilha de destinatarios
    contatos_path = os.getenv("REGIONAIS_CONTATOS_PATH", "").strip()
    if not contatos_path:
        raise RuntimeError("REGIONAIS_CONTATOS_PATH nao configurado no .env.")
    contatos_sheet = os.getenv("REGIONAIS_CONTATOS_SHEET", "").strip() or None
    recipients = RecipientsService(
        str((base_dir / contatos_path).resolve()),
        sheet_name=contatos_sheet,
    )

    test_emails = [e.strip() for e in SAFE_TEST_TO.split(",") if e.strip()]

    # Cliente Graph (mantido para uso futuro; neste fluxo nao envia)
    mailer = GraphMailClient()

    # Anexos inline da assinatura (GIF)
    signature_attachments = build_signature_inline_attachments(
        base_dir=base_dir,
        gif_relative_path="image/assinatura_gif.gif",
        cid="assinatura_gif",
    )

    meses_slug = [
        "jan", "fev", "mar", "abr", "mai", "jun",
        "jul", "ago", "set", "out", "nov", "dez",
    ]
    hoje_dir = f"{hoje.day:02d}"
    mes_slug = meses_slug[hoje.month - 1]
    export_root = (
        base_dir
        / "exports"
        / f"{hoje.year}"
        / mes_slug
        / hoje_dir
    )
    export_root.mkdir(parents=True, exist_ok=True)
    summary_rows = []

    for item in regionais_ok:
        regional_nome = str(item.get("regional", "")).strip()
        sla = float(item.get("sla", 0.0))

        if sla < 98.0:
            target = "draft"
        elif sla >= 99.0:
            target = "send"
        else:
            continue

        if target == "send":
            subject, html, attachments = build_email_acima_99(
                regional_nome=regional_nome,
                mes_referencia=mes_referencia,
                ano_referencia=ano_referencia,
                sla_percent=sla,
                base_dir=base_dir,
                sla_print_relative_path="image/sla_print.png",  # coloque o print aqui se quiser
                usar_imagem_embutida_base64=True,
            )
        else:
            subject, html, attachments = build_email_abaixo_99(
                regional_nome=regional_nome,
                mes_referencia=mes_referencia,
                ano_referencia=ano_referencia,
                sla_percent=sla,
                base_dir=base_dir,
                sla_print_relative_path="image/sla_print.png",  # coloque o print aqui se quiser
                usar_imagem_embutida_base64=True,
            )

        # Junta anexos (template + assinatura)
        final_attachments = []
        final_attachments.extend(attachments or [])
        final_attachments.extend(signature_attachments or [])

        # Destinatarios da planilha
        to_emails = recipients.get_emails_by_regional(regional_nome)
        if test_emails:
            to_emails = test_emails
        if not to_emails:
            logger.warning("Sem emails na planilha para regional: %s", regional_nome)

        summary_rows.append(
            {
                "regional": regional_nome,
                "sla": f"{sla:.1f}",
                "acao": "enviar" if target == "send" else "rascunho",
                "emails": ";".join(to_emails),
                "assunto": subject,
            }
        )

        if target == "send":
            if DRY_RUN:
                logger.info(
                    "[DRY_RUN] Enviaria: %s | sla=%.1f | para=%s | subject=%s",
                    regional_nome,
                    sla,
                    to_emails,
                    subject,
                )
                continue

            logger.info("Enviando: %s | sla=%.1f | para=%s", regional_nome, sla, to_emails)
            mailer.send_mail(
                to=to_emails,
                subject=subject,
                body_content=html,
                is_html=True,
                attachments=final_attachments,
                # reply_to pode ser configurado no .env (REPLY_TO_GROUP_EMAIL)
            )
            continue

        if DRY_RUN:
            logger.info(
                "[DRY_RUN] Criaria rascunho: %s | sla=%.1f | para=%s | subject=%s",
                regional_nome,
                sla,
                to_emails,
                subject,
            )
            continue

        draft_id = mailer.create_draft(
            to=to_emails,
            subject=subject,
            body_content=html,
            is_html=True,
            attachments=final_attachments,
        )
        logger.info(
            "[RASCUNHO] %s | sla=%.1f | para=%s | draft_id=%s",
            regional_nome,
            sla,
            to_emails,
            draft_id,
        )

    summary_xlsx = export_root / "envio_sla_mes.xlsx"
    pd.DataFrame(summary_rows).to_excel(summary_xlsx, index=False)
    logger.info("Resumo XLSX gerado: %s", summary_xlsx)

    logger.info("Finalizado.")


if __name__ == "__main__":
    main()