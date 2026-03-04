import os
import logging
from pathlib import Path
from datetime import datetime

from dotenv import load_dotenv

from clients.graph_client import GraphMailClient
from services.email_templates import build_email_acima_99
from services.email_signature import build_signature_inline_attachments


# ======================================================
# LOG
# ======================================================
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger("sla_mensal")

# Carrega .env
load_dotenv()

DRY_RUN = os.getenv("DRY_RUN", "True").strip().lower() in ("1", "true", "yes", "y", "on")


def main():
    base_dir = Path(__file__).resolve().parent

    # =========================
    # EXEMPLO (MVP)
    # =========================
    # No começo vamos rodar com itens “mockados”.
    # Depois você pluga o Grafana/Zabbix aqui.
    regionais_ok = [
        {"regional": "Integrada Rio de Janeiro", "sla": 100.0},
        # {"regional": "Integrada X", "sla": 99.2},
    ]

    # Referência do mês (ex.: Janeiro 2026)
    # (depois a gente automatiza com "mês anterior")
    mes_referencia = "Janeiro"
    ano_referencia = "2026"

    # Destinatários (por enquanto fixo; depois vem da sua planilha por regional)
    # Você pode testar com seu e-mail.
    to_emails = [
        os.getenv("TEAMS_TEST_TO", "").strip() or "roosevelt.pimentel@gpssa.com.br"
    ]

    # Cliente Graph
    mailer = GraphMailClient()

    # Anexos inline da assinatura (GIF)
    signature_attachments = build_signature_inline_attachments(
        base_dir=base_dir,
        gif_relative_path="image/assinatura_gif.gif",
        cid="assinatura_gif",
    )

    for item in regionais_ok:
        regional_nome = item["regional"]
        sla = float(item["sla"])

        subject, html, attachments = build_email_acima_99(
            regional_nome=regional_nome,
            mes_referencia=mes_referencia,
            ano_referencia=ano_referencia,
            sla_percent=sla,
            base_dir=base_dir,
            grafana_print_relative_path="image/grafana_print.png",  # coloque o print aqui se quiser
            usar_imagem_embutida_base64=True,
        )

        # Junta anexos (template + assinatura)
        final_attachments = []
        final_attachments.extend(attachments or [])
        final_attachments.extend(signature_attachments or [])

        if DRY_RUN:
            logger.info("[DRY_RUN] Enviaria: %s | para=%s | subject=%s", regional_nome, to_emails, subject)
            continue

        logger.info("Enviando: %s | para=%s | subject=%s", regional_nome, to_emails, subject)

        mailer.send_mail(
            to=to_emails,
            subject=subject,
            body_content=html,
            is_html=True,
            attachments=final_attachments,
            # reply_to pode ser configurado no .env (REPLY_TO_GROUP_EMAIL)
        )

    logger.info("Finalizado.")


if __name__ == "__main__":
    main()