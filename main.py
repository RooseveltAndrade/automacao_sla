# Ajuste sys.path para garantir imports robustos
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))

import os
import logging
import pandas as pd
from datetime import datetime, timedelta
import warnings
import urllib3
import unicodedata
import re
from collections import defaultdict
from pathlib import Path

from dotenv import load_dotenv

# Integração: Importa função principal do downloader de PDFs
from scripts.fortianalyzer_api_modelo1 import main as baixar_pdfs_fortianalyzer

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

# Carrega .env sem sobrescrever variaveis ja definidas no processo.
load_dotenv(override=False)

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

    def normalize_match(value: str) -> str:
        text = str(value or "").strip()
        text = unicodedata.normalize("NFKD", text)
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        chunks = []
        for ch in text.upper():
            chunks.append(ch if ch.isalnum() else " ")
        return " ".join("".join(chunks).split())

    def extract_report_identity(value: str) -> str:
        text = str(value or "").strip()
        if text.lower().endswith(".pdf"):
            text = Path(text).stem
        match = re.match(r"^(.*?)-\d{4}-\d{2}-\d{2}-\d{4}-\d{4}(?:_\d+)?$", text)
        if match:
            text = match.group(1)
        return normalize_match(text)

    def matches_report_identity(candidate: str, expected: str) -> bool:
        if not candidate or not expected:
            return False
        return candidate == expected or candidate.startswith(expected + " ")

    def index_pdf_results(pdf_results: list[dict]) -> tuple[dict[str, list[Path]], dict[str, list[Path]]]:
        by_regional: dict[str, list[Path]] = defaultdict(list)
        by_forti: dict[str, list[Path]] = defaultdict(list)
        for item in pdf_results or []:
            pdf_path = Path(str(item.get("pdf") or "")).resolve()
            if not pdf_path.exists():
                continue

            regional_planilha = normalize_match(item.get("regional_planilha", ""))
            regional_forti = extract_report_identity(item.get("regional_forti", ""))
            report_identity = extract_report_identity(item.get("report_name") or item.get("filename") or pdf_path.name)
            if regional_planilha:
                by_regional[regional_planilha].append(pdf_path)
            if regional_forti:
                by_forti[regional_forti].append(pdf_path)
            elif report_identity:
                by_forti[report_identity].append(pdf_path)
        return by_regional, by_forti

    def dedupe_paths(paths: list[Path], pdf_metadata_by_path: dict[str, dict]) -> list[Path]:
        unique_by_identity: dict[str, Path] = {}
        rank_by_identity: dict[str, tuple[str, str]] = {}
        for path in paths:
            resolved_path = path.resolve()
            metadata = pdf_metadata_by_path.get(str(resolved_path), {})
            identity = extract_report_identity(metadata.get("regional_forti") or metadata.get("report_name") or path.name)
            if not identity:
                identity = str(resolved_path)

            rank = (
                str(metadata.get("report_name") or path.name),
                path.name,
            )
            current_rank = rank_by_identity.get(identity)
            if current_rank is None or rank > current_rank:
                unique_by_identity[identity] = resolved_path
                rank_by_identity[identity] = rank

        return list(unique_by_identity.values())

    def resolve_pdf_paths_for_regional(
        regional_nome: str,
        *,
        recipients_service: RecipientsService,
        by_regional: dict[str, list[Path]],
        by_forti: dict[str, list[Path]],
        pdf_dir: Path,
    ) -> list[Path]:
        normalized_regional = normalize_match(regional_nome)
        resolved = list(by_regional.get(normalized_regional, []))

        forti_name = recipients_service.get_forti_name_by_regional(regional_nome)
        normalized_forti = extract_report_identity(forti_name or "")
        if normalized_forti:
            resolved.extend(by_forti.get(normalized_forti, []))
            if not resolved:
                for report_identity, paths in by_forti.items():
                    if matches_report_identity(report_identity, normalized_forti):
                        resolved.extend(paths)

        if not resolved and normalized_forti and pdf_dir.exists():
            for pdf_path in pdf_dir.glob("*.pdf"):
                if matches_report_identity(extract_report_identity(pdf_path.name), normalized_forti):
                    resolved.append(pdf_path.resolve())

        return dedupe_paths(resolved, pdf_metadata_by_path)

    def build_pdf_metadata(pdf_results: list[dict]) -> dict[str, dict]:
        metadata: dict[str, dict] = {}
        for item in pdf_results or []:
            pdf_path = Path(str(item.get("pdf") or "")).resolve()
            if not pdf_path.exists():
                continue
            metadata[str(pdf_path)] = item
        return metadata

    # Passo 1: Baixar os PDFs do FortiAnalyzer antes de qualquer processamento
    pdf_results = []
    try:
        print("[INFO] Iniciando download dos relatórios PDF do FortiAnalyzer...")
        pdf_results = baixar_pdfs_fortianalyzer() or []
        print("[INFO] Download dos relatórios PDF concluído.")
    except Exception as e:
        print(f"[ERRO] Falha ao baixar relatórios do FortiAnalyzer: {e}")
        # Dependendo da criticidade, pode-se abortar ou apenas logar o erro
        # raise

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

    hoje_dir = f"{hoje.day:02d}"
    mes_slug = hoje.strftime("%b").lower()
    export_root = (
        base_dir
        / "exports"
        / f"{hoje.year}"
        / mes_slug
        / hoje_dir
    )
    export_root.mkdir(parents=True, exist_ok=True)
    pdf_dir = export_root / "pdf_do_mês"
    pdfs_by_regional, pdfs_by_forti = index_pdf_results(pdf_results)
    pdf_metadata_by_path = build_pdf_metadata(pdf_results)
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
        pdf_paths = resolve_pdf_paths_for_regional(
            regional_nome,
            recipients_service=recipients,
            by_regional=pdfs_by_regional,
            by_forti=pdfs_by_forti,
            pdf_dir=pdf_dir,
        )
        for pdf_path in pdf_paths:
            final_attachments.append(GraphMailClient.make_file_attachment(pdf_path))
        final_attachments.extend(signature_attachments or [])
        if pdf_paths:
            logger.info(
                "PDF(s) anexado(s) para %s: %s",
                regional_nome,
                ", ".join(path.name for path in pdf_paths),
            )
        else:
            logger.warning("Nenhum PDF localizado para a regional: %s", regional_nome)

        # Destinatarios da planilha
        original_to_emails = recipients.get_emails_by_regional(regional_nome)
        to_emails = list(original_to_emails)
        if test_emails:
            to_emails = test_emails
        if not to_emails:
            logger.warning("Sem emails na planilha para regional: %s", regional_nome)

        pdf_meta = [pdf_metadata_by_path.get(str(path.resolve()), {}) for path in pdf_paths]
        summary_row = {
            "regional": regional_nome,
            "sla": f"{sla:.1f}",
            "acao": "enviar" if target == "send" else "rascunho",
            "emails_originais": ";".join(original_to_emails),
            "emails_utilizados": ";".join(to_emails),
            "safe_test_to_aplicado": ";".join(test_emails),
            "assunto": subject,
            "anexos_pdf": ";".join(path.name for path in pdf_paths),
            "anexos_pdf_paths": ";".join(str(path) for path in pdf_paths),
            "anexos_pdf_tids": ";".join(str(meta.get("tid", "")) for meta in pdf_meta if meta),
            "anexos_pdf_reports": ";".join(str(meta.get("report_name", "")) for meta in pdf_meta if meta),
            "resultado": "pendente",
            "draft_id": "",
        }

        if target == "send":
            if DRY_RUN:
                summary_row["resultado"] = "dry_run_send"
                summary_rows.append(summary_row)
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
            summary_row["resultado"] = "enviado"
            summary_rows.append(summary_row)
            continue

        if DRY_RUN:
            summary_row["resultado"] = "dry_run_draft"
            summary_rows.append(summary_row)
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
        summary_row["resultado"] = "rascunho_criado"
        summary_row["draft_id"] = draft_id
        summary_rows.append(summary_row)
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