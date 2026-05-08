# Ajuste sys.path para garantir imports robustos
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))

import os
import html
import logging
import pandas as pd
from datetime import datetime, timedelta
import warnings
import urllib3
import unicodedata
import re
from collections import Counter, defaultdict
from pathlib import Path

from dotenv import load_dotenv

# Integração: Importa função principal do downloader de PDFs
from scripts.fortianalyzer_api_modelo1 import main as baixar_pdfs_fortianalyzer

from clients.graph_client import GraphMailClient
from clients.teams_client import TeamsClient
from clients.zabbix_client import ZabbixClient
from services.email_templates import build_email_acima_99, build_email_abaixo_99
from services.email_signature import build_signature_html, build_signature_inline_attachments
from services.recipients_service import EMAIL_FIELD_CANDIDATES, RecipientsService
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


def parse_env_list(name: str) -> list[str]:
    raw_value = str(os.getenv(name, "")).strip()
    return [item.strip() for item in raw_value.split(",") if item.strip()]


def html_table(headers: list[str], rows: list[list[str]]) -> str:
    head_html = "".join(f"<th>{html.escape(header)}</th>" for header in headers)
    body_lines = []
    for row in rows:
        cols_html = "".join(f"<td>{html.escape(str(col or ''))}</td>" for col in row)
        body_lines.append(f"<tr>{cols_html}</tr>")
    body_html = "".join(body_lines) or (
        f"<tr><td colspan='{len(headers)}'>Nenhum registro.</td></tr>"
    )
    return (
        "<table border='1' cellpadding='6' cellspacing='0' "
        "style='border-collapse:collapse; margin:0 0 18px 0;'>"
        f"<thead><tr>{head_html}</tr></thead>"
        f"<tbody>{body_html}</tbody>"
        "</table>"
    )


def spacer(height: int = 10) -> str:
    return f"<div style='height:{height}px; line-height:{height}px;'>&nbsp;</div>"


def display_result_label(value: str) -> str:
    raw_value = str(value or "").strip()
    mapping = {
        "rascunho_criado": "SLA não atingido",
        "dry_run_draft": "SLA não atingido",
    }
    return mapping.get(raw_value, raw_value)


def get_summary_email_recipients(test_emails: list[str]) -> list[str]:
    test_override = parse_env_list("SUMMARY_SAFE_TEST_TO")
    if test_override:
        return test_override
    configured = parse_env_list("SUMMARY_EMAIL_RECIPIENTS")
    if configured:
        return configured
    return list(test_emails)


def build_summary_teams_message(
    summary_rows: list[dict],
    issue_rows: list[dict],
    *,
    mes_referencia: str,
    ano_referencia: str,
) -> str:
    saudacao = "Bom dia a todos" if datetime.now().hour < 12 else "Boa tarde a todos"
    execution_rows = [
        [
            str(item.get("regional", "")),
            str(item.get("sla", "")),
            display_result_label(item.get("resultado", "")),
            str(item.get("email_diretor", "")),
            str(item.get("email_gerente", "")),
            str(item.get("email_apoio_1", "")),
            str(item.get("email_apoio_2", "")),
        ]
        for item in summary_rows
    ]
    issue_table_rows = [
        [
            str(item.get("regional", "")),
            str(item.get("tipo", "")),
            str(item.get("emails_nao_encaminhados", "")),
            str(item.get("email_diretor", "")),
            str(item.get("email_gerente", "")),
            str(item.get("email_apoio_1", "")),
            str(item.get("email_apoio_2", "")),
            str(item.get("detalhe", "")),
        ]
        for item in issue_rows
    ]

    lines = [
        "<div>",
        f"<p><strong>{html.escape(saudacao)}</strong></p>",
        f"<p>Resumo do envio de SLA encaminhado no ciclo atual ({html.escape(mes_referencia)}/{html.escape(ano_referencia)}).</p>",
        spacer(),
        html_table([
            "Regional",
            "SLA (%)",
            "Resultado",
            "EMAIL_DIRETOR",
            "EMAIL_GERENTE",
            "EMAIL_APOIO_1",
            "EMAIL_APOIO_2",
        ], execution_rows),
        spacer(),
        "<p><strong>Problemas detectados</strong></p>",
        spacer(),
        html_table([
            "Regional",
            "Tipo",
            "Emails Nao Encaminhados",
            "EMAIL_DIRETOR",
            "EMAIL_GERENTE",
            "EMAIL_APOIO_1",
            "EMAIL_APOIO_2",
            "Detalhe",
        ], issue_table_rows),
        spacer(),
        "<p><strong>IMPORTANTE!</strong> Esta e uma mensagem automatica. Por favor, nao responda.</p>",
        "</div>",
    ]
    return "\n".join(lines)


def build_summary_message(
    summary_rows: list[dict],
    issue_rows: list[dict],
    *,
    mes_referencia: str,
    ano_referencia: str,
    total_regionais: int,
    dry_run: bool,
    safe_test_to: list[str],
) -> tuple[str, str, str]:
    counts = Counter(str(row.get("resultado") or "desconhecido") for row in summary_rows)
    sent_count = counts.get("enviado", 0)
    draft_count = counts.get("rascunho_criado", 0)
    dry_send_count = counts.get("dry_run_send", 0)
    dry_draft_count = counts.get("dry_run_draft", 0)
    issue_count = len(issue_rows)
    saudacao = "Bom dia a todos," if datetime.now().hour < 12 else "Boa tarde a todos,"

    subject = (
        f"[Sumario Executivo] SLA Mensal {mes_referencia}/{ano_referencia} | "
        f"enviados={sent_count} | rascunhos={draft_count} | problemas={issue_count}"
    )

    execution_rows = [
        [
            str(item.get("regional", "")),
            f"{mes_referencia}/{ano_referencia}",
            str(item.get("sla", "")),
            str(item.get("acao", "")),
            display_result_label(item.get("resultado", "")),
            str(item.get("email_diretor", "")),
            str(item.get("email_gerente", "")),
            str(item.get("email_apoio_1", "")),
            str(item.get("email_apoio_2", "")),
        ]
        for item in summary_rows
    ]
    issue_table_rows = [
        [
            str(item.get("regional", "")),
            str(item.get("tipo", "")),
            str(item.get("acao", "")),
            str(item.get("emails_nao_encaminhados", "")),
            str(item.get("email_diretor", "")),
            str(item.get("email_gerente", "")),
            str(item.get("email_apoio_1", "")),
            str(item.get("email_apoio_2", "")),
            str(item.get("detalhe", "")),
        ]
        for item in issue_rows
    ]
    signature_html = build_signature_html(
        titulo="Governanca de TI",
        subtitulo="Infraestrutura Regional | SLA Mensal",
        email=str(os.getenv("REPLY_TO_GROUP_EMAIL", "governanca.ti@gpssa.com.br")).strip() or "governanca.ti@gpssa.com.br",
        teams_1_label="Fale com Roosevelt Pimentel no Teams",
        teams_1_link="https://teams.microsoft.com/l/chat/0/0?users=roosevelt.pimentel@gpssa.com.br",
        teams_2_label="",
        teams_2_link="",
        teams_3_label="",
        teams_3_link="",
    )
    summary_indicator_rows = [
        ["Regionais avaliadas", str(total_regionais)],
        ["Emails enviados para quem atingiu SLA", str(sent_count + dry_send_count)],
        ["Rascunhos criados para quem nao atingiu o SLA", str(draft_count + dry_draft_count)],
    ]

    html_parts = [
        "<div style='font-family:Arial, sans-serif; font-size:14px; color:#1a1a1a; line-height:1.4;'>",
        f"<p>{html.escape(saudacao)}</p>",
        spacer(),
        "<p>Este e um e-mail automatico de relatorio, contendo o resumo do SLA encaminhado no ciclo atual. Caso haja duvidas ou necessidade de esclarecimentos, estou a disposicao.</p>",
        spacer(),
        "<p><strong>Resumo do envio</strong></p>",
        spacer(),
        html_table(["Indicador", "Valor"], summary_indicator_rows),
        spacer(),
        "<p><strong>Resumo por regional</strong></p>",
        spacer(),
        html_table(
            [
                "Regional",
                "Mes",
                "SLA (%)",
                "Acao",
                "Resultado",
                "EMAIL_DIRETOR",
                "EMAIL_GERENTE",
                "EMAIL_APOIO_1",
                "EMAIL_APOIO_2",
            ],
            execution_rows,
        ),
        spacer(),
        "<p><strong>Problemas e alertas locais</strong></p>",
        spacer(),
        html_table(
            [
                "Regional",
                "Tipo",
                "Acao",
                "Emails Nao Encaminhados",
                "EMAIL_DIRETOR",
                "EMAIL_GERENTE",
                "EMAIL_APOIO_1",
                "EMAIL_APOIO_2",
                "Detalhe",
            ],
            issue_table_rows,
        ),
        spacer(),
        signature_html,
        "</div>",
    ]

    text_lines = [
        f"Sumario da execucao SLA {mes_referencia}/{ano_referencia}",
        "",
        f"Regionais avaliadas: {total_regionais}",
        f"Emails enviados para quem atingiu SLA: {sent_count + dry_send_count}",
        f"Rascunhos criados para quem nao atingiu o SLA: {draft_count + dry_draft_count}",
        "",
        "Resumo por regional:",
    ]
    for item in summary_rows:
        text_lines.append(
            " - "
            f"{item.get('regional', '')} | SLA={item.get('sla', '')}% | "
            f"acao={item.get('acao', '')} | resultado={item.get('resultado', '')} | "
            f"resultado_exibicao={display_result_label(item.get('resultado', ''))} | "
            f"email_diretor={item.get('email_diretor', '')} | "
            f"email_gerente={item.get('email_gerente', '')} | "
            f"email_apoio_1={item.get('email_apoio_1', '')} | "
            f"email_apoio_2={item.get('email_apoio_2', '')}"
        )
    text_lines.append("")
    text_lines.append("Problemas e alertas locais:")
    if issue_rows:
        for item in issue_rows:
            text_lines.append(
                " - "
                f"{item.get('regional', '')} | tipo={item.get('tipo', '')} | "
                f"acao={item.get('acao', '')} | emails_nao_encaminhados={item.get('emails_nao_encaminhados', '')} | detalhe={item.get('detalhe', '')}"
            )
    else:
        text_lines.append(" - Nenhum problema local detectado.")
    return subject, "\n".join(text_lines), "\n".join(html_parts)


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

    def empty_email_snapshot() -> dict[str, str]:
        return {key: "" for key in EMAIL_FIELD_CANDIDATES if key != "email_apoio"}

    def build_issue_row(
        *,
        regional: str,
        issue_type: str,
        action: str,
        detail: str,
        email_snapshot: dict[str, str],
        failed_emails: list[str],
    ) -> dict:
        row = {
            "regional": regional,
            "tipo": issue_type,
            "acao": action,
            "emails_nao_encaminhados": ";".join(failed_emails),
            "detalhe": detail,
        }
        row.update(email_snapshot)
        return row

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
    teams_client = TeamsClient()

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
    issue_rows = []

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
        email_snapshot = recipients.get_email_fields_by_regional(regional_nome) or empty_email_snapshot()
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
            "problemas": "",
        }
        summary_row.update(email_snapshot)

        if not pdf_paths:
            issue_rows.append(
                build_issue_row(
                    regional=regional_nome,
                    issue_type="sem_pdf",
                    action=target,
                    detail="Nenhum PDF localizado para a regional.",
                    email_snapshot=email_snapshot,
                    failed_emails=to_emails,
                )
            )
            summary_row["problemas"] = "Nenhum PDF localizado para a regional."

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
            if not to_emails:
                summary_row["resultado"] = "sem_destinatario"
                summary_row["problemas"] = "Sem destinatarios validos na planilha."
                issue_rows.append(
                    build_issue_row(
                        regional=regional_nome,
                        issue_type="sem_destinatario",
                        action=target,
                        detail="Sem destinatarios validos na planilha.",
                        email_snapshot=email_snapshot,
                        failed_emails=original_to_emails,
                    )
                )
                summary_rows.append(summary_row)
                continue

            try:
                mailer.send_mail(
                    to=to_emails,
                    subject=subject,
                    body_content=html,
                    is_html=True,
                    attachments=final_attachments,
                    # reply_to pode ser configurado no .env (REPLY_TO_GROUP_EMAIL)
                )
                summary_row["resultado"] = "enviado"
            except Exception as exc:
                summary_row["resultado"] = "erro_envio"
                summary_row["problemas"] = str(exc)
                issue_rows.append(
                    build_issue_row(
                        regional=regional_nome,
                        issue_type="erro_envio",
                        action=target,
                        detail=str(exc),
                        email_snapshot=email_snapshot,
                        failed_emails=to_emails,
                    )
                )
                logger.exception("Falha ao enviar email da regional %s", regional_nome)
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

        if not to_emails:
            summary_row["resultado"] = "sem_destinatario"
            summary_row["problemas"] = "Sem destinatarios validos na planilha."
            issue_rows.append(
                build_issue_row(
                    regional=regional_nome,
                    issue_type="sem_destinatario",
                    action=target,
                    detail="Sem destinatarios validos na planilha.",
                    email_snapshot=email_snapshot,
                    failed_emails=original_to_emails,
                )
            )
            summary_rows.append(summary_row)
            continue

        try:
            draft_id = mailer.create_draft(
                to=to_emails,
                subject=subject,
                body_content=html,
                is_html=True,
                attachments=final_attachments,
            )
            summary_row["resultado"] = "rascunho_criado"
            summary_row["draft_id"] = draft_id
        except Exception as exc:
            summary_row["resultado"] = "erro_rascunho"
            summary_row["problemas"] = str(exc)
            issue_rows.append(
                build_issue_row(
                    regional=regional_nome,
                    issue_type="erro_rascunho",
                    action=target,
                    detail=str(exc),
                    email_snapshot=email_snapshot,
                    failed_emails=to_emails,
                )
            )
            logger.exception("Falha ao criar rascunho da regional %s", regional_nome)
        summary_rows.append(summary_row)
        if summary_row["resultado"] == "rascunho_criado":
            logger.info(
                "[RASCUNHO] %s | sla=%.1f | para=%s | draft_id=%s",
                regional_nome,
                sla,
                to_emails,
                summary_row["draft_id"],
            )

    summary_xlsx = export_root / "envio_sla_mes.xlsx"
    pd.DataFrame(summary_rows).to_excel(summary_xlsx, index=False)
    logger.info("Resumo XLSX gerado: %s", summary_xlsx)

    summary_recipients = get_summary_email_recipients(test_emails)
    if summary_recipients:
        try:
            summary_subject, summary_text, summary_html = build_summary_message(
                summary_rows,
                issue_rows,
                mes_referencia=mes_referencia,
                ano_referencia=ano_referencia,
                total_regionais=len(regionais_ok),
                dry_run=DRY_RUN,
                safe_test_to=test_emails,
            )
            summary_attachments = [
                GraphMailClient.make_file_attachment(summary_xlsx),
                *build_signature_inline_attachments(
                    base_dir=base_dir,
                    gif_relative_path="image/assinatura_gif.gif",
                    cid="assinatura_gif",
                ),
            ]
            mailer.send_mail(
                to=summary_recipients,
                subject=summary_subject,
                body_content=summary_html,
                is_html=True,
                attachments=summary_attachments,
            )
            logger.info("Sumario executivo enviado para %s", summary_recipients)
        except Exception:
            logger.exception("Falha ao enviar sumario executivo do SLA")
    else:
        logger.info("SUMMARY_EMAIL_RECIPIENTS nao configurado. Sumario executivo nao sera enviado.")

    teams_summary_recipients = teams_client.get_summary_recipients()
    if teams_summary_recipients:
        try:
            teams_summary_html = build_summary_teams_message(
                summary_rows,
                issue_rows,
                mes_referencia=mes_referencia,
                ano_referencia=ano_referencia,
            )
            teams_client.send_summary(teams_summary_recipients, teams_summary_html)
        except Exception:
            logger.exception("Falha ao enviar sumario executivo do SLA para o Teams")
    else:
        logger.info("TEAMS_SUMMARY_RECIPIENTS nao configurado. Sumario do Teams nao sera enviado.")

    logger.info("Finalizado.")


if __name__ == "__main__":
    main()