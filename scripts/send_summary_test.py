import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT_DIR))

from main import build_summary_message, build_summary_teams_message, parse_env_list  # noqa: E402
from clients.graph_client import GraphMailClient  # noqa: E402
from clients.teams_client import TeamsClient  # noqa: E402
from services.email_signature import build_signature_inline_attachments  # noqa: E402
from services.recipients_service import EMAIL_FIELD_CANDIDATES, RecipientsService  # noqa: E402


def _latest_summary_xlsx() -> Path:
    files = sorted(ROOT_DIR.glob("exports/*/*/*/envio_sla_mes.xlsx"), key=lambda path: path.stat().st_mtime)
    if not files:
        raise RuntimeError("Nenhum envio_sla_mes.xlsx encontrado em exports/.")
    return files[-1]


def _parse_issue_rows(df: pd.DataFrame) -> list[dict]:
    rows = []
    for _, row in df.iterrows():
        problema = str(row.get("problemas", "") or "").strip()
        resultado = str(row.get("resultado", "") or "").strip()
        if not problema and resultado not in {"erro_envio", "erro_rascunho", "sem_destinatario"}:
            continue
        rows.append(
            {
                "regional": str(row.get("regional", "") or ""),
                "tipo": resultado or "alerta",
                "acao": str(row.get("acao", "") or ""),
                "emails_nao_encaminhados": str(row.get("emails_utilizados", "") or ""),
                "email_diretor": str(row.get("email_diretor", "") or ""),
                "email_gerente": str(row.get("email_gerente", "") or ""),
                "email_apoio_1": str(row.get("email_apoio_1", "") or ""),
                "email_apoio_2": str(row.get("email_apoio_2", "") or ""),
                "detalhe": problema or resultado,
            }
        )
    return rows


def _enrich_email_columns(summary_rows: list[dict]) -> None:
    recipients = RecipientsService(str((ROOT_DIR / "data" / "Lideres.xlsx").resolve()), sheet_name="MODELOPY")
    for row in summary_rows:
        regional = str(row.get("regional", "") or "").strip()
        snapshot = recipients.get_email_fields_by_regional(regional) if regional else {}
        for key in EMAIL_FIELD_CANDIDATES:
            if key == "email_apoio":
                continue
            row[key] = str(row.get(key, "") or snapshot.get(key, "") or "")


def main() -> None:
    summary_xlsx = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else _latest_summary_xlsx().resolve()
    df = pd.read_excel(summary_xlsx)
    summary_rows = df.fillna("").to_dict("records")
    _enrich_email_columns(summary_rows)
    issue_rows = _parse_issue_rows(df)
    if not issue_rows:
        for row in summary_rows:
            if not str(row.get("problemas", "") or "").strip():
                continue
            issue_rows.append(
                {
                    "regional": str(row.get("regional", "") or ""),
                    "tipo": str(row.get("resultado", "") or "alerta"),
                    "acao": str(row.get("acao", "") or ""),
                    "emails_nao_encaminhados": str(row.get("emails_utilizados", "") or ""),
                    "email_diretor": str(row.get("email_diretor", "") or ""),
                    "email_gerente": str(row.get("email_gerente", "") or ""),
                    "email_apoio_1": str(row.get("email_apoio_1", "") or ""),
                    "email_apoio_2": str(row.get("email_apoio_2", "") or ""),
                    "detalhe": str(row.get("problemas", "") or row.get("resultado", "") or ""),
                }
            )

    today = datetime.now()
    first_day_this_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_day_prev_month = first_day_this_month - timedelta(seconds=1)
    meses = [
        "Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
    ]
    mes_referencia = meses[last_day_prev_month.month - 1]
    ano_referencia = str(last_day_prev_month.year)

    mailer = GraphMailClient()
    teams_client = TeamsClient()

    summary_subject, _summary_text, summary_html = build_summary_message(
        summary_rows,
        issue_rows,
        mes_referencia=mes_referencia,
        ano_referencia=ano_referencia,
        total_regionais=len(summary_rows),
        dry_run=False,
        safe_test_to=[],
    )

    email_recipients = parse_env_list("SUMMARY_SAFE_TEST_TO") or parse_env_list("SUMMARY_EMAIL_RECIPIENTS")
    if email_recipients:
        attachments = [
            GraphMailClient.make_file_attachment(summary_xlsx),
            *build_signature_inline_attachments(
                base_dir=ROOT_DIR,
                gif_relative_path="image/assinatura_gif.gif",
                cid="assinatura_gif",
            ),
        ]
        mailer.send_mail(
            to=email_recipients,
            subject=summary_subject,
            body_content=summary_html,
            is_html=True,
            attachments=attachments,
        )

    teams_recipients = teams_client.get_summary_recipients()
    if teams_recipients:
        teams_html = build_summary_teams_message(
            summary_rows,
            issue_rows,
            mes_referencia=mes_referencia,
            ano_referencia=ano_referencia,
        )
        teams_client.send_summary(teams_recipients, teams_html)

    print(f"Resumo reenviado usando {summary_xlsx}")


if __name__ == "__main__":
    main()