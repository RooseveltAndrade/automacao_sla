import logging
import base64
from pathlib import Path
from typing import Tuple, List, Dict, Any, Optional

from services.email_signature import build_signature_html

logger = logging.getLogger(__name__)


def _read_png_as_b64(path: Path) -> Optional[str]:
    if not path.exists():
        return None
    try:
        return base64.b64encode(path.read_bytes()).decode("ascii")
    except Exception as e:
        logger.warning("Não foi possível carregar imagem (%s): %s", path, e)
        return None


def _build_sla_table(
    *,
    regional_nome: str,
    mes_referencia: str,
    ano_referencia: str,
    sla_percent: float,
    base_dir: Path,
    sla_print_relative_path: str,
    usar_imagem_embutida_base64: bool,
) -> str:
    print_path = (base_dir / sla_print_relative_path).resolve()
    img_html = ""

    if usar_imagem_embutida_base64:
        b64 = _read_png_as_b64(print_path)
        if b64:
            img_html = (
                "<p style='margin:0 0 12px 0;'>"
                f"<img src='data:image/png;base64,{b64}' "
                "alt='Demonstrativo SLA' style='max-width:680px; height:auto; border:1px solid #ddd;'>"
                "</p>"
            )

    if not img_html:
        img_html = (
            "<div style='height:10px; line-height:10px;'>&nbsp;</div>"
            "<table role='presentation' cellpadding='0' cellspacing='0' "
            "style='border-collapse:collapse; margin:0 0 12px 0; font-size:13px;'>"
            "<tr>"
            "<th style='border:1px solid #ddd; padding:6px 10px; background:#f6f6f6; text-align:left;'>Regional</th>"
            "<th style='border:1px solid #ddd; padding:6px 10px; background:#f6f6f6; text-align:left;'>Mes</th>"
            "<th style='border:1px solid #ddd; padding:6px 10px; background:#f6f6f6; text-align:left;'>SLA (%)</th>"
            "</tr>"
            "<tr>"
            f"<td style='border:1px solid #ddd; padding:6px 10px;'>{regional_nome}</td>"
            f"<td style='border:1px solid #ddd; padding:6px 10px;'>{mes_referencia}/{ano_referencia}</td>"
            f"<td style='border:1px solid #ddd; padding:6px 10px;'><b>{sla_percent:.1f}%</b></td>"
            "</tr>"
            "</table>"
            "<div style='height:10px; line-height:10px;'>&nbsp;</div>"
        )

    return img_html


def build_email_acima_99(
    *,
    regional_nome: str,
    mes_referencia: str,
    ano_referencia: str,
    sla_percent: float,
    base_dir: Path,
    sla_print_relative_path: str = "image/sla_print.png",
    usar_imagem_embutida_base64: bool = True,
) -> Tuple[str, str, List[Dict[str, Any]]]:
    """
    Retorna (subject, html, attachments_graph).

    - Por padrao, tenta embutir um print do SLA em base64 no HTML
      (mais simples e funciona bem em muitos casos).
    - Se você preferir CID inline no futuro, adaptamos para anexos inline.
    """
    subject = f"[SLA OK] {regional_nome} - {sla_percent:.1f}% ({mes_referencia}/{ano_referencia})"

    # Texto do seu modelo
    header_html = (
        "<p style='margin:0 0 10px 0;'>Prezados(a), boa tarde!</p>"
        "<p style='margin:0 0 10px 0;'>"
        "Conforme acompanhamento mensal dos indicadores de disponibilidade das infraestruturas regionais, "
        f"informamos que, no mês de <b>{mes_referencia}</b> de <b>{ano_referencia}</b>, "
        f"a integrada <b>{regional_nome}</b> atingiu o SLA acordado de <b>99%</b>. "
        f"SLA apurado: <b>{sla_percent:.1f}%</b>."
        "</p>"
        "<p style='margin:0 0 10px 0;'>"
        "Abaixo, segue o demonstrativo do SLA da regional:"
        "</p>"
    )

    img_html = _build_sla_table(
        regional_nome=regional_nome,
        mes_referencia=mes_referencia,
        ano_referencia=ano_referencia,
        sla_percent=sla_percent,
        base_dir=base_dir,
        sla_print_relative_path=sla_print_relative_path,
        usar_imagem_embutida_base64=usar_imagem_embutida_base64,
    )

    footer_html = (
        "<p style='margin:0 0 12px 0;'>"
        "Nos colocamos à disposição para apoiar tecnicamente a regional tanto na análise dos indicadores de disponibilidade "
        "quanto na interpretação dos dados de segurança."
        "</p>"
        "<p style='margin:0 0 12px 0;'><b>IMPORTANTE!</b> Esta é uma mensagem automática. Por favor, não responda.</p>"
        "<p style='margin:0 0 8px 0;'>Atenciosamente,</p>"
    )

    assinatura_html = build_signature_html(
        titulo="Governança / SLA",
        subtitulo="Indicadores de Disponibilidade | Infraestrutura",
        email="governanca.ti@gpssa.com.br",
        assinatura_gif_cid="assinatura_gif",  # se não tiver gif, só não anexar
    )

    html = (
        "<div style='font-family:Arial,sans-serif;font-size:14px;color:#111;line-height:1.45;'>"
        f"{header_html}"
        f"{img_html}"
        f"{footer_html}"
        f"{assinatura_html}"
        "</div>"
    )

    # Aqui a gente NÃO adiciona anexos, porque o print está indo em base64 no HTML.
    # Os anexos inline da assinatura (GIF) serão adicionados no main.py.
    attachments: List[Dict[str, Any]] = []
    return subject, html, attachments


def build_email_abaixo_99(
    *,
    regional_nome: str,
    mes_referencia: str,
    ano_referencia: str,
    sla_percent: float,
    base_dir: Path,
    sla_print_relative_path: str = "image/sla_print.png",
    usar_imagem_embutida_base64: bool = True,
) -> Tuple[str, str, List[Dict[str, Any]]]:
    """
    Retorna (subject, html, attachments_graph) para SLA abaixo da meta.
    """
    subject = f"SLA Não atingido - {regional_nome} ({mes_referencia}/{ano_referencia})"

    header_html = (
        "<p style='margin:0 0 10px 0;'>Prezados, boa tarde.</p>"
        "<p style='margin:0 0 10px 0;'>"
        "Conforme acompanhamento mensal dos indicadores de disponibilidade das infraestruturas regionais, "
        f"informamos que, no mês de <b>{mes_referencia}</b> de <b>{ano_referencia}</b>, "
        f"a regional <b>{regional_nome}</b> não atingiu o SLA acordado de <b>99%</b>. "
        f"SLA apurado: <b>{sla_percent:.1f}%</b>."
        "</p>"
        "<p style='margin:0 0 10px 0;'>"
        "Abaixo, segue o demonstrativo do SLA da regional:"
        "</p>"
    )

    img_html = _build_sla_table(
        regional_nome=regional_nome,
        mes_referencia=mes_referencia,
        ano_referencia=ano_referencia,
        sla_percent=sla_percent,
        base_dir=base_dir,
        sla_print_relative_path=sla_print_relative_path,
        usar_imagem_embutida_base64=usar_imagem_embutida_base64,
    )

    footer_html = (
        "<p style='margin:0 0 12px 0;'>"
        "Nos colocamos à disposição para apoiar tecnicamente a regional tanto na análise dos indicadores de disponibilidade "
        "quanto na interpretação dos dados de segurança."
        "</p>"
        "<p style='margin:0 0 12px 0;'><b>IMPORTANTE!</b> Esta é uma mensagem automática. Por favor, não responda.</p>"
        "<p style='margin:0 0 8px 0;'>Atenciosamente,</p>"
    )

    assinatura_html = build_signature_html(
        titulo="Governança / SLA",
        subtitulo="Indicadores de Disponibilidade | Infraestrutura",
        email="governanca.ti@gpssa.com.br",
        assinatura_gif_cid="assinatura_gif",
    )

    html = (
        "<div style='font-family:Arial,sans-serif;font-size:14px;color:#111;line-height:1.45;'>"
        f"{header_html}"
        f"{img_html}"
        f"{footer_html}"
        f"{assinatura_html}"
        "</div>"
    )

    attachments: List[Dict[str, Any]] = []
    return subject, html, attachments