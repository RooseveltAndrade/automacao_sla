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


def build_email_acima_99(
    *,
    regional_nome: str,
    mes_referencia: str,
    ano_referencia: str,
    sla_percent: float,
    base_dir: Path,
    grafana_print_relative_path: str = "image/grafana_print.png",
    usar_imagem_embutida_base64: bool = True,
) -> Tuple[str, str, List[Dict[str, Any]]]:
    """
    Retorna (subject, html, attachments_graph).

    - Por padrão, tenta embutir um print do Grafana em base64 no HTML
      (mais simples e funciona bem em muitos casos).
    - Se você preferir CID inline no futuro, adaptamos para anexos inline.
    """
    subject = f"[SLA OK] {regional_nome} - {sla_percent:.1f}% ({mes_referencia}/{ano_referencia})"

    # Texto do seu modelo
    header_html = (
        "<p style='margin:0 0 10px 0;'>Prezados, boa tarde!</p>"
        "<p style='margin:0 0 10px 0;'>"
        "Conforme acompanhamento mensal dos indicadores de disponibilidade das infraestruturas regionais, "
        f"informamos que, no mês de <b>{mes_referencia}</b> de <b>{ano_referencia}</b>, "
        f"a integrada <b>{regional_nome}</b> atingiu o SLA acordado de <b>99%</b>."
        "</p>"
        "<p style='margin:0 0 10px 0;'>"
        "Abaixo, segue a imagem com o demonstrativo do SLA da regional:"
        "</p>"
    )

    # Imagem (opcional)
    print_path = (base_dir / grafana_print_relative_path).resolve()
    img_html = "<p style='margin:0 0 12px 0; color:#777;'>(sem imagem do demonstrativo)</p>"

    if usar_imagem_embutida_base64:
        b64 = _read_png_as_b64(print_path)
        if b64:
            img_html = (
                "<p style='margin:0 0 12px 0;'>"
                f"<img src='data:image/png;base64,{b64}' "
                "alt='Demonstrativo SLA - Grafana' style='max-width:680px; height:auto; border:1px solid #ddd;'>"
                "</p>"
            )

    footer_html = (
        "<p style='margin:0 0 12px 0;'>"
        "Nos colocamos à disposição para apoiar tecnicamente a regional tanto na análise dos indicadores de disponibilidade "
        "quanto na interpretação dos dados de segurança."
        "</p>"
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