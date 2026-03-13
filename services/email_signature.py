import logging
import base64
from pathlib import Path
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)


def _read_file_b64(path: Path) -> Optional[str]:
    if not path.exists():
        return None
    try:
        return base64.b64encode(path.read_bytes()).decode("ascii")
    except Exception as e:
        logger.warning("Falha ao ler arquivo para base64 (%s): %s", path, e)
        return None


def build_signature_html(
    *,
    titulo: str = "Governança / SLA",
    subtitulo: str = "Indicadores de Disponibilidade | Infraestrutura",
    email: str = "governanca.ti@gpssa.com.br",
    teams_1_label: str = "Fale com o Felipe B. Oliveira no Teams",
    teams_1_link: str = "https://teams.microsoft.com/l/chat/0/0?users=felipe.b.oliveira@gpssa.com.br",
    teams_2_label: str = "Fale com o Michel Santos no Teams",
    teams_2_link: str = "https://teams.microsoft.com/l/chat/0/0?users=michel_santos@gpssa.com.br",
    teams_3_label: str = "Fale com o Joao Gama no Teams",
    teams_3_link: str = "https://teams.microsoft.com/l/chat/0/0?users=joao.gama@gpssa.com.br",
    assinatura_gif_cid: Optional[str] = "assinatura_gif",
) -> str:
    """
    Retorna HTML da assinatura.
    Se assinatura_gif_cid existir, será usado <img src="cid:...">.
    """
    img_html = ""
    if assinatura_gif_cid:
        img_html = (
            f"<td style='vertical-align:middle; padding-right:8px;'>"
            f"<img src='cid:{assinatura_gif_cid}' alt='Assinatura Grupo GPS' "
            f"width='145' style='width:145px; max-width:145px; height:auto; display:block;'>"
            f"</td>"
        )

    signature_html = (
        "<table role='presentation' cellpadding='0' cellspacing='0' "
        "style='margin-top:14px; border-collapse:collapse;'>"
        "<tr>"
        f"{img_html}"
        "<td style='vertical-align:middle; font-family:Arial, sans-serif; color:#1f3352; text-align:left;'>"
        f"<div style='font-size:18px; line-height:1.15; font-weight:700; margin:0;'>{titulo}</div>"
        f"<div style='margin-top:2px; color:#666666; font-size:13px; line-height:1.15;'>{subtitulo}</div>"
        f"<div style='margin-top:4px; line-height:1.15;'>"
        f"<a href='mailto:{email}' style='color:#1f4e9a; font-size:13px;'>{email}</a>"
        f"</div>"
        "<div style='margin-top:6px; font-size:12.5px; line-height:1.15;'>"
        f"<a href='{teams_1_link}' style='color:#1f4e9a;'>{teams_1_label}</a> | "
        f"<a href='{teams_2_link}' style='color:#1f4e9a;'>{teams_2_label}</a> | "
        f"<a href='{teams_3_link}' style='color:#1f4e9a;'>{teams_3_label}</a>"
        "</div>"
        "</td>"
        "</tr>"
        "</table>"
    )
    return signature_html


def build_signature_inline_attachments(
    *,
    base_dir: Path,
    gif_relative_path: str = "image/assinatura_gif.gif",
    cid: str = "assinatura_gif",
) -> List[Dict[str, Any]]:
    """
    Retorna lista de anexos inline (formato Graph) para a assinatura.
    """
    gif_path = (base_dir / gif_relative_path).resolve()
    b64 = _read_file_b64(gif_path)
    if not b64:
        return []

    return [
        {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": gif_path.name,
            "contentType": "image/gif",
            "contentBytes": b64,
            "isInline": True,
            "contentId": cid,
        }
    ]