import os
import logging
import base64
import mimetypes
from pathlib import Path
from typing import Optional, List, Dict, Any

import requests
import msal


logger = logging.getLogger(__name__)


class GraphMailClient:
    """
    Cliente simples para enviar e-mail via Microsoft Graph.
    - Usa Client Credentials (app registration)
    - Suporta HTML/texto
    - Suporta Reply-To
    - Suporta anexos inline (CID) e anexos comuns
    """

    def __init__(self):
        self.tenant_id = os.getenv("M365_TENANT_ID", "").strip()
        self.client_id = os.getenv("M365_CLIENT_ID", "").strip()
        self.client_secret = os.getenv("M365_CLIENT_SECRET", "").strip()
        self.sender_upn = os.getenv("M365_SENDER_UPN", "").strip()
        self.reply_to = os.getenv("REPLY_TO_GROUP_EMAIL", "").strip() or None
        self.use_cache_for_draft = os.getenv("GRAPH_USE_AUTH_CACHE_FOR_DRAFT", "False").strip().lower() in (
            "1", "true", "yes", "y", "on"
        )
        self.cache_path = os.getenv("GRAPH_AUTH_CACHE_PATH", "").strip()
        self.delegated_scopes = self._normalize_scopes(
            os.getenv("GRAPH_DELEGATED_SCOPES", "Mail.ReadWrite").strip()
        )

        missing = [k for k, v in {
            "M365_TENANT_ID": self.tenant_id,
            "M365_CLIENT_ID": self.client_id,
            "M365_CLIENT_SECRET": self.client_secret,
            "M365_SENDER_UPN": self.sender_upn,
        }.items() if not v]
        if missing:
            raise RuntimeError(f"Faltando variáveis no .env: {', '.join(missing)}")

        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scope = ["https://graph.microsoft.com/.default"]

    def _get_token(self) -> str:
        app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            client_credential=self.client_secret,
        )
        result = app.acquire_token_for_client(scopes=self.scope)
        token = result.get("access_token")
        if not token:
            raise RuntimeError(f"Falha ao obter token Graph: {result}")
        return token

    def _get_delegated_token_from_cache(self) -> str:
        if not self.cache_path:
            raise RuntimeError("GRAPH_AUTH_CACHE_PATH nao configurado no .env.")

        cache_file = Path(self.cache_path)
        if not cache_file.exists():
            raise RuntimeError(f"Cache MSAL nao encontrado: {cache_file}")

        cache = msal.SerializableTokenCache()
        cache.deserialize(cache_file.read_text(encoding="utf-8", errors="ignore"))

        app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=cache,
        )

        accounts = app.get_accounts()
        if not accounts:
            raise RuntimeError("Nenhuma conta encontrada no cache MSAL.")

        account = None
        if self.sender_upn:
            for acc in accounts:
                if str(acc.get("username", "")).strip().lower() == self.sender_upn.lower():
                    account = acc
                    break

        if account is None:
            account = accounts[0]

        result = app.acquire_token_silent(self.delegated_scopes, account=account)
        token = (result or {}).get("access_token")
        if not token:
            raise RuntimeError("Falha ao obter token delegado do cache MSAL.")

        return token

    @staticmethod
    def _normalize_scopes(raw_scopes: str) -> List[str]:
        scopes = []
        for part in raw_scopes.replace(",", " ").split():
            scope = part.strip()
            if not scope:
                continue
            if scope.startswith("http://") or scope.startswith("https://"):
                scopes.append(scope)
            else:
                scopes.append(f"https://graph.microsoft.com/{scope}")
        return scopes

    @staticmethod
    def _make_attachment(
        *,
        name: str,
        content_type: str,
        data_b64: str,
        is_inline: bool = False,
        content_id: Optional[str] = None
    ) -> Dict[str, Any]:
        att = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": name,
            "contentType": content_type,
            "contentBytes": data_b64,
        }
        if is_inline:
            att["isInline"] = True
            if content_id:
                att["contentId"] = content_id
        return att

    @classmethod
    def make_file_attachment(
        cls,
        file_path: str | Path,
        *,
        name: Optional[str] = None,
        content_type: Optional[str] = None,
        is_inline: bool = False,
        content_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        path = Path(file_path)
        data_b64 = base64.b64encode(path.read_bytes()).decode("ascii")
        guessed_type = content_type or mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        return cls._make_attachment(
            name=name or path.name,
            content_type=guessed_type,
            data_b64=data_b64,
            is_inline=is_inline,
            content_id=content_id,
        )

    def send_mail(
        self,
        to: List[str],
        subject: str,
        body_content: str,
        *,
        is_html: bool = True,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
        attachments: Optional[List[Dict[str, Any]]] = None,
        reply_to: Optional[str] = None,
        save_to_sent_items: bool = True,
    ) -> None:
        token = self._get_token()
        url = f"https://graph.microsoft.com/v1.0/users/{self.sender_upn}/sendMail"

        # Normaliza listas
        to = [e.strip() for e in (to or []) if str(e).strip()]
        cc = [e.strip() for e in (cc or []) if str(e).strip()]
        bcc = [e.strip() for e in (bcc or []) if str(e).strip()]

        if not to:
            raise ValueError("Lista de destinatários (to) está vazia.")

        message: Dict[str, Any] = {
            "subject": subject,
            "body": {"contentType": "HTML" if is_html else "Text", "content": body_content},
            "toRecipients": [{"emailAddress": {"address": d}} for d in to],
        }

        if cc:
            message["ccRecipients"] = [{"emailAddress": {"address": d}} for d in cc]
        if bcc:
            message["bccRecipients"] = [{"emailAddress": {"address": d}} for d in bcc]

        # Reply-To (prioriza parâmetro, senão usa env)
        reply_to_final = (reply_to or self.reply_to)
        if reply_to_final:
            message["replyTo"] = [{"emailAddress": {"address": reply_to_final}}]

        if attachments:
            message["attachments"] = attachments

        payload = {
            "message": message,
            "saveToSentItems": save_to_sent_items,
        }

        r = requests.post(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
            timeout=60,
        )

        if r.status_code != 202:
            raise RuntimeError(f"Erro sendMail: {r.status_code} - {r.text}")

        logger.info("Email enviado via Graph para %s | subject=%s", ", ".join(to), subject)

    def create_draft(
        self,
        to: List[str],
        subject: str,
        body_content: str,
        *,
        is_html: bool = True,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
        attachments: Optional[List[Dict[str, Any]]] = None,
        reply_to: Optional[str] = None,
    ) -> str:
        use_delegated = self.use_cache_for_draft
        token = self._get_delegated_token_from_cache() if use_delegated else self._get_token()
        if use_delegated:
            url = "https://graph.microsoft.com/v1.0/me/messages"
        else:
            url = f"https://graph.microsoft.com/v1.0/users/{self.sender_upn}/messages"

        # Normaliza listas
        to = [e.strip() for e in (to or []) if str(e).strip()]
        cc = [e.strip() for e in (cc or []) if str(e).strip()]
        bcc = [e.strip() for e in (bcc or []) if str(e).strip()]

        if not to:
            raise ValueError("Lista de destinatários (to) está vazia.")

        message: Dict[str, Any] = {
            "subject": subject,
            "body": {"contentType": "HTML" if is_html else "Text", "content": body_content},
            "toRecipients": [{"emailAddress": {"address": d}} for d in to],
        }

        if cc:
            message["ccRecipients"] = [{"emailAddress": {"address": d}} for d in cc]
        if bcc:
            message["bccRecipients"] = [{"emailAddress": {"address": d}} for d in bcc]

        # Reply-To (prioriza parâmetro, senão usa env)
        reply_to_final = (reply_to or self.reply_to)
        if reply_to_final:
            message["replyTo"] = [{"emailAddress": {"address": reply_to_final}}]

        if attachments:
            message["attachments"] = attachments

        r = requests.post(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=message,
            timeout=60,
        )

        if r.status_code != 201:
            raise RuntimeError(f"Erro createDraft: {r.status_code} - {r.text}")

        data = r.json()
        draft_id = data.get("id", "")
        logger.info("Rascunho criado via Graph para %s | subject=%s", ", ".join(to), subject)
        return draft_id