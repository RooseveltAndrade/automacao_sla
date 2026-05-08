import atexit
import logging
import os
from pathlib import Path

import msal
import requests


logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
ROOT_DIR = Path(__file__).resolve().parent.parent
SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Chat.ReadWrite",
    "https://graph.microsoft.com/ChatMessage.Send",
]


def _parse_env_list(name: str) -> list[str]:
    raw_value = str(os.getenv(name, "")).strip()
    return [item.strip() for item in raw_value.split(",") if item.strip()]


class TeamsClient:
    def __init__(self):
        self.tenant_id = str(os.getenv("M365_TENANT_ID", "")).strip()
        self.client_id = str(os.getenv("M365_CLIENT_ID", "")).strip()
        self.sender_upn = str(os.getenv("M365_SENDER_UPN", "")).strip()
        raw_cache_path = str(os.getenv("GRAPH_AUTH_CACHE_PATH", "")).strip()
        self.cache_path = str((ROOT_DIR / raw_cache_path).resolve()) if raw_cache_path and not Path(raw_cache_path).is_absolute() else raw_cache_path
        self.enabled = str(os.getenv("ENABLE_TEAMS_SUMMARY", "False")).strip().lower() in (
            "1",
            "true",
            "yes",
            "y",
            "on",
        )
        self._token_cache = None
        self._sender_cache = None
        self._user_cache: dict[str, dict] = {}

    def is_enabled(self) -> bool:
        return self.enabled

    def get_summary_recipients(self) -> list[str]:
        test_override = _parse_env_list("TEAMS_SUMMARY_SAFE_TEST_TO")
        if test_override:
            return test_override
        return _parse_env_list("TEAMS_SUMMARY_RECIPIENTS")

    def _load_cache(self) -> msal.SerializableTokenCache:
        cache = msal.SerializableTokenCache()
        cache_file = Path(self.cache_path)
        if cache_file.exists():
            cache.deserialize(cache_file.read_text(encoding="utf-8", errors="ignore"))

        def save_cache() -> None:
            if cache.has_state_changed and self.cache_path:
                cache_file.parent.mkdir(parents=True, exist_ok=True)
                cache_file.write_text(cache.serialize(), encoding="utf-8")

        atexit.register(save_cache)
        return cache

    def _get_token(self) -> str | None:
        if self._token_cache:
            return self._token_cache

        if not all([self.tenant_id, self.client_id, self.sender_upn, self.cache_path]):
            logger.error(
                "Teams summary: configuracao incompleta. Defina M365_TENANT_ID, M365_CLIENT_ID, M365_SENDER_UPN e GRAPH_AUTH_CACHE_PATH."
            )
            return None

        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=authority,
            token_cache=self._load_cache(),
        )

        accounts = app.get_accounts(username=self.sender_upn)
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            token = (result or {}).get("access_token")
            if token:
                self._token_cache = token
                return token

        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            token = (result or {}).get("access_token")
            if token:
                self._token_cache = token
                return token

        logger.error(
            "Teams summary: token delegado indisponivel no cache. Rode scripts/graph_login_cache.py com escopos de chat habilitados."
        )
        return None

    def _graph_get(self, url: str, token: str) -> requests.Response:
        return requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)

    def _graph_post(self, url: str, token: str, payload: dict) -> requests.Response:
        return requests.post(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
            timeout=60,
        )

    def _resolve_user(self, token: str, email: str) -> dict | None:
        cache_key = email.lower()
        if cache_key in self._user_cache:
            return self._user_cache[cache_key]

        response = self._graph_get(
            f"{GRAPH_BASE}/users/{email}?$select=id,displayName,userPrincipalName",
            token,
        )
        if response.status_code != 200:
            logger.error(
                "Teams summary: falha ao resolver usuario %s: HTTP %s | %s",
                email,
                response.status_code,
                response.text,
            )
            return None

        data = response.json()
        self._user_cache[cache_key] = data
        return data

    def _resolve_sender(self, token: str) -> dict | None:
        if self._sender_cache:
            return self._sender_cache
        self._sender_cache = self._resolve_user(token, self.sender_upn)
        return self._sender_cache

    def _get_or_create_chat(self, token: str, sender_id: str, target_id: str) -> str | None:
        payload = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"{GRAPH_BASE}/users('{sender_id}')",
                },
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"{GRAPH_BASE}/users('{target_id}')",
                },
            ],
        }
        response = self._graph_post(f"{GRAPH_BASE}/chats", token, payload)
        if response.status_code not in (200, 201, 202):
            logger.error(
                "Teams summary: falha ao criar/obter chat 1:1: HTTP %s | %s",
                response.status_code,
                response.text,
            )
            return None

        chat_id = response.json().get("id")
        if not chat_id:
            logger.error("Teams summary: chat 1:1 retornado sem id.")
            return None
        return chat_id

    def send_summary(self, recipients: list[str], html_message: str) -> bool:
        if not self.is_enabled():
            return True

        recipients = [item for item in recipients if item]
        if not recipients:
            logger.info("Teams summary: nenhum destinatario configurado.")
            return False

        token = self._get_token()
        if not token:
            return False

        sender = self._resolve_sender(token)
        if not sender or not sender.get("id"):
            logger.error("Teams summary: nao foi possivel resolver o remetente.")
            return False

        sender_id = sender["id"]
        success = True
        for recipient in recipients:
            user = self._resolve_user(token, recipient)
            if not user or not user.get("id"):
                success = False
                continue
            chat_id = self._get_or_create_chat(token, sender_id, user["id"])
            if not chat_id:
                success = False
                continue
            payload = {"body": {"contentType": "html", "content": html_message}}
            response = self._graph_post(f"{GRAPH_BASE}/chats/{chat_id}/messages", token, payload)
            if response.status_code not in (200, 201, 202):
                logger.error(
                    "Teams summary: falha ao enviar mensagem para %s: HTTP %s | %s",
                    recipient,
                    response.status_code,
                    response.text,
                )
                success = False
                continue
            logger.info("Teams summary enviado para %s", recipient)

        return success