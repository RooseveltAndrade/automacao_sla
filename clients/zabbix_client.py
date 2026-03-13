import os
from typing import Any, Dict, List, Optional

import requests


class ZabbixClient:
    """
    Cliente simples para Zabbix API via token.
    """

    def __init__(self) -> None:
        self.base_url = os.getenv("ZABBIX_URL", "").strip()
        self.api_token = os.getenv("ZABBIX_TOKEN", "").strip()
        self.verify_ssl = os.getenv("ZABBIX_VERIFY_SSL", "True").strip().lower() in (
            "1", "true", "yes", "y", "on"
        )

        missing = [
            name
            for name, value in {
                "ZABBIX_URL": self.base_url,
                "ZABBIX_TOKEN": self.api_token,
            }.items()
            if not value
        ]
        if missing:
            raise RuntimeError(
                "Faltando variaveis no .env: " + ", ".join(missing)
            )

        self.api_url = self._normalize_api_url(self.base_url)
        self._request_id = 1

    def host_get(self, *, output: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        resp = self.do_request(
            "host.get",
            {"output": output or ["hostid", "name"]},
        )
        result = resp.get("result", [])
        return result if isinstance(result, list) else []

    def do_request(self, method: str, params: Dict[str, Any]) -> Dict[str, Any]:
        payload = {
            "jsonrpc": "2.0",
            "method": method,
            "params": params or {},
            "auth": self.api_token,
            "id": self._request_id,
        }
        self._request_id += 1

        r = requests.post(
            self.api_url,
            headers={"Content-Type": "application/json"},
            json=payload,
            verify=self.verify_ssl,
            timeout=60,
        )
        r.raise_for_status()
        data = r.json()
        if "error" in data:
            raise RuntimeError(data["error"])
        return data

    @staticmethod
    def _normalize_api_url(url: str) -> str:
        if url.endswith("/api_jsonrpc.php"):
            return url
        if url.endswith("/"):
            return url + "api_jsonrpc.php"
        if url.endswith("/zabbix"):
            return url + "/api_jsonrpc.php"
        return url + "/api_jsonrpc.php"
