from typing import Any, Dict, List, Optional


class SlaService:
    """
    Busca SLAs por regional no Zabbix e retorna no formato padrao.
    """

    def __init__(self, zabbix_client) -> None:
        self.zabbix = zabbix_client

    def _extract_result(self, response: Dict[str, Any]) -> Any:
        return response.get("result", response)

    def get_regional_slas(
        self,
        *,
        name_prefix: str = "V",
        name_contains: str = "REGIONAL",
    ) -> List[Dict[str, Any]]:
        resp = self.zabbix.do_request(
            "sla.get",
            {
                "output": ["slaid", "name", "status", "slo"],
            },
        )
        result = self._extract_result(resp)
        slas = result if isinstance(result, list) else []

        filtered = []
        for sla in slas:
            name = str(sla.get("name", "")).strip()
            name_upper = name.upper()
            if name_contains and name_contains.upper() not in name_upper:
                continue
            if name_prefix and not name_upper.startswith(name_prefix.upper()):
                continue
            filtered.append(sla)
        return filtered

    def get_sli_for_sla(
        self,
        *,
        slaid: str,
        period_from: int,
        period_to: int,
        sla_name: str,
    ) -> Optional[float]:
        resp = self.zabbix.do_request(
            "sla.getsli",
            {
                "slaid": slaid,
                "period_from": period_from,
                "period_to": period_to,
            },
        )
        result = self._extract_result(resp)
        if not isinstance(result, dict):
            return None

        service_ids = result.get("serviceids") or []
        sli_matrix = result.get("sli") or []
        if not service_ids or not sli_matrix:
            return None

        sli_row = sli_matrix[0] if isinstance(sli_matrix, list) else []
        if not isinstance(sli_row, list) or not sli_row:
            return None

        service_ids_str = [str(sid) for sid in service_ids]
        sli_map: Dict[str, Optional[float]] = {}
        for idx, sid in enumerate(service_ids_str):
            if idx < len(sli_row) and isinstance(sli_row[idx], dict):
                sli_map[sid] = _to_float(sli_row[idx].get("sli"))

        # Busca nomes dos services para achar o que corresponde ao SLA
        try:
            svc_resp = self.zabbix.do_request(
                "service.get",
                {
                    "serviceids": service_ids_str,
                    "output": ["serviceid", "name"],
                },
            )
            services = self._extract_result(svc_resp)
        except Exception:
            services = []

        if isinstance(services, list):
            for svc in services:
                if not isinstance(svc, dict):
                    continue
                if str(svc.get("name", "")).strip().upper() == sla_name.strip().upper():
                    svc_id = str(svc.get("serviceid"))
                    return sli_map.get(svc_id)

        # Fallback: media simples dos SLIs retornados
        values = [v for v in sli_map.values() if v is not None]
        if not values:
            return None
        return sum(values) / len(values)

    def get_regionals_sla(
        self,
        *,
        time_from: int,
        time_to: int,
        name_prefix: str = "V",
        name_contains: str = "REGIONAL",
    ) -> List[Dict[str, Any]]:
        regionals = self.get_regional_slas(
            name_prefix=name_prefix,
            name_contains=name_contains,
        )

        out: List[Dict[str, Any]] = []
        for sla in regionals:
            slaid = str(sla.get("slaid"))
            name = str(sla.get("name", "")).strip()
            sli = self.get_sli_for_sla(
                slaid=slaid,
                period_from=time_from,
                period_to=time_to,
                sla_name=name,
            )
            out.append(
                {
                    "regional": name,
                    "sla": sli if sli is not None else 0.0,
                    "periodo": "previous_month",
                }
            )
        return out


def _to_float(value: Any) -> Optional[float]:
    try:
        if value is None:
            return None
        return float(value)
    except Exception:
        return None
