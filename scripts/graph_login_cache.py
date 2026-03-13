import os
from pathlib import Path

import msal
from dotenv import load_dotenv

def _normalize_scopes(raw_scopes: str) -> list[str]:
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

def main() -> None:
    load_dotenv(override=True)
    tenant_id = os.getenv("M365_TENANT_ID", "").strip()
    client_id = os.getenv("M365_CLIENT_ID", "").strip()
    cache_path = os.getenv("GRAPH_AUTH_CACHE_PATH", "").strip()
    scopes_raw = os.getenv("GRAPH_DELEGATED_SCOPES", "Mail.ReadWrite").strip()

    if not tenant_id or not client_id:
        raise RuntimeError("M365_TENANT_ID e M365_CLIENT_ID sao obrigatorios.")
    if not cache_path:
        raise RuntimeError("GRAPH_AUTH_CACHE_PATH nao configurado.")

    scopes = _normalize_scopes(scopes_raw)
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    cache_file = Path(cache_path)
    cache_file.parent.mkdir(parents=True, exist_ok=True)

    cache = msal.SerializableTokenCache()
    if cache_file.exists():
        cache.deserialize(cache_file.read_text(encoding="utf-8", errors="ignore"))

    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )

    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError("Falha ao iniciar device flow.")

    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Falha no login: {result}")

    print("Scopes concedidos:", result.get("scope", ""))

    cache_file.write_text(cache.serialize(), encoding="utf-8")
    print(f"Cache salvo em: {cache_file}")

if __name__ == "__main__":
    main()
