from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import msal

SCOPES = ["https://graph.microsoft.com/Mail.Read", "offline_access"]
AUTHORITY = "https://login.microsoftonline.com/common"

# Redacted/placeholder. Real client_id comes from credentials.json. This is
# never used in production flows; auth_run always reads credentials.json.
DEFAULT_CLIENT_ID = ""


def credentials_path(state_dir: Path) -> Path:
    in_state = state_dir / "credentials.json"
    in_cwd = Path.cwd() / "credentials.json"
    if in_cwd.exists():
        return in_cwd
    return in_state


def token_path(state_dir: Path, account: str = "default") -> Path:
    return state_dir / account / "token.json"


def load_client_id(state_dir: Path, client_file: str | None = None) -> str:
    creds_path = credentials_path(state_dir)
    if client_file:
        creds_path = Path(client_file)
    if not creds_path.exists():
        raise SystemExit(
            f"No app credentials found at {creds_path}.\n"
            "Register an app in the Microsoft Entra / Azure portal (or the Azure CLI), "
            "enable the Mail.Read delegated permission, and save its Application "
            "(client) ID in credentials.json as {\"client_id\": \"<id>\"} (see README "
            "and the outlook-triage-setup skill)."
        )
    raw = creds_path.read_text(encoding="utf-8")
    try:
        data = json.loads(raw)
        if isinstance(data, dict) and data.get("client_id"):
            return str(data["client_id"]).strip()
        # Some users paste the bare client ID as the file contents.
        text = raw.strip().strip('"')
        if text and text.count("-") == 4:
            return text
    except json.JSONDecodeError:
        text = raw.strip().strip('"')
        if text and text.count("-") == 4:
            return text
    raise SystemExit(
        f"No 'client_id' found in {creds_path}. Save the app's Application (client) "
        "ID as JSON: {\"client_id\": \"<id>\"}."
    )


def build_app(state_dir: Path, client_id: str, account: str = "default") -> msal.PublicClientApplication:
    state_dir.mkdir(parents=True, exist_ok=True)
    tok_path = token_path(state_dir, account)
    tok_path.parent.mkdir(parents=True, exist_ok=True)
    cache = msal.SerializableTokenCache()
    if tok_path.exists():
        try:
            cache.deserialize(tok_path.read_text(encoding="utf-8"))
        except Exception:
            cache = msal.SerializableTokenCache()
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=AUTHORITY,
        token_cache=cache,
    )
    app._tok_path = tok_path  # type: ignore[attr-defined]
    return app


def _save_cache(app: msal.PublicClientApplication) -> None:
    cache = getattr(app, "token_cache", None)
    tok_path = getattr(app, "_tok_path", None)
    if cache and tok_path:
        tok_path.write_text(cache.serialize(), encoding="utf-8")
        try:
            os.chmod(tok_path, 0o600)
        except OSError:
            pass


def get_access_token(app: msal.PublicClientApplication) -> str:
    """Return a valid access token, re-authenticating via device flow if needed."""
    result = app.acquire_token_silent(SCOPES, account=None)
    if result:
        _save_cache(app)
        return result["access_token"]
    result = authenticate(app)
    return result["access_token"]


def authenticate(app: msal.PublicClientApplication) -> dict[str, Any]:
    """Run the OAuth 2.0 device code flow interactively. Prints URL + code."""
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_message" not in flow:
        raise SystemExit("Device flow failed to start: " + str(flow))
    print("\nTo sign in, open the URL below and enter the code:")
    print(flow["verification_uri"])
    print(f"Code: {flow['user_code']}\n")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise SystemExit("Authentication failed: " + json.dumps(result.get("error_description") or result))
    _save_cache(app)
    return result
