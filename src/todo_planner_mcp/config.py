"""Runtime configuration.

Sources, highest priority first:
  1. Environment variables (TODO_MCP_*)
  2. Saved settings chosen at login (%USERPROFILE%\\.todo-planner-mcp\\settings.json)

Only non-secret identifiers are stored. This is a public client (no client secret);
user tokens live in a separate OS-encrypted cache (see auth.py).
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

_GUID = re.compile(r"^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$")
_DOMAIN = re.compile(r"^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$")

# Least privilege: To Do + Planner read/write, plus basic profile.
DEFAULT_SCOPES = ["User.Read", "Tasks.ReadWrite"]
# Read-only access to the user's own mailbox settings, used only to detect their time zone.
# Requested automatically with the Microsoft-published app (dynamic consent); with your own app
# registration add it to the registration and to TODO_MCP_EXTRA_SCOPES.
TIMEZONE_SCOPE = "MailboxSettings.Read"
AUTO_TZ = "auto"

# Microsoft-published multi-tenant public client ("Microsoft Graph Command Line Tools").
MICROSOFT_APP_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

AppChoice = Literal["own", "microsoft"]
APP_LABELS: dict[str, str] = {
    "own": "Own app registration (your Entra client ID)",
    "microsoft": "Microsoft-published app (Microsoft Graph Command Line Tools, no registration)",
}


class ConfigError(RuntimeError):
    pass


def env(name: str) -> str:
    """Read a TODO_MCP_* variable. Blank values and unexpanded placeholders such as
    '${user_config.client_id}' (an optional extension setting left empty) count as unset."""
    value = os.environ.get(name, "").strip()
    return "" if value.startswith("${") else value


NOT_CONFIGURED = (
    "Sign-in method not chosen yet. Ask the user which app to sign in with, then call 'login' with "
    "app='own' (their Entra app registration: needs client_id and tenant) or app='microsoft' "
    "(Microsoft Graph Command Line Tools: no registration, tenant optional). "
    "Or run `uv run todo-planner-mcp login` in a terminal."
)


def data_dir() -> Path:
    # Deliberately NOT under %APPDATA%/%LOCALAPPDATA%: packaged (MSIX) apps such as Claude Desktop
    # silently redirect AppData writes into a private per-app copy, so a login done in a terminal and
    # one done from inside Claude would otherwise land in different files.
    override = env("TODO_MCP_DATA_DIR")
    return Path(override) if override else Path.home() / ".todo-planner-mcp"


def settings_path() -> Path:
    return data_dir() / "settings.json"


def _default_cache_path() -> Path:
    override = env("TODO_MCP_CACHE_PATH")
    return Path(override) if override else data_dir() / "msal_token_cache.bin"


@dataclass(frozen=True)
class Settings:
    app: str
    client_id: str
    tenant_id: str
    timezone: str
    auto_login: bool = True
    scopes: list[str] = field(default_factory=lambda: list(DEFAULT_SCOPES))
    # Consented at interactive sign-in, but never required: features using them degrade gracefully.
    optional_scopes: list[str] = field(default_factory=list)
    cache_path: Path = field(default_factory=_default_cache_path)
    graph_base: str = "https://graph.microsoft.com/v1.0"

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"


def validate(client_id: str, tenant_id: str, timezone: str) -> None:
    if not _GUID.match(client_id):
        raise ConfigError("Client ID must be the Entra app's Application (client) ID GUID.")
    if not (_GUID.match(tenant_id) or tenant_id == "organizations" or _DOMAIN.match(tenant_id)):
        raise ConfigError("Tenant must be a tenant ID GUID, a verified domain (e.g. contoso.com.au) or 'organizations'.")
    if timezone == AUTO_TZ:
        return
    try:
        ZoneInfo(timezone)
    except (ZoneInfoNotFoundError, ValueError) as exc:
        raise ConfigError(
            f"Time zone '{timezone}' is not valid. Use 'auto' or an IANA zone (e.g. Australia/Sydney)."
        ) from exc


def read_saved() -> dict:
    try:
        return json.loads(settings_path().read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as exc:
        raise ConfigError(f"Saved settings at {settings_path()} are unreadable: {exc}") from exc


def save_settings(app: AppChoice, client_id: str | None, tenant_id: str | None, timezone: str | None) -> dict:
    """Validate and persist the sign-in choice. Returns the saved dict."""
    if app not in APP_LABELS:
        raise ConfigError("app must be 'own' or 'microsoft'.")
    saved = read_saved()
    if app == "microsoft":
        client_id = MICROSOFT_APP_CLIENT_ID
        tenant_id = (tenant_id or "organizations").strip()
    else:
        client_id = (client_id or "").strip()
        tenant_id = (tenant_id or "").strip()
        if not tenant_id:
            raise ConfigError("With your own app registration, the tenant ID or domain is required.")
    timezone = (timezone or saved.get("timezone") or AUTO_TZ).strip()
    validate(client_id, tenant_id, timezone)

    data = {
        "app": app, "client_id": client_id, "tenant_id": tenant_id, "timezone": timezone,
        "auto_login": bool(saved.get("auto_login", True)),
    }
    _write(data)
    return data


def set_auto_login(enabled: bool) -> dict:
    """Turn automatic browser sign-in on session expiry on or off."""
    saved = read_saved()
    if not saved:
        raise ConfigError(NOT_CONFIGURED)
    saved["auto_login"] = enabled
    _write(saved)
    return saved


def _write(data: dict) -> None:
    data_dir().mkdir(parents=True, exist_ok=True)
    settings_path().write_text(json.dumps(data, indent=2), encoding="utf-8")


def _env_bool(name: str) -> bool | None:
    raw = env(name).lower()
    if not raw:
        return None
    if raw in {"1", "true", "yes", "on"}:
        return True
    if raw in {"0", "false", "no", "off"}:
        return False
    raise ConfigError(f"{name} must be on/off (got '{raw}').")


def load_settings() -> Settings:
    saved = read_saved()
    client_id = env("TODO_MCP_CLIENT_ID") or saved.get("client_id", "")
    tenant_id = env("TODO_MCP_TENANT_ID") or saved.get("tenant_id", "")
    timezone = env("TODO_MCP_TIMEZONE") or saved.get("timezone") or AUTO_TZ
    extra = env("TODO_MCP_EXTRA_SCOPES").split()

    if not client_id:
        raise ConfigError(NOT_CONFIGURED)
    if client_id == MICROSOFT_APP_CLIENT_ID:
        app = "microsoft"
        tenant_id = tenant_id or "organizations"
    else:
        app = "own"
    validate(client_id, tenant_id, timezone)

    auto_login = _env_bool("TODO_MCP_AUTO_LOGIN")
    return Settings(
        app=app,
        client_id=client_id,
        tenant_id=tenant_id,
        timezone=timezone,
        auto_login=bool(saved.get("auto_login", True)) if auto_login is None else auto_login,
        scopes=list(dict.fromkeys(DEFAULT_SCOPES + extra)),
        optional_scopes=[TIMEZONE_SCOPE] if app == "microsoft" and TIMEZONE_SCOPE not in extra else [],
    )
