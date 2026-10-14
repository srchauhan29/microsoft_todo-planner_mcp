"""Entry point: stdio MCP server, plus `login` / `logout` / `status` CLI commands.

stdout is the MCP protocol channel, so all logging goes to stderr.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from typing import Literal

import anyio

from . import app as app_module
from . import planner_tools, todo_tools  # noqa: F401  (registers tools)
from .app import READ_ONLY, WRITE, ctx, mcp, reset_ctx, resolve_timezone, tool
from .auth import AuthRequiredError
from .config import APP_LABELS, env, NOT_CONFIGURED, ConfigError, read_saved, save_settings, set_auto_login, settings_path

log = logging.getLogger("todo_planner_mcp")


def _apply_choice(app: str, client_id: str | None, tenant: str | None, timezone: str | None) -> dict:
    """Save the sign-in choice. If the app or tenant changed, clear cached tokens from the old one."""
    if env("TODO_MCP_CLIENT_ID"):
        raise ConfigError(
            "The client ID is fixed by the TODO_MCP_CLIENT_ID environment variable (e.g. in "
            "claude_desktop_config.json). Remove it there to choose the app at login."
        )
    before = read_saved()
    saved = save_settings(app, client_id, tenant, timezone)  # type: ignore[arg-type]
    changed = (before.get("client_id"), before.get("tenant_id")) != (saved["client_id"], saved["tenant_id"])
    app_module._ctx = None  # callers that hold an open HTTP client use reset_ctx() first
    if changed:
        ctx().tokens.logout()
    return saved


# ---------------------------------------------------------------- MCP tools

@tool(READ_ONLY)
async def auth_status() -> dict:
    """Show which sign-in app and Microsoft 365 account are in use and whether a valid token is available."""
    try:
        c = ctx()
    except ConfigError:
        return {"configured": False, "next": NOT_CONFIGURED, "options": APP_LABELS}
    user = c.tokens.signed_in_user()
    try:
        await anyio.to_thread.run_sync(lambda: c.tokens.get_token(interactive=False))
        token_ok = True
    except AuthRequiredError:
        token_ok = False
    return {
        "configured": True,
        "app": APP_LABELS[c.settings.app],
        "clientId": c.settings.client_id,
        "tenant": c.settings.tenant_id,
        "signedInAs": user,
        "tokenValid": token_ok,
        "scopes": c.settings.scopes,
        "timezoneSetting": c.settings.timezone,
        **(dict(zip(("timezone", "timezoneSource"), await resolve_timezone())) if token_ok else {}),
        "autoLogin": c.settings.auto_login,
    }


@tool(WRITE)
async def auto_login(enabled: bool) -> dict:
    """Turn automatic sign-in on or off. When on (default), an expired Microsoft 365 session
    opens the browser sign-in automatically and the pending request continues afterwards."""
    if env("TODO_MCP_AUTO_LOGIN"):
        raise ConfigError("Auto sign-in is fixed by the TODO_MCP_AUTO_LOGIN environment variable.")
    await anyio.to_thread.run_sync(set_auto_login, enabled)
    await reset_ctx()
    return {"autoLogin": enabled}


@tool(WRITE)
async def login(
    app: Literal["own", "microsoft"] | None = None,
    client_id: str | None = None,
    tenant: str | None = None,
    timezone: str | None = None,
) -> dict:
    """Sign in to Microsoft 365. Opens the system browser on this computer for sign-in/MFA.

    On first use, or when the user wants to switch, ASK THE USER which app to sign in with:
      1. app='own'       - their own Entra app registration. Needs client_id (GUID) and tenant
                           (tenant ID or verified domain).
      2. app='microsoft' - Microsoft-published 'Microsoft Graph Command Line Tools' app, no
                           registration needed. tenant is optional (default 'organizations' =
                           any work account); suggest pinning their domain to avoid signing in
                           to the wrong (e.g. customer) tenant.
    timezone: 'auto' (default - the user's Microsoft 365 time zone, else this computer's) or an
    IANA zone such as Australia/Sydney. Omit app to reuse the saved choice."""
    if app:
        await reset_ctx()
        await anyio.to_thread.run_sync(_apply_choice, app, client_id, tenant, timezone)
    c = ctx()
    result = await anyio.to_thread.run_sync(c.tokens.login_interactive)
    return {**result, "app": APP_LABELS[c.settings.app], "tenantSetting": c.settings.tenant_id}


# ---------------------------------------------------------------- CLI

def _ask(prompt: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip()
    return value or (default or "")


def _interactive_choice() -> None:
    saved = read_saved()
    current = {"own": "1", "microsoft": "2"}.get(saved.get("app", ""))
    print("\nHow do you want to sign in to Microsoft 365?")
    print(f"  1. {APP_LABELS['own']}")
    print(f"  2. {APP_LABELS['microsoft']}")
    if current:
        print(f"Current: option {current} (tenant {saved.get('tenant_id')})")
    choice = _ask("Choose 1 or 2", current)
    while choice not in {"1", "2"}:
        choice = _ask("Please enter 1 or 2", current)

    same = saved.get("app") == ("own" if choice == "1" else "microsoft")
    if choice == "2":
        print("Tip: enter your domain to make sure you can only sign in to your own tenant.")
    for attempt in range(3):
        if choice == "1":
            client_id = _ask("Application (client) ID", saved.get("client_id") if same else None)
            tenant = _ask("Tenant ID or verified domain (e.g. contoso.com.au)", saved.get("tenant_id") if same else None)
            app = "own"
        else:
            prior = saved.get("tenant_id") if same else None
            tenant = _ask("Tenant ID or domain (Enter = any work account)", prior or "organizations")
            client_id, app = None, "microsoft"
        timezone = _ask("Time zone ('auto' = detect from Microsoft 365, or e.g. Australia/Sydney)", saved.get("timezone") or "auto")
        try:
            _apply_choice(app, client_id, tenant, timezone)
            break
        except ConfigError as exc:
            if attempt == 2:
                raise
            print(f"  {exc} Please try again.")
    print(f"Saved to {settings_path()}\n")


def _cli(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(prog="todo-planner-mcp", description="No command = run the MCP server.")
    sub = parser.add_subparsers(dest="command", required=True)
    p_login = sub.add_parser("login", help="Choose the sign-in app and sign in")
    p_login.add_argument("--app", choices=["own", "microsoft"], help="Skip the menu: 'own' or 'microsoft'")
    p_login.add_argument("--client-id", help="Client ID (with --app own)")
    p_login.add_argument("--tenant", help="Tenant ID or domain")
    p_login.add_argument("--timezone", help="'auto' (default) or an IANA time zone, e.g. Australia/Sydney")
    sub.add_parser("logout", help="Remove cached tokens")
    sub.add_parser("status", help="Show sign-in status")
    p_auto = sub.add_parser("auto-login", help="Open browser sign-in automatically when the session expires")
    p_auto.add_argument("state", choices=["on", "off"])
    args = parser.parse_args(argv)

    try:
        if args.command == "login":
            if args.app:
                _apply_choice(args.app, args.client_id, args.tenant, args.timezone)
            elif env("TODO_MCP_CLIENT_ID"):
                print("Using client ID from TODO_MCP_CLIENT_ID environment variable.")
            elif sys.stdin.isatty():
                _interactive_choice()
            elif not read_saved():
                raise ConfigError("No sign-in app chosen. Re-run with --app own|microsoft.")
            c = ctx()
            print(f"Opening browser to sign in ({APP_LABELS[c.settings.app]}, tenant {c.settings.tenant_id})...")
            result = c.tokens.login_interactive()
            print(f"Signed in as {result['signedInAs']} (tenant {result['tenantId']}).")
        elif args.command == "logout":
            print(f"Removed {ctx().tokens.logout()} cached account(s).")
        elif args.command == "auto-login":
            set_auto_login(args.state == "on")
            print(f"Auto sign-in on session expiry: {args.state}")
        elif args.command == "status":
            c = ctx()
            try:
                c.tokens.get_token(interactive=False)
                state = "token valid"
            except AuthRequiredError:
                state = "sign-in required"
            print(f"App:     {APP_LABELS[c.settings.app]}")
            print(f"Tenant:  {c.settings.tenant_id}")
            print(f"Auto sign-in on expiry: {'on' if c.settings.auto_login else 'off'}")
            if state == "token valid":
                name, source = anyio.run(resolve_timezone)
                print(f"Time zone: {name} (from {source}; setting '{c.settings.timezone}')")
            else:
                print(f"Time zone: setting '{c.settings.timezone}' (resolved after sign-in)")
            print(f"Account: {c.tokens.signed_in_user() or '(none)'} - {state}")
    except (ConfigError, AuthRequiredError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except (KeyboardInterrupt, EOFError):
        print("\nCancelled.", file=sys.stderr)
        return 130
    return 0


def main() -> None:
    logging.basicConfig(
        level=getattr(logging, (env("TODO_MCP_LOG_LEVEL") or "INFO").upper(), logging.INFO),
        stream=sys.stderr,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    # httpx logs full request URLs at INFO; keep it quiet.
    logging.getLogger("httpx").setLevel(logging.WARNING)
    if len(sys.argv) > 1:
        logging.getLogger().setLevel(logging.WARNING)  # keep the CLI output clean
        sys.exit(_cli(sys.argv[1:]))
    log.info("Starting todo-planner MCP server (stdio)")
    mcp.run()


if __name__ == "__main__":
    main()
