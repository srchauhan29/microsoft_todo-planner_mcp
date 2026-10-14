"""Delegated (signed-in user) authentication with MSAL.

- Public client app: no client secret is stored anywhere.
- Tokens are cached with msal-extensions, encrypted at rest with the
  OS user's key (DPAPI on Windows, Keychain on macOS, libsecret on Linux).
- Interactive sign-in uses the system browser with a localhost redirect,
  so the tenant's MFA and Conditional Access policies apply as normal.
- Auto sign-in: when the session can't be refreshed silently (expired, revoked,
  Conditional Access sign-in frequency, password change), the browser sign-in
  opens automatically and the pending tool call continues once it completes.
"""

from __future__ import annotations

import logging
import threading
import time

import msal
from msal_extensions import PersistedTokenCache, build_encrypted_persistence

from .config import Settings

log = logging.getLogger(__name__)

AUTO_LOGIN_TIMEOUT = 180  # seconds the browser sign-in waits for the user
AUTO_LOGIN_COOLDOWN = 120  # seconds before re-opening after an abandoned sign-in


class AuthRequiredError(RuntimeError):
    """Raised when no valid token is available without user interaction."""


class TokenProvider:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        settings.cache_path.parent.mkdir(parents=True, exist_ok=True)
        persistence = build_encrypted_persistence(str(settings.cache_path))
        self._cache = PersistedTokenCache(persistence)
        self._app = msal.PublicClientApplication(
            settings.client_id,
            authority=settings.authority,
            token_cache=self._cache,
        )
        self._lock = threading.Lock()
        self._auto_login_blocked_until = 0.0

    def _account(self):
        accounts = self._app.get_accounts()
        return accounts[0] if accounts else None

    def signed_in_user(self) -> str | None:
        account = self._account()
        return account.get("username") if account else None

    def _silent_locked(self) -> str | None:
        account = self._account()
        if not account:
            return None
        result = self._app.acquire_token_silent(self._settings.scopes, account=account)
        if result and "access_token" in result:
            return result["access_token"]
        if result and "error" in result:
            log.warning("Silent token acquisition failed: %s", result.get("error"))
        return None

    def _interactive_locked(self, timeout: int) -> dict:
        account = self._account()
        result = self._app.acquire_token_interactive(
            self._settings.scopes + self._settings.optional_scopes,
            login_hint=account.get("username") if account else None,
            prompt=None if account else "select_account",
            timeout=timeout,
        )
        if "access_token" not in result:
            # error_description can include correlation IDs useful for support; never includes tokens.
            raise AuthRequiredError(f"Sign-in failed: {result.get('error')}: {result.get('error_description')}")
        return result

    def try_silent(self, scopes: list[str]) -> str | None:
        """Token for extra scopes if already consented; never prompts."""
        with self._lock:
            account = self._account()
            if not account:
                return None
            result = self._app.acquire_token_silent(scopes, account=account)
            return result.get("access_token") if result else None

    def get_token(self, interactive: bool = True) -> str:
        """Return an access token.

        Tries silent refresh first. If that fails and auto sign-in is enabled, opens the
        browser sign-in and waits for it (one window at a time; concurrent callers wait on
        the lock and then reuse the new token). interactive=False never opens a browser.
        """
        with self._lock:
            token = self._silent_locked()
            if token:
                return token
            if not (interactive and self._settings.auto_login):
                raise AuthRequiredError(
                    "Not signed in to Microsoft 365 (or the session expired). "
                    "Call the 'login' tool, or run `uv run todo-planner-mcp login` in a terminal."
                )
            if time.monotonic() < self._auto_login_blocked_until:
                raise AuthRequiredError(
                    "Microsoft 365 sign-in was opened recently but not completed. "
                    "Call the 'login' tool when the user is ready to sign in."
                )
            log.info("Session expired or missing; opening browser sign-in automatically")
            try:
                result = self._interactive_locked(AUTO_LOGIN_TIMEOUT)
            except Exception as exc:
                self._auto_login_blocked_until = time.monotonic() + AUTO_LOGIN_COOLDOWN
                raise AuthRequiredError(
                    f"The Microsoft 365 sign-in window was opened but sign-in did not complete ({exc}). "
                    "Ask the user to sign in, then call the 'login' tool."
                ) from exc
            self._auto_login_blocked_until = 0.0
            return result["access_token"]

    def login_interactive(self, timeout: int = AUTO_LOGIN_TIMEOUT) -> dict[str, str]:
        """Open the system browser for sign-in. Returns the signed-in username and tenant ID."""
        with self._lock:
            result = self._interactive_locked(timeout)
            self._auto_login_blocked_until = 0.0
        claims = result.get("id_token_claims", {})
        return {"signedInAs": claims.get("preferred_username", "unknown"), "tenantId": claims.get("tid", "unknown")}

    def logout(self) -> int:
        """Remove all cached accounts/tokens. Returns number of accounts removed."""
        with self._lock:
            accounts = self._app.get_accounts()
            for acct in accounts:
                self._app.remove_account(acct)
            return len(accounts)
