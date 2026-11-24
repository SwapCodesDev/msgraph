"""Authentication handler for Microsoft Graph using MSAL Python."""

import sys
from pathlib import Path
import msal
from config import Config


def _load_token_cache(cache_path: Path) -> msal.SerializableTokenCache:
    """Loads serializable token cache from disk if available."""
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except Exception as e:
            print(f"[Warning] Failed to read token cache: {e}. A new cache will be created.")
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache, cache_path: Path):
    """Saves serializable token cache to disk if it was modified."""
    if cache.has_state_changed:
        try:
            with open(cache_path, "w", encoding="utf-8") as f:
                f.write(cache.serialize())
        except Exception as e:
            print(f"[Warning] Failed to write token cache: {e}")


class GraphAuthenticator:
    """Manages Microsoft Entra ID authentication for Microsoft Graph API."""

    def __init__(self):
        self.flow = Config.AUTH_FLOW
        self.client_id = Config.AZURE_CLIENT_ID
        self.tenant_id = Config.AZURE_TENANT_ID
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.cache_path = Config.TOKEN_CACHE_FILE
        self.token_cache = _load_token_cache(self.cache_path)

    def get_access_token(self, device_code_callback=None) -> str:
        """Acquires a valid access token according to the configured auth flow."""
        errors = Config.validate()
        if errors:
            print("[Error] Configuration validation failed:")
            for err in errors:
                print(f"  - {err}")
            sys.exit(1)

        if self.flow == "client_credentials":
            return self._acquire_token_client_credentials()
        elif self.flow == "device_code":
            return self._acquire_token_device_code(device_code_callback=device_code_callback)
        else:
            raise ValueError(f"Unsupported auth flow: {self.flow}")

    def _acquire_token_client_credentials(self) -> str:
        """Acquires app-only token using Client ID and Client Secret (headless daemon)."""
        app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=Config.AZURE_CLIENT_SECRET,
            authority=self.authority,
            token_cache=self.token_cache,
        )

        # Look in cache first
        result = app.acquire_token_silent(scopes=Config.APP_SCOPES, account=None)

        if not result:
            # Fetch new token using client secret
            result = app.acquire_token_for_client(scopes=Config.APP_SCOPES)

        if "access_token" in result:
            _save_token_cache(self.token_cache, self.cache_path)
            return result["access_token"]

        error_description = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Client credentials authentication failed: {error_description}")

    def _acquire_token_device_code(self, device_code_callback=None) -> str:
        """Acquires delegated token via Device Code Flow (interactive login once, cached afterwards)."""
        app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self.token_cache,
        )

        # Check existing cached accounts
        accounts = app.get_accounts()
        if accounts:
            chosen_account = accounts[0]
            result = app.acquire_token_silent(
                scopes=Config.DELEGATED_SCOPES,
                account=chosen_account,
            )
            if result and "access_token" in result:
                _save_token_cache(self.token_cache, self.cache_path)
                return result["access_token"]

        # No valid cached token, initiate Device Code flow
        flow = app.initiate_device_flow(scopes=Config.DELEGATED_SCOPES)

        if "user_code" not in flow:
            raise RuntimeError(f"Failed to initiate device flow: {flow.get('error_description', flow)}")

        if device_code_callback:
            device_code_callback(flow)
        else:
            print("=" * 60)
            print("Microsoft Graph Authentication Required (First-time or expired)")
            print("=" * 60)
            print(f"\n{flow['message']}\n")
            print("Waiting for authentication to complete in your browser...")

        result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            _save_token_cache(self.token_cache, self.cache_path)
            if not device_code_callback:
                print("\n[Success] Successfully authenticated and saved token cache.")
            return result["access_token"]

        error_desc = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Device code authentication failed: {error_desc}")

    def clear_cache(self):
        """Clears the local token cache file and memory."""
        if self.cache_path.exists():
            try:
                self.cache_path.unlink()
                print("[Info] Token cache deleted.")
            except Exception as e:
                print(f"[Warning] Failed to delete cache file: {e}")
        self.token_cache = msal.SerializableTokenCache()
