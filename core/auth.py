# Outlook OAuth
import os
import msal
from .config import TOKEN_CACHE_FILE, ensure_config_dir, log

class OutlookAuth:
    def __init__(self, client_id, scopes, redirect_uri, authority):
        self.client_id = client_id
        self.scopes = scopes
        self.redirect_uri = redirect_uri
        self.authority = authority
        self._cache = msal.SerializableTokenCache()
        self._app = None
        self._load_cache()
        self._build_app()

    def _load_cache(self):
        if os.path.exists(TOKEN_CACHE_FILE):
            with open(TOKEN_CACHE_FILE, "r") as f:
                self._cache.deserialize(f.read())

    def _save_cache(self):
        if self._cache.has_state_changed:
            ensure_config_dir()
            with open(TOKEN_CACHE_FILE, "w") as f:
                f.write(self._cache.serialize())

    def _build_app(self):
        self._app = msal.PublicClientApplication(
            self.client_id,
            authority=self.authority,
            token_cache=self._cache,
        )

    def get_token_silent(self):
        accounts = self._app.get_accounts()
        if not accounts:
            return None
        result = self._app.acquire_token_silent(self.scopes, account=accounts[0])
        if result and "access_token" in result:
            self._save_cache()
            return result["access_token"]
        return None

    def get_token_interactive(self):
        result = self._app.acquire_token_interactive(
            scopes=self.scopes,
            port=8400,
        )
        if result and "access_token" in result:
            self._save_cache()
            return result["access_token"]
        error = result.get("error_description", result.get("error", "Unknown error"))
        raise Exception(f"Authentication failed: {error}")

    def get_token(self):
        token = self.get_token_silent()
        if token:
            return token
        return self.get_token_interactive()

    def logout(self):
        for account in self._app.get_accounts():
            self._app.remove_account(account)
        self._save_cache()
        if os.path.exists(TOKEN_CACHE_FILE):
            os.remove(TOKEN_CACHE_FILE)

