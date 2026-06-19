"""
Outlook / Microsoft OAuth2 using MSAL.
Flow:
  1. /auth/outlook/start  → redirect user to Microsoft login
  2. Microsoft redirects back to /auth/outlook/callback with code
  3. Exchange code for tokens, store encrypted in DB
"""

import msal
from config import get_settings

SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/User.Read",
    "offline_access",
]


def _build_msal_app() -> msal.ConfidentialClientApplication:
    s = get_settings()
    return msal.ConfidentialClientApplication(
        client_id=s.outlook_client_id,
        client_credential=s.outlook_client_secret,
        authority=f"https://login.microsoftonline.com/{s.outlook_tenant_id}",
    )


def get_auth_url(state: str) -> str:
    """Return the Microsoft login URL to redirect the user to."""
    app = _build_msal_app()
    return app.get_authorization_request_url(
        scopes=SCOPES,
        state=state,
        redirect_uri=get_settings().outlook_redirect_uri,
    )


def exchange_code(code: str) -> dict:
    """
    Exchange an auth code for tokens.
    Returns dict with access_token, refresh_token, etc.
    Raises ValueError on failure.
    """
    app = _build_msal_app()
    result = app.acquire_token_by_authorization_code(
        code=code,
        scopes=SCOPES,
        redirect_uri=get_settings().outlook_redirect_uri,
    )
    if "error" in result:
        raise ValueError(f"Token exchange failed: {result.get('error_description', result['error'])}")
    return result


def refresh_tokens(refresh_token: str) -> dict:
    """
    Use a refresh token to get a new access token.
    Returns updated token dict.
    Raises ValueError on failure.
    """
    app = _build_msal_app()
    result = app.acquire_token_by_refresh_token(
        refresh_token=refresh_token,
        scopes=SCOPES,
    )
    if "error" in result:
        raise ValueError(f"Token refresh failed: {result.get('error_description', result['error'])}")
    return result
