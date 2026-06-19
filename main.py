"""
Entry point.
- /mcp/sse       → MCP SSE endpoint (Bearer token required)
- /mcp/messages  → MCP message POST endpoint
- /              → Web UI (Basic Auth)
- /auth/outlook/ → OAuth2 callbacks
- /api/accounts  → Account management REST API
"""

import secrets
import uuid
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, HTTPException, Depends, Form
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.templating import Jinja2Templates
from starlette.middleware.base import BaseHTTPMiddleware

from mcp.server.sse import SseServerTransport

import database
import outlook_auth
from config import get_settings
from mcp_tools import server as mcp_server

settings = get_settings()

# ── OAuth state store (in-memory, ephemeral) ──────────────────────────────────
_oauth_states: dict[str, str] = {}  # state → provider


# ── Lifespan ──────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    await database.init_db()
    yield


# ── App ───────────────────────────────────────────────────────────────────────

app = FastAPI(lifespan=lifespan, docs_url=None, redoc_url=None)
templates = Jinja2Templates(directory="templates")


# ── Bearer token auth (MCP endpoints) ────────────────────────────────────────

bearer_scheme = HTTPBearer(auto_error=False)


def require_bearer(credentials: HTTPAuthorizationCredentials = Depends(bearer_scheme)):
    if not credentials or not secrets.compare_digest(
        credentials.credentials, settings.mcp_bearer_token
    ):
        raise HTTPException(status_code=401, detail="Invalid or missing Bearer token")


# ── Basic auth middleware (Web UI) ────────────────────────────────────────────

class WebUIAuth(BaseHTTPMiddleware):
    WEB_PATHS = {"/", "/api/accounts"}

    async def dispatch(self, request: Request, call_next):
        # Only protect web UI paths, not MCP or OAuth
        if not any(request.url.path.startswith(p) for p in ["/", "/api/"]):
            return await call_next(request)
        if request.url.path.startswith("/auth/") or request.url.path.startswith("/mcp/"):
            return await call_next(request)

        auth = request.headers.get("Authorization", "")
        if auth.startswith("Basic "):
            import base64
            try:
                decoded = base64.b64decode(auth[6:]).decode()
                _, password = decoded.split(":", 1)
                if secrets.compare_digest(password, settings.web_ui_password):
                    return await call_next(request)
            except Exception:
                pass

        return HTMLResponse(
            status_code=401,
            content="Unauthorized",
            headers={"WWW-Authenticate": 'Basic realm="MCP Email Manager"'},
        )


app.add_middleware(WebUIAuth)


# ── MCP SSE ───────────────────────────────────────────────────────────────────

sse_transport = SseServerTransport("/mcp/messages")


@app.get("/mcp/sse", dependencies=[Depends(require_bearer)])
async def mcp_sse(request: Request):
    async with sse_transport.connect_sse(
        request.scope, request.receive, request._send
    ) as streams:
        await mcp_server.run(
            streams[0],
            streams[1],
            mcp_server.create_initialization_options(),
        )


@app.post("/mcp/messages", dependencies=[Depends(require_bearer)])
async def mcp_messages(request: Request):
    return await sse_transport.handle_post_message(
        request.scope, request.receive, request._send
    )


# ── Web UI ────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    accounts = await database.list_accounts()
    return templates.TemplateResponse(
        "index.html",
        {"request": request, "accounts": accounts, "base_url": settings.base_url},
    )


# ── Account management API ────────────────────────────────────────────────────

@app.get("/api/accounts")
async def api_list_accounts():
    return await database.list_accounts()


@app.delete("/api/accounts/{email:path}")
async def api_delete_account(email: str):
    deleted = await database.delete_account(email)
    if not deleted:
        raise HTTPException(status_code=404, detail="Account not found")
    return {"ok": True}


@app.post("/api/accounts/imap")
async def api_add_imap_account(
    email: str = Form(...),
    password: str = Form(...),
    imap_host: str = Form(...),
    imap_port: int = Form(993),
    smtp_host: str = Form(...),
    smtp_port: int = Form(587),
    display_name: str = Form(""),
    provider: str = Form("imap"),  # "imap" or "icloud"
):
    data = {
        "provider": provider,
        "username": email,
        "password": password,
        "imap_host": imap_host,
        "imap_port": imap_port,
        "smtp_host": smtp_host,
        "smtp_port": smtp_port,
    }
    await database.upsert_account(email, provider, data, display_name or email)
    return RedirectResponse("/", status_code=303)


# ── Outlook OAuth2 ────────────────────────────────────────────────────────────

@app.get("/auth/outlook/start")
async def outlook_start():
    state = str(uuid.uuid4())
    _oauth_states[state] = "outlook"
    url = outlook_auth.get_auth_url(state)
    return RedirectResponse(url)


@app.get("/auth/outlook/callback")
async def outlook_callback(request: Request):
    code = request.query_params.get("code")
    state = request.query_params.get("state")
    error = request.query_params.get("error")

    if error:
        return HTMLResponse(f"<p>OAuth error: {error}</p><a href='/'>Back</a>", status_code=400)
    if not code or state not in _oauth_states:
        return HTMLResponse("<p>Invalid state.</p><a href='/'>Back</a>", status_code=400)

    _oauth_states.pop(state)

    try:
        tokens = outlook_auth.exchange_code(code)
    except ValueError as e:
        return HTMLResponse(f"<p>Token error: {e}</p><a href='/'>Back</a>", status_code=400)

    # Get user email from Microsoft Graph
    import httpx
    async with httpx.AsyncClient() as client:
        r = await client.get(
            "https://graph.microsoft.com/v1.0/me",
            headers={"Authorization": f"Bearer {tokens['access_token']}"},
        )
        user = r.json()

    email = user.get("mail") or user.get("userPrincipalName", "")
    display_name = user.get("displayName", email)

    data = {**tokens, "provider": "outlook"}
    await database.upsert_account(email, "outlook", data, display_name)

    return RedirectResponse("/", status_code=303)
