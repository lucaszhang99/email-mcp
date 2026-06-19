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
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials, HTTPBasic, HTTPBasicCredentials
from fastapi.templating import Jinja2Templates

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


# ── Basic auth (Web UI) ───────────────────────────────────────────────────────

basic_security = HTTPBasic()


def require_basic_auth(credentials: HTTPBasicCredentials = Depends(basic_security)):
    if not secrets.compare_digest(credentials.password, settings.web_ui_password):
        raise HTTPException(
            status_code=401,
            detail="Incorrect password",
            headers={"WWW-Authenticate": "Basic"},
        )


# ── MCP SSE ───────────────────────────────────────────────────────────────────

sse_transport = SseServerTransport("/mcp/messages")


class _TransportHandledResponse(Response):
    """The MCP SSE transport writes the ASGI response itself. Returning this
    no-op Response stops FastAPI from sending a second one, which would raise
    'Unexpected ASGI message http.response.start ... after response already
    completed'."""

    async def __call__(self, scope, receive, send) -> None:
        return


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
    return _TransportHandledResponse()


@app.post("/mcp/messages", dependencies=[Depends(require_bearer)])
async def mcp_messages(request: Request):
    await sse_transport.handle_post_message(
        request.scope, request.receive, request._send
    )
    return _TransportHandledResponse()


# ── Web UI ────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index(request: Request, _=Depends(require_basic_auth)):
    accounts = await database.list_accounts()
    return templates.TemplateResponse(
        request,
        "index.html",
        {"accounts": accounts, "base_url": settings.base_url},
    )


# ── Account management API ────────────────────────────────────────────────────

@app.get("/api/accounts", dependencies=[Depends(require_basic_auth)])
async def api_list_accounts():
    return await database.list_accounts()


@app.delete("/api/accounts/{email:path}", dependencies=[Depends(require_basic_auth)])
async def api_delete_account(email: str):
    deleted = await database.delete_account(email)
    if not deleted:
        raise HTTPException(status_code=404, detail="Account not found")
    return {"ok": True}


@app.post("/api/accounts/imap", dependencies=[Depends(require_basic_auth)])
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
