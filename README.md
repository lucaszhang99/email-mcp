# mcp-email

A self-hosted MCP server that gives Claude (and ChatGPT) access to your email accounts via HTTP/SSE. Supports Outlook OAuth2, iCloud, and generic IMAP/SMTP.

## Features

- **MCP SSE transport** — works as a custom connector in Claude, ChatGPT developer mode, and any MCP-compatible client
- **Outlook / Hotmail** — OAuth2 via Microsoft Graph API
- **iCloud** — IMAP with App-Specific Password
- **Generic IMAP** — any email provider
- **Encrypted storage** — OAuth tokens stored in SQLite encrypted with Fernet (AES-128)
- **Web UI** — manage accounts at `http://your-domain.com`
- **Docker + Caddy** — one-command deploy with auto HTTPS

## Tools exposed to Claude

| Tool | Description |
|------|-------------|
| `list_accounts` | List all connected email accounts |
| `list_emails` | List recent emails from a folder |
| `read_email` | Read full email content |
| `send_email` | Send an email |
| `search_emails` | Search across an account |
| `delete_email` | Move email to trash |

## Setup

### 1. Clone & configure

```bash
git clone https://github.com/lucaszhang99/email-mcp.git
cd email-mcp
cp .env.example .env
```

Edit `.env`:

```env
MCP_BEARER_TOKEN=<random strong token>
ENCRYPTION_KEY=<generate below>
BASE_URL=https://your-domain.com
WEB_UI_PASSWORD=<your password>

OUTLOOK_CLIENT_ID=<from Azure Portal>
OUTLOOK_CLIENT_SECRET=<from Azure Portal>
OUTLOOK_TENANT_ID=common
```

Generate `ENCRYPTION_KEY`:
```bash
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

### 2. Azure App Registration (Outlook)

1. Go to [portal.azure.com](https://portal.azure.com) → App registrations → New registration
2. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
3. Add redirect URI: `https://your-domain.com/auth/outlook/callback`
4. API permissions → Microsoft Graph → Delegated:
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `User.Read`, `email`, `offline_access`
5. Certificates & secrets → New client secret → copy the value

### 3. Run locally

```bash
pip install -r requirements.txt
uvicorn main:app --port 8080 --reload
```

Open `http://localhost:8080`, log in with `WEB_UI_PASSWORD`, click **+ Outlook** to authorize.

### 4. Deploy with Docker

Edit `Caddyfile` — replace `your-domain.com` with your actual domain, then:

```bash
docker compose up -d
```

Caddy handles HTTPS automatically via Let's Encrypt.

#### Cloudflare Tunnel (alternative, no VPS domain needed)

```bash
cloudflared tunnel --url http://localhost:8080
```

Set `BASE_URL` in `.env` to the tunnel URL, update the Azure redirect URI to match.

## Connect to Claude

Settings → Integrations → Add custom connector:

- **URL**: `https://your-domain.com/mcp/sse`
- **Header**: `Authorization: Bearer <MCP_BEARER_TOKEN>`

## Adding iCloud

1. Go to [appleid.apple.com](https://appleid.apple.com) → Sign-In and Security → App-Specific Passwords → Generate
2. In the Web UI click **+ iCloud**, enter your iCloud email and the app-specific password

## Project structure

```
mcp-email/
├── main.py           # FastAPI app — MCP SSE endpoint + Web UI + OAuth callbacks
├── mcp_tools.py      # MCP tool definitions
├── providers.py      # Outlook (Graph API) + IMAP/SMTP implementations
├── outlook_auth.py   # Microsoft OAuth2 flow (MSAL)
├── database.py       # SQLite + Fernet encrypted storage
├── config.py         # Settings (pydantic-settings)
├── templates/
│   └── index.html    # Web UI
├── Dockerfile
├── docker-compose.yml
└── Caddyfile
```
