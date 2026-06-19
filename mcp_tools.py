"""
MCP tool definitions.
Tools exposed to Claude / ChatGPT:
  - list_accounts
  - list_emails
  - read_email
  - send_email
  - search_emails
  - delete_email
"""

from mcp.server import Server
from mcp.types import Tool, TextContent
import json

import database
import providers

server = Server("email-mcp")


# ── Tool schemas ──────────────────────────────────────────────────────────────

@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name="list_accounts",
            description="List all connected email accounts.",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        Tool(
            name="list_emails",
            description="List recent emails from an account's folder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email": {"type": "string", "description": "Email address of the account"},
                    "folder": {"type": "string", "description": "Folder name (default: inbox)", "default": "inbox"},
                    "limit": {"type": "integer", "description": "Max emails to return (default: 20)", "default": 20},
                },
                "required": ["email"],
            },
        ),
        Tool(
            name="read_email",
            description="Read the full content of a specific email.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email": {"type": "string", "description": "Account email address"},
                    "message_id": {"type": "string", "description": "Message ID from list_emails"},
                },
                "required": ["email", "message_id"],
            },
        ),
        Tool(
            name="send_email",
            description="Send an email from a connected account.",
            inputSchema={
                "type": "object",
                "properties": {
                    "from_email": {"type": "string", "description": "Sender account email address"},
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses",
                    },
                    "subject": {"type": "string", "description": "Email subject"},
                    "body": {"type": "string", "description": "Email body (plain text or HTML)"},
                    "html": {"type": "boolean", "description": "True if body is HTML", "default": False},
                },
                "required": ["from_email", "to", "subject", "body"],
            },
        ),
        Tool(
            name="search_emails",
            description="Search emails in a connected account.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email": {"type": "string", "description": "Account email address"},
                    "query": {"type": "string", "description": "Search query"},
                    "limit": {"type": "integer", "description": "Max results (default: 20)", "default": 20},
                },
                "required": ["email", "query"],
            },
        ),
        Tool(
            name="delete_email",
            description="Delete (trash) an email from a connected account.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email": {"type": "string", "description": "Account email address"},
                    "message_id": {"type": "string", "description": "Message ID from list_emails"},
                },
                "required": ["email", "message_id"],
            },
        ),
    ]


# ── Tool handlers ─────────────────────────────────────────────────────────────

@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    try:
        result = await _dispatch(name, arguments)
    except Exception as e:
        result = {"error": str(e)}

    return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]


async def _dispatch(name: str, args: dict):
    if name == "list_accounts":
        return await database.list_accounts()

    email = args.get("email") or args.get("from_email")

    if name == "list_emails":
        p = await providers.get_provider(email)
        return await p.list_emails(
            folder=args.get("folder", "inbox"),
            limit=args.get("limit", 20),
        )

    if name == "read_email":
        p = await providers.get_provider(email)
        return await p.read_email(args["message_id"])

    if name == "send_email":
        p = await providers.get_provider(email)
        return await p.send_email(
            to=args["to"],
            subject=args["subject"],
            body=args["body"],
            html=args.get("html", False),
        )

    if name == "search_emails":
        p = await providers.get_provider(email)
        return await p.search_emails(
            query=args["query"],
            limit=args.get("limit", 20),
        )

    if name == "delete_email":
        p = await providers.get_provider(email)
        return await p.delete_email(args["message_id"])

    raise ValueError(f"Unknown tool: {name}")
