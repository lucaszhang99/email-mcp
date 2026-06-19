"""
Email providers: Outlook (Microsoft Graph API) and generic IMAP/SMTP.
Each provider exposes: list_emails, read_email, send_email, search_emails.
"""

import asyncio
import base64
import email as email_lib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional

import httpx
import aioimaplib
import aiosmtplib

import database
import outlook_auth


# ── Helpers ──────────────────────────────────────────────────────────────────

def _truncate(text: str, max_len: int = 200) -> str:
    return text[:max_len] + "…" if len(text) > max_len else text


# ── Outlook (Microsoft Graph) ─────────────────────────────────────────────────

class OutlookProvider:
    GRAPH = "https://graph.microsoft.com/v1.0"

    def __init__(self, email: str):
        self.email = email

    async def _headers(self) -> dict:
        data = await database.get_account_data(self.email)
        if not data:
            raise ValueError(f"Account not found: {self.email}")

        # Refresh token if needed (MSAL doesn't expose expiry reliably here,
        # so we always try refresh; Graph will 401 if access_token is expired)
        if "refresh_token" not in data:
            raise ValueError("No refresh token stored. Re-authenticate the account.")

        try:
            refreshed = outlook_auth.refresh_tokens(data["refresh_token"])
            data.update(refreshed)
            await database.upsert_account(self.email, "outlook", data)
        except Exception:
            pass  # Use existing access_token; let Graph return 401 if stale

        return {
            "Authorization": f"Bearer {data['access_token']}",
            "Content-Type": "application/json",
        }

    async def _get(self, path: str, params: dict = None) -> dict:
        async with httpx.AsyncClient() as client:
            r = await client.get(
                f"{self.GRAPH}{path}",
                headers=await self._headers(),
                params=params,
                timeout=30,
            )
            r.raise_for_status()
            return r.json()

    async def _post(self, path: str, body: dict) -> httpx.Response:
        async with httpx.AsyncClient() as client:
            r = await client.post(
                f"{self.GRAPH}{path}",
                headers=await self._headers(),
                json=body,
                timeout=30,
            )
            r.raise_for_status()
            return r

    async def list_emails(self, folder: str = "inbox", limit: int = 20) -> list[dict]:
        data = await self._get(
            f"/me/mailFolders/{folder}/messages",
            params={
                "$top": limit,
                "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview",
                "$orderby": "receivedDateTime desc",
            },
        )
        return [
            {
                "id": m["id"],
                "subject": m.get("subject", "(no subject)"),
                "from": m.get("from", {}).get("emailAddress", {}).get("address", ""),
                "received": m.get("receivedDateTime", ""),
                "is_read": m.get("isRead", False),
                "preview": _truncate(m.get("bodyPreview", ""), 150),
            }
            for m in data.get("value", [])
        ]

    async def read_email(self, message_id: str) -> dict:
        m = await self._get(
            f"/me/messages/{message_id}",
            params={"$select": "id,subject,from,toRecipients,receivedDateTime,body,isRead"},
        )
        return {
            "id": m["id"],
            "subject": m.get("subject", "(no subject)"),
            "from": m.get("from", {}).get("emailAddress", {}).get("address", ""),
            "to": [r["emailAddress"]["address"] for r in m.get("toRecipients", [])],
            "received": m.get("receivedDateTime", ""),
            "body": m.get("body", {}).get("content", ""),
            "is_read": m.get("isRead", False),
        }

    async def send_email(self, to: list[str], subject: str, body: str, html: bool = False) -> bool:
        await self._post(
            "/me/sendMail",
            {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "HTML" if html else "Text",
                        "content": body,
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": addr}} for addr in to
                    ],
                }
            },
        )
        return True

    async def search_emails(self, query: str, limit: int = 20) -> list[dict]:
        data = await self._get(
            "/me/messages",
            params={
                "$search": f'"{query}"',
                "$top": limit,
                "$select": "id,subject,from,receivedDateTime,bodyPreview",
            },
        )
        return [
            {
                "id": m["id"],
                "subject": m.get("subject", "(no subject)"),
                "from": m.get("from", {}).get("emailAddress", {}).get("address", ""),
                "received": m.get("receivedDateTime", ""),
                "preview": _truncate(m.get("bodyPreview", ""), 150),
            }
            for m in data.get("value", [])
        ]

    async def delete_email(self, message_id: str) -> bool:
        async with httpx.AsyncClient() as client:
            r = await client.delete(
                f"{self.GRAPH}/me/messages/{message_id}",
                headers=await self._headers(),
                timeout=30,
            )
            r.raise_for_status()
        return True


# ── Generic IMAP / SMTP (iCloud, custom) ─────────────────────────────────────

class ImapProvider:
    def __init__(self, email: str):
        self.email = email

    async def _get_config(self) -> dict:
        data = await database.get_account_data(self.email)
        if not data:
            raise ValueError(f"Account not found: {self.email}")
        return data

    async def list_emails(self, folder: str = "INBOX", limit: int = 20) -> list[dict]:
        cfg = await self._get_config()
        client = aioimaplib.IMAP4_SSL(cfg["imap_host"], cfg.get("imap_port", 993))
        await client.wait_hello_from_server()
        await client.login(cfg["username"], cfg["password"])
        await client.select(folder)

        _, data = await client.search("ALL")
        uids = data[0].split()
        uids = uids[-limit:]  # most recent N

        emails = []
        for uid in reversed(uids):
            _, msg_data = await client.fetch(uid.decode(), "(RFC822.HEADER)")
            raw = msg_data[1]
            msg = email_lib.message_from_bytes(raw)
            emails.append({
                "id": uid.decode(),
                "subject": msg.get("Subject", "(no subject)"),
                "from": msg.get("From", ""),
                "received": msg.get("Date", ""),
                "is_read": False,
                "preview": "",
            })

        await client.logout()
        return emails

    async def read_email(self, message_id: str) -> dict:
        cfg = await self._get_config()
        client = aioimaplib.IMAP4_SSL(cfg["imap_host"], cfg.get("imap_port", 993))
        await client.wait_hello_from_server()
        await client.login(cfg["username"], cfg["password"])
        await client.select("INBOX")

        _, msg_data = await client.fetch(message_id, "(RFC822)")
        raw = msg_data[1]
        msg = email_lib.message_from_bytes(raw)

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="replace")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="replace")

        await client.logout()
        return {
            "id": message_id,
            "subject": msg.get("Subject", "(no subject)"),
            "from": msg.get("From", ""),
            "to": msg.get("To", ""),
            "received": msg.get("Date", ""),
            "body": body,
            "is_read": True,
        }

    async def send_email(self, to: list[str], subject: str, body: str, html: bool = False) -> bool:
        cfg = await self._get_config()

        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = self.email
        msg["To"] = ", ".join(to)
        msg.attach(MIMEText(body, "html" if html else "plain"))

        await aiosmtplib.send(
            msg,
            hostname=cfg["smtp_host"],
            port=cfg.get("smtp_port", 587),
            username=cfg["username"],
            password=cfg["password"],
            start_tls=True,
        )
        return True

    async def search_emails(self, query: str, limit: int = 20) -> list[dict]:
        cfg = await self._get_config()
        client = aioimaplib.IMAP4_SSL(cfg["imap_host"], cfg.get("imap_port", 993))
        await client.wait_hello_from_server()
        await client.login(cfg["username"], cfg["password"])
        await client.select("INBOX")

        _, data = await client.search("TEXT", query)
        uids = data[0].split()[-limit:]

        emails = []
        for uid in reversed(uids):
            _, msg_data = await client.fetch(uid.decode(), "(RFC822.HEADER)")
            msg = email_lib.message_from_bytes(msg_data[1])
            emails.append({
                "id": uid.decode(),
                "subject": msg.get("Subject", "(no subject)"),
                "from": msg.get("From", ""),
                "received": msg.get("Date", ""),
                "preview": "",
            })

        await client.logout()
        return emails

    async def delete_email(self, message_id: str) -> bool:
        cfg = await self._get_config()
        client = aioimaplib.IMAP4_SSL(cfg["imap_host"], cfg.get("imap_port", 993))
        await client.wait_hello_from_server()
        await client.login(cfg["username"], cfg["password"])
        await client.select("INBOX")
        await client.store(message_id, "+FLAGS", "\\Deleted")
        await client.expunge()
        await client.logout()
        return True


# ── Factory ───────────────────────────────────────────────────────────────────

async def get_provider(email: str):
    """Return the correct provider instance for the given email account."""
    data = await database.get_account_data(email)
    if not data:
        raise ValueError(f"No account found for {email}")

    provider = data.get("provider")
    if provider == "outlook":
        return OutlookProvider(email)
    elif provider in ("imap", "icloud"):
        return ImapProvider(email)
    else:
        raise ValueError(f"Unknown provider: {provider}")
