"""
Encrypted SQLite storage for email account credentials and OAuth tokens.
All sensitive fields are encrypted with Fernet (AES-128-CBC) using ENCRYPTION_KEY.
"""

import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from cryptography.fernet import Fernet
from sqlalchemy import Column, String, DateTime, Text, select, delete
from sqlalchemy.ext.asyncio import AsyncSession, create_async_engine, async_sessionmaker
from sqlalchemy.orm import DeclarativeBase

from config import get_settings


# ── ORM ─────────────────────────────────────────────────────────────────────

class Base(DeclarativeBase):
    pass


class Account(Base):
    __tablename__ = "accounts"

    id = Column(String, primary_key=True, default=lambda: str(uuid.uuid4()))
    email = Column(String, nullable=False, unique=True)
    provider = Column(String, nullable=False)        # gmail | outlook | imap
    display_name = Column(String, nullable=True)
    encrypted_data = Column(Text, nullable=False)    # JSON encrypted with Fernet
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


# ── Engine ───────────────────────────────────────────────────────────────────

_engine = None
_session_factory = None


async def init_db():
    global _engine, _session_factory
    settings = get_settings()

    # Auto-create the data/ directory so SQLite can write the file
    db_path = settings.database_url.replace("sqlite+aiosqlite:///", "")
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)

    _engine = create_async_engine(settings.database_url, echo=False)
    _session_factory = async_sessionmaker(_engine, expire_on_commit=False)
    async with _engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)


def get_session() -> AsyncSession:
    return _session_factory()


# ── Encryption ───────────────────────────────────────────────────────────────

def _fernet() -> Fernet:
    return Fernet(get_settings().encryption_key.encode())


def encrypt(data: dict) -> str:
    return _fernet().encrypt(json.dumps(data).encode()).decode()


def decrypt(token: str) -> dict:
    return json.loads(_fernet().decrypt(token.encode()))


# ── CRUD ─────────────────────────────────────────────────────────────────────

async def upsert_account(
    email: str,
    provider: str,
    data: dict,
    display_name: Optional[str] = None,
) -> Account:
    """Create or update an account, encrypting sensitive data."""
    async with get_session() as session:
        result = await session.execute(select(Account).where(Account.email == email))
        account = result.scalar_one_or_none()

        if account:
            account.encrypted_data = encrypt(data)
            account.display_name = display_name or account.display_name
            account.updated_at = datetime.utcnow()
        else:
            account = Account(
                email=email,
                provider=provider,
                display_name=display_name or email,
                encrypted_data=encrypt(data),
            )
            session.add(account)

        await session.commit()
        return account


async def get_account(email: str) -> Optional[tuple[Account, dict]]:
    """Returns (account, decrypted_data) or None."""
    async with get_session() as session:
        result = await session.execute(select(Account).where(Account.email == email))
        account = result.scalar_one_or_none()
        if not account:
            return None
        return account, decrypt(account.encrypted_data)


async def list_accounts() -> list[dict]:
    """Returns all accounts (without sensitive data)."""
    async with get_session() as session:
        result = await session.execute(select(Account))
        accounts = result.scalars().all()
        return [
            {
                "id": a.id,
                "email": a.email,
                "provider": a.provider,
                "display_name": a.display_name,
                "created_at": a.created_at.isoformat(),
            }
            for a in accounts
        ]


async def delete_account(email: str) -> bool:
    async with get_session() as session:
        result = await session.execute(
            delete(Account).where(Account.email == email)
        )
        await session.commit()
        return result.rowcount > 0


async def get_account_data(email: str) -> Optional[dict]:
    """Returns only the decrypted data dict."""
    result = await get_account(email)
    if result is None:
        return None
    _, data = result
    return data
