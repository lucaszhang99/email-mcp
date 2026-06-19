from pydantic_settings import BaseSettings, SettingsConfigDict
from functools import lru_cache


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8")

    # MCP
    mcp_bearer_token: str
    encryption_key: str
    base_url: str = "http://localhost:8080"
    web_ui_password: str = "admin"

    # Gmail
    gmail_client_id: str = ""
    gmail_client_secret: str = ""

    # Outlook
    outlook_client_id: str = ""
    outlook_client_secret: str = ""
    outlook_tenant_id: str = "common"

    # DB
    database_url: str = "sqlite+aiosqlite:///./data/accounts.db"

    @property
    def gmail_redirect_uri(self) -> str:
        return f"{self.base_url}/auth/gmail/callback"

    @property
    def outlook_redirect_uri(self) -> str:
        return f"{self.base_url}/auth/outlook/callback"


@lru_cache
def get_settings() -> Settings:
    return Settings()
