"""Application settings using Pydantic Settings."""

import logging
import os
import warnings

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application configuration loaded from environment variables."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # Model configuration
    litellm_model: str = "vertex_ai.gemini-3.0-flash-preview"

    # LiteLLM Proxy (env: OPENAI_API_KEY, OPENAI_API_BASE)
    openai_api_key: str = ""
    openai_api_base: str = ""

    # Session storage
    session_service_type: str = "sqlite"  # postgresql, sqlite, inmemory
    postgres_uri: str | None = None
    sqlite_db_path: str = "./data/sessions.db"

    # Application
    app_name: str = "bash_skills_agent"
    log_level: str = "INFO"

    # Container
    container_image: str = "bash-skills-agent-sandbox:latest"
    container_memory: str = "512m"
    container_network: str = "none"
    container_timeout: int = 120
    user_files_path: str = "/tmp/agent-workspace"

    # Skills
    skills_dir: str = ""

    # MCP
    mcp_servers: str = ""

    def configure_litellm_proxy(self) -> None:
        """Set OPENAI_* env vars so litellm picks them up."""
        if self.openai_api_base:
            os.environ["OPENAI_API_BASE"] = self.openai_api_base
        if self.openai_api_key:
            os.environ["OPENAI_API_KEY"] = self.openai_api_key

    def configure_logging(self) -> None:
        """Configure logging for the application."""
        logging.basicConfig(level=self.log_level)
        # Suppress noisy Pydantic serialization warnings from LiteLLM
        warnings.filterwarnings("ignore", message="Pydantic serializer warnings")


# Global settings instance
settings = Settings()
settings.configure_litellm_proxy()
settings.configure_logging()
