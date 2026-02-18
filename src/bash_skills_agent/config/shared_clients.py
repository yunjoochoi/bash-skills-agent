"""Shared infrastructure: singleton services and per-session resources."""

from __future__ import annotations

import json
import logging
import os

from google.adk.sessions import (
    DatabaseSessionService,
    InMemorySessionService,
)
from google.adk.tools.mcp_tool.mcp_toolset import McpToolset
from google.adk.tools.mcp_tool.mcp_session_manager import (
    SseConnectionParams,
    StdioConnectionParams,
    StreamableHTTPConnectionParams,
)

from .settings import settings

logger = logging.getLogger(__name__)

# Global singleton instances
_session_service = None
_container_managers: dict[str, object] = {}  # session_id → ContainerManager


def get_session_service():
    """Return shared SessionService instance based on configuration.

    Session service type is determined by settings.session_service_type:
    - postgresql: Uses DatabaseSessionService with PostgreSQL
    - sqlite: Uses DatabaseSessionService with SQLite
    - inmemory: Uses InMemorySessionService

    Returns:
        Configured session service instance

    Raises:
        ValueError: If session_service_type is unknown or required config is missing
    """
    global _session_service
    if _session_service is None:
        if settings.session_service_type == "postgresql":
            if not settings.postgres_uri:
                raise ValueError(
                    "postgres_uri required when session_service_type=postgresql"
                )
            _session_service = DatabaseSessionService(db_uri=settings.postgres_uri)
        elif settings.session_service_type == "sqlite":
            # Ensure parent directory exists for SQLite database
            db_dir = os.path.dirname(settings.sqlite_db_path)
            if db_dir and not os.path.exists(db_dir):
                os.makedirs(db_dir, exist_ok=True)

            db_uri = f"sqlite:///{settings.sqlite_db_path}"
            _session_service = DatabaseSessionService(db_uri=db_uri)
        elif settings.session_service_type == "inmemory":
            _session_service = InMemorySessionService()
        else:
            raise ValueError(
                f"Unknown session_service_type: {settings.session_service_type}"
            )
    return _session_service



def resolve_skills_dir() -> str:
    """Resolve the skills directory path from settings or default."""
    skills_dir = settings.skills_dir
    if not skills_dir:
        skills_dir = str(os.path.join(os.path.dirname(__file__), "..", "skills"))
    return os.path.abspath(skills_dir)


def get_session_workspace(session_id: str) -> str:
    """Return session-isolated workspace path on host."""
    workspace = os.path.join(os.path.abspath(settings.user_files_path), session_id)
    os.makedirs(workspace, exist_ok=True)
    return workspace


def get_container_manager(session_id: str):
    """Return per-session ContainerManager (lazy-created)."""
    if session_id not in _container_managers:
        from ..container.manager import ContainerManager

        _container_managers[session_id] = ContainerManager(
            image=settings.container_image,
            workspace_dir=get_session_workspace(session_id),
            skills_dir=resolve_skills_dir(),
            memory=settings.container_memory,
            network=settings.container_network,
        )
    return _container_managers[session_id]


async def shutdown_container(session_id: str | None = None) -> None:
    """Stop and remove sandbox container(s)."""
    if session_id:
        mgr = _container_managers.pop(session_id, None)
        if mgr:
            await mgr.stop()
    else:
        for mgr in _container_managers.values():
            await mgr.stop()
        _container_managers.clear()


def build_mcp_toolsets() -> list:
    """Parse MCP_SERVERS JSON and build McpToolset instances."""
    raw = settings.mcp_servers
    if not raw:
        return []

    try:
        servers = json.loads(raw)
    except json.JSONDecodeError:
        logger.error("Invalid MCP_SERVERS JSON: %s", raw)
        return []

    toolsets = []
    for cfg in servers:
        server_type = cfg.get("type", "stdio")
        try:
            if server_type == "stdio":
                params = StdioConnectionParams(
                    command=cfg["command"],
                    args=cfg.get("args", []),
                )
            elif server_type == "sse":
                params = SseConnectionParams(url=cfg["url"])
            elif server_type == "http":
                params = StreamableHTTPConnectionParams(url=cfg["url"])
            else:
                logger.warning("Unknown MCP server type: %s", server_type)
                continue

            toolsets.append(McpToolset(connection_params=params))
            logger.info(
                "MCP toolset created: %s (%s)",
                cfg.get("command", cfg.get("url")),
                server_type,
            )
        except Exception:
            logger.exception("Failed to create MCP toolset for %s", cfg)

    return toolsets
