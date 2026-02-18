"""Shared infrastructure: singleton services and per-session resources."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import TYPE_CHECKING

from google.adk.sessions import (
    DatabaseSessionService,
    InMemorySessionService,
)
from google.adk.tools.mcp_tool.mcp_session_manager import (
    SseConnectionParams,
    StdioConnectionParams,
    StreamableHTTPConnectionParams,
)
from google.adk.tools.mcp_tool.mcp_toolset import McpToolset

from .settings import settings

if TYPE_CHECKING:
    from ..container.manager import ContainerManager

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
            db_path = Path(settings.sqlite_db_path)
            db_path.parent.mkdir(parents=True, exist_ok=True)

            db_uri = f"sqlite:///{db_path}"
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
    if settings.skills_dir:
        return str(Path(settings.skills_dir).resolve())
    return str((Path(__file__).parent.parent / "skills").resolve())


def get_session_workspace(session_id: str) -> str:
    """Return session-isolated workspace path on host."""
    workspace = Path(settings.user_files_path).resolve() / session_id
    workspace.mkdir(parents=True, exist_ok=True)
    return str(workspace)


def get_container_manager(session_id: str) -> "ContainerManager":
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
