"""Bash tool — runs commands inside the Docker sandbox container."""

from __future__ import annotations

from typing import Any

from google.adk.tools.tool_context import ToolContext


async def bash(
    command: str,
    timeout: int = 120,
    tool_context: ToolContext | None = None,
) -> dict[str, Any]:
    """Run a shell command in the sandboxed Docker container.

    The container mounts:
    - /workspace: user files (read/write)
    - /skills: skill scripts (read-only)

    Args:
        command: Shell command to execute
        timeout: Max seconds to wait (default 120)
        tool_context: ADK tool context

    Returns:
        Dict with exit_code, stdout, stderr
    """
    from ..config.shared_clients import get_container_manager

    session_id = tool_context.session.id if tool_context else "default"
    mgr = get_container_manager(session_id)
    result = await mgr.exec(command, timeout=timeout)

    return {
        "exit_code": result.exit_code,
        "stdout": result.stdout,
        "stderr": result.stderr,
    }
