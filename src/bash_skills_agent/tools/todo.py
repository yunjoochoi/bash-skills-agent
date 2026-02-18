"""Todo list tool — session state-backed task tracking."""

from __future__ import annotations

from typing import Any

from google.adk.tools.tool_context import ToolContext

_TODO_STATE_KEY = "_todo_list"


async def todo_write(
    todos: list[dict[str, str]],
    tool_context: ToolContext,
) -> dict[str, Any]:
    """Save or update the todo list in session state.

    Each item should have: content, status (pending|in_progress|completed), activeForm.

    Args:
        todos: List of todo items
        tool_context: ADK tool context for session state access

    Returns:
        Dict with the updated todo list
    """
    validated = []
    for item in todos:
        validated.append({
            "content": item.get("content", ""),
            "status": item.get("status", "pending"),
            "activeForm": item.get("activeForm", ""),
        })

    tool_context.state[_TODO_STATE_KEY] = validated
    return {"todos": validated, "count": len(validated)}
