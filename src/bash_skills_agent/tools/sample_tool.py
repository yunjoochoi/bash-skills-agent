"""Sample tool for ADK agent."""

import logging

from google.adk.tools.tool_context import ToolContext

logger = logging.getLogger(__name__)


def sample_tool(
    query: str,
    tool_context: ToolContext | None = None,
) -> dict[str, str]:
    """Echo back the user's query. A minimal example tool.

    Args:
        query: The user's input to echo back.

    Returns:
        A dict with the echoed response.
    """
    if tool_context:
        count = tool_context.state.get("sample_tool_count", 0)
        tool_context.state["sample_tool_count"] = count + 1

    logger.info("Sample tool called with query: %s", query)

    return {"status": "success", "response": f"Received: {query}"}
