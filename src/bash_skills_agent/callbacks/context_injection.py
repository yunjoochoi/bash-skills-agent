"""After-tool callback for injecting runtime context variables into skill content."""

from datetime import datetime
from typing import Any

from google.adk.tools.base_tool import BaseTool
from google.adk.tools.tool_context import ToolContext


def _inject_context_variables(content: str) -> str:
    """Replace context variable placeholders in skill content.

    Supported variables:
        - {today}: Current date in YYYY-MM-DD format
    """
    today = datetime.now().strftime("%Y-%m-%d")
    return content.replace("{today}", today)


def after_tool_modifier(
    tool: BaseTool,
    args: dict[str, Any],
    tool_context: ToolContext,
    tool_response: dict,
) -> dict | None:
    """Inject context variables into skill content after loading.

    ADK wraps FunctionTool returns as {"result": <value>}.
    Handles both str and dict response formats.

    Args:
        tool: The tool that was called.
        args: Arguments passed to the tool.
        tool_context: The tool context.
        tool_response: The tool's response (dict with "result" key).

    Returns:
        Modified response dict, or None to keep original.
    """
    if getattr(tool, "name", None) != "read_skill":
        return None

    if isinstance(tool_response, dict) and "result" in tool_response:
        tool_response["result"] = _inject_context_variables(
            str(tool_response["result"])
        )
        return tool_response

    if isinstance(tool_response, str):
        return {"result": _inject_context_variables(tool_response)}

    return None
