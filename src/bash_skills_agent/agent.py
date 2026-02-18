"""
Root agent definition — assembles all tools, skills, and MCP connectors.

adk web src --port 8080 2>&1 | tee logs/agent.log

"""

from __future__ import annotations

import logging

from google.adk.agents import Agent
from google.adk.models.lite_llm import LiteLlm

from .prompt import ROOT_AGENT_PROMPT
from .config.settings import settings
from .config.shared_clients import build_mcp_toolsets
from .utils import load_all_skill_frontmatters
from .tools import (
    bash,
    edit_file,
    glob_search,
    grep_search,
    read_file,
    search_web,
    skill_tool,
    todo_write,
    web_fetch,
    write_file,
)
from .callbacks import (
    after_tool_modifier,
    extract_files_callback,
    export_outputs_callback,
    log_reasoning_callback,
)
from .sub_agents.greeter import greeter_agent

logger = logging.getLogger(__name__)


def create_root_agent() -> Agent:
    """Create the root agent with all tools, skills, and MCP connectors."""
    tools: list = [
        # File tools
        read_file,
        write_file,
        edit_file,
        glob_search,
        grep_search,
        # Container execution
        bash,
        # Web
        search_web,
        web_fetch,
        # Task management
        todo_write,
        # Skills
        skill_tool,
    ]

    # MCP connectors
    tools.extend(build_mcp_toolsets())

    # Build instruction with dynamic skill context
    skill_context = load_all_skill_frontmatters()
    instruction = ROOT_AGENT_PROMPT.format(skill_context=skill_context)

    return Agent(
        name="main_assistant",
        model=LiteLlm(model=settings.litellm_model),
        instruction=instruction,
        description="AI assistant with file editing, code execution, skills, and web search",
        tools=tools,
        sub_agents=[greeter_agent],
        before_agent_callback=extract_files_callback,
        after_agent_callback=export_outputs_callback,
        after_model_callback=log_reasoning_callback,
        after_tool_callback=after_tool_modifier,
    )


# ADK CLI entry point
root_agent = create_root_agent()
