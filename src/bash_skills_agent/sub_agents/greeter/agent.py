"""Greeter sub-agent definition."""

from google.adk.agents import Agent
from google.adk.models.lite_llm import LiteLlm

from bash_skills_agent.config import settings
from .prompt import GREETER_PROMPT

greeter_agent = Agent(
    name="greeter",
    model=LiteLlm(model=settings.litellm_model),
    instruction=GREETER_PROMPT,
    description="Handles greetings and getting-started guidance.",
)

# adk web compatibility: `adk web src/<pkg>/sub_agents/` runs each sub-agent
# as a standalone agent. adk web expects a `root_agent` variable in agent.py.
root_agent = greeter_agent
