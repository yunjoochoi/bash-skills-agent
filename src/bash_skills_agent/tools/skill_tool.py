"""Skill loading tool for progressive disclosure."""

import os

from google.adk.tools import FunctionTool

SKILLS_DIR = os.path.join(os.path.dirname(__file__), "..", "skills")


def _list_available_skills() -> list[str]:
    """List available skill directories."""
    if not os.path.isdir(SKILLS_DIR):
        return []
    return [
        d
        for d in os.listdir(SKILLS_DIR)
        if os.path.isfile(os.path.join(SKILLS_DIR, d, "SKILL.md"))
    ]


def read_skill(skill_name: str) -> str:
    """Load a skill's detailed instructions for use in the current response.

    Args:
        skill_name: Name of the skill directory to load.

    Returns:
        The skill instructions as a string.
    """
    available = _list_available_skills()

    if skill_name not in available:
        return f"Unknown skill: {skill_name}. Available skills: {', '.join(available)}"

    skill_path = os.path.join(SKILLS_DIR, skill_name, "SKILL.md")

    try:
        with open(skill_path, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return f"Skill file not found: {skill_name}"
    except Exception as e:
        return f"Error loading skill: {e}"


skill_tool = FunctionTool(func=read_skill)