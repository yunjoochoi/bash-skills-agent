"""Skill loading tool for progressive disclosure."""

from google.adk.tools import FunctionTool

from ..utils import SKILLS_BASE_DIR


def _list_available_skills() -> list[str]:
    """List available skill directories."""
    if not SKILLS_BASE_DIR.is_dir():
        return []
    return [
        d.name
        for d in sorted(SKILLS_BASE_DIR.iterdir())
        if d.is_dir() and (d / "SKILL.md").is_file()
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

    skill_path = SKILLS_BASE_DIR / skill_name / "SKILL.md"

    try:
        return skill_path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return f"Skill file not found: {skill_name}"
    except Exception as e:
        return f"Error loading skill: {e}"


skill_tool = FunctionTool(func=read_skill)