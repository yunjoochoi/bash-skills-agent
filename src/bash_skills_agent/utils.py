"""Utility functions for agent initialization."""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

import frontmatter

from .config.settings import settings

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------


def setup_logging(level: int | None = None) -> None:
    """Set up global logging configuration.

    Args:
        level: Logging level override (default: from settings.log_level).
    """
    if level is None:
        level_map = {
            "DEBUG": logging.DEBUG,
            "INFO": logging.INFO,
            "WARNING": logging.WARNING,
            "ERROR": logging.ERROR,
            "CRITICAL": logging.CRITICAL,
        }
        level = level_map.get(settings.log_level.upper(), logging.INFO)

    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(levelname)s - %(name)s - %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


setup_logging()

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Skill frontmatter loader
# ---------------------------------------------------------------------------

SKILLS_BASE_DIR = Path(__file__).parent / "skills"


def load_skill_frontmatter(skill_name: str) -> str:
    """Load name + description from a skill's SKILL.md as formatted string.

    Args:
        skill_name: Name of the skill directory.

    Returns:
        Formatted string like "- skill-name: description text".
    """
    skill_path = SKILLS_BASE_DIR / skill_name / "SKILL.md"
    with open(skill_path, encoding="utf-8") as f:
        post = frontmatter.load(f)
    return f"- {post['name']}: {post['description']}"


def load_all_skill_frontmatters() -> str:
    """Discover all skills and return their frontmatters as a formatted string.

    Returns:
        Newline-joined string of all skill descriptions,
        or "No skills available." if none found.
    """
    if not SKILLS_BASE_DIR.is_dir():
        return "No skills available."

    lines: list[str] = []
    for name in sorted(os.listdir(SKILLS_BASE_DIR)):
        if (SKILLS_BASE_DIR / name / "SKILL.md").is_file():
            try:
                lines.append(load_skill_frontmatter(name))
            except Exception:
                logger.warning("Failed to load skill frontmatter: %s", name)

    return "\n".join(lines) if lines else "No skills available."
