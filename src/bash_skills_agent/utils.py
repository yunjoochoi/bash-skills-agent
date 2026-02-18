"""Utility functions for agent initialization and session management."""

import json
import logging
import os
import sys
from typing import Optional

import frontmatter
from google.adk.events import Event

from .config.settings import settings

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

LOG_LEVEL_MAP = {
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL,
}


def setup_logging(
    level: int | None = None,
    format_string: Optional[str] = None,
) -> None:
    """Set up global logging configuration.

    Args:
        level: Logging level override (default: from settings.log_level)
        format_string: Custom format string (optional)
    """
    if level is None:
        level = LOG_LEVEL_MAP.get(settings.log_level.upper(), logging.INFO)

    if format_string is None:
        format_string = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"

    logging.basicConfig(
        level=level,
        format=format_string,
        handlers=[logging.StreamHandler(sys.stdout)],
    )


def get_logger(name: str, level: Optional[int] = None) -> logging.Logger:
    """Get a logger instance with the specified name.

    Args:
        name: Logger name (typically __name__ of the module)
        level: Optional logging level override

    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name)

    if level is not None:
        logger.setLevel(level)

    return logger


# Initialize default logging configuration
setup_logging()


# ---------------------------------------------------------------------------
# Skill frontmatter loader
# ---------------------------------------------------------------------------

SKILLS_BASE_DIR = os.path.join(os.path.dirname(__file__), "skills")


def load_skill_frontmatter(skill_name: str) -> str:
    """Load name + description from a skill's SKILL.md as formatted string.

    Used at agent init time to embed lightweight metadata in instructions.

    Args:
        skill_name: Name of the skill directory.

    Returns:
        Formatted string like "- skill-name: description text".
    """
    skill_path = os.path.join(SKILLS_BASE_DIR, skill_name, "SKILL.md")

    with open(skill_path, "r", encoding="utf-8") as f:
        post = frontmatter.load(f)

    return f"- {post['name']}: {post['description']}"


def load_all_skill_frontmatters() -> str:
    """Discover all skills and return their frontmatters as a formatted string.

    Scans SKILLS_BASE_DIR for subdirectories containing SKILL.md.

    Returns:
        Newline-joined string of all skill descriptions,
        or "No skills available." if none found.
    """
    if not os.path.isdir(SKILLS_BASE_DIR):
        return "No skills available."

    lines = []
    for name in sorted(os.listdir(SKILLS_BASE_DIR)):
        skill_md = os.path.join(SKILLS_BASE_DIR, name, "SKILL.md")
        if os.path.isfile(skill_md):
            try:
                lines.append(load_skill_frontmatter(name))
            except Exception:
                pass

    return "\n".join(lines) if lines else "No skills available."


# ---------------------------------------------------------------------------
# Session utilities
# ---------------------------------------------------------------------------


async def ensure_session_exists(
    session_service,
    app_name: str,
    user_id: str,
    session_id: str,
    initial_state: dict | None = None,
) -> dict | None:
    """Ensure session exists, create if not found.

    Args:
        session_service: The session service instance
        app_name: Application name
        user_id: User identifier
        session_id: Session identifier
        initial_state: Initial state to set if creating new session

    Returns:
        Initial state dict if session was created, None if already existed
    """
    session = await session_service.get_session(
        app_name=app_name,
        user_id=user_id,
        session_id=session_id,
    )

    if not session:
        await session_service.create_session(
            app_name=app_name,
            user_id=user_id,
            session_id=session_id,
        )
        return initial_state or {}
    return None


def track_token_usage(event: Event, usage_map: dict[str, list[int]]) -> dict[str, list[int]]:
    """Track token usage from event metadata.

    Args:
        event: The event containing usage metadata
        usage_map: Dictionary mapping author to [input_tokens, output_tokens]

    Returns:
        Updated usage_map with new token counts
    """
    if event.usage_metadata:
        author = event.author
        if author not in usage_map:
            usage_map[author] = [0, 0]
        usage_map[author][0] += event.usage_metadata.prompt_token_count or 0
        usage_map[author][1] += event.usage_metadata.candidates_token_count or 0
    return usage_map


# ---------------------------------------------------------------------------
# Event utilities
# ---------------------------------------------------------------------------


def sse_frame(event_type: str, data: dict) -> str:
    """Format data as Server-Sent Events frame.

    Args:
        event_type: Type of the event
        data: Data to send in the event

    Returns:
        Formatted SSE frame string
    """
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    lines = [f"event: {event_type}"]
    for part in data_json.split("\n"):
        lines.append(f"data: {part}")
    return "\n".join(lines) + "\n\n"
