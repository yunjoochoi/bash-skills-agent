"""Web fetch tool — retrieve and extract text content from a URL."""

from __future__ import annotations

import logging
from typing import Any

import httpx
from bs4 import BeautifulSoup

from ..config.settings import settings

logger = logging.getLogger(__name__)

_STRIP_TAGS = ("script", "style", "nav", "footer", "header", "noscript")


def _parse_html(html: str) -> tuple[str, str]:
    """Parse HTML and return (clean_text, title)."""
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.string.strip() if soup.title and soup.title.string else ""

    for tag in soup(_STRIP_TAGS):
        tag.decompose()

    text = soup.get_text(separator="\n", strip=True)
    lines = [line for line in text.splitlines() if line.strip()]
    return "\n".join(lines), title


async def web_fetch(
    url: str,
    max_length: int = 0,
) -> dict[str, Any]:
    """Fetch a web page and return its text content.

    Args:
        url: The URL to fetch.
        max_length: Maximum character length of returned content.
                    Defaults to settings.web_fetch_max_length.

    Returns:
        Dictionary with url, title, and content (or error).
    """
    if max_length <= 0:
        max_length = settings.web_fetch_max_length

    try:
        async with httpx.AsyncClient(
            timeout=settings.web_fetch_timeout,
            follow_redirects=True,
            headers={"User-Agent": "Mozilla/5.0 (compatible; BashSkillsAgentBot/1.0)"},
        ) as client:
            resp = await client.get(url)
            resp.raise_for_status()

        content_type = resp.headers.get("content-type", "")

        if "html" in content_type:
            text, title = _parse_html(resp.text)
        else:
            text, title = resp.text, ""

        if len(text) > max_length:
            text = text[:max_length] + f"\n\n... (truncated, {len(text)} total chars)"

        logger.info("Fetched %s (%d chars)", url, len(text))
        return {"url": url, "title": title, "content": text}

    except httpx.HTTPStatusError as e:
        return {"error": f"HTTP {e.response.status_code}: {url}"}
    except Exception as e:
        logger.error("Fetch failed for %s: %s", url, e)
        return {"error": f"Fetch failed: {e}"}
