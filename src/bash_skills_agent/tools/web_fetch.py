"""Web fetch tool - retrieve and extract text content from a URL."""

import logging
from typing import Any

import httpx
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

_TIMEOUT = 15
_MAX_CONTENT_LENGTH = 50_000


def _html_to_text(html: str) -> str:
    """Extract clean text from HTML using BeautifulSoup."""
    soup = BeautifulSoup(html, "html.parser")

    for tag in soup(["script", "style", "nav", "footer", "header", "noscript"]):
        tag.decompose()

    text = soup.get_text(separator="\n", strip=True)

    # Collapse excessive blank lines
    lines = [line for line in text.splitlines() if line.strip()]
    return "\n".join(lines)


async def web_fetch(
    url: str,
    max_length: int = _MAX_CONTENT_LENGTH,
) -> dict[str, Any]:
    """Fetch a web page and return its text content.

    Args:
        url: The URL to fetch
        max_length: Maximum character length of returned content (default: 50000)

    Returns:
        Dictionary with url, title, and content (or error)
    """
    try:
        async with httpx.AsyncClient(
            timeout=_TIMEOUT,
            follow_redirects=True,
            headers={"User-Agent": "Mozilla/5.0 (compatible; BashSkillsAgentBot/1.0)"},
        ) as client:
            resp = await client.get(url)
            resp.raise_for_status()

        content_type = resp.headers.get("content-type", "")

        if "html" in content_type:
            text = _html_to_text(resp.text)
            soup = BeautifulSoup(resp.text, "html.parser")
            title = soup.title.string.strip() if soup.title and soup.title.string else ""
        else:
            text = resp.text
            title = ""

        if len(text) > max_length:
            text = text[:max_length] + f"\n\n... (truncated, {len(text)} total chars)"

        logger.info("Fetched %s (%d chars)", url, len(text))
        return {"url": url, "title": title, "content": text}

    except httpx.HTTPStatusError as e:
        return {"error": f"HTTP {e.response.status_code}: {url}"}
    except Exception as e:
        logger.error("Fetch failed for %s: %s", url, e)
        return {"error": f"Fetch failed: {e}"}
