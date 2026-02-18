"""Web search tool using DuckDuckGo."""

import logging
from typing import Any

from ..config.settings import settings

logger = logging.getLogger(__name__)


async def search_web(
    query: str,
    num_results: int = 5,
    region: str = "",
    allowed_domains: list[str] | None = None,
    blocked_domains: list[str] | None = None,
) -> dict[str, Any]:
    """Search the web using DuckDuckGo.

    IMPORTANT: Always write queries in English for best results.
    Translate Korean queries to English before searching.

    Args:
        query: Search query string (use English)
        num_results: Number of results to return (default: 5, max: 10)
        region: Region code for localized results (default: kr-kr)
        allowed_domains: Only include results from these domains
        blocked_domains: Exclude results from these domains

    Returns:
        Dictionary containing search results or error information
    """
    try:
        from duckduckgo_search import DDGS
    except ImportError:
        return {"error": "duckduckgo-search not installed"}

    if not region:
        region = settings.web_search_region
    num_results = min(num_results, 10)

    # Fetch extra results when filtering, to compensate for filtered-out entries
    fetch_count = num_results * 3 if allowed_domains or blocked_domains else num_results

    try:
        results = []
        for r in DDGS().text(query, region=region, max_results=fetch_count):
            url = r.get("href", "")
            if allowed_domains and not any(d in url for d in allowed_domains):
                continue
            if blocked_domains and any(d in url for d in blocked_domains):
                continue
            results.append({
                "title": r.get("title", ""),
                "url": url,
                "snippet": r.get("body", ""),
            })
            if len(results) >= num_results:
                break

        logger.info("Search completed: %d results for '%s'", len(results), query)
        return {"query": query, "results": results, "total_results": len(results)}

    except Exception as e:
        logger.error("Search failed: %s", e)
        return {"error": f"Search failed: {e}"}
