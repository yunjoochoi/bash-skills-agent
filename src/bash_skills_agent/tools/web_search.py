"""Web search tool using DuckDuckGo."""

import logging
from typing import Any

logger = logging.getLogger(__name__)


async def search_web(
    query: str,
    num_results: int = 5,
    region: str = "kr-kr",
) -> dict[str, Any]:
    """Search the web using DuckDuckGo.

    IMPORTANT: Always write queries in English for best results.
    Translate Korean queries to English before searching.

    Args:
        query: Search query string (use English)
        num_results: Number of results to return (default: 5, max: 10)
        region: Region code for localized results (default: kr-kr)

    Returns:
        Dictionary containing search results or error information
    """
    try:
        from duckduckgo_search import DDGS
    except ImportError:
        return {"error": "duckduckgo-search not installed. Run: pip install duckduckgo-search"}

    num_results = min(num_results, 10)

    try:
        results = []
        for r in DDGS().text(query, region=region, max_results=num_results):
            results.append({
                "title": r.get("title", ""),
                "url": r.get("href", ""),
                "snippet": r.get("body", ""),
            })

        logger.info("Search completed: %d results for '%s'", len(results), query)
        return {"query": query, "results": results, "total_results": len(results)}

    except Exception as e:
        logger.error("Search failed: %s", e)
        return {"error": f"Search failed: {e}"}
