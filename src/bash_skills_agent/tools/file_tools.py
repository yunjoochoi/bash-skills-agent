"""Host filesystem tools — read, write, edit, glob, grep.

Paths starting with /workspace/ are translated to the host session workspace.
"""

from __future__ import annotations

import fnmatch
import os
import re
from pathlib import Path
from typing import Any

from google.adk.tools.tool_context import ToolContext


def _resolve_path(file_path: str, tool_context: ToolContext | None) -> Path:
    """Translate /workspace/ paths to host session workspace.

    Relative paths (not starting with /) are treated as relative to /workspace/.
    """
    # Treat relative paths as relative to /workspace/
    if not file_path.startswith("/") and tool_context:
        file_path = f"/workspace/{file_path}"

    if file_path.startswith("/workspace") and tool_context:
        from ..config.shared_clients import get_session_workspace

        host_workspace = get_session_workspace(tool_context.session.id)
        relative = file_path.replace("/workspace", "", 1).lstrip("/")
        return Path(host_workspace) / relative
    return Path(file_path)


async def read_file(
    file_path: str,
    offset: int = 0,
    limit: int = 2000,
    tool_context: ToolContext | None = None,
) -> dict[str, Any]:
    """Read a file from the workspace.

    Args:
        file_path: Path to the file (/workspace/... or absolute)
        offset: Line number to start reading from (0-based)
        limit: Maximum number of lines to read

    Returns:
        Dict with content, total_lines, and path info
    """
    p = _resolve_path(file_path, tool_context)
    if not p.is_file():
        return {"error": f"File not found: {file_path}"}

    try:
        lines = p.read_text(encoding="utf-8", errors="replace").splitlines()
        total = len(lines)
        selected = lines[offset : offset + limit]
        content = "\n".join(
            f"{i + offset + 1:>6}\t{line}" for i, line in enumerate(selected)
        )
        return {"content": content, "total_lines": total, "path": str(p.resolve())}
    except Exception as e:
        return {"error": str(e)}


async def write_file(
    file_path: str,
    content: str,
    tool_context: ToolContext | None = None,
) -> dict[str, Any]:
    """Write content to a file (creates or overwrites).

    Args:
        file_path: Path to the file (/workspace/... or absolute)
        content: Content to write

    Returns:
        Dict with success status and path
    """
    p = _resolve_path(file_path, tool_context)
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")
        return {"success": True, "path": str(p.resolve()), "bytes_written": len(content.encode("utf-8"))}
    except Exception as e:
        return {"error": str(e)}


async def edit_file(
    file_path: str,
    old_text: str,
    new_text: str,
    tool_context: ToolContext | None = None,
) -> dict[str, Any]:
    """Replace an exact text block in a file.

    Args:
        file_path: Path to the file (/workspace/... or absolute)
        old_text: Text to find (must be unique in the file)
        new_text: Replacement text

    Returns:
        Dict with success status
    """
    p = _resolve_path(file_path, tool_context)
    if not p.is_file():
        return {"error": f"File not found: {file_path}"}

    try:
        text = p.read_text(encoding="utf-8")
        count = text.count(old_text)
        if count == 0:
            return {"error": "old_text not found in file"}
        if count > 1:
            return {"error": f"old_text found {count} times — must be unique"}

        p.write_text(text.replace(old_text, new_text, 1), encoding="utf-8")
        return {"success": True, "path": str(p.resolve())}
    except Exception as e:
        return {"error": str(e)}


async def glob_search(
    pattern: str,
    path: str = "/workspace",
    tool_context: ToolContext | None = None,
) -> dict[str, Any]:
    """Find files matching a glob pattern.

    Args:
        pattern: Glob pattern (e.g., "**/*.py")
        path: Directory to search in (default: /workspace)

    Returns:
        Dict with list of matching file paths
    """
    p = _resolve_path(path, tool_context)
    if not p.is_dir():
        return {"error": f"Directory not found: {path}"}

    try:
        matches = sorted(str(f) for f in p.glob(pattern) if f.is_file())
        return {"matches": matches[:200], "total": len(matches)}
    except Exception as e:
        return {"error": str(e)}


async def grep_search(
    pattern: str,
    path: str = "/workspace",
    glob: str | None = None,
    tool_context: ToolContext | None = None,
) -> dict[str, Any]:
    """Search file contents with regex.

    Args:
        pattern: Regex pattern to search for
        path: File or directory to search in (default: /workspace)
        glob: Optional glob filter for file names (e.g., "*.py")

    Returns:
        Dict with matching files and lines
    """
    p = _resolve_path(path, tool_context)
    try:
        regex = re.compile(pattern)
    except re.error as e:
        return {"error": f"Invalid regex: {e}"}

    results: list[dict] = []

    def _search_file(fp: Path) -> None:
        try:
            for i, line in enumerate(fp.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
                if regex.search(line):
                    results.append({"file": str(fp), "line": i, "text": line.rstrip()[:500]})
        except Exception:
            pass

    if p.is_file():
        _search_file(p)
    elif p.is_dir():
        for root, _, files in os.walk(p):
            for name in files:
                if glob and not fnmatch.fnmatch(name, glob):
                    continue
                _search_file(Path(root) / name)
                if len(results) >= 500:
                    break
            if len(results) >= 500:
                break
    else:
        return {"error": f"Path not found: {path}"}

    return {"matches": results, "total": len(results)}
