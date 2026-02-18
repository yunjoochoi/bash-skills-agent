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
    replace_all: bool = False,
    tool_context: ToolContext | None = None,
) -> dict[str, Any]:
    """Replace an exact text block in a file.

    Args:
        file_path: Path to the file (/workspace/... or absolute)
        old_text: Text to find (must be unique unless replace_all=True)
        new_text: Replacement text
        replace_all: Replace all occurrences (default: False)

    Returns:
        Dict with success status and replacement count
    """
    p = _resolve_path(file_path, tool_context)
    if not p.is_file():
        return {"error": f"File not found: {file_path}"}

    try:
        text = p.read_text(encoding="utf-8")
        count = text.count(old_text)
        if count == 0:
            return {"error": "old_text not found in file"}
        if not replace_all and count > 1:
            return {"error": f"old_text found {count} times — must be unique (use replace_all=True to replace all)"}

        if replace_all:
            p.write_text(text.replace(old_text, new_text), encoding="utf-8")
        else:
            p.write_text(text.replace(old_text, new_text, 1), encoding="utf-8")
        return {"success": True, "path": str(p.resolve()), "replacements": count if replace_all else 1}
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
    output_mode: str = "content",
    ignore_case: bool = False,
    context_lines: int = 0,
    head_limit: int = 0,
    tool_context: ToolContext | None = None,
) -> dict[str, Any]:
    """Search file contents with regex.

    Args:
        pattern: Regex pattern to search for
        path: File or directory to search in (default: /workspace)
        glob: Optional glob filter for file names (e.g., "*.py")
        output_mode: "content" (matching lines), "files_with_matches" (file paths only),
                     "count" (match counts per file)
        ignore_case: Ignore case when matching (default: False)
        context_lines: Number of lines to show before and after each match (default: 0)
        head_limit: Limit output to first N entries (default: 0 = unlimited, capped at 500)

    Returns:
        Dict with matches (format depends on output_mode) and total count
    """
    p = _resolve_path(path, tool_context)
    flags = re.IGNORECASE if ignore_case else 0
    try:
        regex = re.compile(pattern, flags)
    except re.error as e:
        return {"error": f"Invalid regex: {e}"}

    max_results = min(head_limit, 500) if head_limit > 0 else 500

    # Collect target files
    files_to_search: list[Path] = []
    if p.is_file():
        files_to_search.append(p)
    elif p.is_dir():
        for root_dir, _, filenames in os.walk(p):
            for name in filenames:
                if glob and not fnmatch.fnmatch(name, glob):
                    continue
                files_to_search.append(Path(root_dir) / name)
    else:
        return {"error": f"Path not found: {path}"}

    if output_mode == "files_with_matches":
        matched_files: list[str] = []
        for fp in files_to_search:
            try:
                text = fp.read_text(encoding="utf-8", errors="replace")
                if regex.search(text):
                    matched_files.append(str(fp))
                    if len(matched_files) >= max_results:
                        break
            except Exception:
                pass
        return {"matches": matched_files, "total": len(matched_files)}

    if output_mode == "count":
        counts: list[dict] = []
        for fp in files_to_search:
            try:
                text = fp.read_text(encoding="utf-8", errors="replace")
                n = len(regex.findall(text))
                if n > 0:
                    counts.append({"file": str(fp), "count": n})
                    if len(counts) >= max_results:
                        break
            except Exception:
                pass
        return {"matches": counts, "total": len(counts)}

    # output_mode == "content" (default)
    results: list[dict] = []
    for fp in files_to_search:
        try:
            lines = fp.read_text(
                encoding="utf-8", errors="replace",
            ).splitlines()
            for i, line in enumerate(lines):
                if regex.search(line):
                    entry: dict[str, Any] = {
                        "file": str(fp),
                        "line": i + 1,
                        "text": line.rstrip()[:500],
                    }
                    if context_lines > 0:
                        start = max(0, i - context_lines)
                        end = min(len(lines), i + context_lines + 1)
                        ctx = [
                            f"{j + 1}: {lines[j].rstrip()[:500]}"
                            for j in range(start, end)
                        ]
                        entry["context"] = ctx
                    results.append(entry)
                    if len(results) >= max_results:
                        break
        except Exception:
            pass
        if len(results) >= max_results:
            break

    return {"matches": results, "total": len(results)}
