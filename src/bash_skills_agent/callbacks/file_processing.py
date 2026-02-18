"""Agent callbacks for pre/post processing."""

import base64
import logging
import os
import uuid

from google.adk.agents.callback_context import CallbackContext
from google.genai import types

logger = logging.getLogger(__name__)


_turn_counter: dict[str, int] = {}  # session_id → turn count


async def log_reasoning_callback(
    callback_context: CallbackContext,
    llm_response,
) -> None:
    """Log turn info and reasoning content from LLM response."""
    session_id = callback_context.session.id
    _turn_counter[session_id] = _turn_counter.get(session_id, 0) + 1
    turn = _turn_counter[session_id]

    # Log tool calls in this turn
    content = getattr(llm_response, "content", None)
    tool_calls = []
    if content and content.parts:
        for part in content.parts:
            fc = getattr(part, "function_call", None)
            if fc:
                args = dict(fc.args) if fc.args else {}
                # For bash, show the command; for others show key args
                if fc.name == "bash":
                    summary = args.get("command", "")[:200]
                elif fc.name in ("read_file", "write_file", "edit_file"):
                    summary = args.get("file_path", "")
                elif fc.name == "grep_search":
                    summary = f"pattern={args.get('pattern', '')} path={args.get('path', '')}"
                else:
                    summary = str(args)[:150]
                tool_calls.append(f"{fc.name}({summary})")
            if getattr(part, "thought", False) and part.text:
                logger.info(
                    "\n===== Reasoning (%s) =====\n%s\n=============================",
                    callback_context.agent_name,
                    part.text,
                )

    if tool_calls:
        for tc in tool_calls:
            logger.info("[Turn %d] %s", turn, tc)
    else:
        logger.info("[Turn %d] Final response", turn)

    return None


# =====================================================================
# File extraction callback
# =====================================================================

async def _process_part(
    callback_context: CallbackContext,
    part: types.Part,
    uploaded: list[dict],
    workspace: str,
) -> types.Part:
    """Extract DOCX from a Part, write to session workspace, save as artifact."""
    name = None
    data_raw = None

    if inline_data := getattr(part, "inline_data", None):
        name = getattr(inline_data, "display_name", None) or f"upload_{uuid.uuid4().hex[:8]}.docx"
        data_raw = getattr(inline_data, "data", None)
    elif file_data := getattr(part, "file_data", None):
        if getattr(file_data, "data", None) is not None:
            name = (
                getattr(file_data, "file_uri", None)
                or getattr(file_data, "display_name", None)
                or f"upload_{uuid.uuid4().hex[:8]}.docx"
            )
            data_raw = file_data.data

    if not (name and data_raw):
        return part

    data = base64.b64decode(data_raw) if isinstance(data_raw, str) else data_raw
    if not (name.lower().endswith(".docx") or data[:4] == b"PK\x03\x04"):
        return part

    safe_name = os.path.basename(name)
    if not safe_name.lower().endswith(".docx"):
        safe_name += ".docx"

    if any(a["original_name"] == name for a in uploaded):
        return types.Part(text=f"[File already uploaded: {name}]")

    # Deduplicate filename
    file_path = os.path.join(workspace, safe_name)
    if os.path.exists(file_path):
        stem, ext = os.path.splitext(safe_name)
        i = 1
        while os.path.exists(os.path.join(workspace, f"{stem}_{i}{ext}")):
            i += 1
        safe_name = f"{stem}_{i}{ext}"
        file_path = os.path.join(workspace, safe_name)

    # Write to session workspace → mounted as /workspace in the container
    with open(file_path, "wb") as f:
        f.write(data)
    logger.info("Wrote file to workspace: %s", file_path)

    # Save as ADK artifact (versioned)
    version = None
    try:
        docx_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        artifact_part = types.Part(inline_data=types.Blob(mime_type=docx_mime, data=data))
        version = await callback_context.save_artifact(safe_name, artifact_part)
    except ValueError as e:
        logger.warning("Artifact service not available: %s", e)

    uploaded.append({
        "original_name": name,
        "file_name": safe_name,
        "version": version,
        "path": f"/workspace/{safe_name}",
    })
    return types.Part(text=f"[File uploaded: {name} → /workspace/{safe_name}]")


async def extract_files_callback(
    callback_context: CallbackContext,
) -> None:
    """Before-agent callback: extract binary DOCX files to session workspace + artifact."""
    from ..config.shared_clients import get_session_workspace

    session = callback_context.session
    uploaded: list[dict] = list(session.state.get("uploaded_files", []))
    processed_idx: int = session.state.get("_files_processed_idx", 0)

    workspace = get_session_workspace(session.id)
    events = session.events

    # Process ALL events every turn — Part replacements are in-memory only
    # and don't persist to DB, so binary Parts must be replaced each turn.
    # The dedup check in _process_part prevents re-writing files to disk.
    for event in events:
        if event.author != "user" or not event.content:
            continue
        event.content.parts = [
            await _process_part(callback_context, part, uploaded, workspace)
            for part in event.content.parts
        ]

    # Record mtimes of uploaded files so export_outputs_callback skips them
    mtimes: dict[str, float] = dict(session.state.get("_workspace_mtimes", {}))
    for info in uploaded:
        fp = os.path.join(workspace, info["file_name"])
        if os.path.isfile(fp):
            mtimes[info["file_name"]] = os.path.getmtime(fp)
    callback_context.state["_workspace_mtimes"] = mtimes

    callback_context.state["_files_processed_idx"] = len(events)
    if uploaded:
        callback_context.state["uploaded_files"] = uploaded


# =====================================================================
# Output artifact callback
# =====================================================================

_ARTIFACT_EXTENSIONS = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".pdf": "application/pdf",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".csv": "text/csv",
    ".json": "application/json",
    ".zip": "application/zip",
}


def _log_token_usage(session) -> None:
    """Compute and log cumulative token usage from session events."""
    prompt_total = 0
    candidates_total = 0
    for event in session.events:
        if meta := getattr(event, "usage_metadata", None):
            prompt_total += meta.prompt_token_count or 0
            candidates_total += meta.candidates_token_count or 0
    if prompt_total or candidates_total:
        total = prompt_total + candidates_total
        logger.info(
            "Token usage: %s input + %s output = %s total",
            f"{prompt_total:,}", f"{candidates_total:,}", f"{total:,}",
        )


async def export_outputs_callback(
    callback_context: CallbackContext,
) -> None:
    """After-agent callback: save new/modified workspace files as downloadable artifacts.

    Tracks file mtimes via session state. Only saves files that changed since
    the last run, avoiding duplicate artifact versions.
    """
    from ..config.shared_clients import get_session_workspace

    session = callback_context.session
    workspace = get_session_workspace(session.id)

    if not os.path.isdir(workspace):
        return

    # {filename: mtime} from previous run
    prev_mtimes: dict[str, float] = session.state.get("_workspace_mtimes", {})
    new_mtimes: dict[str, float] = {}

    for name in os.listdir(workspace):
        fp = os.path.join(workspace, name)
        if not os.path.isfile(fp):
            continue

        ext = os.path.splitext(name)[1].lower()
        if ext not in _ARTIFACT_EXTENSIONS:
            continue

        mtime = os.path.getmtime(fp)
        new_mtimes[name] = mtime

        # Skip if unchanged
        if name in prev_mtimes and prev_mtimes[name] == mtime:
            continue

        try:
            with open(fp, "rb") as f:
                data = f.read()
            mime = _ARTIFACT_EXTENSIONS[ext]
            artifact_part = types.Part(inline_data=types.Blob(mime_type=mime, data=data))
            version = await callback_context.save_artifact(name, artifact_part)
            logger.info("Exported artifact: %s (v%d, %d bytes)", name, version, len(data))
        except Exception:
            logger.exception("Failed to export artifact: %s", name)

    callback_context.state["_workspace_mtimes"] = new_mtimes

    # Log token usage summary
    _log_token_usage(session)
