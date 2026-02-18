#!/usr/bin/env python3
"""Generate run distribution prompts for edits that need multi-style runs.

For each paragraph REPLACE/INSERT edit, determines if the target has
multiple run styles (RSTs). If so, builds a prompt for the LLM to
distribute text across those styles — mirroring the original docx-agent's
_generate_runs() flow.

Usage:
    python3 generate_run_prompts.py <work_dir>

Arguments:
    work_dir  Directory containing analysis.json and edits.json.

Outputs:
    JSON to stdout:
    {
      "prompts": [
        {
          "edit_index": 0,
          "target_id": "b5",
          "prompt": "Run styles:\n  RS0: [italic, ...]..."
        },
        ...
      ]
    }

    Empty prompts list = no edits need run distribution.
"""

import copy
import json
import os
import re
import sys
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NAMESPACES = {"w": W_NS}

# Register prefixes so ET.tostring() uses w: / r: instead of ns0: / ns1:
ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------

def _describe_run_rpr(run):
    """Human-readable description of a run's rPr formatting."""
    r_pr = run.find("w:rPr", NAMESPACES)
    if r_pr is None:
        return "default"

    descriptions = []
    for child in r_pr:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        val = child.get(f"{{{W_NS}}}val")
        off = val in ("0", "false")
        if tag == "b" and not off:
            descriptions.append("bold")
        elif tag == "i" and not off:
            descriptions.append("italic")
        elif tag == "u":
            descriptions.append("underline")
        elif tag == "sz":
            descriptions.append(f"size:{val or ''}")
        elif tag == "color":
            descriptions.append(f"color:{val or ''}")
        elif tag == "rFonts":
            font = child.get(f"{{{W_NS}}}ascii", "")
            if font:
                descriptions.append(f"font:{font}")

    return ", ".join(descriptions) if descriptions else "default"


def _build_rst_from_run(run):
    """Build RST dict from a <w:r> element.

    Returns:
        Dict with rpr_xml (template with {{content}}) and display_description.
    """
    template_run = copy.deepcopy(run)
    for t in template_run.findall("w:t", NAMESPACES):
        t.text = "{{content}}"
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    rpr_xml = ET.tostring(template_run, encoding="unicode")
    desc = _describe_run_rpr(run)
    return {"rpr_xml": rpr_xml, "display_description": desc}


def _extract_table_paragraph_xml(table_xml, row_idx, col_idx, para_idx):
    """Extract specific paragraph XML from table block."""
    try:
        root = ET.fromstring(table_xml)
        rows = root.findall(".//w:tr", NAMESPACES)
        if row_idx >= len(rows):
            return ""
        cells = rows[row_idx].findall(".//w:tc", NAMESPACES)
        if col_idx >= len(cells):
            return ""
        paragraphs = cells[col_idx].findall("w:p", NAMESPACES)
        if para_idx >= len(paragraphs):
            return ""
        return ET.tostring(paragraphs[para_idx], encoding="unicode")
    except ET.ParseError:
        return ""


def _extract_runs_from_xml(paragraph_xml):
    """Extract RST map + original segments from paragraph XML.

    Returns:
        (rst_map, original_segments) or (None, None) if single-style.
        rst_map: {"RS0": {"rpr_xml": ..., "display_description": ...}, ...}
        original_segments: [("text", "RS0"), ...]
    """
    try:
        root = ET.fromstring(paragraph_xml)
    except ET.ParseError:
        return None, None

    runs = root.findall(".//w:r", NAMESPACES)

    rst_map = {}
    original_segments = []
    fp_to_alias = {}

    for run in runs:
        texts = [t.text for t in run.findall("w:t", NAMESPACES) if t.text]
        if not any(t.strip() for t in texts):
            continue

        rpr_elem = run.find("w:rPr", NAMESPACES)
        fp = ET.tostring(rpr_elem, encoding="unicode") if rpr_elem is not None else ""

        if fp in fp_to_alias:
            alias = fp_to_alias[fp]
        else:
            alias = f"RS{len(rst_map)}"
            fp_to_alias[fp] = alias
            rst_map[alias] = _build_rst_from_run(run)

        original_segments.append(("".join(texts), alias))

    if len(rst_map) < 2:
        return None, None

    return rst_map, original_segments


def _build_prompt(new_text, rst_map, semantic_tag="", original_segments=None):
    """Build run distribution prompt — same format as original docx-agent."""
    style_lines = []
    for alias, rst in rst_map.items():
        desc = rst.get("display_description", "default")
        style_lines.append(f"  {alias}: [{desc}]")
    runs_section = "Run styles:\n" + "\n".join(style_lines)

    if original_segments is not None:
        seg_lines = [f'  "{text}" -> {alias}' for text, alias in original_segments]
        runs_section += "\n\nOriginal text distribution:\n" + "\n".join(seg_lines)
    elif semantic_tag:
        runs_section += f"\nSemantic context: {semantic_tag}"

    return (
        f"{runs_section}\n\n"
        f'New text: "{new_text}"\n\n'
        f"Task:\n"
        f"Distribute the new text across the run styles.\n"
        f"- Preserve formatting for key information (dates, numbers, terms)\n"
        f"- Match the original distribution pattern when possible\n"
        f"- Use run_style alias (RS0, RS1...) for each text segment\n\n"
        f"Output Format (JSON only):\n"
        f'{{"runs": [{{"text": "...", "run_style": "RS0"}}, ...]}}'
    )


# -------------------------------------------------------------------
# Main logic
# -------------------------------------------------------------------

def generate_prompts(work_dir):
    """Generate run distribution prompts for all applicable edits."""
    analysis_path = os.path.join(work_dir, "analysis.json")
    edits_path = os.path.join(work_dir, "edits.json")

    if not os.path.isfile(edits_path):
        print(
            f"Warning: {edits_path} not found. "
            "Run this AFTER writing edits.json (Step 3.5).",
            file=sys.stderr,
        )
        return {"prompts": []}

    with open(analysis_path, "r", encoding="utf-8") as f:
        analysis = json.load(f)

    with open(edits_path, "r", encoding="utf-8") as f:
        edits_data = json.load(f)

    edits = edits_data.get("edits", [])
    blocks = analysis.get("blocks", [])
    pst_dict = analysis.get("paragraph_style_templates", {})
    style_alias_map = analysis.get("style_alias_map", {})

    id_to_block = {b["id"]: b for b in blocks}

    prompts = []

    for i, edit in enumerate(edits):
        action = edit.get("action", "")
        if action not in ("replace", "insert_after", "insert_before"):
            continue

        # Skip table-level edits
        edit_unit = edit.get("edit_unit")
        if edit_unit in ("table", "row", "column"):
            continue

        target_id = edit.get("target_id", "")
        new_text = edit.get("new_text", "")
        semantic_tag = edit.get("semantic_tag", "")

        if not new_text:
            continue

        # Skip if edit already has runs specified
        if edit.get("runs"):
            continue

        base_id = target_id.split(":")[0]
        block = id_to_block.get(base_id)

        if action == "replace" and block:
            # REPLACE: extract original run distribution from block XML
            xml = block.get("xml", "")
            paragraph_xml = xml

            # For table cell paragraphs, extract specific paragraph
            cell_match = re.match(r"b\d+:r(\d+)c(\d+)p(\d+)", target_id)
            if cell_match:
                r_idx = int(cell_match.group(1))
                c_idx = int(cell_match.group(2))
                p_idx = int(cell_match.group(3))
                extracted = _extract_table_paragraph_xml(xml, r_idx, c_idx, p_idx)
                if extracted:
                    paragraph_xml = extracted

            rst_map, original_segments = _extract_runs_from_xml(paragraph_xml)
            if rst_map:
                prompt = _build_prompt(
                    new_text, rst_map,
                    original_segments=original_segments,
                )
                # Include rpr_xml templates so apply_edits can assemble runs
                rst_templates_out = {
                    alias: {"rpr_xml": rst["rpr_xml"]}
                    for alias, rst in rst_map.items()
                }
                prompts.append({
                    "edit_index": i,
                    "target_id": target_id,
                    "rst_aliases": list(rst_map.keys()),
                    "run_style_templates": rst_templates_out,
                    "prompt": prompt,
                })

        elif action in ("insert_after", "insert_before"):
            # INSERT: use PST run_style_templates
            style_alias = edit.get("style_alias", "")
            style_key = style_alias_map.get(style_alias, "")

            pst = pst_dict.get(style_key)
            if not pst:
                continue

            rst_templates = pst.get("run_style_templates", {})
            if len(rst_templates) < 2:
                continue

            # Build rst_map from PST templates
            rst_map = {}
            for alias, rst in rst_templates.items():
                rst_map[alias] = {
                    "rpr_xml": rst.get("rpr_xml", ""),
                    "display_description": rst.get("display_description", "default"),
                }

            prompt = _build_prompt(
                new_text, rst_map, semantic_tag=semantic_tag,
            )
            rst_templates_out = {
                alias: {"rpr_xml": rst["rpr_xml"]}
                for alias, rst in rst_map.items()
            }
            prompts.append({
                "edit_index": i,
                "target_id": target_id,
                "rst_aliases": list(rst_map.keys()),
                "run_style_templates": rst_templates_out,
                "prompt": prompt,
            })

    return {"prompts": prompts}


def main():
    if len(sys.argv) < 2:
        print("Usage: generate_run_prompts.py <work_dir>", file=sys.stderr)
        sys.exit(1)

    work_dir = sys.argv[1]
    result = generate_prompts(work_dir)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
