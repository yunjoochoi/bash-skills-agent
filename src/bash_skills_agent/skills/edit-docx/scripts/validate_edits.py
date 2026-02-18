#!/usr/bin/env python3
"""Pre-apply validation of edits.json against analysis.json.

Validates target_id existence, style alias references, table field
completeness, paragraph newline rules, and optional run distribution
specs — all before apply_edits.py touches the XML.

Usage:
    python3 validate_edits.py <work_dir>

Reads:
    <work_dir>/analysis.json
    <work_dir>/edits.json

Outputs:
    JSON to stdout: {"valid": bool, "errors": [...], "warnings": [...]}
    Exit code: 0 if valid, 1 if errors found.
"""

import json
import os
import re
import sys
from xml.etree import ElementTree as ET

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

# target_id patterns:
#   b5          -> block-level
#   b5:r2       -> table row
#   b5:r2c1     -> table cell
#   b5:r2c1p0   -> cell paragraph
#   b5:c1       -> table column
#   b5:p1       -> SDT/TOC paragraph
_RE_BLOCK = re.compile(r"^b(\d+)$")
_RE_ROW = re.compile(r"^b(\d+):r(\d+)$")
_RE_CELL = re.compile(r"^b(\d+):r(\d+)c(\d+)$")
_RE_CELL_PARA = re.compile(r"^b(\d+):r(\d+)c(\d+)p(\d+)$")
_RE_COL = re.compile(r"^b(\d+):c(\d+)$")
_RE_SDT_PARA = re.compile(r"^b(\d+):p(\d+)$")

VALID_ACTIONS = {"replace", "insert_after", "insert_before", "delete"}
PARAGRAPH_TAGS = {"H1", "H2", "H3", "BODY", "LIST", "TITLE", "SUBTITLE", "OTHER"}
TABLE_TAG = "TBL"


# -------------------------------------------------------------------
# Data loading
# -------------------------------------------------------------------

def load_inputs(work_dir):
    """Load analysis.json and edits.json from *work_dir*."""
    analysis_path = os.path.join(work_dir, "analysis.json")
    edits_path = os.path.join(work_dir, "edits.json")

    with open(analysis_path, "r", encoding="utf-8") as f:
        analysis = json.load(f)
    with open(edits_path, "r", encoding="utf-8") as f:
        edits_data = json.load(f)

    return analysis, edits_data.get("edits", [])


# -------------------------------------------------------------------
# Block index builder
# -------------------------------------------------------------------

class _TableMeta:
    """Lightweight metadata extracted from a table block's XML."""

    __slots__ = ("row_count", "col_counts", "cell_para_counts")

    def __init__(self, xml_str):
        ns = NAMESPACES["w"]
        root = ET.fromstring(xml_str)
        rows = root.findall(f".//{{{ns}}}tr")
        self.row_count = len(rows)
        self.col_counts = []
        self.cell_para_counts = {}
        for ri, row in enumerate(rows):
            cells = row.findall(f"{{{ns}}}tc")
            self.col_counts.append(len(cells))
            for ci, cell in enumerate(cells):
                paras = cell.findall(f"{{{ns}}}p")
                self.cell_para_counts[(ri, ci)] = len(paras)


class _SdtMeta:
    """Lightweight metadata extracted from an SDT block's XML."""

    __slots__ = ("para_count",)

    def __init__(self, xml_str):
        ns = NAMESPACES["w"]
        root = ET.fromstring(xml_str)
        content = root.find(f"{{{ns}}}sdtContent")
        if content is not None:
            self.para_count = len(content.findall(f"{{{ns}}}p"))
        else:
            self.para_count = 0


class BlockIndex:
    """Pre-parsed lookup structures for fast validation."""

    def __init__(self, blocks):
        self.id_to_block = {}
        self.table_meta = {}
        self.sdt_meta = {}

        for b in blocks:
            bid = b["id"]
            self.id_to_block[bid] = b
            btype = b.get("type", "")
            xml_str = b.get("xml", "")
            if not xml_str:
                continue
            try:
                if btype == "tbl":
                    self.table_meta[bid] = _TableMeta(xml_str)
                elif btype == "sdt":
                    self.sdt_meta[bid] = _SdtMeta(xml_str)
            except ET.ParseError:
                pass


# -------------------------------------------------------------------
# Individual validators
# -------------------------------------------------------------------

def _err(edit_idx, target_id, check, msg):
    return {"edit_index": edit_idx, "target_id": target_id,
            "check": check, "level": "error", "message": msg}


def _warn(edit_idx, target_id, check, msg):
    return {"edit_index": edit_idx, "target_id": target_id,
            "check": check, "level": "warning", "message": msg}


def validate_target_ids(edits, index):
    """Check every target_id references an existing block / row / cell."""
    issues = []
    for i, edit in enumerate(edits):
        tid = edit.get("target_id", "")
        action = edit.get("action", "")

        # Block-level
        m = _RE_BLOCK.match(tid)
        if m:
            bid = f"b{m.group(1)}"
            if bid not in index.id_to_block:
                issues.append(_err(i, tid, "target_id",
                                   f"Block {bid} does not exist"))
            continue

        # Table row
        m = _RE_ROW.match(tid)
        if m:
            bid, ri = f"b{m.group(1)}", int(m.group(2))
            if bid not in index.id_to_block:
                issues.append(_err(i, tid, "target_id",
                                   f"Block {bid} does not exist"))
            elif bid not in index.table_meta:
                issues.append(_err(i, tid, "target_id",
                                   f"{bid} is not a table"))
            else:
                tm = index.table_meta[bid]
                if action == "delete" and ri >= tm.row_count:
                    issues.append(_err(
                        i, tid, "target_id",
                        f"Row {ri} out of range (table has {tm.row_count} rows: 0-{tm.row_count - 1})"))
                elif action not in ("insert_after", "insert_before") and ri >= tm.row_count:
                    issues.append(_err(
                        i, tid, "target_id",
                        f"Row {ri} out of range (table has {tm.row_count} rows: 0-{tm.row_count - 1})"))
            continue

        # Table cell (no paragraph)
        m = _RE_CELL.match(tid)
        if m:
            bid, ri, ci = f"b{m.group(1)}", int(m.group(2)), int(m.group(3))
            if bid not in index.table_meta:
                issues.append(_err(i, tid, "target_id",
                                   f"{bid} is not a table"))
            else:
                tm = index.table_meta[bid]
                if ri >= tm.row_count:
                    issues.append(_err(i, tid, "target_id",
                                       f"Row {ri} out of range"))
                elif ci >= tm.col_counts[ri]:
                    issues.append(_err(i, tid, "target_id",
                                       f"Col {ci} out of range in row {ri}"))
            continue

        # Table cell paragraph
        m = _RE_CELL_PARA.match(tid)
        if m:
            bid = f"b{m.group(1)}"
            ri, ci, pi = int(m.group(2)), int(m.group(3)), int(m.group(4))
            if bid not in index.table_meta:
                issues.append(_err(i, tid, "target_id",
                                   f"{bid} is not a table"))
            else:
                tm = index.table_meta[bid]
                if ri >= tm.row_count:
                    issues.append(_err(i, tid, "target_id",
                                       f"Row {ri} out of range"))
                elif ci >= tm.col_counts[ri]:
                    issues.append(_err(i, tid, "target_id",
                                       f"Col {ci} out of range in row {ri}"))
                elif pi >= tm.cell_para_counts.get((ri, ci), 0):
                    issues.append(_err(
                        i, tid, "target_id",
                        f"Para {pi} out of range in cell ({ri},{ci})"))
            continue

        # Table column
        m = _RE_COL.match(tid)
        if m:
            bid, ci = f"b{m.group(1)}", int(m.group(2))
            if bid not in index.table_meta:
                issues.append(_err(i, tid, "target_id",
                                   f"{bid} is not a table"))
            else:
                tm = index.table_meta[bid]
                first_row_cols = tm.col_counts[0] if tm.col_counts else 0
                if action == "delete" and ci >= first_row_cols:
                    issues.append(_err(
                        i, tid, "target_id",
                        f"Col {ci} out of range ({first_row_cols} cols)"))
            continue

        # SDT / TOC paragraph
        m = _RE_SDT_PARA.match(tid)
        if m:
            bid, pi = f"b{m.group(1)}", int(m.group(2))
            if bid not in index.id_to_block:
                issues.append(_err(i, tid, "target_id",
                                   f"Block {bid} does not exist"))
            elif bid not in index.sdt_meta:
                issues.append(_err(i, tid, "target_id",
                                   f"{bid} is not an SDT/TOC block"))
            else:
                sm = index.sdt_meta[bid]
                if pi >= sm.para_count:
                    issues.append(_err(
                        i, tid, "target_id",
                        f"Para {pi} out of range (SDT has {sm.para_count} paras: 0-{sm.para_count - 1})"))
            continue

        issues.append(_err(i, tid, "target_id",
                           f"Unrecognised target_id format: {tid}"))

    return issues


def validate_semantic_tags(edits, index):
    """Warn when edit semantic_tag disagrees with the actual block type."""
    issues = []
    for i, edit in enumerate(edits):
        action = edit.get("action", "")
        if action not in ("replace", "delete"):
            continue

        tid = edit.get("target_id", "")
        m = _RE_BLOCK.match(tid)
        if not m:
            continue

        bid = f"b{m.group(1)}"
        block = index.id_to_block.get(bid)
        if not block:
            continue

        expected = block.get("semantic_tag", "")
        actual = edit.get("semantic_tag", "")
        if expected and actual and expected != actual:
            issues.append(_warn(
                i, tid, "semantic_tag",
                f"Edit tag '{actual}' differs from block tag '{expected}'"))

    return issues


def validate_style_aliases(edits, alias_map):
    """Verify every referenced alias exists in the style_alias_map."""
    issues = []
    for i, edit in enumerate(edits):
        tid = edit.get("target_id", "")
        action = edit.get("action", "")
        tag = edit.get("semantic_tag", "")

        if action == "delete":
            continue

        # Paragraph style_alias
        if tag in PARAGRAPH_TAGS:
            sa = edit.get("style_alias")
            if sa and sa not in alias_map:
                issues.append(_err(i, tid, "style_alias",
                                   f"Style alias '{sa}' not in alias map"))

        # Table aliases
        if tag == TABLE_TAG:
            tsa = edit.get("table_style_alias")
            if tsa and tsa not in alias_map:
                issues.append(_err(i, tid, "table_style_alias",
                                   f"Table style alias '{tsa}' not in alias map"))

            for rs in (edit.get("row_style_aliases") or []):
                if rs not in alias_map:
                    issues.append(_err(i, tid, "row_style_alias",
                                       f"Row style alias '{rs}' not in alias map"))

            for row_cs in (edit.get("cell_style_aliases") or []):
                if isinstance(row_cs, list):
                    for cs in row_cs:
                        if cs not in alias_map:
                            issues.append(_err(
                                i, tid, "cell_style_alias",
                                f"Cell style alias '{cs}' not in alias map"))
                elif isinstance(row_cs, str):
                    if row_cs not in alias_map:
                        issues.append(_err(
                            i, tid, "cell_style_alias",
                            f"Cell style alias '{row_cs}' not in alias map"))

    return issues


def validate_table_fields(edits, index):
    """Check required fields for table edits."""
    issues = []
    for i, edit in enumerate(edits):
        tid = edit.get("target_id", "")
        tag = edit.get("semantic_tag", "")
        action = edit.get("action", "")
        eu = edit.get("edit_unit")

        if tag != TABLE_TAG:
            continue
        if action == "delete" and not eu:
            continue

        # edit_unit is required for all non-block-level table edits
        if not eu and action != "delete":
            m = _RE_BLOCK.match(tid)
            if not m:
                issues.append(_err(i, tid, "edit_unit",
                                   "edit_unit required for table edits"))

        # Row INSERT must have RS + CS
        if eu == "row" and action in ("insert_after", "insert_before"):
            if not edit.get("row_style_aliases"):
                issues.append(_err(i, tid, "row_style_aliases",
                                   "row_style_aliases required for row INSERT"))
            if not edit.get("cell_style_aliases"):
                issues.append(_err(i, tid, "cell_style_aliases",
                                   "cell_style_aliases required for row INSERT"))

        # Table INSERT must have table_style_alias
        if eu == "table" and action in ("insert_after", "insert_before"):
            if not edit.get("table_style_alias"):
                issues.append(_err(i, tid, "table_style_alias",
                                   "table_style_alias required for table INSERT"))

        # Cell count in new_text vs cell_style_aliases
        if eu == "row" and action in ("insert_after", "insert_before"):
            new_text = edit.get("new_text", "")
            if new_text and "|" in new_text:
                cell_count = len([c.strip() for c in new_text.split("|")])
                cs_list = edit.get("cell_style_aliases") or []
                if cs_list and isinstance(cs_list[0], list):
                    cs_count = len(cs_list[0])
                    if cell_count != cs_count:
                        issues.append(_err(
                            i, tid, "cell_count",
                            f"new_text has {cell_count} cells but "
                            f"cell_style_aliases[0] has {cs_count}"))

    return issues


def validate_newlines(edits):
    """Flag \\n in paragraph new_text (not allowed outside TBL/TOC)."""
    issues = []
    for i, edit in enumerate(edits):
        tid = edit.get("target_id", "")
        tag = edit.get("semantic_tag", "")
        action = edit.get("action", "")
        new_text = edit.get("new_text", "")

        if action == "delete" or not new_text:
            continue
        if tag == TABLE_TAG:
            continue
        if "\n" in new_text:
            issues.append(_err(
                i, tid, "newline",
                "new_text contains '\\n' — split into separate edits"))

    return issues


def validate_column_counts(edits, index):
    """Warn when row INSERT cell count differs from table width."""
    issues = []
    for i, edit in enumerate(edits):
        tid = edit.get("target_id", "")
        eu = edit.get("edit_unit")
        action = edit.get("action", "")
        new_text = edit.get("new_text", "")

        if eu != "row" or action not in ("insert_after", "insert_before"):
            continue
        if not new_text or "|" not in new_text:
            continue

        m = _RE_ROW.match(tid) or _RE_BLOCK.match(tid)
        if not m:
            continue
        bid = f"b{m.group(1)}"
        tm = index.table_meta.get(bid)
        if not tm or not tm.col_counts:
            continue

        text_cols = len([c.strip() for c in new_text.split("|")])
        table_cols = tm.col_counts[0]
        if text_cols != table_cols:
            issues.append(_warn(
                i, tid, "column_count",
                f"new_text has {text_cols} cells but table has {table_cols} columns"))

    return issues


def validate_runs(edits, alias_map, paragraph_style_templates):
    """Validate optional runs spec: aliases exist, text sums to new_text."""
    issues = []
    for i, edit in enumerate(edits):
        tid = edit.get("target_id", "")
        runs_spec = edit.get("runs")
        if not runs_spec or not isinstance(runs_spec, list):
            continue

        new_text = edit.get("new_text", "")
        style_alias = edit.get("style_alias", "")

        # Resolve PST to get available RSTs
        style_key = alias_map.get(style_alias, "")
        pst = paragraph_style_templates.get(style_key, {})
        rst_dict = pst.get("run_style_templates", {})

        # Check each run's alias
        for j, spec in enumerate(runs_spec):
            rs = spec.get("run_style", "")
            if rs and rst_dict and rs not in rst_dict:
                issues.append(_err(
                    i, tid, "runs",
                    f"runs[{j}].run_style '{rs}' not found in RST pool"))

        # Check concatenated text matches new_text
        concat = "".join(spec.get("text", "") for spec in runs_spec)
        if concat != new_text:
            issues.append(_warn(
                i, tid, "runs",
                f"Concatenated runs text does not match new_text"))

    return issues


# -------------------------------------------------------------------
# Orchestrator
# -------------------------------------------------------------------

def validate(work_dir):
    """Run all validations and return structured result."""
    analysis, edits = load_inputs(work_dir)
    blocks = analysis.get("blocks", [])
    alias_map = analysis.get("style_alias_map", {})
    pst = analysis.get("paragraph_style_templates", {})
    index = BlockIndex(blocks)

    errors = []
    warnings = []

    all_issues = []
    all_issues.extend(validate_target_ids(edits, index))
    all_issues.extend(validate_semantic_tags(edits, index))
    all_issues.extend(validate_style_aliases(edits, alias_map))
    all_issues.extend(validate_table_fields(edits, index))
    all_issues.extend(validate_newlines(edits))
    all_issues.extend(validate_column_counts(edits, index))
    all_issues.extend(validate_runs(edits, alias_map, pst))

    for issue in all_issues:
        if issue.get("level") == "warning":
            warnings.append(issue)
        else:
            errors.append(issue)

    return {"valid": len(errors) == 0, "errors": errors, "warnings": warnings}


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 validate_edits.py <work_dir>", file=sys.stderr)
        sys.exit(1)

    work_dir = sys.argv[1]

    if not os.path.isdir(work_dir):
        print(f"Error: work_dir not found: {work_dir}", file=sys.stderr)
        sys.exit(1)

    result = validate(work_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result["valid"] else 1)


if __name__ == "__main__":
    main()
