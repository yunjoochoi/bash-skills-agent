#!/usr/bin/env python3
"""Apply edits to an extracted DOCX document.xml."""

import copy
import json
import os
import re
import sys
from xml.etree import ElementTree as ET

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

for _pfx, _uri in NAMESPACES.items():
    ET.register_namespace(_pfx, _uri)

_NON_TEXT_INDICATORS = (
    "AlternateContent",
    "<w:drawing",
    "<w:pict",
    "<wp:anchor",
    "<wp:inline",
    "<v:shape",
    "<v:group",
)


def escape_xml(text):
    """Escape XML reserved characters."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _ensure_t_space_preserve(run_xml):
    """Ensure <w:t> elements have xml:space='preserve'."""
    import re
    def _add_attr(m):
        tag = m.group(0)
        if "xml:space" in tag:
            return tag
        return tag[:-1] + ' xml:space="preserve">'
    return re.sub(r"<w:t(?:\s[^>]*)?>", _add_attr, run_xml)


def _resolve_style_key(style_alias, style_alias_map, fallback=""):
    """Resolve a style alias (S1, S2...) to a style_key via the alias map."""
    if not style_alias:
        return fallback
    return style_alias_map.get(style_alias, fallback)


def _build_run_xml_from_template(new_text, run_style_templates):
    """Build a single <w:r> XML using the first run style template."""
    if not run_style_templates:
        return []
    first_rst = next(iter(run_style_templates.values()))
    rpr_xml = first_rst.get("rpr_xml", "")
    if not rpr_xml:
        return []
    escaped = escape_xml(new_text)
    return [_ensure_t_space_preserve(rpr_xml.replace("{{content}}", escaped))]


def _build_run_xmls_from_spec(runs_spec, run_style_templates):
    """Build <w:r> XML strings from an explicit runs specification."""
    if not run_style_templates:
        return []
    first_rst = next(iter(run_style_templates.values()))
    result = []
    for spec in runs_spec:
        text = spec.get("text", "")
        rs_alias = spec.get("run_style", "")
        rst = run_style_templates.get(rs_alias, first_rst)
        rpr_xml = rst.get("rpr_xml", "")
        if not rpr_xml:
            continue
        escaped = escape_xml(text)
        result.append(_ensure_t_space_preserve(rpr_xml.replace("{{content}}", escaped)))
    return result


def generate_new_blocks(edits, analysis):
    """Generate new_block dicts from edits + analysis (Phase A)."""
    style_alias_map = analysis.get("style_alias_map", {})
    paragraph_style_templates = analysis.get("paragraph_style_templates", {})
    blocks = analysis.get("blocks", [])

    id_to_block = {b["id"]: b for b in blocks}

    new_blocks = []

    for edit in edits:
        action = edit.get("action", "")
        target_id = edit.get("target_id", "")
        semantic_tag = edit.get("semantic_tag", "")
        new_text = edit.get("new_text", "")
        style_alias = edit.get("style_alias")
        edit_unit = edit.get("edit_unit")

        base_id = target_id.split(":")[0]

        if action == "delete":
            nb = {
                "action": "delete",
                "target_id": target_id,
                "style_key": "",
                "content": "",
                "run_xmls": [],
                "edit_unit": edit_unit,
                "row_style_aliases": None,
                "cell_style_aliases": None,
            }
            new_blocks.append(nb)
            continue

        if action in ("insert_after", "insert_before"):
            if semantic_tag == "TBL" or edit_unit:
                style_key = ""
                tbl_alias = edit.get("table_style_alias") or style_alias
                if edit_unit == "table" and tbl_alias:
                    style_key = _resolve_style_key(
                        tbl_alias, style_alias_map,
                    )
                nb = {
                    "action": action,
                    "target_id": target_id,
                    "style_key": style_key,
                    "content": new_text or "",
                    "run_xmls": [],
                    "edit_unit": edit_unit,
                    "row_style_aliases": edit.get("row_style_aliases"),
                    "cell_style_aliases": edit.get("cell_style_aliases"),
                }
                new_blocks.append(nb)
                continue

            style_key = _resolve_style_key(
                style_alias, style_alias_map,
            )
            if not style_key:
                if paragraph_style_templates:
                    style_key = next(iter(paragraph_style_templates))

            run_xmls = []
            runs_spec = edit.get("runs")
            edit_rst = edit.get("run_style_templates")
            if not edit_rst:
                pst = paragraph_style_templates.get(style_key, {})
                edit_rst = pst.get("run_style_templates", {})
            if edit_rst and new_text:
                if runs_spec and isinstance(runs_spec, list):
                    run_xmls = _build_run_xmls_from_spec(runs_spec, edit_rst)
                else:
                    run_xmls = _build_run_xml_from_template(new_text, edit_rst)

            nb = {
                "action": action,
                "target_id": target_id,
                "style_key": style_key,
                "content": new_text or "",
                "run_xmls": run_xmls,
                "edit_unit": None,
                "row_style_aliases": None,
                "cell_style_aliases": None,
            }
            new_blocks.append(nb)
            continue

        if action == "replace":
            if semantic_tag == "TBL" or edit_unit:
                style_key = ""
                if style_alias:
                    style_key = _resolve_style_key(
                        style_alias, style_alias_map,
                    )
                nb = {
                    "action": "replace",
                    "target_id": target_id,
                    "style_key": style_key,
                    "content": new_text or "",
                    "run_xmls": [],
                    "edit_unit": edit_unit,
                    "row_style_aliases": edit.get("row_style_aliases"),
                    "cell_style_aliases": edit.get("cell_style_aliases"),
                }
                new_blocks.append(nb)
                continue

            block = id_to_block.get(base_id)
            style_key = ""
            if style_alias:
                style_key = _resolve_style_key(
                    style_alias, style_alias_map,
                )
            if not style_key and block:
                style_key = block.get("style_key", "")

            run_xmls = []
            if block and new_text:
                xml_str = block.get("xml", "")
                if not _has_non_text_content(xml_str):
                    edit_rst = edit.get("run_style_templates")
                    if not edit_rst:
                        pst = paragraph_style_templates.get(style_key, {})
                        edit_rst = pst.get("run_style_templates", {})
                    if edit_rst:
                        runs_spec = edit.get("runs")
                        if runs_spec and isinstance(runs_spec, list):
                            run_xmls = _build_run_xmls_from_spec(
                                runs_spec, edit_rst,
                            )
                        else:
                            run_xmls = _build_run_xml_from_template(
                                new_text, edit_rst,
                            )

            nb = {
                "action": "replace",
                "target_id": target_id,
                "style_key": style_key,
                "content": new_text or "",
                "run_xmls": run_xmls,
                "edit_unit": edit_unit,
                "row_style_aliases": edit.get("row_style_aliases"),
                "cell_style_aliases": edit.get("cell_style_aliases"),
            }
            new_blocks.append(nb)

    return new_blocks


def _has_non_text_content(xml_str):
    """Detect images/drawings in paragraph XML."""
    for indicator in _NON_TEXT_INDICATORS:
        if indicator in xml_str:
            return True
    return False


def _extract_bookmarks_xml(xml_str):
    """Extract bookmarkStart/End XML strings from a paragraph."""
    try:
        root = ET.fromstring(xml_str)
    except ET.ParseError:
        return [], []
    ns_w = NAMESPACES["w"]
    starts = [ET.tostring(e, encoding="unicode")
              for e in root.findall(f"{{{ns_w}}}bookmarkStart")]
    ends = [ET.tostring(e, encoding="unicode")
            for e in root.findall(f"{{{ns_w}}}bookmarkEnd")]
    return starts, ends


def _inject_bookmarks_into_para(para_xml, bk_starts, bk_ends):
    """Inject bookmark elements into a paragraph XML string."""
    if not bk_starts and not bk_ends:
        return para_xml
    try:
        root = ET.fromstring(para_xml)
        ns_w = NAMESPACES["w"]

        ppr = root.find(f"{{{ns_w}}}pPr")
        children = list(root)
        insert_idx = (children.index(ppr) + 1) if ppr is not None else 0

        for bk_xml_str in reversed(bk_starts):
            try:
                root.insert(insert_idx, ET.fromstring(bk_xml_str))
            except ET.ParseError:
                pass

        for bk_xml_str in bk_ends:
            try:
                root.append(ET.fromstring(bk_xml_str))
            except ET.ParseError:
                pass

        return ET.tostring(root, encoding="unicode")
    except ET.ParseError:
        return para_xml


def _normalize_newlines(text):
    """Normalize literal backslash-n to actual newlines."""
    if "\n" not in text and "\\n" in text:
        return text.replace("\\n", "\n")
    return text


def _replace_text_preserving_structure(original_xml, new_text):
    """In-place text replacement preserving paragraph structure."""
    try:
        root = ET.fromstring(original_xml)
    except ET.ParseError:
        return original_xml

    first_set = False
    for run in root.findall(".//w:r", NAMESPACES):
        for t in run.findall("w:t", NAMESPACES):
            if not first_set:
                t.text = new_text
                t.set(
                    "{http://www.w3.org/XML/1998/namespace}space",
                    "preserve",
                )
                first_set = True
            else:
                t.text = ""

    return ET.tostring(root, encoding="unicode")


def apply_mapping_to_blocks(blocks, new_blocks):
    """Mark blocks with internal edit flags from new_blocks."""
    blocks = [dict(b) for b in blocks]
    id_to_idx = {b["id"]: i for i, b in enumerate(blocks)}

    for nb in new_blocks:
        target_id = nb["target_id"].split(":")[0]
        if target_id not in id_to_idx:
            print(f"  [WARN] Target ID not found: {nb['target_id']}")
            continue

        idx = id_to_idx[target_id]
        block_type = blocks[idx].get("type", "")
        is_table = block_type == "tbl"
        is_sdt = block_type == "sdt"

        sub_coord = (
            nb["target_id"][len(target_id) + 1:]
            if ":" in nb["target_id"]
            else ""
        )

        action = nb["action"]

        if action == "replace":
            blocks[idx]["_replaced"] = True
            if "_replacements" not in blocks[idx]:
                blocks[idx]["_replacements"] = []
            blocks[idx]["_replacements"].append({
                "style_key": nb.get("style_key", ""),
                "content": nb.get("content", ""),
                "original_target_id": nb["target_id"],
                "edit_unit": nb.get("edit_unit"),
                "run_xmls": list(nb.get("run_xmls") or []),
                "row_style_aliases": nb.get("row_style_aliases"),
                "cell_style_aliases": nb.get("cell_style_aliases"),
            })

        elif action == "delete":
            sdt_para_match = re.match(r"p(\d+)$", sub_coord)
            if is_sdt and sdt_para_match:
                p_idx = int(sdt_para_match.group(1))
                blocks[idx].setdefault("_sdt_entry_deletions", []).append(
                    p_idx,
                )
                blocks[idx]["_replaced"] = True
                blocks[idx].setdefault("_replacements", [])
                continue

            row_match = re.match(r"r(\d+)$", sub_coord)
            col_match = re.match(r"c(\d+)$", sub_coord)

            eu = nb.get("edit_unit")
            is_row_del = (eu == "row") if eu else (is_table and row_match is not None)
            is_col_del = (eu == "column") if eu else (is_table and col_match is not None)

            if is_row_del and row_match:
                r_idx = int(row_match.group(1))
                blocks[idx].setdefault("_row_deletions", []).append(r_idx)
                blocks[idx]["_replaced"] = True
                blocks[idx].setdefault("_replacements", [])
            elif is_col_del and col_match:
                c_idx = int(col_match.group(1))
                blocks[idx].setdefault("_col_deletions", []).append(c_idx)
                blocks[idx]["_replaced"] = True
                blocks[idx].setdefault("_replacements", [])
            else:
                cell_para_del = re.match(r"r(\d+)c(\d+)p(\d+)$", sub_coord)
                if is_table and cell_para_del:
                    blocks[idx].setdefault("_para_deletions", []).append({
                        "row_idx": int(cell_para_del.group(1)),
                        "col_idx": int(cell_para_del.group(2)),
                        "para_idx": int(cell_para_del.group(3)),
                    })
                    blocks[idx]["_replaced"] = True
                    blocks[idx].setdefault("_replacements", [])
                else:
                    blocks[idx]["_deleted"] = True

        elif action == "insert_after":
            sdt_para_match = re.match(r"p(\d+)$", sub_coord)
            if is_sdt and sdt_para_match:
                p_idx = int(sdt_para_match.group(1))
                insert_list = blocks[idx].setdefault(
                    "_sdt_entry_inserts", [],
                )
                insert_list.append({
                    "para_idx": p_idx,
                    "content": nb.get("content", ""),
                    "insert_after": True,
                    "_insert_order": len(insert_list),
                })
                blocks[idx]["_replaced"] = True
                blocks[idx].setdefault("_replacements", [])
                continue

            col_match_ia = re.match(r"c(\d+)$", sub_coord)
            eu_ia = nb.get("edit_unit")
            is_col_insert_ia = (
                (eu_ia == "column")
                if eu_ia
                else (is_table and col_match_ia is not None)
            )

            if is_col_insert_ia and col_match_ia:
                c_idx = int(col_match_ia.group(1))
                cs_aliases = nb.get("cell_style_aliases") or []
                if cs_aliases and isinstance(cs_aliases[0], list):
                    cs_aliases = [row[0] for row in cs_aliases if row]
                blocks[idx].setdefault("_col_inserts_after", []).append({
                    "col_idx": c_idx,
                    "content": nb.get("content", ""),
                    "original_target_id": nb["target_id"],
                    "cell_style_aliases": cs_aliases,
                    "style_key": nb.get("style_key", ""),
                })
                blocks[idx]["_replaced"] = True
                blocks[idx].setdefault("_replacements", [])
            elif is_col_insert_ia and not col_match_ia:
                raise ValueError(
                    f"Column insert_after requires bN:cN target_id format, "
                    f"got '{nb['target_id']}'. sub_coord '{sub_coord}' does "
                    f"not match 'c(\\d+)$'."
                )
            else:
                blocks[idx].setdefault("_inserts_after", []).append({
                    "style_key": nb.get("style_key", ""),
                    "content": nb.get("content", ""),
                    "original_target_id": nb["target_id"],
                    "edit_unit": nb.get("edit_unit"),
                    "row_style_aliases": nb.get("row_style_aliases"),
                    "cell_style_aliases": nb.get("cell_style_aliases"),
                })

        elif action == "insert_before":
            sdt_para_match = re.match(r"p(\d+)$", sub_coord)
            if is_sdt and sdt_para_match:
                p_idx = int(sdt_para_match.group(1))
                insert_list = blocks[idx].setdefault(
                    "_sdt_entry_inserts", [],
                )
                insert_list.append({
                    "para_idx": p_idx,
                    "content": nb.get("content", ""),
                    "insert_after": False,
                    "_insert_order": len(insert_list),
                })
                blocks[idx]["_replaced"] = True
                blocks[idx].setdefault("_replacements", [])
                continue

            col_match_ib = re.match(r"c(\d+)$", sub_coord)
            eu_ib = nb.get("edit_unit")
            is_col_insert_ib = (
                (eu_ib == "column")
                if eu_ib
                else (is_table and col_match_ib is not None)
            )

            if is_col_insert_ib and col_match_ib:
                c_idx = int(col_match_ib.group(1))
                cs_aliases = nb.get("cell_style_aliases") or []
                if cs_aliases and isinstance(cs_aliases[0], list):
                    cs_aliases = [row[0] for row in cs_aliases if row]
                blocks[idx].setdefault("_col_inserts_before", []).append({
                    "col_idx": c_idx,
                    "content": nb.get("content", ""),
                    "original_target_id": nb["target_id"],
                    "cell_style_aliases": cs_aliases,
                    "style_key": nb.get("style_key", ""),
                })
                blocks[idx]["_replaced"] = True
                blocks[idx].setdefault("_replacements", [])
            elif is_col_insert_ib and not col_match_ib:
                raise ValueError(
                    f"Column insert_before requires bN:cN target_id format, "
                    f"got '{nb['target_id']}'. sub_coord '{sub_coord}' does "
                    f"not match 'c(\\d+)$'."
                )
            else:
                blocks[idx].setdefault("_inserts_before", []).append({
                    "style_key": nb.get("style_key", ""),
                    "content": nb.get("content", ""),
                    "original_target_id": nb["target_id"],
                    "edit_unit": nb.get("edit_unit"),
                    "row_style_aliases": nb.get("row_style_aliases"),
                    "cell_style_aliases": nb.get("cell_style_aliases"),
                })

    return blocks


def _build_block_xml(block_spec, paragraph_style_templates, table_style_templates,
                     style_alias_map):
    """Generate XML for a new/replacement block."""
    ns = NAMESPACES["w"]
    style_key = block_spec.get("style_key", "")
    content = block_spec.get("content", "")
    run_xmls = block_spec.get("run_xmls", [])

    if style_key in table_style_templates:
        return _build_table_from_template(
            table_style_templates[style_key],
            content,
            row_style_aliases=block_spec.get("row_style_aliases") or [],
            cell_style_aliases_per_row=block_spec.get("cell_style_aliases") or [],
            paragraph_style_templates=paragraph_style_templates,
            table_style_key=style_key,
            style_alias_map=style_alias_map,
        )

    if style_key not in paragraph_style_templates:
        if paragraph_style_templates:
            style_key = next(iter(paragraph_style_templates))
        else:
            return None

    template = paragraph_style_templates[style_key]

    if run_xmls:
        ppr = template.get("ppr_xml_template") or f'<w:pPr xmlns:w="{ns}"/>'
        return (
            f'<w:p xmlns:w="{ns}">'
            f'{ppr}{"".join(run_xmls)}</w:p>'
        )

    rst_dict = template.get("run_style_templates", {})
    if rst_dict:
        first_rst = next(iter(rst_dict.values()))
        escaped = escape_xml(content)
        run_xml = _ensure_t_space_preserve(
            first_rst.get("rpr_xml", "").replace("{{content}}", escaped)
        )
        ppr = template.get("ppr_xml_template") or f'<w:pPr xmlns:w="{ns}"/>'
        return f'<w:p xmlns:w="{ns}">{ppr}{run_xml}</w:p>'

    escaped = escape_xml(content)
    ppr = template.get("ppr_xml_template") or f'<w:pPr xmlns:w="{ns}"/>'
    return (
        f'<w:p xmlns:w="{ns}">{ppr}'
        f'<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r></w:p>'
    )


def _append_paragraphs_from_content(body_parts, template, content, ns):
    """Append one or more paragraphs from template + plain content."""
    rst_dict = template.get("run_style_templates", {})
    ppr = template.get("ppr_xml_template") or f'<w:pPr xmlns:w="{ns}"/>'

    def _build_para(text):
        escaped = escape_xml(text)
        if rst_dict:
            first_rst = next(iter(rst_dict.values()))
            run_xml = _ensure_t_space_preserve(
                first_rst.get("rpr_xml", "").replace(
                    "{{content}}", escaped,
                )
            )
        else:
            run_xml = (
                f'<w:r xmlns:w="{ns}">'
                f'<w:t xml:space="preserve">{escaped}</w:t></w:r>'
            )
        return f'<w:p xmlns:w="{ns}">{ppr}{run_xml}</w:p>'

    if "\n" in content and "|" not in content:
        lines = [ln.strip() for ln in content.split("\n") if ln.strip()]
        for line in lines:
            body_parts.append(_build_para(line))
    else:
        body_parts.append(_build_para(content))


def _table_replace_paragraph(table_xml, r_idx, c_idx, p_idx, new_text,
                             run_xmls=None, style_key="",
                             paragraph_style_templates=None):
    """Replace text in a specific paragraph within a table cell."""
    try:
        root = ET.fromstring(table_xml)
        xml_rows = root.findall(".//w:tr", NAMESPACES)
        if r_idx >= len(xml_rows):
            return table_xml

        xml_cells = xml_rows[r_idx].findall(".//w:tc", NAMESPACES)
        if c_idx >= len(xml_cells):
            return table_xml

        target_cell = xml_cells[c_idx]
        paragraphs = target_cell.findall("w:p", NAMESPACES)
        if p_idx >= len(paragraphs):
            return table_xml

        target_para = paragraphs[p_idx]

        if run_xmls:
            existing_runs = target_para.findall("w:r", NAMESPACES)
            for run in existing_runs:
                target_para.remove(run)
            for run_xml_str in run_xmls:
                try:
                    new_run = ET.fromstring(run_xml_str)
                    target_para.append(new_run)
                except ET.ParseError:
                    pass
        else:
            para_runs = target_para.findall("w:r", NAMESPACES)

            if "\n" in new_text:
                lines = [ln for ln in new_text.split("\n") if ln.strip()]
                if not lines:
                    lines = [""]

                first_escaped = escape_xml(lines[0])
                first_text_set = False
                for run in para_runs:
                    for t in run.findall("w:t", NAMESPACES):
                        if not first_text_set:
                            t.text = first_escaped
                            t.set(
                                "{http://www.w3.org/XML/1998/namespace}space",
                                "preserve",
                            )
                            first_text_set = True
                        else:
                            t.text = ""

                insert_ref = target_para
                for extra_line in lines[1:]:
                    new_para = _clone_paragraph_with_text(
                        target_para, escape_xml(extra_line),
                    )
                    children = list(target_cell)
                    pos = next(
                        (i for i, c in enumerate(children) if c is insert_ref),
                        len(children) - 1,
                    )
                    target_cell.insert(pos + 1, new_para)
                    insert_ref = new_para
            else:
                escaped_text = escape_xml(new_text)
                first_text_set = False
                for run in para_runs:
                    for t in run.findall("w:t", NAMESPACES):
                        if not first_text_set:
                            t.text = escaped_text
                            if (
                                t.get("{http://www.w3.org/XML/1998/namespace}space")
                                is None
                            ):
                                t.set(
                                    "{http://www.w3.org/XML/1998/namespace}space",
                                    "preserve",
                                )
                            first_text_set = True
                        else:
                            t.text = ""

        return ET.tostring(root, encoding="unicode")

    except ET.ParseError:
        return table_xml


def _clone_paragraph_with_text(source_para, escaped_text):
    """Clone a paragraph, keep first run style, replace text."""
    new_para = copy.deepcopy(source_para)
    ns_w = NAMESPACES["w"]
    runs = new_para.findall(f"{{{ns_w}}}r")
    if runs:
        first_run = runs[0]
        for run in runs[1:]:
            new_para.remove(run)
        t_elements = first_run.findall(f"{{{ns_w}}}t")
        if t_elements:
            t_elements[0].text = escaped_text
            t_elements[0].set(
                "{http://www.w3.org/XML/1998/namespace}space", "preserve",
            )
            for t in t_elements[1:]:
                first_run.remove(t)
        else:
            t_elem = ET.SubElement(first_run, f"{{{ns_w}}}t")
            t_elem.text = escaped_text
            t_elem.set(
                "{http://www.w3.org/XML/1998/namespace}space", "preserve",
            )
    else:
        new_run = ET.SubElement(new_para, f"{{{ns_w}}}r")
        t_elem = ET.SubElement(new_run, f"{{{ns_w}}}t")
        t_elem.text = escaped_text
        t_elem.set(
            "{http://www.w3.org/XML/1998/namespace}space", "preserve",
        )
    return new_para


def _table_add_row(table_xml, after_r_idx, row_contents,
                   row_style_alias="", cell_style_aliases=None,
                   paragraph_style_templates=None,
                   table_style_templates=None,
                   table_style_key="", style_alias_map=None):
    """Add a new row to a table after the specified row."""
    cell_style_aliases = cell_style_aliases or []
    paragraph_style_templates = paragraph_style_templates or {}
    table_style_templates = table_style_templates or {}
    style_alias_map = style_alias_map or {}

    try:
        root = ET.fromstring(table_xml)
        xml_rows = root.findall(".//w:tr", NAMESPACES)

        if after_r_idx >= len(xml_rows):
            return table_xml

        ns_w = NAMESPACES["w"]
        target_row = xml_rows[after_r_idx]

        num_cols = max(
            len(cell_style_aliases), len(row_contents),
        ) if (cell_style_aliases or row_contents) else len(
            target_row.findall("w:tc", NAMESPACES),
        )

        tbl_grid = root.find("w:tblGrid", NAMESPACES)
        total_width = _get_table_total_width(root)
        col_widths = _extract_column_widths(tbl_grid, num_cols, total_width)

        new_row = ET.Element(f"{{{ns_w}}}tr")
        _apply_row_style(new_row, row_style_alias, style_alias_map)

        tst = table_style_templates.get(table_style_key, {})
        tst_cells = []
        if tst:
            rs_trpr = style_alias_map.get(row_style_alias, "")
            tst_row_styles = tst.get("row_styles", {})
            cst = tst.get("cell_style_templates", {})
            for idx_key in sorted(
                tst_row_styles.keys(),
                key=lambda x: int(x) if x.isdigit() else 999,
            ):
                if tst_row_styles[idx_key].get(
                    "tr_pr_xml_template", ""
                ) == rs_trpr:
                    tst_cells = cst.get(idx_key, [])
                    break

        for col_idx in range(num_cols):
            cs_alias = (
                cell_style_aliases[col_idx]
                if col_idx < len(cell_style_aliases)
                else (cell_style_aliases[-1] if cell_style_aliases else "")
            )
            cell_text = (
                row_contents[col_idx]
                if col_idx < len(row_contents) else ""
            )
            cell_ps = None
            if tst_cells:
                tst_ci = min(col_idx, len(tst_cells) - 1)
                cell_ps = tst_cells[tst_ci].get("paragraph_styles")
            new_cell = _build_cell_from_alias(
                cs_alias, cell_text, col_widths[col_idx],
                paragraph_style_templates, style_alias_map,
                cell_para_styles=cell_ps,
            )
            new_row.append(new_cell)

        parent = root
        for child_idx, child in enumerate(list(parent)):
            if child is target_row:
                parent.insert(child_idx + 1, new_row)
                break
        else:
            for tbl in root.iter():
                children = list(tbl)
                for child_idx, child in enumerate(children):
                    if child is target_row:
                        tbl.insert(child_idx + 1, new_row)
                        break

        return ET.tostring(root, encoding="unicode")

    except ET.ParseError:
        return table_xml


def _table_delete_row(table_xml, r_idx):
    """Delete a specific row from a table."""
    try:
        root = ET.fromstring(table_xml)
        xml_rows = root.findall(".//w:tr", NAMESPACES)
        if r_idx >= len(xml_rows):
            return table_xml

        target_row = xml_rows[r_idx]
        for tbl in root.iter():
            if target_row in list(tbl):
                tbl.remove(target_row)
                break

        return ET.tostring(root, encoding="unicode")

    except ET.ParseError:
        return table_xml


def _table_delete_column(table_xml, c_idx):
    """Delete a specific column from a table."""
    try:
        root = ET.fromstring(table_xml)
        tbl_grid = root.find("w:tblGrid", NAMESPACES)
        xml_rows = root.findall("w:tr", NAMESPACES)

        if not xml_rows:
            return table_xml

        first_row_cells = xml_rows[0].findall("w:tc", NAMESPACES)
        if c_idx >= len(first_row_cells):
            return table_xml

        if len(first_row_cells) <= 1:
            return table_xml

        for tr in xml_rows:
            cells = tr.findall("w:tc", NAMESPACES)
            if c_idx < len(cells):
                tr.remove(cells[c_idx])

        if tbl_grid is not None:
            grid_cols = tbl_grid.findall("w:gridCol", NAMESPACES)
            if c_idx < len(grid_cols):
                tbl_grid.remove(grid_cols[c_idx])

        return ET.tostring(root, encoding="unicode")

    except ET.ParseError:
        return table_xml


def _table_delete_paragraph(table_xml, row_idx, col_idx, para_idx):
    """Delete a specific paragraph from a table cell."""
    try:
        root = ET.fromstring(table_xml)
        xml_rows = root.findall(".//w:tr", NAMESPACES)
        if row_idx >= len(xml_rows):
            return table_xml

        xml_cells = xml_rows[row_idx].findall(".//w:tc", NAMESPACES)
        if col_idx >= len(xml_cells):
            return table_xml

        target_cell = xml_cells[col_idx]
        paragraphs = target_cell.findall("w:p", NAMESPACES)
        if para_idx >= len(paragraphs):
            return table_xml

        if len(paragraphs) <= 1:
            return table_xml

        target_cell.remove(paragraphs[para_idx])
        return ET.tostring(root, encoding="unicode")

    except ET.ParseError:
        return table_xml


def _table_add_column(table_xml, after_col_idx, col_contents,
                      cell_style_aliases, paragraph_style_templates,
                      style_alias_map, tst=None):
    """Add a new column to a table after the specified column."""
    try:
        root = ET.fromstring(table_xml)
        tbl_grid = root.find("w:tblGrid", NAMESPACES)
        xml_rows = root.findall("w:tr", NAMESPACES)

        if not xml_rows:
            return table_xml

        first_row_cells = xml_rows[0].findall("w:tc", NAMESPACES)
        current_col_count = len(first_row_cells)

        if after_col_idx >= current_col_count:
            return table_xml

        total_width = _get_table_total_width(root)
        insert_pos = after_col_idx + 1 if after_col_idx >= 0 else 0

        original_gc_widths = _extract_column_widths(
            tbl_grid, current_col_count, total_width,
        )

        min_col_w = 400
        new_col_max_chars = max(
            (_estimate_text_width(c) for c in col_contents),
            default=2,
        )
        new_col_width = new_col_max_chars * 200 + 200
        new_col_width = max(new_col_width, min_col_w)
        new_col_width = min(new_col_width, total_width * 3 // 10)

        original_total = sum(original_gc_widths) or total_width
        remaining = total_width - new_col_width
        if remaining < min_col_w * current_col_count:
            remaining = total_width * 7 // 10
            new_col_width = total_width - remaining

        scale = remaining / original_total
        adjusted_widths = [
            max(min_col_w, int(w * scale)) for w in original_gc_widths
        ]
        rounding_diff = remaining - sum(adjusted_widths)
        if adjusted_widths:
            max_i = adjusted_widths.index(max(adjusted_widths))
            adjusted_widths[max_i] += rounding_diff

        all_widths = list(adjusted_widths)
        all_widths.insert(insert_pos, new_col_width)

        ns_w = NAMESPACES["w"]
        if tbl_grid is not None:
            new_grid_col = ET.Element(f"{{{ns_w}}}gridCol")
            new_grid_col.set(f"{{{ns_w}}}w", str(new_col_width))
            tbl_grid.insert(insert_pos, new_grid_col)

            for i, gc in enumerate(
                tbl_grid.findall("w:gridCol", NAMESPACES),
            ):
                if i < len(all_widths):
                    gc.set(f"{{{ns_w}}}w", str(all_widths[i]))

        cst = tst.get("cell_style_templates", {}) if tst else {}
        for r_idx, tr in enumerate(xml_rows):
            cs_alias = (
                cell_style_aliases[r_idx]
                if r_idx < len(cell_style_aliases)
                else cell_style_aliases[-1] if cell_style_aliases else ""
            )
            cell_content = (
                col_contents[r_idx] if r_idx < len(col_contents) else ""
            )
            cell_ps = None
            if cst:
                tc_xml_ref = style_alias_map.get(cs_alias, "")
                for _row_key, cells in cst.items():
                    for cell_entry in cells:
                        if cell_entry.get("tc_xml_template") == tc_xml_ref:
                            cell_ps = cell_entry.get("paragraph_styles")
                            break
                    if cell_ps:
                        break
            new_cell = _build_cell_from_alias(
                cs_alias, cell_content, new_col_width,
                paragraph_style_templates, style_alias_map,
                cell_para_styles=cell_ps,
            )

            row_children = list(tr)
            cell_positions = [
                i for i, child in enumerate(row_children)
                if child.tag == f"{{{ns_w}}}tc"
            ]

            if insert_pos < len(cell_positions):
                tr.insert(cell_positions[insert_pos], new_cell)
            else:
                tr.append(new_cell)

        _update_all_cell_widths(root, all_widths)

        return ET.tostring(root, encoding="unicode")

    except ET.ParseError:
        return table_xml


def _table_insert_column_paragraph(table_xml, col_idx, col_contents,
                                   style_key, paragraph_style_templates):
    """Insert a new paragraph into each cell of a specific column."""
    try:
        root = ET.fromstring(table_xml)
        xml_rows = root.findall(".//w:tr", NAMESPACES)

        if not xml_rows:
            return table_xml

        ns_w = NAMESPACES["w"]
        for r_idx, tr in enumerate(xml_rows):
            cells = tr.findall("w:tc", NAMESPACES)
            if col_idx >= len(cells):
                continue
            cell_text = (
                col_contents[r_idx]
                if r_idx < len(col_contents)
                else ""
            )
            if not cell_text:
                continue
            template = paragraph_style_templates.get(style_key)
            if template:
                ppr = template.get("ppr_xml_template", "")
                rst = template.get("run_style_templates", {})
                first_rst = next(iter(rst.values()), None) if rst else None
                if first_rst:
                    rpr_xml = first_rst.get("rpr_xml", "")
                    esc = escape_xml(cell_text)
                    run_xml = _ensure_t_space_preserve(
                        rpr_xml.replace("{{content}}", esc)
                    )
                    para_xml = f'<w:p xmlns:w="{ns_w}">{ppr}{run_xml}</w:p>'
                else:
                    esc = escape_xml(cell_text)
                    para_xml = (
                        f'<w:p xmlns:w="{ns_w}">{ppr}'
                        f'<w:r><w:t xml:space="preserve">{esc}</w:t></w:r>'
                        f'</w:p>'
                    )
            else:
                esc = escape_xml(cell_text)
                para_xml = (
                    f'<w:p xmlns:w="{ns_w}">'
                    f'<w:r><w:t xml:space="preserve">{esc}</w:t></w:r>'
                    f'</w:p>'
                )
            new_para = ET.fromstring(para_xml)
            cells[col_idx].append(new_para)

        return ET.tostring(root, encoding="unicode")

    except ET.ParseError:
        return table_xml


def _get_table_total_width(root):
    """Extract total table width from tblPr/tblW."""
    tbl_pr = root.find("w:tblPr", NAMESPACES)
    if tbl_pr is not None:
        tbl_w = tbl_pr.find("w:tblW", NAMESPACES)
        if tbl_w is not None:
            w_val = tbl_w.get(f"{{{NAMESPACES['w']}}}w")
            if w_val and w_val.isdigit():
                return int(w_val)
    return 9000


def _extract_column_widths(tbl_grid, num_cols, total_width):
    """Extract per-column widths from tblGrid."""
    ns_w = NAMESPACES["w"]
    grid_cols = (
        tbl_grid.findall("w:gridCol", NAMESPACES)
        if tbl_grid is not None
        else []
    )
    default_width = total_width // max(num_cols, 1)

    col_widths = []
    for i in range(num_cols):
        if i < len(grid_cols):
            w_val = grid_cols[i].get(f"{{{ns_w}}}w")
            col_widths.append(
                int(w_val) if w_val and w_val.isdigit() else default_width,
            )
        else:
            col_widths.append(default_width)
    return col_widths


def _estimate_text_width(text):
    """Estimate visual text width in half-width character units."""
    width = 0
    for ch in text:
        cp = ord(ch)
        if (
            0x1100 <= cp <= 0x11FF
            or 0x3000 <= cp <= 0x9FFF
            or 0xAC00 <= cp <= 0xD7AF
            or 0xF900 <= cp <= 0xFAFF
            or 0xFF01 <= cp <= 0xFF60
        ):
            width += 2
        else:
            width += 1
    return max(width, 1)


def _update_all_cell_widths(root, col_widths):
    """Update tcW in all table cells to match col_widths."""
    ns_w = NAMESPACES["w"]
    for tr in root.findall("w:tr", NAMESPACES):
        for c_i, tc in enumerate(tr.findall("w:tc", NAMESPACES)):
            w = (
                col_widths[c_i]
                if c_i < len(col_widths)
                else col_widths[-1] if col_widths else 9000
            )
            tc_pr = tc.find(f"{{{ns_w}}}tcPr")
            if tc_pr is not None:
                tc_w = tc_pr.find(f"{{{ns_w}}}tcW")
                if tc_w is not None:
                    tc_w.set(f"{{{ns_w}}}w", str(w))


def _apply_row_style(row, row_style_alias, style_alias_map):
    """Apply RS alias trPr to a row element."""
    ns_w = NAMESPACES["w"]
    tr_pr_xml = style_alias_map.get(row_style_alias, "")

    if tr_pr_xml and tr_pr_xml.strip():
        old_trpr = row.find(f"{{{ns_w}}}trPr")
        if old_trpr is not None:
            row.remove(old_trpr)
        try:
            new_trpr = ET.fromstring(tr_pr_xml)
            hdr = new_trpr.find(f"{{{ns_w}}}tblHeader")
            if hdr is not None:
                new_trpr.remove(hdr)
            row.insert(0, new_trpr)
        except ET.ParseError:
            pass


def _build_cell_from_alias(cs_alias, text, col_width,
                           paragraph_style_templates, style_alias_map,
                           cell_para_styles=None):
    """Build a complete <w:tc> element from CS alias."""
    ns_w = NAMESPACES["w"]
    tc_xml = style_alias_map.get(cs_alias, "")
    text = _normalize_newlines(text)
    lines = text.split("\n") if "\n" in text else [text]

    cell_ppr = ""
    cell_rst_xml = ""
    if cell_para_styles:
        ps0 = cell_para_styles[0]
        cell_ppr = ps0.get("ppr_xml_template", "")
        rst_dict = ps0.get("run_style_templates", {})
        if rst_dict:
            cell_rst_xml = next(iter(rst_dict.values())).get("rpr_xml", "")

    if tc_xml and "{{content}}" in tc_xml:
        paras_xml = ""
        for line in lines:
            esc = escape_xml(line)
            if cell_rst_xml:
                run_xml = _ensure_t_space_preserve(
                    cell_rst_xml.replace("{{content}}", esc)
                )
                paras_xml += f'<w:p xmlns:w="{ns_w}">{cell_ppr}{run_xml}</w:p>'
            else:
                paras_xml += (
                    f'<w:p xmlns:w="{ns_w}"><w:r>'
                    f'<w:t xml:space="preserve">{esc}</w:t>'
                    f"</w:r></w:p>"
                )
        assembled = tc_xml.replace("{{content}}", paras_xml)
        try:
            cell = ET.fromstring(assembled)
        except ET.ParseError:
            cell = _minimal_cell(ns_w, escape_xml(text))
    else:
        cell = _minimal_cell(ns_w, escape_xml(text))

    tc_pr = cell.find(f"{{{ns_w}}}tcPr")
    if tc_pr is None:
        tc_pr = ET.Element(f"{{{ns_w}}}tcPr")
        cell.insert(0, tc_pr)

    tc_w = tc_pr.find(f"{{{ns_w}}}tcW")
    if tc_w is None:
        tc_w = ET.SubElement(tc_pr, f"{{{ns_w}}}tcW")
    tc_w.set(f"{{{ns_w}}}w", str(col_width))
    tc_w.set(f"{{{ns_w}}}type", "dxa")

    return cell


def _minimal_cell(ns_w, escaped_text):
    """Create a minimal <w:tc> element with one paragraph."""
    cell = ET.Element(f"{{{ns_w}}}tc")
    tc_pr = ET.SubElement(cell, f"{{{ns_w}}}tcPr")
    ET.SubElement(tc_pr, f"{{{ns_w}}}tcW")
    para = ET.SubElement(cell, f"{{{ns_w}}}p")
    run = ET.SubElement(para, f"{{{ns_w}}}r")
    t_elem = ET.SubElement(run, f"{{{ns_w}}}t")
    t_elem.text = escaped_text
    return cell


def _build_table_from_template(template, content,
                               row_style_aliases=None,
                               cell_style_aliases_per_row=None,
                               paragraph_style_templates=None,
                               table_style_key="",
                               style_alias_map=None):
    """Generate table XML from template using RS/CS style aliases."""
    row_style_aliases = row_style_aliases or []
    cell_style_aliases_per_row = cell_style_aliases_per_row or []
    paragraph_style_templates = paragraph_style_templates or {}
    style_alias_map = style_alias_map or {}

    try:
        rows_content = []
        normalized = _normalize_newlines(content.strip())
        for line in normalized.split("\n"):
            cells = [c.strip() for c in line.split("|")]
            rows_content.append(cells)

        if not rows_content:
            return template.get("tbl_xml_template", "")

        num_cols = max(len(row) for row in rows_content)
        ns_w = NAMESPACES["w"]

        root = ET.fromstring(template["tbl_xml_template"])

        for old_tr in root.findall("w:tr", NAMESPACES):
            root.remove(old_tr)

        total_width = _get_table_total_width(root)
        tbl_grid = root.find("w:tblGrid", NAMESPACES)
        if tbl_grid is not None:
            existing_gc = tbl_grid.findall("w:gridCol", NAMESPACES)
            if len(existing_gc) != num_cols:
                for gc in existing_gc:
                    tbl_grid.remove(gc)
                col_w = total_width // num_cols
                remainder = total_width - col_w * num_cols
                for i in range(num_cols):
                    gc = ET.Element(f"{{{ns_w}}}gridCol")
                    w = col_w + (1 if i < remainder else 0)
                    gc.set(f"{{{ns_w}}}w", str(w))
                    tbl_grid.append(gc)
        col_widths = _extract_column_widths(tbl_grid, num_cols, total_width)

        cst = template.get("cell_style_templates", {})
        tst_row_styles = template.get("row_styles", {})
        rs_to_tst_row = {}
        for idx_key in sorted(
            tst_row_styles.keys(),
            key=lambda x: int(x) if x.isdigit() else 999,
        ):
            trpr = tst_row_styles[idx_key].get("tr_pr_xml_template", "")
            for rs_a in set(row_style_aliases):
                if rs_a and rs_a not in rs_to_tst_row:
                    if style_alias_map.get(rs_a, "") == trpr:
                        rs_to_tst_row[rs_a] = idx_key

        for row_idx, row_cells in enumerate(rows_content):
            new_row = ET.Element(f"{{{ns_w}}}tr")

            rs_alias = (
                row_style_aliases[row_idx]
                if row_idx < len(row_style_aliases)
                else (row_style_aliases[-1] if row_style_aliases else "")
            )
            _apply_row_style(new_row, rs_alias, style_alias_map)

            tst_row_key = rs_to_tst_row.get(rs_alias)
            tst_cells = cst.get(tst_row_key, []) if tst_row_key else []

            row_cs = (
                cell_style_aliases_per_row[row_idx]
                if row_idx < len(cell_style_aliases_per_row)
                else (
                    cell_style_aliases_per_row[-1]
                    if cell_style_aliases_per_row
                    else []
                )
            )

            for col_idx in range(num_cols):
                cs_alias = (
                    row_cs[col_idx]
                    if col_idx < len(row_cs)
                    else (row_cs[-1] if row_cs else "")
                )
                cell_text = (
                    row_cells[col_idx]
                    if col_idx < len(row_cells)
                    else ""
                )
                cell_w = (
                    col_widths[col_idx]
                    if col_idx < len(col_widths)
                    else col_widths[-1] if col_widths else 9000
                )
                cell_ps = None
                if tst_cells:
                    tst_ci = min(col_idx, len(tst_cells) - 1)
                    cell_ps = tst_cells[tst_ci].get("paragraph_styles")
                new_cell = _build_cell_from_alias(
                    cs_alias, cell_text, cell_w,
                    paragraph_style_templates, style_alias_map,
                    cell_para_styles=cell_ps,
                )
                new_row.append(new_cell)

            root.append(new_row)

        return ET.tostring(root, encoding="unicode")

    except ET.ParseError:
        return template.get("tbl_xml_template", "")


def assemble_document_xml(blocks, paragraph_style_templates,
                          table_style_templates,
                          style_alias_map):
    """Assemble document.xml body content from modified blocks."""
    body_parts = []
    block_id_to_parts_idx = {}

    for block in blocks:
        if block.get("_deleted"):
            continue

        for insert in block.get("_inserts_before", []):
            content = insert.get("content", "")
            if "\n" in content and "|" not in content:
                lines = [ln.strip() for ln in content.split("\n") if ln.strip()]
                for line in lines:
                    line_insert = dict(insert)
                    line_insert["content"] = line
                    xml = _build_block_xml(
                        line_insert, paragraph_style_templates,
                        table_style_templates, style_alias_map,
                    )
                    if xml:
                        body_parts.append(xml)
            else:
                xml = _build_block_xml(
                    insert, paragraph_style_templates,
                    table_style_templates, style_alias_map,
                )
                if xml:
                    body_parts.append(xml)

        if block.get("_replaced") and "_replacements" in block:
            replacements = block["_replacements"]

            block_type = block.get("type", "")
            is_table = block_type == "tbl"

            if is_table:
                current_xml = block["xml"]

                for r_idx in sorted(
                    block.get("_row_deletions", []), reverse=True,
                ):
                    current_xml = _table_delete_row(current_xml, r_idx)

                for c_idx in sorted(
                    block.get("_col_deletions", []), reverse=True,
                ):
                    current_xml = _table_delete_column(current_xml, c_idx)

                para_deletions = block.get("_para_deletions", [])
                for pd in sorted(
                    para_deletions,
                    key=lambda x: (x["row_idx"], x["col_idx"], x["para_idx"]),
                    reverse=True,
                ):
                    current_xml = _table_delete_paragraph(
                        current_xml,
                        pd["row_idx"], pd["col_idx"], pd["para_idx"],
                    )

                tbl_style_key = block.get("style_key", "")
                tbl_tst = table_style_templates.get(tbl_style_key, {})

                col_inserts_after = block.get("_col_inserts_after", [])
                for col_insert in sorted(
                    col_inserts_after, key=lambda x: x["col_idx"],
                ):
                    c_idx = col_insert["col_idx"]
                    content = col_insert.get("content", "")
                    content = _normalize_newlines(content) if content else ""
                    col_contents = (
                        [c.strip() for c in content.split("\n")]
                        if content else []
                    )
                    cs_aliases = col_insert.get("cell_style_aliases", [])
                    if cs_aliases:
                        current_xml = _table_add_column(
                            current_xml, c_idx, col_contents, cs_aliases,
                            paragraph_style_templates, style_alias_map,
                            tst=tbl_tst,
                        )
                    else:
                        current_xml = _table_insert_column_paragraph(
                            current_xml, c_idx, col_contents,
                            style_key=col_insert.get("style_key", ""),
                            paragraph_style_templates=paragraph_style_templates,
                        )

                col_inserts_before = block.get("_col_inserts_before", [])
                for col_insert in sorted(
                    col_inserts_before, key=lambda x: x["col_idx"],
                ):
                    c_idx = col_insert["col_idx"]
                    content = col_insert.get("content", "")
                    content = _normalize_newlines(content) if content else ""
                    col_contents = (
                        [c.strip() for c in content.split("\n")]
                        if content else []
                    )
                    cs_aliases = col_insert.get("cell_style_aliases", [])
                    if cs_aliases:
                        current_xml = _table_add_column(
                            current_xml, c_idx - 1, col_contents, cs_aliases,
                            paragraph_style_templates, style_alias_map,
                            tst=tbl_tst,
                        )
                    else:
                        current_xml = _table_insert_column_paragraph(
                            current_xml, max(0, c_idx - 1), col_contents,
                            style_key=col_insert.get("style_key", ""),
                            paragraph_style_templates=paragraph_style_templates,
                        )

                for replacement in replacements:
                    original_target_id = replacement.get(
                        "original_target_id", "",
                    )

                    cell_para_match = re.match(
                        r"b\d+:r(\d+)c(\d+)p(\d+)", original_target_id,
                    )
                    row_match = re.match(
                        r"b\d+:r(\d+)$", original_target_id,
                    )

                    if cell_para_match:
                        r_idx = int(cell_para_match.group(1))
                        c_idx = int(cell_para_match.group(2))
                        p_idx = int(cell_para_match.group(3))
                        run_xmls = (
                            replacement.get("run_xmls")
                            if replacement.get("run_xmls")
                            else None
                        )
                        current_xml = _table_replace_paragraph(
                            current_xml, r_idx, c_idx, p_idx,
                            replacement["content"],
                            run_xmls=run_xmls,
                        )
                    elif row_match:
                        r_idx = int(row_match.group(1))
                        content = replacement["content"]
                        row_contents = (
                            [c.strip() for c in content.split("|")]
                            if content else []
                        )
                        try:
                            root = ET.fromstring(current_xml)
                            xml_rows = root.findall(".//w:tr", NAMESPACES)
                            if r_idx < len(xml_rows):
                                target_row = xml_rows[r_idx]
                                cells = target_row.findall("w:tc", NAMESPACES)
                                for ci, cell in enumerate(cells):
                                    cell_text = (
                                        row_contents[ci]
                                        if ci < len(row_contents) else ""
                                    )
                                    first_text_set = False
                                    for t in cell.findall(".//w:t", NAMESPACES):
                                        if not first_text_set:
                                            t.text = cell_text
                                            first_text_set = True
                                        else:
                                            t.text = ""
                            current_xml = ET.tostring(root, encoding="unicode")
                        except ET.ParseError:
                            pass

                body_parts.append(current_xml)

            else:
                block_id_to_parts_idx[block["id"]] = len(body_parts)
                replacement = replacements[0]

                style_key = replacement.get("style_key", "")
                original_style_key = block.get("style_key", "")

                effective_style_key = original_style_key

                if effective_style_key not in paragraph_style_templates:
                    body_parts.append(block["xml"])
                    continue

                template = paragraph_style_templates[effective_style_key]
                run_xmls = replacement.get("run_xmls", [])
                content = replacement["content"]
                ns = NAMESPACES["w"]

                original_xml = block.get("xml", "")
                if _has_non_text_content(original_xml):
                    body_parts.append(
                        _replace_text_preserving_structure(
                            original_xml, content,
                        ),
                    )
                else:
                    bk_starts, bk_ends = _extract_bookmarks_xml(
                        original_xml,
                    )

                    if run_xmls:
                        ppr = (
                            template.get("ppr_xml_template")
                            or f'<w:pPr xmlns:w="{ns}"/>'
                        )
                        bk_s = "".join(bk_starts)
                        bk_e = "".join(bk_ends)
                        xml = (
                            f'<w:p xmlns:w="{ns}">'
                            f'{ppr}{bk_s}{"".join(run_xmls)}{bk_e}'
                            f'</w:p>'
                        )
                        body_parts.append(xml)
                    else:
                        start_idx = len(body_parts)
                        _append_paragraphs_from_content(
                            body_parts, template, content, ns,
                        )
                        if ((bk_starts or bk_ends)
                                and len(body_parts) > start_idx):
                            body_parts[start_idx] = (
                                _inject_bookmarks_into_para(
                                    body_parts[start_idx],
                                    bk_starts, bk_ends,
                                )
                            )
        else:
            block_id_to_parts_idx[block["id"]] = len(body_parts)
            body_parts.append(block["xml"])

        for insert in block.get("_inserts_after", []):
            original_target_id = insert.get("original_target_id", "")
            insert_edit_unit = insert.get("edit_unit")

            row_add_match = re.match(r"b\d+:r(\d+)$", original_target_id)
            cell_para_match = re.match(
                r"b\d+:r(\d+)c(\d+)p(\d+)", original_target_id,
            )

            is_row_add = (
                insert_edit_unit == "row"
                if insert_edit_unit
                else (row_add_match is not None)
            )

            if is_row_add:
                if row_add_match:
                    r_idx = int(row_add_match.group(1))
                else:
                    r_match = re.search(r"r(\d+)", original_target_id)
                    if not r_match:
                        continue
                    r_idx = int(r_match.group(1))

                content = insert.get("content", "")
                row_contents = (
                    [c.strip() for c in content.split("|")]
                    if content else []
                )

                cell_style_aliases_nested = insert.get(
                    "cell_style_aliases",
                ) or []
                cell_style_aliases = (
                    cell_style_aliases_nested[0]
                    if cell_style_aliases_nested
                    else []
                )

                if cell_style_aliases and body_parts:
                    row_style_aliases = insert.get(
                        "row_style_aliases",
                    ) or []
                    row_style_alias = (
                        row_style_aliases[0]
                        if row_style_aliases else ""
                    )
                    modified_xml = _table_add_row(
                        body_parts[-1], r_idx, row_contents,
                        row_style_alias=row_style_alias,
                        cell_style_aliases=cell_style_aliases,
                        paragraph_style_templates=paragraph_style_templates,
                        table_style_templates=table_style_templates,
                        table_style_key=block.get("style_key", ""),
                        style_alias_map=style_alias_map,
                    )
                    body_parts[-1] = modified_xml

            elif cell_para_match:
                pass
            else:
                content = insert.get("content", "")
                if "\n" in content and "|" not in content:
                    lines = [
                        ln.strip() for ln in content.split("\n") if ln.strip()
                    ]
                    for line in lines:
                        line_insert = dict(insert)
                        line_insert["content"] = line
                        xml = _build_block_xml(
                            line_insert, paragraph_style_templates,
                            table_style_templates, style_alias_map,
                        )
                        if xml:
                            body_parts.append(xml)
                else:
                    xml = _build_block_xml(
                        insert, paragraph_style_templates,
                        table_style_templates, style_alias_map,
                    )
                    if xml:
                        body_parts.append(xml)

    return "".join(body_parts)


def wrap_document_body(body_content, original_document_xml):
    """Replace body content in document.xml preserving namespace declarations."""
    body_pattern = re.compile(
        r"(<w:body[^>]*>)[\s\S]*?(</w:body>)",
        re.IGNORECASE,
    )

    match = body_pattern.search(original_document_xml)
    if match:
        prefix = original_document_xml[:match.start(1)]
        body_open = match.group(1)
        body_close = match.group(2)
        suffix = original_document_xml[match.end(2):]
        return f"{prefix}{body_open}{body_content}{body_close}{suffix}"

    return original_document_xml


def save_document_xml(document_xml, extracted_path):
    """Write assembled document.xml to the extracted directory."""
    output_path = os.path.join(extracted_path, "word", "document.xml")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(document_xml)
    return output_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 apply_edits.py <work_dir>")
        sys.exit(1)

    work_dir = sys.argv[1]

    analysis_path = os.path.join(work_dir, "analysis.json")
    edits_path = os.path.join(work_dir, "edits.json")

    if not os.path.isfile(analysis_path):
        print(f"Error: {analysis_path} not found")
        sys.exit(1)
    if not os.path.isfile(edits_path):
        print(f"Error: {edits_path} not found")
        sys.exit(1)

    with open(analysis_path, "r", encoding="utf-8") as f:
        analysis = json.load(f)

    with open(edits_path, "r", encoding="utf-8") as f:
        edits_data = json.load(f)

    edits = edits_data.get("edits", [])
    if not edits:
        print("No edits to apply")
        sys.exit(0)

    extracted_path = analysis.get("extracted_path", "")
    if not extracted_path:
        extracted_path = os.path.join(work_dir, "extracted")

    document_xml_path = os.path.join(extracted_path, "word", "document.xml")
    if not os.path.isfile(document_xml_path):
        print(f"Error: {document_xml_path} not found")
        sys.exit(1)

    with open(document_xml_path, "r", encoding="utf-8") as f:
        original_document_xml = f.read()

    paragraph_style_templates = analysis.get("paragraph_style_templates", {})
    table_style_templates = analysis.get("table_style_templates", {})
    style_alias_map = analysis.get("style_alias_map", {})
    blocks = analysis.get("blocks", [])

    print(f"Phase A: Mapping {len(edits)} edits...")
    new_blocks = generate_new_blocks(edits, analysis)
    print(f"  Generated {len(new_blocks)} new blocks")

    print("Phase B: Assembling document XML...")

    marked_blocks = apply_mapping_to_blocks(blocks, new_blocks)

    body_content = assemble_document_xml(
        marked_blocks,
        paragraph_style_templates,
        table_style_templates,
        style_alias_map,
    )

    document_xml = wrap_document_body(body_content, original_document_xml)

    output_path = save_document_xml(document_xml, extracted_path)

    print(f"Applied {len(edits)} edits successfully")
    print(f"Output: {output_path}")


if __name__ == "__main__":
    main()
