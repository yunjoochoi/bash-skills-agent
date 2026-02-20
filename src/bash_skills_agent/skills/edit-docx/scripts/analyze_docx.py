#!/usr/bin/env python3
"""Phase 1: Standalone DOCX Analyzer — python3 analyze_docx.py <docx_path> <work_dir>

Outputs text_merge to stdout and analysis.json to <work_dir>.
"""

import copy
import json
import os
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W_NS = NAMESPACES["w"]

for _prefix, _uri in NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)

_TCPR_LAYOUT_TAGS = frozenset({
    f"{{{W_NS}}}tcW",
    f"{{{W_NS}}}gridSpan",
    f"{{{W_NS}}}vMerge",
    f"{{{W_NS}}}hMerge",
})

_BOOLEAN_TAGS = {"b", "bCs", "i", "iCs", "strike", "dstrike", "caps", "smallCaps"}


class AnalyzerState:
    """Mutable state for table style deduplication caches."""

    __slots__ = (
        "table_style_templates",
        "_table_counter",
        "_row_style_cache",
        "_cell_style_cache",
        "_rs_counter",
        "_cs_counter",
        "_row_style_map",
        "_cell_style_map",
    )

    def __init__(self):
        self.table_style_templates = {}
        self._table_counter = 0
        self._row_style_cache = {}
        self._cell_style_cache = {}
        self._rs_counter = 0
        self._cs_counter = 0
        self._row_style_map = {}
        self._cell_style_map = {}


def _extract_paragraph_text(p_element):
    """Return concatenated text from all w:t elements in a paragraph."""
    texts = []
    for t in p_element.findall(".//w:t", NAMESPACES):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _extract_table_text(tbl_element):
    """Extract table text with [rNcN] coordinates, rows separated by |."""
    rows = []
    for r_idx, tr in enumerate(tbl_element.findall(".//w:tr", NAMESPACES)):
        cells = []
        for c_idx, tc in enumerate(tr.findall(".//w:tc", NAMESPACES)):
            paragraphs = tc.findall("w:p", NAMESPACES)

            if len(paragraphs) <= 1:
                cell_text = ""
                for t in tc.findall(".//w:t", NAMESPACES):
                    if t.text:
                        cell_text += t.text
                cells.append(f"[r{r_idx}c{c_idx}] {cell_text}")
            else:
                cell_lines = [f"[r{r_idx}c{c_idx}]"]
                for p_idx, p in enumerate(paragraphs):
                    p_text = ""
                    for t in p.findall(".//w:t", NAMESPACES):
                        if t.text:
                            p_text += t.text
                    if p_text.strip():
                        cell_lines.append(
                            f"    [r{r_idx}c{c_idx}p{p_idx}] {p_text}"
                        )
                cells.append("\n".join(cell_lines))
        rows.append(" | ".join(cells))
    return "\n".join(rows)


def _extract_sdt_text(sdt_element):
    """Extract text from SDT element, showing per-paragraph indices for multi-paragraph blocks."""
    alias = ""
    sdt_pr = sdt_element.find("w:sdtPr", NAMESPACES)
    if sdt_pr is not None:
        alias_elem = sdt_pr.find("w:alias", NAMESPACES)
        if alias_elem is not None:
            alias = alias_elem.get(f"{{{NAMESPACES['w']}}}val", "")

    sdt_content = sdt_element.find("w:sdtContent", NAMESPACES)
    if sdt_content is None:
        texts = []
        for t in sdt_element.findall(".//w:t", NAMESPACES):
            if t.text:
                texts.append(t.text)
        content_preview = " ".join(texts)[:100]
        if alias:
            return f"[{alias}] {content_preview}"
        return f"[SDT] {content_preview}"

    paragraphs = sdt_content.findall(".//w:p", NAMESPACES)
    if len(paragraphs) <= 1:
        texts = []
        for t in sdt_element.findall(".//w:t", NAMESPACES):
            if t.text:
                texts.append(t.text)
        content = " ".join(texts)[:100]
        if alias:
            return f"[{alias}] {content}"
        return f"[SDT] {content}"

    lines = [f"[{alias or 'SDT'}]"]
    for p_idx, p in enumerate(paragraphs):
        p_text = ""
        for t in p.findall(".//w:t", NAMESPACES):
            if t.text:
                p_text += t.text
        if p_text.strip():
            lines.append(f"  [p{p_idx}] {p_text}")

    return "\n".join(lines)


def _is_toc_sdt(sdt_element):
    """Detect if SDT is a TOC (docPartGallery / alias / PAGEREF)."""
    w_ns = NAMESPACES["w"]
    sdt_pr = sdt_element.find("w:sdtPr", NAMESPACES)
    if sdt_pr is not None:
        doc_part = sdt_pr.find(".//w:docPartGallery", NAMESPACES)
        if doc_part is not None:
            val = doc_part.get(f"{{{w_ns}}}val", "")
            if "Table of Contents" in val:
                return True

        alias_elem = sdt_pr.find("w:alias", NAMESPACES)
        if alias_elem is not None:
            val = alias_elem.get(f"{{{w_ns}}}val", "")
            if "TOC" in val.upper():
                return True

    instr_texts = sdt_element.findall(".//w:instrText", NAMESPACES)
    for instr in instr_texts:
        if instr.text and "PAGEREF" in instr.text:
            return True

    return False


def _element_to_key_part(elem):
    """Convert XML element to deterministic key part string (recursive)."""
    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
    val_attr = elem.get(f"{{{W_NS}}}val")
    if (
        tag in _BOOLEAN_TAGS
        and val_attr in ("0", "false")
        and len(elem.attrib) == 1
        and len(elem) == 0
    ):
        return ""

    attr_parts = []
    for attr_name in sorted(elem.attrib.keys()):
        local_name = attr_name.split("}")[-1] if "}" in attr_name else attr_name
        attr_val = elem.attrib[attr_name]
        if local_name == "val":
            attr_parts.append(attr_val)
        else:
            attr_parts.append(f"{local_name}{attr_val}")

    child_parts = []
    for child in sorted(elem, key=lambda x: x.tag):
        child_part = _element_to_key_part(child)
        if child_part:
            child_parts.append(child_part)

    all_sub = attr_parts + child_parts
    if all_sub:
        return f"{tag}-{'_'.join(all_sub)}"
    return tag


def _extract_p_style(p_element):
    """Return w:pStyle value or 'Normal'."""
    p_pr = p_element.find("w:pPr", NAMESPACES)
    if p_pr is not None:
        p_style_elem = p_pr.find("w:pStyle", NAMESPACES)
        if p_style_elem is not None:
            return p_style_elem.get(f"{{{W_NS}}}val", "Normal")
    return "Normal"


def _build_ppr_key(p_element):
    """Build deterministic key from pPr children (excluding pStyle, rPr)."""
    p_pr = p_element.find("w:pPr", NAMESPACES)
    if p_pr is None:
        return ""

    skip_tags = {f"{{{W_NS}}}pStyle", f"{{{W_NS}}}rPr"}
    parts = []

    for child in sorted(p_pr, key=lambda x: x.tag):
        if child.tag in skip_tags:
            continue
        part = _element_to_key_part(child)
        if part:
            parts.append(part)

    return "_".join(parts)


def _generate_style_key(p_element):
    """Return '{pStyle}_{ppr_key}' style key (no rPr)."""
    p_style = _extract_p_style(p_element)
    ppr_key = _build_ppr_key(p_element)
    return f"{p_style}_{ppr_key}" if ppr_key else p_style


def _build_rpr_key(run):
    """Build deterministic key from rPr children, or 'default' if none."""
    r_pr = run.find("w:rPr", NAMESPACES)
    if r_pr is None:
        return "default"

    parts = []
    for child in sorted(r_pr, key=lambda x: x.tag):
        part = _element_to_key_part(child)
        if part:
            parts.append(part)

    return "_".join(parts) if parts else "default"


def _run_has_text(run):
    """True if run has at least one non-empty w:t."""
    for t in run.findall("w:t", NAMESPACES):
        if t.text and t.text.strip():
            return True
    return False


def _describe_run_styles(run):
    """Human-readable run style description (e.g. 'bold, size:28')."""
    r_pr = run.find("w:rPr", NAMESPACES)
    if r_pr is None:
        return "default"

    descriptions = []
    for child in sorted(r_pr, key=lambda x: x.tag):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "b":
            val = child.get(f"{{{W_NS}}}val")
            if val not in ("0", "false"):
                descriptions.append("bold")
        elif tag == "i":
            val = child.get(f"{{{W_NS}}}val")
            if val not in ("0", "false"):
                descriptions.append("italic")
        elif tag == "u":
            val = child.get(f"{{{W_NS}}}val", "single")
            descriptions.append(f"underline:{val}")
        elif tag == "sz":
            val = child.get(f"{{{W_NS}}}val", "")
            descriptions.append(f"size:{val}")
        elif tag == "szCs":
            pass  # Skip szCs (complex script size, redundant with sz)
        elif tag == "color":
            val = child.get(f"{{{W_NS}}}val", "")
            descriptions.append(f"color:{val}")
        elif tag == "rFonts":
            ascii_font = child.get(f"{{{W_NS}}}ascii", "")
            ea_font = child.get(f"{{{W_NS}}}eastAsia", "")
            if ascii_font:
                descriptions.append(f"font:{ascii_font}")
            elif ea_font:
                descriptions.append(f"font:{ea_font}")
        elif tag == "highlight":
            val = child.get(f"{{{W_NS}}}val", "")
            descriptions.append(f"highlight:{val}")
        elif tag == "rStyle":
            val = child.get(f"{{{W_NS}}}val", "")
            descriptions.append(f"rStyle:{val}")
        elif tag == "lang":
            pass  # Skip lang
        else:
            # Generic: include tag name
            val = child.get(f"{{{W_NS}}}val", "")
            if val:
                descriptions.append(f"{tag}:{val}")
            else:
                descriptions.append(tag)

    return ", ".join(descriptions) if descriptions else "default"


def _build_run_style_template(run):
    """Deep-copy run, replace w:t with {{content}}, return template dict."""
    rpr_key = _build_rpr_key(run)
    template_run = copy.deepcopy(run)
    for t in template_run.findall("w:t", NAMESPACES):
        t.text = "{{content}}"
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    rpr_xml = ET.tostring(template_run, encoding="unicode")
    display_desc = _describe_run_styles(run)

    return {
        "rpr_key": rpr_key,
        "rpr_xml": rpr_xml,
        "display_description": display_desc,
    }


def _build_ppr_xml(p_element):
    """Return serialized w:pPr XML string, or '' if no pPr."""
    p_pr = p_element.find("w:pPr", NAMESPACES)
    if p_pr is None:
        return ""
    return ET.tostring(p_pr, encoding="unicode")


def _extract_table_xml_template(tbl_element):
    """Return table shell XML (tblPr + tblGrid, no rows)."""
    template = copy.deepcopy(tbl_element)
    for tr in template.findall("w:tr", NAMESPACES):
        template.remove(tr)
    return ET.tostring(template, encoding="unicode")


def _extract_row_xml_template(tr_element):
    """Return dict with tr_pr_xml_template from row element."""
    tr_pr = tr_element.find("w:trPr", NAMESPACES)
    tr_pr_xml = ET.tostring(tr_pr, encoding="unicode") if tr_pr is not None else ""
    return {"tr_pr_xml_template": tr_pr_xml}


def _extract_cell_xml_template(tc_element):
    """Return cell XML shell: visual tcPr + {{content}} (layout tags stripped)."""
    template = copy.deepcopy(tc_element)

    for p in template.findall("w:p", NAMESPACES):
        template.remove(p)

    tc_pr = template.find("w:tcPr", NAMESPACES)
    if tc_pr is not None:
        for child in list(tc_pr):
            if child.tag in _TCPR_LAYOUT_TAGS:
                tc_pr.remove(child)
        tc_pr.tail = "{{content}}"
    else:
        template.text = "{{content}}"

    return ET.tostring(template, encoding="unicode")


def _extract_cell_style(tc_element):
    """Extract cell style: paragraph_styles + tc_xml_template."""
    tc_xml_template = _extract_cell_xml_template(tc_element)
    paragraph_styles = []

    for p in tc_element.findall("w:p", NAMESPACES):
        p_key = _generate_style_key(p)
        ppr_xml = _build_ppr_xml(p)

        run_templates = {}
        seen = set()
        for run in p.findall("w:r", NAMESPACES):
            if not _run_has_text(run):
                continue
            rst = _build_run_style_template(run)
            if rst["rpr_key"] not in seen:
                run_templates[f"RS{len(run_templates)}"] = rst
                seen.add(rst["rpr_key"])

        paragraph_styles.append({
            "paragraph_style_key": p_key,
            "ppr_xml_template": ppr_xml,
            "run_style_templates": run_templates,
        })

    return {
        "paragraph_styles": paragraph_styles,
        "tc_xml_template": tc_xml_template,
    }


def _extract_table_hierarchy(tbl_element, state):
    """Extract hierarchical table style (T:RS:CS) with deduplication."""
    state._table_counter += 1
    table_key = f"T{state._table_counter}"
    tbl_xml_template = _extract_table_xml_template(tbl_element)

    row_styles = {}
    cell_templates = {}
    row_style_aliases = []
    cell_style_map = {}

    for r_idx, tr in enumerate(tbl_element.findall("w:tr", NAMESPACES)):
        row_tmpl = _extract_row_xml_template(tr)
        fp = row_tmpl["tr_pr_xml_template"]

        if fp not in state._row_style_cache:
            rs_alias = f"RS{state._rs_counter}"
            state._row_style_cache[fp] = rs_alias
            state._row_style_map[rs_alias] = fp
            state._rs_counter += 1

        row_style_aliases.append(state._row_style_cache[fp])
        row_styles[r_idx] = row_tmpl

        row_cells = []
        for c_idx, tc in enumerate(tr.findall("w:tc", NAMESPACES)):
            cell_tmpl = _extract_cell_style(tc)
            fp_c = cell_tmpl["tc_xml_template"]

            if fp_c not in state._cell_style_cache:
                cs_alias = f"CS{state._cs_counter}"
                state._cell_style_cache[fp_c] = cs_alias
                state._cell_style_map[cs_alias] = fp_c
                state._cs_counter += 1

            cell_style_map[f"r{r_idx}c{c_idx}"] = state._cell_style_cache[fp_c]
            row_cells.append(cell_tmpl)

        cell_templates[r_idx] = row_cells

    table_template = {
        "table_style_key": table_key,
        "tbl_xml_template": tbl_xml_template,
        "row_styles": {str(k): v for k, v in row_styles.items()},
        "cell_style_templates": {str(k): v for k, v in cell_templates.items()},
    }

    return table_template, row_style_aliases, cell_style_map



def _build_style_lookup(extracted_path):
    """Parse styles.xml → {styleId: {name, outline_lvl, based_on}}."""
    styles_xml_path = os.path.join(extracted_path, "word", "styles.xml")
    if not os.path.exists(styles_xml_path):
        return {}

    lookup = {}
    try:
        tree = ET.parse(styles_xml_path)
        root = tree.getroot()
        for style in root.findall(".//w:style", NAMESPACES):
            style_id = style.get(f"{{{W_NS}}}styleId", "")
            if not style_id:
                continue

            name_elem = style.find("w:name", NAMESPACES)
            name = name_elem.get(f"{{{W_NS}}}val", "") if name_elem is not None else ""

            outline_lvl = None
            outline_elem = style.find(".//w:outlineLvl", NAMESPACES)
            if outline_elem is not None:
                try:
                    outline_lvl = int(outline_elem.get(f"{{{W_NS}}}val", "0"))
                except ValueError:
                    outline_lvl = None

            based_on_elem = style.find("w:basedOn", NAMESPACES)
            based_on = (
                based_on_elem.get(f"{{{W_NS}}}val", "")
                if based_on_elem is not None else None
            )

            lookup[style_id] = {
                "name": name,
                "outline_lvl": outline_lvl,
                "based_on": based_on,
            }
    except ET.ParseError:
        pass

    return lookup


def _infer_semantic_info(block, styles_xml_path):
    """Infer semantic tag (H1..H9, BODY, LIST, TITLE, SUBTITLE, TOC)."""
    try:
        p_element = ET.fromstring(block["xml"])
    except ET.ParseError:
        return "BODY"

    outline_elem = p_element.find(".//w:outlineLvl", NAMESPACES)
    if outline_elem is not None:
        try:
            lvl = int(outline_elem.get(f"{{{W_NS}}}val", "0"))
            return f"H{lvl + 1}"
        except ValueError:
            pass

    style_id = _extract_p_style(p_element)
    style_lookup = _build_style_lookup(styles_xml_path)
    if style_id in style_lookup:
        info = style_lookup[style_id]
        if info.get("outline_lvl") is not None:
            return f"H{info['outline_lvl'] + 1}"

    name_lower = style_id.lower()
    name_from_lookup = ""
    if style_id in style_lookup:
        name_from_lookup = style_lookup[style_id].get("name", "").lower()

    check_names = [name_lower, name_from_lookup]
    for name in check_names:
        if not name:
            continue
        if "heading" in name:
            for i in range(1, 10):
                if str(i) in name:
                    return f"H{i}"
            return "H1"
        if name == "title":
            return "TITLE"
        if "subtitle" in name:
            return "SUBTITLE"
        if "list" in name or "bullet" in name or "number" in name:
            return "LIST"
        if "toc" in name:
            return "TOC"

    return "BODY"


def _parse_numbering_xml(extracted_path):
    """Parse numbering.xml → {abstract_nums, num_map}, or None."""
    numbering_path = os.path.join(extracted_path, "word", "numbering.xml")
    if not os.path.exists(numbering_path):
        return None

    try:
        tree = ET.parse(numbering_path)
    except ET.ParseError:
        return None

    root = tree.getroot()
    w = W_NS

    abstract_nums = {}
    for an in root.findall(f".//{{{w}}}abstractNum"):
        aid = an.get(f"{{{w}}}abstractNumId")
        if not aid:
            continue
        levels = {}
        for lvl in an.findall(f"{{{w}}}lvl"):
            ilvl_str = lvl.get(f"{{{w}}}ilvl")
            if ilvl_str is None:
                continue
            ilvl = int(ilvl_str)

            start_elem = lvl.find(f"{{{w}}}start")
            start = (
                int(start_elem.get(f"{{{w}}}val", "1"))
                if start_elem is not None else 1
            )

            fmt_elem = lvl.find(f"{{{w}}}numFmt")
            num_fmt = (
                fmt_elem.get(f"{{{w}}}val", "decimal")
                if fmt_elem is not None else "decimal"
            )

            text_elem = lvl.find(f"{{{w}}}lvlText")
            lvl_text = (
                text_elem.get(f"{{{w}}}val", "")
                if text_elem is not None else ""
            )

            levels[ilvl] = {
                "start": start,
                "numFmt": num_fmt,
                "lvlText": lvl_text,
            }

        abstract_nums[aid] = levels

    num_map = {}
    for num in root.findall(f".//{{{w}}}num"):
        nid = num.get(f"{{{w}}}numId")
        if not nid:
            continue
        anid_elem = num.find(f"{{{w}}}abstractNumId")
        anid = (
            anid_elem.get(f"{{{w}}}val")
            if anid_elem is not None else None
        )

        overrides = {}
        for ov in num.findall(f"{{{w}}}lvlOverride"):
            ov_ilvl = ov.get(f"{{{w}}}ilvl")
            if ov_ilvl is None:
                continue
            start_ov = ov.find(f"{{{w}}}startOverride")
            if start_ov is not None:
                overrides[int(ov_ilvl)] = int(
                    start_ov.get(f"{{{w}}}val", "1")
                )

        num_map[nid] = {"abstractNumId": anid, "overrides": overrides}

    return {"abstract_nums": abstract_nums, "num_map": num_map}


def _format_number(value, num_fmt):
    """Format counter value per Word numFmt (decimal, lowerLetter, lowerRoman, etc.)."""
    if num_fmt == "decimal":
        return str(value)
    if num_fmt == "lowerLetter":
        return chr(ord("a") + (value - 1) % 26) if value >= 1 else str(value)
    if num_fmt == "upperLetter":
        return chr(ord("A") + (value - 1) % 26) if value >= 1 else str(value)
    if num_fmt in ("lowerRoman", "upperRoman"):
        pairs = [
            (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
            (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
            (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
        ]
        result = ""
        remaining = value
        for threshold, numeral in pairs:
            while remaining >= threshold:
                result += numeral
                remaining -= threshold
        return result.upper() if num_fmt == "upperRoman" else result
    if num_fmt == "bullet":
        return ""
    return str(value)


def _compute_effective_numbering(blocks, numbering_defs):
    """Compute numbering prefix per block → {block_id: prefix_string}."""
    if numbering_defs is None:
        return {}

    abstract_nums = numbering_defs["abstract_nums"]
    num_map = numbering_defs["num_map"]

    counters = {}
    used = {}
    last_ilvl = {}
    result = {}

    for block in blocks:
        if block["type"] != "p":
            continue

        style_key = block.get("style_key", "")
        if "numPr" not in style_key:
            continue

        try:
            p_elem = ET.fromstring(block["xml"])
        except ET.ParseError:
            continue

        p_pr = p_elem.find("w:pPr", NAMESPACES)
        if p_pr is None:
            continue

        num_pr = p_pr.find("w:numPr", NAMESPACES)
        if num_pr is None:
            continue

        ilvl_elem = num_pr.find("w:ilvl", NAMESPACES)
        numid_elem = num_pr.find("w:numId", NAMESPACES)
        if ilvl_elem is None or numid_elem is None:
            continue

        ilvl = int(ilvl_elem.get(f"{{{W_NS}}}val", "0"))
        num_id = numid_elem.get(f"{{{W_NS}}}val", "0")

        if num_id == "0":
            continue

        if num_id not in num_map:
            continue
        abstract_num_id = num_map[num_id]["abstractNumId"]
        if abstract_num_id is None or abstract_num_id not in abstract_nums:
            continue

        levels = abstract_nums[abstract_num_id]
        overrides = num_map[num_id].get("overrides", {})

        if num_id not in counters:
            counters[num_id] = {}
            used[num_id] = {}
            for lvl_idx, lvl_def in levels.items():
                start = overrides.get(lvl_idx, lvl_def["start"])
                counters[num_id][lvl_idx] = start - 1
                used[num_id][lvl_idx] = False
            last_ilvl[num_id] = -1

        prev_ilvl = last_ilvl.get(num_id, -1)
        if ilvl <= prev_ilvl:
            for reset_lvl in range(ilvl + 1, 9):
                if reset_lvl in levels:
                    start = overrides.get(reset_lvl, levels[reset_lvl]["start"])
                    counters[num_id][reset_lvl] = start - 1
                    used[num_id][reset_lvl] = False

        if ilvl not in counters[num_id]:
            counters[num_id][ilvl] = 0
        counters[num_id][ilvl] += 1
        used[num_id][ilvl] = True
        last_ilvl[num_id] = ilvl

        if ilvl not in levels:
            continue

        lvl_def = levels[ilvl]
        num_fmt = lvl_def["numFmt"]
        lvl_text = lvl_def["lvlText"]

        if num_fmt == "bullet":
            result[block["id"]] = lvl_text if lvl_text else ""
        else:
            prefix = lvl_text
            for ref_lvl in range(9):
                placeholder = f"%{ref_lvl + 1}"
                if placeholder in prefix:
                    if used[num_id].get(ref_lvl, False):
                        counter_val = counters[num_id].get(ref_lvl, 0)
                    else:
                        ref_start = overrides.get(
                            ref_lvl,
                            levels.get(ref_lvl, {}).get("start", 1),
                        )
                        counter_val = ref_start
                    ref_fmt = levels.get(ref_lvl, {}).get("numFmt", "decimal")
                    formatted = _format_number(counter_val, ref_fmt)
                    prefix = prefix.replace(placeholder, formatted)
            result[block["id"]] = prefix

    return result


def extract_docx_xml(docx_path, output_dir):
    """Extract .xml and .rels files from DOCX to output_dir."""
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")

    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    with zipfile.ZipFile(docx_path, "r") as zip_ref:
        for filename in zip_ref.namelist():
            if filename.endswith(".xml") or filename.endswith(".rels"):
                target_path = os.path.join(output_dir, filename)
                os.makedirs(os.path.dirname(target_path), exist_ok=True)
                with open(target_path, "wb") as f:
                    f.write(zip_ref.read(filename))

    return output_dir


def parse_document_blocks(extracted_path, state):
    """Parse document.xml into content blocks (p, tbl, sdt)."""
    document_xml_path = os.path.join(extracted_path, "word", "document.xml")
    if not os.path.exists(document_xml_path):
        raise FileNotFoundError(f"document.xml not found: {document_xml_path}")

    tree = ET.parse(document_xml_path)
    root = tree.getroot()

    body = root.find(".//w:body", NAMESPACES)
    if body is None:
        return []

    blocks = []
    block_id = 0

    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            text = _extract_paragraph_text(child)
            xml_str = ET.tostring(child, encoding="unicode")
            style_key = _generate_style_key(child)

            blocks.append({
                "id": f"b{block_id}",
                "type": "p",
                "text": text,
                "xml": xml_str,
                "style_key": style_key,
                "semantic_tag": "BODY",
                "row_style_aliases": None,
                "cell_style_map": None,
            })
            block_id += 1

        elif tag == "tbl":
            text = _extract_table_text(child)
            xml_str = ET.tostring(child, encoding="unicode")

            table_tmpl, row_aliases, cell_alias_map = (
                _extract_table_hierarchy(child, state)
            )
            state.table_style_templates[table_tmpl["table_style_key"]] = table_tmpl

            blocks.append({
                "id": f"b{block_id}",
                "type": "tbl",
                "text": text,
                "xml": xml_str,
                "style_key": table_tmpl["table_style_key"],
                "semantic_tag": "TBL",
                "row_style_aliases": row_aliases,
                "cell_style_map": cell_alias_map,
            })
            block_id += 1

        elif tag == "sectPr":
            continue

        elif tag == "sdt":
            is_toc = _is_toc_sdt(child)
            text = _extract_sdt_text(child)
            xml_str = ET.tostring(child, encoding="unicode")

            semantic_tag = "TOC" if is_toc else "SDT"
            style_key = "toc_default" if is_toc else "sdt_default"

            blocks.append({
                "id": f"b{block_id}",
                "type": "sdt",
                "text": text,
                "xml": xml_str,
                "style_key": style_key,
                "semantic_tag": semantic_tag,
                "row_style_aliases": None,
                "cell_style_map": None,
            })
            block_id += 1

    for block in blocks:
        if block["type"] == "p":
            block["semantic_tag"] = _infer_semantic_info(block, extracted_path)

    return blocks


def build_paragraph_style_templates(parsed_result):
    """Build deduplicated paragraph style templates from parsed blocks."""
    templates = {}

    for block in parsed_result:
        if block["type"] != "p":
            continue

        try:
            p_element = ET.fromstring(block["xml"])

            if block["style_key"] not in templates:
                ppr_xml = _build_ppr_xml(p_element)
                templates[block["style_key"]] = {
                    "paragraph_style_key": block["style_key"],
                    "ppr_xml_template": ppr_xml,
                    "run_style_templates": {},
                }

            template = templates[block["style_key"]]
            existing_rpr_keys = {
                r["rpr_key"]
                for r in template["run_style_templates"].values()
            }
            for run in p_element.findall("w:r", NAMESPACES):
                if not _run_has_text(run):
                    continue
                rst = _build_run_style_template(run)
                if rst["rpr_key"] not in existing_rpr_keys:
                    alias = f"RS{len(template['run_style_templates'])}"
                    template["run_style_templates"][alias] = rst
                    existing_rpr_keys.add(rst["rpr_key"])

        except ET.ParseError:
            pass

    return templates


def _generate_hierarchical_table_text(
    block, style_key_to_alias, style_alias_map, alias_counter, state
):
    """Generate hierarchical table text with [bN:rNcN|CS] [pN|S] format."""
    if block["row_style_aliases"] is None:
        return {"lines": [], "next_alias_counter": alias_counter}

    for rs_alias, tr_pr_xml in state._row_style_map.items():
        if rs_alias not in style_alias_map:
            style_alias_map[rs_alias] = tr_pr_xml
    for cs_alias, tc_xml in state._cell_style_map.items():
        if cs_alias not in style_alias_map:
            style_alias_map[cs_alias] = tc_xml

    block_id = block["id"]
    lines = []
    tbl_element = ET.fromstring(block["xml"])
    rows = tbl_element.findall("w:tr", NAMESPACES)

    for r_idx, tr in enumerate(rows):
        rs_alias = (
            block["row_style_aliases"][r_idx]
            if r_idx < len(block["row_style_aliases"])
            else f"RS{r_idx}"
        )

        lines.append(f"  [{block_id}:r{r_idx}|{rs_alias}]")

        cells = tr.findall("w:tc", NAMESPACES)
        for c_idx, tc in enumerate(cells):
            cell_key = f"r{r_idx}c{c_idx}"
            cs_alias = (
                block["cell_style_map"].get(cell_key, f"CS{c_idx}")
                if block["cell_style_map"]
                else f"CS{c_idx}"
            )

            paragraphs = tc.findall("w:p", NAMESPACES)

            cell_header = f"[{block_id}:r{r_idx}c{c_idx}|{cs_alias}]"
            first_para_in_cell = True

            for p_idx, p in enumerate(paragraphs):
                p_text = "".join(
                    t.text for t in p.findall(".//w:t", NAMESPACES) if t.text
                )
                if not p_text.strip():
                    continue

                p_style_key = _generate_style_key(p)
                if p_style_key not in style_key_to_alias:
                    p_alias = f"S{alias_counter}"
                    style_key_to_alias[p_style_key] = p_alias
                    style_alias_map[p_alias] = p_style_key
                    alias_counter += 1
                p_alias = style_key_to_alias[p_style_key]

                if first_para_in_cell:
                    lines.append(
                        f"    {cell_header} [p{p_idx}|{p_alias}] {p_text}"
                    )
                    first_para_in_cell = False
                else:
                    lines.append(
                        f"    [p{p_idx}|{p_alias}] {p_text}"
                    )

    return {
        "lines": lines,
        "next_alias_counter": alias_counter,
    }



def generate_text_merge(parsed_result, state, num_prefix_map=None):
    """Generate text_merge string with block markers and style aliases."""
    lines = []
    style_alias_map = {}
    style_key_to_alias = {}
    alias_counter = 1

    for block in parsed_result:
        if not block["text"] or not block["text"].strip():
            continue

        if block["type"] == "tbl":
            alias = block["style_key"]
            if alias not in style_alias_map:
                style_alias_map[alias] = alias
                style_key_to_alias[alias] = alias
        else:
            style_key = block["style_key"]
            if style_key not in style_key_to_alias:
                alias = f"S{alias_counter}"
                style_key_to_alias[style_key] = alias
                style_alias_map[alias] = style_key
                alias_counter += 1
            alias = style_key_to_alias[style_key]

        has_numpr = (
            block["type"] == "p"
            and "numPr" in block.get("style_key", "")
        )
        if has_numpr and num_prefix_map:
            prefix = num_prefix_map.get(block["id"], "")
            numpr_flag = f'|numPr:"{prefix}"' if prefix else "|numPr"
        elif has_numpr:
            numpr_flag = "|numPr"
        else:
            numpr_flag = ""

        if block["semantic_tag"]:
            block_marker = f"[{block['id']}:{block['semantic_tag']}|{alias}{numpr_flag}]"
        else:
            block_marker = f"[{block['id']}|{alias}{numpr_flag}]"

        if block["type"] == "p":
            lines.append(f"{block_marker} {block['text']}")
        elif block["type"] == "tbl":
            lines.append(block_marker)
            table_lines = _generate_hierarchical_table_text(
                block, style_key_to_alias, style_alias_map, alias_counter,
                state,
            )
            alias_counter = table_lines["next_alias_counter"]
            for row_line in table_lines["lines"]:
                lines.append(row_line)
        elif block["type"] == "sdt":
            if block["semantic_tag"] == "TOC":
                lines.append(block_marker)
                try:
                    root = ET.fromstring(block["xml"])
                    sdt_content = root.find("w:sdtContent", NAMESPACES)
                    if sdt_content is not None:
                        for p_idx, p in enumerate(
                            sdt_content.findall(".//w:p", NAMESPACES)
                        ):
                            p_text = "".join(
                                t.text
                                for t in p.findall(".//w:t", NAMESPACES)
                                if t.text
                            )
                            if p_text.strip():
                                lines.append(
                                    f"  [{block['id']}:p{p_idx}] {p_text}"
                                )
                except ET.ParseError:
                    lines.append(f"  {block['text']}")
            else:
                lines.append(f"{block_marker} {block['text']}")

    return "\n".join(lines), style_alias_map


def analyze(docx_path, work_dir):
    """Run complete analysis: extract → parse → templates → text_merge."""
    state = AnalyzerState()
    extracted_path = os.path.join(work_dir, "extracted")

    extract_docx_xml(docx_path, extracted_path)
    parsed_result = parse_document_blocks(extracted_path, state)
    paragraph_templates = build_paragraph_style_templates(parsed_result)
    numbering_defs = _parse_numbering_xml(extracted_path)
    num_prefix_map = _compute_effective_numbering(parsed_result, numbering_defs)
    text_merge, style_alias_map = generate_text_merge(
        parsed_result, state, num_prefix_map
    )

    serializable_blocks = list(parsed_result)
    serializable_table_templates = {}
    for tkey, tval in state.table_style_templates.items():
        serializable_table_templates[tkey] = dict(tval)

    output = {
        "text_merge": text_merge,
        "blocks": serializable_blocks,
        "style_alias_map": style_alias_map,
        "paragraph_style_templates": paragraph_templates,
        "table_style_templates": serializable_table_templates,
        "extracted_path": extracted_path,
    }

    return output


def main():
    """CLI entry point: python3 analyze_docx.py <docx_path> <work_dir>"""
    if len(sys.argv) != 3:
        print(
            f"Usage: {sys.argv[0]} <docx_path> <work_dir>",
            file=sys.stderr,
        )
        sys.exit(1)

    docx_path = sys.argv[1]
    work_dir = sys.argv[2]

    if not os.path.exists(docx_path):
        print(f"Error: DOCX file not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(work_dir, exist_ok=True)

    result = analyze(docx_path, work_dir)

    analysis_path = os.path.join(work_dir, "analysis.json")
    with open(analysis_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(result["text_merge"])


if __name__ == "__main__":
    main()
