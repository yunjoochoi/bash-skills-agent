#!/usr/bin/env python3
"""Phase 1: Standalone DOCX Analyzer for the editing pipeline.

Ported from docx_agent.tools.readers.docx_analyzer.DocxAnalyzer.

Usage:
    python3 analyze_docx.py <docx_path> <work_dir>

Outputs:
    - Prints text_merge to stdout
    - Saves analysis.json to <work_dir>/analysis.json
    - Extracts XML files to <work_dir>/extracted/
"""

import copy
import json
import os
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

# ============================================================================
# Constants
# ============================================================================

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W_NS = NAMESPACES["w"]

# Register namespaces for clean ElementTree output
for _prefix, _uri in NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)

# Layout tags stripped from tcPr (handled by assembler post-processing)
_TCPR_LAYOUT_TAGS = frozenset({
    f"{{{W_NS}}}tcW",
    f"{{{W_NS}}}gridSpan",
    f"{{{W_NS}}}vMerge",
    f"{{{W_NS}}}hMerge",
})

# Boolean-only tags for style key generation
_BOOLEAN_TAGS = {"b", "bCs", "i", "iCs", "strike", "dstrike", "caps", "smallCaps"}

# rPr tags ignored for formatting variance detection.
# These differ per script/language boundary (Korean↔English) and are
# not intentional formatting differences.
_IGNORED_RPR_TAGS = frozenset({
    f"{{{W_NS}}}rFonts",
    f"{{{W_NS}}}lang",
})


# ============================================================================
# State namespace (replaces instance variables from the class)
# ============================================================================

class _AnalyzerState:
    """Mutable state for table/TOC deduplication caches.

    Replaces DocxAnalyzer instance variables.
    """

    __slots__ = (
        "table_style_templates",
        "_table_counter",
        "_row_style_cache",
        "_cell_style_cache",
        "_rs_counter",
        "_cs_counter",
        "_row_style_map",
        "_cell_style_map",
        "_toc_level_cache",
        "_tl_counter",
        "_toc_style_templates",
    )

    def __init__(self):
        # Table style management
        self.table_style_templates = {}   # T1 -> template dict
        self._table_counter = 0

        # RS/CS alias system for hierarchical table styles
        self._row_style_cache = {}   # fingerprint -> RS alias
        self._cell_style_cache = {}  # fingerprint -> CS alias
        self._rs_counter = 0
        self._cs_counter = 0

        # Maps: alias -> template XML (populated during analysis)
        self._row_style_map = {}   # RS0 -> tr_pr_xml
        self._cell_style_map = {}  # CS0 -> tc_xml_template

        # TOC style management
        self._toc_level_cache = {}   # fingerprint -> TL alias
        self._tl_counter = 0
        self._toc_style_templates = {}  # TL0 -> template dict


# ============================================================================
# Text Extraction
# ============================================================================

def _extract_paragraph_text(p_element):
    """Extract text content from paragraph element.

    Args:
        p_element: Paragraph XML element

    Returns:
        Concatenated text from all w:t elements
    """
    texts = []
    for t in p_element.findall(".//w:t", NAMESPACES):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _extract_table_text(tbl_element):
    """Extract text content from table element with cell and paragraph coords.

    Format: Each row on a new line with cells separated by |
    For cells with multiple paragraphs, each paragraph gets its own ID.

    Args:
        tbl_element: Table XML element

    Returns:
        Formatted table text string
    """
    rows = []
    for r_idx, tr in enumerate(tbl_element.findall(".//w:tr", NAMESPACES)):
        cells = []
        for c_idx, tc in enumerate(tr.findall(".//w:tc", NAMESPACES)):
            paragraphs = tc.findall("w:p", NAMESPACES)

            if len(paragraphs) <= 1:
                # Simple cell: single paragraph, inline format
                cell_text = ""
                for t in tc.findall(".//w:t", NAMESPACES):
                    if t.text:
                        cell_text += t.text
                cells.append(f"[r{r_idx}c{c_idx}] {cell_text}")
            else:
                # Complex cell: multiple paragraphs, show each with pN ID
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
    """Extract text content from structured document tag element.

    For TOC and similar multi-paragraph SDT blocks, extracts each paragraph
    with its index for partial editing support.

    Args:
        sdt_element: SDT XML element

    Returns:
        Formatted SDT text string
    """
    # Get SDT alias if available
    alias = ""
    sdt_pr = sdt_element.find("w:sdtPr", NAMESPACES)
    if sdt_pr is not None:
        alias_elem = sdt_pr.find("w:alias", NAMESPACES)
        if alias_elem is not None:
            alias = alias_elem.get(f"{{{NAMESPACES['w']}}}val", "")

    # Find SDT content
    sdt_content = sdt_element.find("w:sdtContent", NAMESPACES)
    if sdt_content is None:
        # Fallback to direct text extraction
        texts = []
        for t in sdt_element.findall(".//w:t", NAMESPACES):
            if t.text:
                texts.append(t.text)
        content_preview = " ".join(texts)[:100]
        if alias:
            return f"[{alias}] {content_preview}"
        return f"[SDT] {content_preview}"

    # Extract paragraphs from SDT content
    paragraphs = sdt_content.findall(".//w:p", NAMESPACES)

    if len(paragraphs) <= 1:
        # Simple SDT: single paragraph, inline format
        texts = []
        for t in sdt_element.findall(".//w:t", NAMESPACES):
            if t.text:
                texts.append(t.text)
        content = " ".join(texts)[:100]
        if alias:
            return f"[{alias}] {content}"
        return f"[SDT] {content}"

    # Multi-paragraph SDT (like TOC): show each paragraph with pN index
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
    """Detect if SDT element is a Table of Contents.

    Checks w:sdtPr for TOC indicators:
    1. w:docPartGallery with "Table of Contents"
    2. w:alias containing "TOC"
    3. PAGEREF instrText in content

    Args:
        sdt_element: SDT XML element

    Returns:
        True if SDT is a TOC
    """
    w_ns = NAMESPACES["w"]
    sdt_pr = sdt_element.find("w:sdtPr", NAMESPACES)
    if sdt_pr is not None:
        # Check docPartGallery
        doc_part = sdt_pr.find(".//w:docPartGallery", NAMESPACES)
        if doc_part is not None:
            val = doc_part.get(f"{{{w_ns}}}val", "")
            if "Table of Contents" in val:
                return True

        # Check alias
        alias_elem = sdt_pr.find("w:alias", NAMESPACES)
        if alias_elem is not None:
            val = alias_elem.get(f"{{{w_ns}}}val", "")
            if "TOC" in val.upper():
                return True

    # Check for PAGEREF fields in content
    instr_texts = sdt_element.findall(".//w:instrText", NAMESPACES)
    for instr in instr_texts:
        if instr.text and "PAGEREF" in instr.text:
            return True

    return False


# ============================================================================
# Style Key System
# ============================================================================

def _element_to_key_part(elem):
    """Convert XML element to deterministic key part string.

    Handles various element patterns:
    - Simple value: <w:jc w:val="center"/> -> "jc-center"
    - Multi-attr: <w:spacing w:after="200" w:line="276"/>
      -> "spacing-after200-line276"
    - Nested: <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
              -> "numPr-ilvl0-numId1"
    - Boolean: <w:b/> -> "b"
    - Boolean false: <w:b w:val="0"/> -> "" (skipped)

    Args:
        elem: XML element

    Returns:
        Key part string, or empty string if element should be skipped
    """
    # Strip namespace prefix for tag name
    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

    # Check for boolean false: <w:b w:val="0"/> or <w:i w:val="false"/>
    val_attr = elem.get(f"{{{W_NS}}}val")
    if (
        tag in _BOOLEAN_TAGS
        and val_attr in ("0", "false")
        and len(elem.attrib) == 1
        and len(elem) == 0
    ):
        return ""

    # Collect attribute values (sorted by local attr name)
    attr_parts = []
    for attr_name in sorted(elem.attrib.keys()):
        local_name = attr_name.split("}")[-1] if "}" in attr_name else attr_name
        attr_val = elem.attrib[attr_name]
        # For simple w:val, use just the value
        if local_name == "val":
            attr_parts.append(attr_val)
        else:
            attr_parts.append(f"{local_name}{attr_val}")

    # Handle nested children recursively
    child_parts = []
    for child in sorted(elem, key=lambda x: x.tag):
        child_part = _element_to_key_part(child)
        if child_part:
            child_parts.append(child_part)

    # Build final part
    all_sub = attr_parts + child_parts
    if all_sub:
        return f"{tag}-{'_'.join(all_sub)}"
    return tag  # Boolean element like <w:b/>


def _extract_p_style(p_element):
    """Extract w:pStyle value from paragraph element.

    Args:
        p_element: Paragraph XML element

    Returns:
        pStyle value or "Normal" if not present
    """
    p_pr = p_element.find("w:pPr", NAMESPACES)
    if p_pr is not None:
        p_style_elem = p_pr.find("w:pStyle", NAMESPACES)
        if p_style_elem is not None:
            return p_style_elem.get(f"{{{W_NS}}}val", "Normal")
    return "Normal"


def _build_ppr_key(p_element):
    """Build deterministic key from ALL pPr children (excluding pStyle, rPr).

    Sorts children by tag name for order independence.
    Sorts attributes alphabetically within each element.

    Args:
        p_element: Paragraph XML element

    Returns:
        Underscore-joined key parts, or empty string if no relevant pPr children
    """
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
    """Generate style key from pPr children only (no rPr).

    Key format: {pStyle}_{ppr_key} or just {pStyle} if no pPr children.

    Args:
        p_element: Paragraph XML element

    Returns:
        Deterministic style key string
    """
    p_style = _extract_p_style(p_element)
    ppr_key = _build_ppr_key(p_element)
    return f"{p_style}_{ppr_key}" if ppr_key else p_style


# ============================================================================
# Template Building
# ============================================================================

def _build_rpr_key(run):
    """Build deterministic key from ALL rPr children.

    Sorts by tag name for order independence.

    Args:
        run: Run XML element (<w:r>)

    Returns:
        Underscore-joined key parts, or "default" if no rPr
    """
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
    """Check if a run element contains actual text content.

    Args:
        run: Run XML element (<w:r>)

    Returns:
        True if at least one <w:t> has non-empty text
    """
    for t in run.findall("w:t", NAMESPACES):
        if t.text and t.text.strip():
            return True
    return False


def _describe_run_styles(run):
    """Generate human-readable description of run styles for LLM prompt.

    Args:
        run: Run XML element (<w:r>)

    Returns:
        Description string like "bold, size:28, font:Arial"
    """
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
    """Create run style template dict from a <w:r> element.

    Deep copies the run, replaces <w:t> text with {{content}}.
    Preserves ALL rPr children (rFonts, sz, b, etc.) in the XML template.

    Args:
        run: Run XML element (<w:r>)

    Returns:
        Dict with rpr_key, rpr_xml, display_description
    """
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
    """Extract <w:pPr>...</w:pPr> XML (no filtering, preserves all children).

    Includes <w:rPr> inside <w:pPr> (paragraph mark default run properties).

    Args:
        p_element: Paragraph XML element

    Returns:
        Serialized <w:pPr> XML string, or empty string if no pPr
    """
    p_pr = p_element.find("w:pPr", NAMESPACES)
    if p_pr is None:
        return ""
    return ET.tostring(p_pr, encoding="unicode")


# ============================================================================
# Table Hierarchy
# ============================================================================

def _extract_table_xml_template(tbl_element):
    """Extract table shell XML: tblPr + tblGrid only (no rows).

    Args:
        tbl_element: Table XML element

    Returns:
        Table shell XML string without w:tr elements
    """
    template = copy.deepcopy(tbl_element)
    for tr in template.findall("w:tr", NAMESPACES):
        template.remove(tr)
    return ET.tostring(template, encoding="unicode")


def _extract_row_xml_template(tr_element):
    """Extract row properties as template dict.

    Args:
        tr_element: Table row XML element

    Returns:
        Dict with tr_pr_xml_template
    """
    tr_pr = tr_element.find("w:trPr", NAMESPACES)
    tr_pr_xml = ET.tostring(tr_pr, encoding="unicode") if tr_pr is not None else ""
    return {"tr_pr_xml_template": tr_pr_xml}


def _extract_cell_xml_template(tc_element):
    """Extract cell XML shell: tcPr (visual only) + {{content}}.

    Matroshka pattern: keeps only visual tcPr properties (shading,
    borders, vAlign), strips layout properties (tcW, gridSpan, vMerge)
    that are handled by assembler post-processing.
    The {{content}} placeholder is filled at assembly time with
    paragraph XML built from the linked ParagraphStyleTemplate.

    Args:
        tc_element: Table cell XML element

    Returns:
        Cell XML shell string
    """
    template = copy.deepcopy(tc_element)

    # Remove all paragraphs - content is injected at assembly time
    for p in template.findall("w:p", NAMESPACES):
        template.remove(p)

    # Strip layout properties from tcPr
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
    """Extract cell style with paragraph hierarchy.

    For each paragraph in the cell, builds a paragraph style template
    dict with pPr XML and unique run style templates.

    Args:
        tc_element: Table cell XML element

    Returns:
        Dict with paragraph_styles and tc_xml_template
    """
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
    """Extract hierarchical table style with T1:R0:C0 format.

    Creates a table style template dict with:
    - Table-level: tblPr + tblGrid shell XML
    - Row-level: trPr XML template
    - Cell-level: paragraph_styles + tc_xml_template

    Deduplicates row/cell styles - identical styles share same alias.

    Args:
        tbl_element: Table XML element
        state: _AnalyzerState instance

    Returns:
        Tuple of (table_template_dict, row_style_aliases, cell_style_map)
    """
    state._table_counter += 1
    table_key = f"T{state._table_counter}"
    tbl_xml_template = _extract_table_xml_template(tbl_element)

    row_styles = {}
    cell_templates = {}
    row_style_aliases = []
    cell_style_map = {}

    for r_idx, tr in enumerate(tbl_element.findall("w:tr", NAMESPACES)):
        # Extract row style and assign RS alias via fingerprint
        row_tmpl = _extract_row_xml_template(tr)
        fp = row_tmpl["tr_pr_xml_template"]

        if fp not in state._row_style_cache:
            rs_alias = f"RS{state._rs_counter}"
            state._row_style_cache[fp] = rs_alias
            state._row_style_map[rs_alias] = fp
            state._rs_counter += 1

        row_style_aliases.append(state._row_style_cache[fp])
        row_styles[r_idx] = row_tmpl

        # Extract cell styles per row
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


# ============================================================================
# TOC Handling
# ============================================================================

def _toc_level_fingerprint(indent_left, borders, bold):
    """Generate fingerprint for TOC level deduplication.

    Primary: w:pBdr (paragraph borders) - different borders = different level.
    Secondary: indent (distinguishes levels with same border style).
    Tabs excluded (not all TOCs have tabs).

    Args:
        indent_left: Left indent in twips
        borders: Frozenset of serialized border elements from w:pBdr
        bold: Whether first run is bold

    Returns:
        Fingerprint string for cache lookup
    """
    border_key = ",".join(sorted(borders)) if borders else "none"
    parts = [f"borders:{border_key}", f"indent:{indent_left}"]
    if bold:
        parts.append("bold:True")
    return "|".join(parts)


def _extract_toc_paragraph_template(p_element):
    """Extract TOC paragraph template preserving structure.

    Deep copies the paragraph, replaces text placeholders:
    - Number text -> {{number}}
    - Title text -> {{title}}
    - Page number -> {{page}}
    - Anchor references -> {{anchor}}

    Args:
        p_element: Paragraph XML element

    Returns:
        Template XML string with placeholders
    """
    template = copy.deepcopy(p_element)
    w_ns = NAMESPACES["w"]

    # Update anchor on hyperlinks
    for hl in template.findall(".//w:hyperlink", NAMESPACES):
        hl.set(f"{{{w_ns}}}anchor", "{{anchor}}")

    # Update PAGEREF instrText
    for instr in template.findall(".//w:instrText", NAMESPACES):
        if instr.text and "PAGEREF" in instr.text:
            instr.text = " PAGEREF {{anchor}} \\h "

    # Replace text content in runs
    hyperlinks = template.findall(".//w:hyperlink", NAMESPACES)
    if hyperlinks and len(hyperlinks) >= 1:
        # Complex structure: first hyperlink has number
        first_hl = hyperlinks[0]
        t_elems = first_hl.findall(".//w:t", NAMESPACES)
        if t_elems:
            t_elems[0].text = "{{number}}"
            for t in t_elems[1:]:
                t.text = ""

    # Replace title/page in PAGEREF field runs
    all_runs = template.findall(".//w:r", NAMESPACES)
    in_pageref = False
    title_set = False
    page_set = False

    for run in all_runs:
        fld_char = run.find("w:fldChar", NAMESPACES)
        if fld_char is not None:
            fld_type = fld_char.get(f"{{{w_ns}}}fldCharType")
            if fld_type == "separate":
                in_pageref = True
                continue
            elif fld_type == "end":
                in_pageref = False
                continue

        if in_pageref:
            t_elems = run.findall("w:t", NAMESPACES)
            for t in t_elems:
                if not title_set:
                    t.text = "{{title}}"
                    title_set = True
                elif not page_set:
                    tab = run.find("w:tab", NAMESPACES)
                    if tab is not None or (t.text and t.text.strip().isdigit()):
                        t.text = "{{page}}"
                        page_set = True
                    else:
                        t.text = ""

    # If no hyperlinks (simple structure), replace all text
    if not hyperlinks:
        runs = template.findall(".//w:r", NAMESPACES)
        first_text_set = False
        for run in runs:
            for t in run.findall("w:t", NAMESPACES):
                if not first_text_set:
                    t.text = "{{number}} {{title}}"
                    first_text_set = True
                else:
                    t.text = ""
        # Append {{page}} as a separate run with <w:tab/> element
        if first_text_set and runs:
            page_run = ET.SubElement(template, f"{{{w_ns}}}r")
            src_rpr = runs[0].find("w:rPr", NAMESPACES)
            if src_rpr is not None:
                page_run.append(copy.deepcopy(src_rpr))
            ET.SubElement(page_run, f"{{{w_ns}}}tab")
            page_t = ET.SubElement(page_run, f"{{{w_ns}}}t")
            page_t.text = "{{page}}"

    return ET.tostring(template, encoding="unicode")


def _extract_toc_entry_levels(sdt_element, state):
    """Extract hierarchical TOC styles from SDT element.

    For each paragraph in sdtContent:
    1. Extract indent, borders, bold for fingerprinting
    2. Assign TL alias via deduplication cache
    3. Build entry_levels mapping (paragraph index -> TL alias)

    Args:
        sdt_element: SDT XML element containing TOC
        state: _AnalyzerState instance

    Returns:
        Dict mapping paragraph index to TL alias
    """
    w_ns = NAMESPACES["w"]

    sdt_content = sdt_element.find("w:sdtContent", NAMESPACES)
    if sdt_content is None:
        return {}

    paragraphs = sdt_content.findall(".//w:p", NAMESPACES)
    if not paragraphs:
        return {}

    entry_levels = {}

    for p_idx, p in enumerate(paragraphs):
        # Skip empty paragraphs
        p_text = ""
        for t in p.findall(".//w:t", NAMESPACES):
            if t.text:
                p_text += t.text
        if not p_text.strip():
            continue

        ppr = p.find("w:pPr", NAMESPACES)

        # Extract indent
        indent_left = 0
        if ppr is not None:
            ind = ppr.find("w:ind", NAMESPACES)
            if ind is not None:
                left_val = ind.get(f"{{{w_ns}}}left", "0")
                try:
                    indent_left = int(left_val)
                except ValueError:
                    indent_left = 0

        # Extract borders for fingerprinting
        borders = frozenset()
        if ppr is not None:
            pbdr = ppr.find("w:pBdr", NAMESPACES)
            if pbdr is not None:
                borders = frozenset(
                    ET.tostring(child, encoding="unicode")
                    for child in pbdr
                )

        # Extract bold from first run for fingerprinting
        bold = False
        first_run = p.find(".//w:r", NAMESPACES)
        if first_run is not None:
            rpr = first_run.find("w:rPr", NAMESPACES)
            if rpr is not None:
                b_elem = rpr.find("w:b", NAMESPACES)
                if b_elem is not None:
                    val = b_elem.get(f"{{{w_ns}}}val", "true")
                    bold = val != "false" and val != "0"

        # Compute fingerprint and assign alias
        fp = _toc_level_fingerprint(indent_left, borders, bold)

        if fp not in state._toc_level_cache:
            alias = f"TL{state._tl_counter}"
            state._toc_level_cache[fp] = alias
            state._tl_counter += 1

            # Extract paragraph template
            xml_template = _extract_toc_paragraph_template(p)

            # Collect run style templates from paragraph
            run_templates = {}
            seen_rpr = set()
            for run in p.findall(".//w:r", NAMESPACES):
                if not _run_has_text(run):
                    continue
                rst = _build_run_style_template(run)
                if rst["rpr_key"] not in seen_rpr:
                    run_templates[f"RS{len(run_templates)}"] = rst
                    seen_rpr.add(rst["rpr_key"])

            state._toc_style_templates[alias] = {
                "toc_style_key": alias,
                "toc_xml_template": xml_template,
                "run_style_templates": run_templates,
                "display_description": (
                    f"TOC Level {state._tl_counter}"
                    f" (indent: {indent_left}twips)"
                ),
            }

        tl_alias = state._toc_level_cache[fp]
        entry_levels[p_idx] = tl_alias

    return entry_levels


def _get_toc_entry_text(block, para_idx):
    """Extract text from a specific TOC paragraph by index.

    Args:
        block: Block dict with 'xml' key containing SDT XML
        para_idx: 0-indexed paragraph in sdtContent

    Returns:
        Entry text (e.g., "1. Introduction | 3")
    """
    try:
        root = ET.fromstring(block["xml"])
        sdt_content = root.find("w:sdtContent", NAMESPACES)
        if sdt_content is None:
            return ""

        paragraphs = sdt_content.findall(".//w:p", NAMESPACES)
        if para_idx >= len(paragraphs):
            return ""

        p = paragraphs[para_idx]
        texts = []
        for t in p.findall(".//w:t", NAMESPACES):
            if t.text:
                texts.append(t.text)
        return " ".join(texts)
    except ET.ParseError:
        return ""


# ============================================================================
# Semantic Tag Inference
# ============================================================================

def _build_style_lookup(extracted_path):
    """Parse styles.xml and build a lookup dict for semantic inference.

    Args:
        extracted_path: Path to extracted XMLs directory

    Returns:
        Dict mapping styleId to style info dict with keys:
        - name: Style name
        - outline_lvl: Outline level (int or None)
        - based_on: basedOn styleId or None
    """
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
    """Determine semantic tag (H1, BODY, LIST, etc.) for a paragraph block.

    Priority:
    1. Paragraph XML outlineLvl (highest priority)
    2. styles.xml outlineLvl
    3. Style name keywords (fallback)

    Args:
        block: Block dict with 'xml' and 'style_key'
        styles_xml_path: Path to extracted XMLs directory (for styles.xml lookup)

    Returns:
        Semantic tag string (H1, H2, ..., BODY, LIST, TITLE, SUBTITLE, OTHER)
    """
    try:
        p_element = ET.fromstring(block["xml"])
    except ET.ParseError:
        return "BODY"

    # 1. Check paragraph-level outlineLvl first
    outline_elem = p_element.find(".//w:outlineLvl", NAMESPACES)
    if outline_elem is not None:
        try:
            lvl = int(outline_elem.get(f"{{{W_NS}}}val", "0"))
            return f"H{lvl + 1}"
        except ValueError:
            pass

    # Extract style ID from paragraph
    style_id = _extract_p_style(p_element)

    # 2. Check styles.xml lookup
    style_lookup = _build_style_lookup(styles_xml_path)
    if style_id in style_lookup:
        info = style_lookup[style_id]
        if info.get("outline_lvl") is not None:
            return f"H{info['outline_lvl'] + 1}"

    # 3. Fallback to name-based inference
    name_lower = style_id.lower()
    name_from_lookup = ""
    if style_id in style_lookup:
        name_from_lookup = style_lookup[style_id].get("name", "").lower()

    check_names = [name_lower, name_from_lookup]
    for name in check_names:
        if not name:
            continue
        if "heading" in name:
            # Try to extract heading level from name
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


# ============================================================================
# Core Pipeline
# ============================================================================

def extract_docx_xml(docx_path, output_dir):
    """Extract XML files from DOCX archive.

    Extracts all .xml and .rels files while preserving directory structure.

    Args:
        docx_path: Path to the DOCX file
        output_dir: Directory to extract XMLs to

    Returns:
        Path to the extracted directory

    Raises:
        FileNotFoundError: If DOCX file does not exist
        zipfile.BadZipFile: If file is not a valid DOCX/ZIP
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")

    # Clean output directory if exists
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
    """Parse document.xml into content blocks.

    Extracts paragraphs (w:p), tables (w:tbl), and SDTs (w:sdt)
    as content blocks with unique IDs. Section properties (w:sectPr)
    are skipped as they are immutable layout anchors.

    Args:
        extracted_path: Path to extracted XMLs directory
        state: _AnalyzerState instance

    Returns:
        List of block dicts

    Raises:
        FileNotFoundError: If document.xml not found
    """
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
                "toc_entry_levels": None,
            })
            block_id += 1

        elif tag == "tbl":
            text = _extract_table_text(child)
            xml_str = ET.tostring(child, encoding="unicode")

            # Extract hierarchical table styles
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
                "toc_entry_levels": None,
            })
            block_id += 1

        elif tag == "sectPr":
            continue  # Section properties are immutable; skip

        elif tag == "sdt":
            is_toc = _is_toc_sdt(child)
            text = _extract_sdt_text(child)
            xml_str = ET.tostring(child, encoding="unicode")

            toc_entry_levels = None
            semantic_tag = "SDT"
            style_key = "sdt_default"

            if is_toc:
                toc_entry_levels = _extract_toc_entry_levels(child, state)
                semantic_tag = "TOC"
                style_key = "toc_default"

            blocks.append({
                "id": f"b{block_id}",
                "type": "sdt",
                "text": text,
                "xml": xml_str,
                "style_key": style_key,
                "semantic_tag": semantic_tag,
                "row_style_aliases": None,
                "cell_style_map": None,
                "toc_entry_levels": toc_entry_levels,
            })
            block_id += 1

    # Infer semantic tags for paragraph blocks
    for block in blocks:
        if block["type"] == "p":
            block["semantic_tag"] = _infer_semantic_info(block, extracted_path)

    return blocks


def build_paragraph_style_templates(parsed_result):
    """Build paragraph style template dictionary from parsed blocks.

    Creates deduplicated paragraph style templates with pPr XML and
    run style templates for each unique paragraph style_key.
    Table blocks are handled separately via table_style_templates.

    Args:
        parsed_result: List of parsed block dicts

    Returns:
        Dictionary mapping style_key to paragraph style template dict
    """
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

            # Collect unique run patterns from every block with this key
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
    """Generate hierarchical table text aligned with target_id format.

    Output format matches edit target_id coordinates directly:

        [b13:r0|RS0]
          [b13:r0c0|CS2] [p0|S1] Header1
          [b13:r0c1|CS3] [p0|S2] Header2
        [b13:r1|RS1]
          [b13:r1c0|CS2] [p0|S1] Data1
          [b13:r1c1|CS3] [p0|S2] Data2

    Args:
        block: Block dict with row_style_aliases and cell_style_map
        style_key_to_alias: Existing style_key -> alias mapping
        style_alias_map: Existing alias -> style_key mapping
        alias_counter: Current alias counter
        state: _AnalyzerState instance

    Returns:
        Dict with 'lines' and 'next_alias_counter'
    """
    if block["row_style_aliases"] is None:
        return {"lines": [], "next_alias_counter": alias_counter}

    # Add RS/CS mappings to style_alias_map
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

        # Row line: [b13:r0|RS0]
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

            # Cell header: [b13:r0c0|CS2] (CS alias at cell level)
            cell_header = f"[{block_id}:r{r_idx}c{c_idx}|{cs_alias}]"
            first_para_in_cell = True

            for p_idx, p in enumerate(paragraphs):
                p_text = "".join(
                    t.text for t in p.findall(".//w:t", NAMESPACES) if t.text
                )
                if not p_text.strip():
                    continue

                # All paragraphs (including p0) get S alias
                p_style_key = _generate_style_key(p)
                if p_style_key not in style_key_to_alias:
                    p_alias = f"S{alias_counter}"
                    style_key_to_alias[p_style_key] = p_alias
                    style_alias_map[p_alias] = p_style_key
                    alias_counter += 1
                p_alias = style_key_to_alias[p_style_key]

                # Format: [b13:r0c0|CS2] [p0|S1] text
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


def _extract_table_text_with_styles(
    table_xml, style_key_to_alias, style_alias_map, alias_counter
):
    """Extract table text with per-cell-paragraph style aliases (legacy fallback).

    Args:
        table_xml: Table XML string
        style_key_to_alias: Existing style_key -> alias mapping
        style_alias_map: Existing alias -> style_key mapping
        alias_counter: Current alias counter

    Returns:
        Dict with 'lines' (list of formatted lines) and 'next_alias_counter'
    """
    tbl_element = ET.fromstring(table_xml)
    rows = []

    for r_idx, tr in enumerate(tbl_element.findall(".//w:tr", NAMESPACES)):
        cells = []
        for c_idx, tc in enumerate(tr.findall(".//w:tc", NAMESPACES)):
            paragraphs = tc.findall("w:p", NAMESPACES)

            if len(paragraphs) <= 1:
                # Simple cell: single paragraph
                p = paragraphs[0] if paragraphs else None
                cell_text = ""
                for t in tc.findall(".//w:t", NAMESPACES):
                    if t.text:
                        cell_text += t.text

                if p is not None:
                    p_style_key = _generate_style_key(p)
                    if p_style_key not in style_key_to_alias:
                        alias = f"S{alias_counter}"
                        style_key_to_alias[p_style_key] = alias
                        style_alias_map[alias] = p_style_key
                        alias_counter += 1
                    p_alias = style_key_to_alias[p_style_key]
                    cells.append(f"[r{r_idx}c{c_idx}|{p_alias}] {cell_text}")
                else:
                    cells.append(f"[r{r_idx}c{c_idx}] {cell_text}")
            else:
                # Complex cell: multiple paragraphs
                cell_lines = [f"[r{r_idx}c{c_idx}]"]
                for p_idx, p in enumerate(paragraphs):
                    p_text = ""
                    for t in p.findall(".//w:t", NAMESPACES):
                        if t.text:
                            p_text += t.text

                    if p_text.strip():
                        p_style_key = _generate_style_key(p)
                        if p_style_key not in style_key_to_alias:
                            alias = f"S{alias_counter}"
                            style_key_to_alias[p_style_key] = alias
                            style_alias_map[alias] = p_style_key
                            alias_counter += 1
                        p_alias = style_key_to_alias[p_style_key]
                        cell_lines.append(
                            f"    [r{r_idx}c{c_idx}p{p_idx}|{p_alias}] {p_text}"
                        )
                cells.append("\n".join(cell_lines))
        rows.append(" | ".join(cells))

    return {
        "lines": rows,
        "next_alias_counter": alias_counter,
    }


def generate_text_merge(parsed_result, state):
    """Generate text merge string for LLM input with style aliases.

    Creates a text representation with block IDs and style aliases.
    Uses short aliases (S1, S2) to save tokens, with mapping stored separately.

    Format: [bN:TAG|S1] text content

    Args:
        parsed_result: List of parsed block dicts
        state: _AnalyzerState instance

    Returns:
        Tuple of (text_merge, style_alias_map)
        - text_merge: Concatenated text with block markers
        - style_alias_map: {"S1": "Heading1_jc-...", "S2": "Normal_jc-..."}
    """
    lines = []
    style_alias_map = {}
    style_key_to_alias = {}
    alias_counter = 1

    for block in parsed_result:
        # Skip blocks with empty or whitespace-only text
        if not block["text"] or not block["text"].strip():
            continue

        # Get or create alias for this style_key
        # TABLE blocks use style_key directly as alias (T1, T2...)
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

        # Build block marker with semantic tag and alias
        if block["semantic_tag"]:
            block_marker = f"[{block['id']}:{block['semantic_tag']}|{alias}]"
        else:
            block_marker = f"[{block['id']}|{alias}]"

        if block["type"] == "p":
            lines.append(f"{block_marker} {block['text']}")
        elif block["type"] == "tbl":
            lines.append(block_marker)
            # Use hierarchical format with inline table metadata
            if block["row_style_aliases"] is not None:
                table_lines = _generate_hierarchical_table_text(
                    block, style_key_to_alias, style_alias_map, alias_counter,
                    state,
                )
                alias_counter = table_lines["next_alias_counter"]
                for row_line in table_lines["lines"]:
                    lines.append(row_line)
            else:
                # Fallback to legacy format
                table_lines = _extract_table_text_with_styles(
                    block["xml"], style_key_to_alias, style_alias_map,
                    alias_counter,
                )
                alias_counter = table_lines["next_alias_counter"]
                for row_line in table_lines["lines"]:
                    lines.append(f"  {row_line}")
        elif block["type"] == "sdt":
            if block["toc_entry_levels"]:
                # TOC with level-aware formatting
                lines.append(block_marker)
                for p_idx, tl_alias in sorted(
                    block["toc_entry_levels"].items(),
                    key=lambda x: x[0],
                ):
                    if tl_alias not in style_alias_map:
                        style_alias_map[tl_alias] = tl_alias
                    entry_text = _get_toc_entry_text(block, p_idx)
                    lines.append(
                        f"  [{block['id']}:p{p_idx}|{tl_alias}] {entry_text}"
                    )
            else:
                lines.append(f"{block_marker} {block['text']}")

    return "\n".join(lines), style_alias_map


def analyze(docx_path, work_dir):
    """Run complete Phase 1 analysis on a DOCX file.

    Combines extraction, parsing, template building, and text merge generation.

    Args:
        docx_path: Path to the DOCX file
        work_dir: Working directory for output files

    Returns:
        Dict containing all analysis results
    """
    state = _AnalyzerState()
    extracted_path = os.path.join(work_dir, "extracted")

    # Step 1: Extract XML from DOCX
    extract_docx_xml(docx_path, extracted_path)

    # Step 2: Parse document blocks
    parsed_result = parse_document_blocks(extracted_path, state)

    # Step 3: Build paragraph style templates
    paragraph_templates = build_paragraph_style_templates(parsed_result)

    # Step 4: Generate text merge
    text_merge, style_alias_map = generate_text_merge(parsed_result, state)

    # Build output dict (all plain dicts, no Pydantic)
    # Convert toc_entry_levels keys from int to str for JSON serialization
    serializable_blocks = []
    for block in parsed_result:
        b = dict(block)
        if b["toc_entry_levels"] is not None:
            b["toc_entry_levels"] = {
                str(k): v for k, v in b["toc_entry_levels"].items()
            }
        serializable_blocks.append(b)

    # Convert table_style_templates row_styles/cell_style_templates keys
    serializable_table_templates = {}
    for tkey, tval in state.table_style_templates.items():
        t = dict(tval)
        # row_styles and cell_style_templates already have str keys from
        # _extract_table_hierarchy
        serializable_table_templates[tkey] = t

    output = {
        "text_merge": text_merge,
        "blocks": serializable_blocks,
        "style_alias_map": style_alias_map,
        "paragraph_style_templates": paragraph_templates,
        "table_style_templates": serializable_table_templates,
        "toc_style_templates": state._toc_style_templates,
        "extracted_path": extracted_path,
    }

    return output


# ============================================================================
# CLI Entry Point
# ============================================================================

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

    # Save analysis.json
    analysis_path = os.path.join(work_dir, "analysis.json")
    with open(analysis_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # Print text_merge to stdout
    print(result["text_merge"])


if __name__ == "__main__":
    main()
