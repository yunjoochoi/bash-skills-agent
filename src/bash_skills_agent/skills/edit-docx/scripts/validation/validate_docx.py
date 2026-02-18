#!/usr/bin/env python3
"""Post-apply validation of the output DOCX file.

Checks ZIP integrity, XML well-formedness, document structure,
namespace prefix pollution, and optionally verifies that edit
content was actually applied.

Usage:
    python3 validate_docx.py <output_docx> [<work_dir>]

Arguments:
    output_docx  Path to the output DOCX file.
    work_dir     Optional. If provided, also verifies content against edits.json.

Outputs:
    JSON to stdout: {"valid": bool, "errors": [...], "warnings": [...]}
    Exit code: 0 if valid, 1 if errors found.
"""

import json
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

REQUIRED_ENTRIES = {"[Content_Types].xml", "word/document.xml"}

_RE_AUTO_NS = re.compile(r"\bns\d+:")


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------

def _err(check, msg):
    return {"check": check, "level": "error", "message": msg}


def _warn(check, msg):
    return {"check": check, "level": "warning", "message": msg}


def _read_zip_entry(zf, name):
    """Read a ZIP entry as UTF-8 string, falling back to latin-1."""
    raw = zf.read(name)
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1")


def _extract_all_text(xml_content):
    """Extract all w:t text nodes from document XML."""
    ns = NAMESPACES["w"]
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError:
        return ""
    parts = []
    for t in root.iter(f"{{{ns}}}t"):
        if t.text:
            parts.append(t.text)
    return " ".join(parts)


# -------------------------------------------------------------------
# Individual validators
# -------------------------------------------------------------------

def validate_zip(path):
    """Verify the file is a valid ZIP archive."""
    issues = []
    try:
        with zipfile.ZipFile(path, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                issues.append(_err("zip", f"Corrupt ZIP entry: {bad}"))
    except zipfile.BadZipFile as e:
        issues.append(_err("zip", f"Invalid ZIP: {e}"))
    except FileNotFoundError:
        issues.append(_err("zip", f"File not found: {path}"))
    return issues


def validate_entries(path):
    """Check that required OOXML entries exist."""
    issues = []
    try:
        with zipfile.ZipFile(path, "r") as zf:
            names = set(zf.namelist())
            for req in REQUIRED_ENTRIES:
                if req not in names:
                    issues.append(_err("entries", f"Missing required entry: {req}"))
    except (zipfile.BadZipFile, FileNotFoundError):
        pass  # Already reported by validate_zip
    return issues


def validate_xml(path):
    """Parse every .xml entry to ensure well-formedness."""
    issues = []
    try:
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                if not name.endswith(".xml"):
                    continue
                try:
                    content = _read_zip_entry(zf, name)
                    ET.fromstring(content)
                except ET.ParseError as e:
                    issues.append(_err("xml", f"{name}: {e}"))
    except (zipfile.BadZipFile, FileNotFoundError):
        pass
    return issues


def validate_structure(path):
    """Verify document.xml has w:document root and w:body child."""
    issues = []
    ns = NAMESPACES["w"]
    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "word/document.xml" not in zf.namelist():
                return issues
            content = _read_zip_entry(zf, "word/document.xml")
            root = ET.fromstring(content)

            if not root.tag.endswith("}document") and root.tag != "document":
                issues.append(_err("structure",
                                   f"Root element is '{root.tag}', expected w:document"))

            body = root.find(f"{{{ns}}}body")
            if body is None:
                issues.append(_err("structure", "w:body element not found"))

    except ET.ParseError:
        pass  # Already reported by validate_xml
    except (zipfile.BadZipFile, FileNotFoundError):
        pass
    return issues


def validate_namespaces(path):
    """Detect auto-generated namespace prefixes (ns0:, ns1:, ...)."""
    issues = []
    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "word/document.xml" not in zf.namelist():
                return issues
            content = _read_zip_entry(zf, "word/document.xml")
            matches = _RE_AUTO_NS.findall(content)
            if matches:
                unique = sorted(set(matches))
                issues.append(_warn(
                    "namespace",
                    f"Auto-generated namespace prefixes found: {', '.join(unique)} "
                    f"({len(matches)} occurrences)"))
    except (zipfile.BadZipFile, FileNotFoundError):
        pass
    return issues


def validate_content(path, work_dir):
    """Verify that edit content appears in the output document.

    Reads edits.json and checks that new_text values are present
    (for INSERT/REPLACE) or absent (for DELETE) in the document text.
    """
    issues = []
    edits_path = os.path.join(work_dir, "edits.json")
    if not os.path.exists(edits_path):
        issues.append(_warn("content", "edits.json not found in work_dir"))
        return issues

    with open(edits_path, "r", encoding="utf-8") as f:
        edits = json.load(f).get("edits", [])

    try:
        with zipfile.ZipFile(path, "r") as zf:
            doc_xml = _read_zip_entry(zf, "word/document.xml")
    except (zipfile.BadZipFile, FileNotFoundError, KeyError):
        return issues

    full_text = _extract_all_text(doc_xml)

    # Also load analysis for DELETE verification
    analysis_path = os.path.join(work_dir, "analysis.json")
    analysis_blocks = {}
    if os.path.exists(analysis_path):
        with open(analysis_path, "r", encoding="utf-8") as f:
            analysis = json.load(f)
        for b in analysis.get("blocks", []):
            analysis_blocks[b["id"]] = b

    for i, edit in enumerate(edits):
        action = edit.get("action", "")
        tid = edit.get("target_id", "")
        new_text = edit.get("new_text", "")
        tag = edit.get("semantic_tag", "")

        if action in ("replace", "insert_after", "insert_before") and new_text:
            # For table edits, check individual cell values
            if tag == "TBL" and "|" in new_text:
                cells = [c.strip() for c in new_text.replace("\n", "|").split("|")]
                missing = [c for c in cells if c and c not in full_text]
                if missing:
                    issues.append(_warn(
                        "content",
                        f"Edit {i} ({tid}): {len(missing)}/{len(cells)} "
                        f"cell values not found in output"))
            else:
                # Check plain text — for multi-line, check each line
                lines = new_text.split("\n") if "\n" in new_text else [new_text]
                for line in lines:
                    clean = line.strip()
                    if clean and clean not in full_text:
                        issues.append(_warn(
                            "content",
                            f"Edit {i} ({tid}): text not found in output: "
                            f"'{clean[:60]}...'"))
                        break

        elif action == "delete":
            base_id = tid.split(":")[0]
            block = analysis_blocks.get(base_id)
            if block:
                old_text = block.get("text", "").strip()
                if old_text and old_text in full_text:
                    issues.append(_warn(
                        "content",
                        f"Edit {i} ({tid}): deleted text still found in output"))

    return issues


# -------------------------------------------------------------------
# Orchestrator
# -------------------------------------------------------------------

def validate(docx_path, work_dir=None):
    """Run all validations and return structured result."""
    errors = []
    warnings = []

    # Structural validations (stop early if ZIP/XML broken)
    structural_checks = [
        validate_zip,
        validate_entries,
        validate_xml,
        validate_structure,
        validate_namespaces,
    ]

    has_structural_error = False
    for check_fn in structural_checks:
        for issue in check_fn(docx_path):
            if issue["level"] == "error":
                errors.append(issue)
                has_structural_error = True
            else:
                warnings.append(issue)

    # Content validation (skip if structure is broken)
    if not has_structural_error and work_dir:
        for issue in validate_content(docx_path, work_dir):
            if issue["level"] == "error":
                errors.append(issue)
            else:
                warnings.append(issue)

    return {"valid": len(errors) == 0, "errors": errors, "warnings": warnings}


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 validate_docx.py <output_docx> [<work_dir>]",
              file=sys.stderr)
        sys.exit(1)

    docx_path = sys.argv[1]
    work_dir = sys.argv[2] if len(sys.argv) >= 3 else None

    result = validate(docx_path, work_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result["valid"] else 1)


if __name__ == "__main__":
    main()
