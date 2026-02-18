#!/usr/bin/env python3
"""Repackage modified XML files back into a DOCX file.

Ported from DocxPackager (docx-agent). Uses only Python stdlib.

Usage:
    python3 repack_docx.py <original_docx> <extracted_dir> <output_docx> [--update-fields]

Arguments:
    original_docx   Path to the original DOCX file (used as ZIP template)
    extracted_dir   Path to the directory containing modified XML/RELS files
    output_docx     Path for the output DOCX file
    --update-fields Inject <w:updateFields w:val="true"/> into settings.xml
                    so Word updates TOC/fields on open
"""

import os
import re
import sys
import zipfile

# Required files that every valid DOCX must contain
REQUIRED_DOCX_FILES = [
    "[Content_Types].xml",
    "word/document.xml",
    "_rels/.rels",
    "word/_rels/document.xml.rels",
]


def package_to_docx(
    original_docx_path: str,
    extracted_path: str,
    output_path: str,
) -> str:
    """Package modified XMLs back into DOCX format.

    Preserves the original ZIP entry order and metadata. For each entry in
    the original DOCX, if a matching .xml or .rels file exists in
    extracted_path, that file replaces the original content. All other
    entries (media, embeddings, etc.) are copied byte-identical.

    Args:
        original_docx_path: Path to the original DOCX file.
        extracted_path: Path to directory with modified XML/RELS files.
        output_path: Desired output DOCX file path.

    Returns:
        The output_path on success.

    Raises:
        FileNotFoundError: If original_docx_path or extracted_path is missing.
    """
    if not os.path.exists(original_docx_path):
        raise FileNotFoundError(f"Original DOCX not found: {original_docx_path}")
    if not os.path.isdir(extracted_path):
        raise FileNotFoundError(f"Extracted directory not found: {extracted_path}")

    # Create output directory if it does not exist
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # Collect the set of modified XML/RELS files from the extracted directory
    modified_files: set[str] = set()
    for root, _, files in os.walk(extracted_path):
        for filename in files:
            if filename.endswith((".xml", ".rels")):
                rel_path = os.path.relpath(
                    os.path.join(root, filename), extracted_path
                )
                modified_files.add(rel_path)

    # Rebuild the DOCX: iterate original entries in order, replacing
    # modified files while keeping everything else byte-identical.
    with zipfile.ZipFile(output_path, "w") as output_zip:
        with zipfile.ZipFile(original_docx_path, "r") as original_zip:
            for info in original_zip.infolist():
                if info.filename in modified_files:
                    # Replace with the modified version, preserving ZipInfo
                    full_path = os.path.join(extracted_path, info.filename)
                    if os.path.exists(full_path):
                        with open(full_path, "rb") as f:
                            output_zip.writestr(info, f.read())
                    else:
                        # File listed but missing on disk; keep original
                        output_zip.writestr(
                            info, original_zip.read(info.filename)
                        )
                else:
                    # Copy original entry as-is (preserves ZipInfo metadata)
                    output_zip.writestr(
                        info, original_zip.read(info.filename)
                    )

    return output_path


def validate_docx(docx_path: str) -> tuple[bool, list[str]]:
    """Validate the structure of a generated DOCX file.

    Checks that all required entries exist and that document.xml contains
    the expected root elements.

    Args:
        docx_path: Path to the DOCX file to validate.

    Returns:
        A tuple of (is_valid, error_messages). is_valid is True when
        error_messages is empty.
    """
    errors: list[str] = []

    if not os.path.exists(docx_path):
        return False, ["File not found"]

    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            file_list = z.namelist()

            # Check required files
            for required in REQUIRED_DOCX_FILES:
                if required not in file_list:
                    errors.append(f"Missing required file: {required}")

            # Basic content validation of document.xml
            if "word/document.xml" in file_list:
                try:
                    content = z.read("word/document.xml")
                    if b"<w:document" not in content:
                        errors.append(
                            "document.xml missing w:document element"
                        )
                    if b"<w:body" not in content:
                        errors.append("document.xml missing w:body element")
                except Exception as e:
                    errors.append(f"Cannot read document.xml: {e}")

    except zipfile.BadZipFile:
        return False, ["Invalid ZIP/DOCX format"]
    except Exception as e:
        return False, [f"Validation error: {e}"]

    return len(errors) == 0, errors


def inject_update_fields(extracted_path: str) -> None:
    """Inject <w:updateFields w:val="true"/> into word/settings.xml.

    Uses string-level manipulation instead of ElementTree to avoid
    namespace rewriting that can break strict OOXML parsers.

    If settings.xml already contains an updateFields element its value is
    flipped to "true". Otherwise the element is inserted before the
    closing </w:settings> tag. If settings.xml does not exist at all a
    minimal one is created.

    Args:
        extracted_path: Path to the extracted XMLs directory.
    """
    settings_path = os.path.join(extracted_path, "word", "settings.xml")
    update_tag = '<w:updateFields w:val="true"/>'

    if os.path.exists(settings_path):
        with open(settings_path, encoding="utf-8") as f:
            content = f.read()

        if "w:updateFields" in content:
            # Flip existing val to "true"
            content = re.sub(
                r'(<w:updateFields[^/]*w:val=")[^"]*(")',
                r"\g<1>true\2",
                content,
            )
            print("Updated existing updateFields to true in settings.xml")
        elif "</w:settings>" in content:
            # Insert before closing tag
            content = content.replace(
                "</w:settings>",
                f"{update_tag}</w:settings>",
            )
            print("Injected updateFields into settings.xml")
        else:
            # Self-closing <w:settings .../> -> open + inject + close
            content = re.sub(
                r"(<w:settings\b[^>]*?)\s*/>",
                rf"\1>{update_tag}</w:settings>",
                content,
            )
            print("Injected updateFields into self-closing settings.xml")

        with open(settings_path, "w", encoding="utf-8") as f:
            f.write(content)
    else:
        # Create a minimal settings.xml with updateFields
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        os.makedirs(os.path.dirname(settings_path), exist_ok=True)
        minimal = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:settings xmlns:w="{ns}">'
            f"{update_tag}</w:settings>"
        )
        with open(settings_path, "w", encoding="utf-8") as f:
            f.write(minimal)
        print("Created settings.xml with updateFields")


def main() -> int:
    """CLI entry point.

    Returns:
        Exit code: 0 on success, 1 on failure.
    """
    # Parse arguments
    args = sys.argv[1:]

    update_fields = False
    if "--update-fields" in args:
        update_fields = True
        args.remove("--update-fields")

    if len(args) != 3:
        print(
            "Usage: python3 repack_docx.py <original_docx> <extracted_dir> "
            "<output_docx> [--update-fields]",
            file=sys.stderr,
        )
        return 1

    original_docx, extracted_dir, output_docx = args

    # Validate inputs exist
    if not os.path.isfile(original_docx):
        print(f"Error: Original DOCX not found: {original_docx}", file=sys.stderr)
        return 1

    if not os.path.isdir(extracted_dir):
        print(
            f"Error: Extracted directory not found: {extracted_dir}",
            file=sys.stderr,
        )
        return 1

    try:
        # Optionally inject updateFields for TOC refresh
        if update_fields:
            inject_update_fields(extracted_dir)

        # Package the DOCX
        result_path = package_to_docx(original_docx, extracted_dir, output_docx)
        print(f"DOCX created: {result_path}")

        # Validate the result
        is_valid, errors = validate_docx(result_path)
        if is_valid:
            file_size = os.path.getsize(result_path)
            print(f"Validation passed ({file_size} bytes)")
        else:
            print("Validation warnings:", file=sys.stderr)
            for error in errors:
                print(f"  - {error}", file=sys.stderr)

        # Report final status
        if is_valid:
            print("SUCCESS")
            return 0
        else:
            print("FAILED: validation errors detected", file=sys.stderr)
            return 1

    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except zipfile.BadZipFile:
        print(
            f"Error: {original_docx} is not a valid ZIP/DOCX file",
            file=sys.stderr,
        )
        return 1
    except Exception as e:
        print(f"Error: Unexpected failure: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
