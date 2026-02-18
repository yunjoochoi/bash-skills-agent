#!/usr/bin/env python3
"""
Create a new DOCX file from a content JSON specification.

Uses ONLY Python stdlib (xml.etree.ElementTree, zipfile, json, os, sys).
No external libraries required.

Usage:
    python3 create_docx.py <content_json> <output_docx>

Content JSON format:
{
  "content": [
    {"type": "heading", "level": 1, "text": "Title"},
    {"type": "paragraph", "text": "Body text."},
    {"type": "bullet_list", "items": ["Item 1", "Item 2"]},
    {"type": "numbered_list", "items": ["Step 1", "Step 2"]},
    {"type": "table", "headers": ["Col1", "Col2"], "rows": [["a", "b"]]}
  ],
  "properties": {
    "title": "My Document",
    "font": "Calibri",
    "font_size": 22
  }
}
"""

import json
import os
import sys
import zipfile

# ---------------------------------------------------------------------------
# OOXML namespaces
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ---------------------------------------------------------------------------
# Embedded template XML strings
# ---------------------------------------------------------------------------

CONTENT_TYPES_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

DOCUMENT_RELS_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    Target="numbering.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>"""

SETTINGS_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
  <w:defaultTabStop w:val="720"/>
</w:settings>"""


def build_styles_xml(font="Calibri", font_size=22):
    """
    Build word/styles.xml with Normal, Heading1-6, ListBullet, ListNumber styles.

    Args:
        font: Font family name (default: Calibri).
        font_size: Font size in half-points for the Normal style (default: 22 = 11pt).

    Returns:
        Complete styles.xml content as a string.
    """
    # Heading sizes in half-points: Heading1=28, Heading2=26, ... Heading6=18
    heading_sizes = {1: 28, 2: 26, 3: 24, 4: 22, 5: 20, 6: 18}

    heading_styles = ""
    for level in range(1, 7):
        sz = heading_sizes[level]
        heading_styles += """
    <w:style w:type="paragraph" w:styleId="Heading{level}">
      <w:name w:val="heading {level}"/>
      <w:basedOn w:val="Normal"/>
      <w:next w:val="Normal"/>
      <w:qFormat/>
      <w:pPr>
        <w:keepNext/>
        <w:keepLines/>
        <w:spacing w:before="240" w:after="60"/>
        <w:outlineLvl w:val="{outline}"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>
        <w:b/>
        <w:bCs/>
        <w:sz w:val="{sz}"/>
        <w:szCs w:val="{sz}"/>
      </w:rPr>
    </w:style>""".format(level=level, outline=level - 1, font=xml_escape(font), sz=sz)

    styles_xml = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>
        <w:sz w:val="{font_size}"/>
        <w:szCs w:val="{font_size}"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>
      <w:sz w:val="{font_size}"/>
      <w:szCs w:val="{font_size}"/>
    </w:rPr>
  </w:style>
{heading_styles}

  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:numPr>
        <w:numId w:val="1"/>
      </w:numPr>
      <w:ind w:left="720" w:hanging="360"/>
    </w:pPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="ListNumber">
    <w:name w:val="List Number"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:numPr>
        <w:numId w:val="2"/>
      </w:numPr>
      <w:ind w:left="720" w:hanging="360"/>
    </w:pPr>
  </w:style>

  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>

  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:basedOn w:val="TableNormal"/>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      </w:tblBorders>
    </w:tblPr>
  </w:style>

</w:styles>""".format(
        font=xml_escape(font),
        font_size=font_size,
        heading_styles=heading_styles,
    )
    return styles_xml


NUMBERING_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Bullet list definition -->
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\u2022"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>

  <!-- Numbered list definition -->
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>

  <!-- Concrete numbering instances -->
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>

</w:numbering>"""


# ---------------------------------------------------------------------------
# XML escape helper
# ---------------------------------------------------------------------------

def xml_escape(text):
    """
    Escape special XML characters in text content.

    Handles: & < > " '
    """
    if not isinstance(text, str):
        text = str(text)
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    text = text.replace("'", "&apos;")
    return text


# ---------------------------------------------------------------------------
# Document XML builders
# ---------------------------------------------------------------------------

def build_heading(item):
    """Build XML for a heading paragraph."""
    level = item.get("level", 1)
    # Clamp level to 1-6
    level = max(1, min(6, int(level)))
    text = xml_escape(item.get("text", ""))
    return (
        '<w:p>'
        '<w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        '</w:p>'
    ).format(level=level, text=text)


def build_paragraph(item):
    """Build XML for a regular paragraph."""
    text = xml_escape(item.get("text", ""))
    return (
        '<w:p>'
        '<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        '</w:p>'
    ).format(text=text)


def build_bullet_list(item):
    """Build XML for a bullet list (one paragraph per item)."""
    items = item.get("items", [])
    paragraphs = []
    for list_item in items:
        text = xml_escape(str(list_item))
        paragraphs.append(
            '<w:p>'
            '<w:pPr>'
            '<w:pStyle w:val="ListBullet"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
            '</w:pPr>'
            '<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
            '</w:p>'.format(text=text)
        )
    return "\n".join(paragraphs)


def build_numbered_list(item):
    """Build XML for a numbered list (one paragraph per item)."""
    items = item.get("items", [])
    paragraphs = []
    for list_item in items:
        text = xml_escape(str(list_item))
        paragraphs.append(
            '<w:p>'
            '<w:pPr>'
            '<w:pStyle w:val="ListNumber"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'
            '</w:pPr>'
            '<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
            '</w:p>'.format(text=text)
        )
    return "\n".join(paragraphs)


def build_table_cell(text, bold=False):
    """Build XML for a single table cell."""
    escaped = xml_escape(str(text))
    bold_xml = "<w:b/><w:bCs/>" if bold else ""
    return (
        '<w:tc>'
        '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>'
        '<w:p>'
        '<w:r>'
        '<w:rPr>{bold_xml}</w:rPr>'
        '<w:t xml:space="preserve">{text}</w:t>'
        '</w:r>'
        '</w:p>'
        '</w:tc>'
    ).format(bold_xml=bold_xml, text=escaped)


def build_table(item):
    """Build XML for a table with optional header row and data rows."""
    headers = item.get("headers", [])
    rows = item.get("rows", [])

    # Determine column count
    num_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if num_cols == 0:
        return ""

    # Table grid columns
    grid_cols = "".join('<w:gridCol w:w="0"/>' for _ in range(num_cols))

    # Build header row if present
    header_row_xml = ""
    if headers:
        header_cells = "".join(build_table_cell(h, bold=True) for h in headers)
        header_row_xml = "<w:tr>{cells}</w:tr>".format(cells=header_cells)

    # Build data rows
    data_rows_xml = ""
    for row in rows:
        cells = "".join(build_table_cell(cell) for cell in row)
        data_rows_xml += "<w:tr>{cells}</w:tr>".format(cells=cells)

    return (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblStyle w:val="TableGrid"/>'
        '<w:tblW w:w="0" w:type="auto"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"'
        ' w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr>'
        '<w:tblGrid>{grid_cols}</w:tblGrid>'
        '{header_row}{data_rows}'
        '</w:tbl>'
        '<w:p/>'
    ).format(
        grid_cols=grid_cols,
        header_row=header_row_xml,
        data_rows=data_rows_xml,
    )


# Map of content type to builder function
CONTENT_BUILDERS = {
    "heading": build_heading,
    "paragraph": build_paragraph,
    "bullet_list": build_bullet_list,
    "numbered_list": build_numbered_list,
    "table": build_table,
}


def build_document_xml(content_items):
    """
    Build the complete word/document.xml from a list of content items.

    Args:
        content_items: List of dicts, each with a "type" key and type-specific fields.

    Returns:
        Complete document.xml content as a string.
    """
    body_parts = []

    for item in content_items:
        item_type = item.get("type", "")
        builder = CONTENT_BUILDERS.get(item_type)
        if builder is None:
            # Skip unknown content types with a warning to stderr
            print(
                "Warning: Unknown content type '{}', skipping.".format(item_type),
                file=sys.stderr,
            )
            continue
        xml_fragment = builder(item)
        if xml_fragment:
            body_parts.append(xml_fragment)

    body_content = "\n".join(body_parts)

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:body>'
        '{body}'
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '</w:sectPr>'
        '</w:body>'
        '</w:document>'
    ).format(body=body_content)

    return document_xml


# ---------------------------------------------------------------------------
# DOCX packaging
# ---------------------------------------------------------------------------

def create_docx(content_json_path, output_docx_path):
    """
    Create a DOCX file from a content JSON specification.

    Args:
        content_json_path: Path to the input JSON file.
        output_docx_path: Path for the output DOCX file.
    """
    # Read and parse the content JSON
    with open(content_json_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    content_items = spec.get("content", [])
    properties = spec.get("properties", {})

    # Extract configurable properties with defaults
    font = properties.get("font", "Calibri")
    font_size = properties.get("font_size", 22)

    # Ensure font_size is an integer
    font_size = int(font_size)

    # Build all the XML parts
    document_xml = build_document_xml(content_items)
    styles_xml = build_styles_xml(font=font, font_size=font_size)

    # Ensure output directory exists
    output_dir = os.path.dirname(output_docx_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # Package everything into a ZIP file with .docx extension
    with zipfile.ZipFile(output_docx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", RELS_XML)
        zf.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_XML)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/numbering.xml", NUMBERING_XML)
        zf.writestr("word/settings.xml", SETTINGS_XML)

    # Report success
    file_size = os.path.getsize(output_docx_path)
    print("Created {} ({} bytes)".format(output_docx_path, file_size))


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    """CLI entry point: parse arguments and create the DOCX file."""
    if len(sys.argv) != 3:
        print(
            "Usage: python3 {} <content_json> <output_docx>".format(sys.argv[0]),
            file=sys.stderr,
        )
        sys.exit(1)

    content_json_path = sys.argv[1]
    output_docx_path = sys.argv[2]

    # Validate input file exists
    if not os.path.isfile(content_json_path):
        print(
            "Error: Input file not found: {}".format(content_json_path),
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        create_docx(content_json_path, output_docx_path)
    except json.JSONDecodeError as e:
        print("Error: Invalid JSON in input file: {}".format(e), file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print("Error: {}".format(e), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
