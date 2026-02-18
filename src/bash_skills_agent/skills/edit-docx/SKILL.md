---
name: edit-docx
description: Edit existing DOCX documents (preserving styles) or create new ones. Use when the user asks to modify, update, or create Word documents.
---

# DOCX Document Editing Skill

## When to Use
- User asks to edit, modify, or update a DOCX/Word document
- User asks to create a new DOCX document
- User uploads a DOCX and wants content changes applied

## A. Editing Workflow (Modify Existing DOCX)

**IMPORTANT: Execute steps IN ORDER. Each step depends on the previous step's output.**

### Step 1: Analyze Document

```
bash(command="python3 /skills/edit-docx/scripts/analyze_docx.py /workspace/<input.docx> /workspace/docx_work")
```

This outputs:
- **stdout**: `text_merge` — document content with block IDs and style aliases
- **`/workspace/docx_work/analysis.json`**: full analysis (blocks, templates, aliases)
- **`/workspace/docx_work/extracted/`**: extracted XML files

### Step 2: Read text_merge & Plan Edits

Read the text_merge output from Step 1. It shows document structure like:

```
[b0:H1|S1] Chapter 1: Introduction
[b1:BODY|S2] This is the first paragraph of body text.
[b2:BODY|S2] Second paragraph with details.
[b3:BODY|S5|numPr:"1-1."] Main requirements
[b4:BODY|S6|numPr:"•"] First bullet item
[b5:TBL|T1]
  [b5:r0|RS0]
    [b5:r0c0|CS0] [p0|S3] Header1
    [b5:r0c1|CS1] [p0|S3] Header2
  [b5:r1|RS1]
    [b5:r1c0|CS0] [p0|S4] Data1
    [b5:r1c1|CS1] [p0|S4] Data2
[b6:TOC|toc_default]
  [b6:p0] 1. Introduction 3
  [b6:p1] 1-1. Background 4
```

- `|numPr:"1-1."` = auto-numbering with computed prefix (Word renders "1-1." before the text)
- `|numPr:"•"` = auto bullet
- `|numPr` (no value) = auto-numbering detected but prefix could not be computed

Based on user's request, write an edit plan JSON to `/workspace/docx_work/edits.json`:

```json
{
  "edits": [
    {"action": "replace", "target_id": "b1", "semantic_tag": "BODY", "new_text": "Updated first paragraph.", "style_alias": "S2"}
  ]
}
```

### Step 3: Validate Edit Plan

```
bash(command="python3 /skills/edit-docx/scripts/validation/validate_edits.py /workspace/docx_work")
```

Validates edits.json against analysis.json. Outputs JSON with errors/warnings.
- If `"valid": false` — fix ONLY the failing edits in edits.json (use `edit_file`, do NOT rewrite the entire file) and re-validate.
- If `"valid": true` — proceed to Step 4.

### Step 3.5: Run Distribution (if needed) — REQUIRES edits.json from Step 2

```
bash(command="python3 /skills/edit-docx/scripts/generate_run_prompts.py /workspace/docx_work")
```

**Prerequisites:** edits.json must exist (from Step 2) and pass validation (Step 3).

If the output has prompts (`"prompts": [...]` non-empty), process each:
1. Read the `prompt` text — it shows run styles (RS0, RS1...) with descriptions
   and the original text distribution (for REPLACE)
2. Distribute the new text across the styles following the pattern
3. Update edits.json: add `"runs"` AND `"run_style_templates"` from the prompt output

```json
{
  "runs": [{"text": "...", "run_style": "RS0"}, {"text": "...", "run_style": "RS1"}],
  "run_style_templates": {"RS0": {"rpr_xml": "..."}, "RS1": {"rpr_xml": "..."}}
}
```

- Copy `run_style_templates` from the prompt output as-is (contains XML templates)
- The concatenation of all `runs[].text` must equal `new_text`
- If prompts is empty → no edits need run distribution, skip to Step 4.

### Step 4: Apply Edits

```
bash(command="python3 /skills/edit-docx/scripts/apply_edits.py /workspace/docx_work")
```

This reads `analysis.json` + `edits.json` and modifies `document.xml`.

### Step 5: Repack DOCX

```
bash(command="python3 /skills/edit-docx/scripts/repack_docx.py /workspace/<input.docx> /workspace/docx_work/extracted /workspace/<output.docx>")
```

This creates the final edited DOCX. Do NOT use `--update-fields` if TOC was manually edited (see Section D).

### Step 6: Validate Output DOCX

```
bash(command="python3 /skills/edit-docx/scripts/validation/validate_docx.py /workspace/<output.docx> /workspace/docx_work")
```

Validates the output DOCX structure and content. Outputs JSON with errors/warnings.
- If `"valid": false` — review errors, fix edits.json, re-run from Step 3.
- If `"valid": true` with warnings only — review warnings, proceed if acceptable.

### Step 7: Re-analyze & Verify

```
bash(command="python3 /skills/edit-docx/scripts/analyze_docx.py /workspace/<output.docx> /workspace/docx_work_verify")
```

Re-run analyze on the OUTPUT DOCX and review the text_merge:
1. Check that all user-requested changes are reflected in the new text_merge
2. Verify numbering is correct (no duplication, correct sequence)
3. Verify TOC entries match the body headings (if TOC was edited)
4. If discrepancies found — identify the cause, fix, and re-run from the appropriate step

---

## B. Edit Plan JSON Format

### Two Edit Types

`semantic_tag` determines the edit type:

#### 1. Paragraph Edit (semantic_tag: H1, H2, H3, BODY, LIST, TITLE, SUBTITLE, OTHER)

```json
{
  "action": "replace",
  "target_id": "b5",
  "semantic_tag": "BODY",
  "new_text": "Updated content",
  "style_alias": "S2",
  "reasoning": "User requested text change"
}
```

Fields: `action`, `target_id`, `semantic_tag`, `new_text`, `style_alias`, `reasoning`

#### 2. Table Edit (semantic_tag: TBL)

```json
{
  "action": "insert_after",
  "target_id": "b13:r2",
  "semantic_tag": "TBL",
  "edit_unit": "row",
  "new_text": "col1 | col2 | col3",
  "row_style_aliases": ["RS0"],
  "cell_style_aliases": [["CS2", "CS3", "CS2"]],
  "reasoning": "Add new data row"
}
```

Fields: `action`, `target_id`, `semantic_tag`, `edit_unit`, `new_text`, `table_style_alias`, `row_style_aliases`, `cell_style_aliases`, `reasoning`

### Action Types

| Action | Value | Description |
|--------|-------|-------------|
| Replace | `"replace"` | Modify existing block content |
| Insert After | `"insert_after"` | Add new block AFTER target |
| Insert Before | `"insert_before"` | Add new block BEFORE target |
| Delete | `"delete"` | Remove existing block |

---

## C. Critical Rules

### semantic_tag is REQUIRED for ALL edits
- Copy TAG from text_merge: `[b5:BODY|S2]` → `"semantic_tag": "BODY"`
- TBL → Table edit, everything else → Paragraph edit
- Valid tags: H1, H2, H3, BODY, LIST, TBL, TITLE, SUBTITLE, OTHER
- TOC blocks are NOT edited via edits.json — see Section D for TOC editing

### Paragraph Rules
- **style_alias is REQUIRED** for REPLACE and INSERT
- **REPLACE**: Use SAME style_alias as the original block
- **INSERT**: Choose appropriate alias from text_merge (heading→S1, body→S2, etc.)
- **DELETE**: No style_alias needed

### Auto-Numbering (`|numPr`)
- `|numPr:"1-5-1."` = Word auto-generates "1-5-1." before the text. The computed prefix is shown in quotes.
- **DO NOT include number prefixes in `new_text`** — Word adds them automatically
- Use the computed prefix when creating TOC entries for these blocks

### One Block = One Edit
- NEVER combine multiple styled elements into one edit
- NEVER use `\n` in paragraph `new_text` (split into separate INSERT edits)
- `new_text` must be COMPLETE content (not a summary!)

### Multi-paragraph Insert Example
To insert a heading + body text after b10:
```json
{
  "edits": [
    {"action": "insert_after", "target_id": "b10", "semantic_tag": "H1", "new_text": "New Section", "style_alias": "S1"},
    {"action": "insert_after", "target_id": "b10", "semantic_tag": "BODY", "new_text": "Body paragraph text.", "style_alias": "S2"}
  ]
}
```

### Table Operations

#### edit_unit Values

| edit_unit | target_id | Description | Required fields |
|-----------|-----------|-------------|-----------------|
| `"cell"` | `b13:r0c0p0` | Edit single cell paragraph | `cell_style_aliases` |
| `"row"` | `b13:r2` | Insert/delete entire row | `row_style_aliases`, `cell_style_aliases` |
| `"column"` | `b13:c1` | Insert/delete entire column | `cell_style_aliases` (for INSERT) |
| `"table"` | `b13` | Insert new table after target | `table_style_alias`, `row_style_aliases`, `cell_style_aliases` |

#### Cell REPLACE
```json
{"action": "replace", "target_id": "b13:r0c0p0", "semantic_tag": "TBL", "edit_unit": "cell", "new_text": "New cell text", "cell_style_aliases": [["CS2"]]}
```

#### Row INSERT
```json
{"action": "insert_after", "target_id": "b13:r2", "semantic_tag": "TBL", "edit_unit": "row", "new_text": "col1 | col2 | col3", "row_style_aliases": ["RS0"], "cell_style_aliases": [["CS2", "CS3", "CS2"]]}
```
- Use `|` to separate cells. Match column count from existing rows.
- Copy RS alias from a similar row, CS aliases from similar cells.

#### Column INSERT
```json
{"action": "insert_after", "target_id": "b13:c0", "semantic_tag": "TBL", "edit_unit": "column", "new_text": "No\n1\n2\n3", "cell_style_aliases": ["CS2", "CS2", "CS2", "CS2"]}
```
- Cell values separated by `\n` (one per row, top to bottom).
- `cell_style_aliases`: Flat list (one CS alias per row). NOT nested.

#### Row/Column DELETE
```json
{"action": "delete", "target_id": "b13:r5", "semantic_tag": "TBL", "edit_unit": "row"}
{"action": "delete", "target_id": "b13:c1", "semantic_tag": "TBL", "edit_unit": "column"}
```

---

## D. TOC (Table of Contents) Editing

TOC editing is done via **direct XML manipulation** using bash, NOT through edits.json.

### TOC Structure in DOCX

The TOC is an `<w:sdt>` block containing `<w:sdtContent>` with paragraphs. Each entry links to a body heading via bookmarks:

**TOC entry** (inside `<w:sdtContent>`):
```xml
<w:hyperlink w:anchor="_Toc12345">
  <w:r><w:t>1. Introduction</w:t></w:r>
</w:hyperlink>
<!-- PAGEREF field code for page number -->
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> PAGEREF _Toc12345 \h </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>3</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

**Body heading** (bookmark target):
```xml
<w:bookmarkStart w:id="0" w:name="_Toc12345"/>
<w:r><w:t>1. Introduction</w:t></w:r>
<w:bookmarkEnd w:id="0"/>
```

### How to Edit TOC

Use `lxml` for TOC XML manipulation. lxml preserves namespace prefixes and handles OOXML reliably.

1. Read the extracted `document.xml` and parse with `lxml.etree`
2. Find the TOC `<w:sdt>` block and the target heading's bookmark
3. Edit operations:
   - **Add entry**: Copy an existing TOC paragraph, change anchor name + text + PAGEREF
   - **Modify entry**: Update text in `<w:t>` and anchor in `<w:hyperlink>` + PAGEREF
   - **Delete entry**: Remove the `<w:p>` from `<w:sdtContent>`
   - **Add bookmark**: Insert `<w:bookmarkStart>` / `<w:bookmarkEnd>` around body heading
4. Save with `tree.write()`

### Key Rules
- TOC `<w:hyperlink w:anchor="name">` must match body `<w:bookmarkStart w:name="name">`
- PAGEREF instrText must reference the same bookmark name
- Bookmark IDs must be unique across the document
- **DO NOT use `--update-fields` flag** when TOC was manually edited — it causes Word to rebuild the entire TOC from heading styles, wiping out manual XML entries
- Page numbers in manually added entries are static placeholders

---

## E. Creating New DOCX

### Step 1: Write Content JSON

Write content specification to a JSON file:

```json
{
  "content": [
    {"type": "heading", "level": 1, "text": "Document Title"},
    {"type": "paragraph", "text": "Regular body text paragraph."},
    {"type": "heading", "level": 2, "text": "Section Header"},
    {"type": "paragraph", "text": "More text here."},
    {"type": "bullet_list", "items": ["First item", "Second item", "Third item"]},
    {"type": "numbered_list", "items": ["Step one", "Step two"]},
    {"type": "table", "headers": ["Name", "Value"], "rows": [["A", "1"], ["B", "2"]]}
  ]
}
```

### Step 2: Generate DOCX

```
bash(command="python3 /skills/edit-docx/scripts/create_docx.py /workspace/content.json /workspace/new_document.docx")
```

### Supported Element Types
- `heading`: level 1-6
- `paragraph`: plain body text
- `bullet_list`: unordered list with items array
- `numbered_list`: ordered list with items array
- `table`: headers + rows arrays

---

## F. Final Checklist

Before writing edits.json, verify:
1. Every edit has `semantic_tag`
2. Paragraph REPLACE/INSERT has `style_alias`
3. Table edits have `edit_unit`
4. Table row INSERT has `row_style_aliases` + `cell_style_aliases`
5. Table INSERT has `table_style_alias`
6. No `\n` in paragraph `new_text`
7. `new_text` is COMPLETE content
8. TOC edits are NOT in edits.json (handle via direct XML — see Section D)

Execution order (MUST follow):
9. Write `edits.json` (Step 2)
11. Run `validation/validate_edits.py` (Step 3) — fix errors and re-validate until valid
12. Run `generate_run_prompts.py` AFTER edits.json is finalized (Step 3.5) — process prompts if any
13. Run `apply_edits.py` (Step 4)
14. Edit TOC XML directly if needed (Section D)
15. Run `repack_docx.py` (Step 5) — **without** `--update-fields` if TOC was manually edited
16. Run `validation/validate_docx.py` (Step 6) — verify output DOCX
17. Re-analyze output DOCX (Step 7) — verify text_merge matches user request
