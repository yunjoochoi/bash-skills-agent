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
[b3:TBL|T1]
  [b3:r0|RS0]
    [b3:r0c0|CS0] [p0|S3] Header1
    [b3:r0c1|CS1] [p0|S3] Header2
  [b3:r1|RS1]
    [b3:r1c0|CS0] [p0|S4] Data1
    [b3:r1c1|CS1] [p0|S4] Data2
[b4:TOC|toc_default]
  [b4:p0|TL0] 1. Introduction | 3
  [b4:p1|TL1] 1-1. Background | 4
```

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
bash(command="python3 /skills/edit-docx/scripts/validate_edits.py /workspace/docx_work")
```

Validates edits.json against analysis.json. Outputs JSON with errors/warnings.
- If `"valid": false` — fix edits.json based on error messages and re-validate.
- If `"valid": true` — proceed to Step 4.

### Step 3.5: Run Distribution (if needed)

```
bash(command="python3 /skills/edit-docx/scripts/generate_run_prompts.py /workspace/docx_work")
```

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

This creates the final edited DOCX. Add `--update-fields` flag if TOC was modified.

### Step 6: Validate Output DOCX

```
bash(command="python3 /skills/edit-docx/scripts/validate_docx.py /workspace/<output.docx> /workspace/docx_work")
```

Validates the output DOCX structure and content. Outputs JSON with errors/warnings.
- If `"valid": false` — review errors, fix edits.json, re-run from Step 3.
- If `"valid": true` with warnings only — review warnings, proceed if acceptable.

---

## B. Edit Plan JSON Format

### Three Edit Types

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

#### 3. TOC Edit (semantic_tag: TOC)

```json
{
  "action": "replace",
  "target_id": "b5:p1",
  "semantic_tag": "TOC",
  "new_text": "1-1. New Title | 4",
  "toc_level_alias": "TL1",
  "anchor_block_id": "b10",
  "reasoning": "Update TOC entry for renamed section"
}
```

Fields: `action`, `target_id`, `semantic_tag`, `new_text`, `toc_level_alias`, `anchor_block_id`, `reasoning`

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
- TBL → Table edit, TOC → TOC edit, everything else → Paragraph edit
- Valid tags: H1, H2, H3, BODY, LIST, TBL, TITLE, SUBTITLE, TOC, OTHER

### Paragraph Rules
- **style_alias is REQUIRED** for REPLACE and INSERT
- **REPLACE**: Use SAME style_alias as the original block
- **INSERT**: Choose appropriate alias from text_merge (heading→S1, body→S2, etc.)
- **DELETE**: No style_alias needed

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

### TOC Rules

All TOC edits operate at the **entry (row) level** using `b5:pN` target_id format.

- `toc_level_alias`: TL alias for level (TL0=level1, TL1=level2, TL2=level3)
- `anchor_block_id`: **REQUIRED** for INSERT and REPLACE — body heading block ID for bookmark linking

**TOC Consistency Rules:**
- TOC INSERT **MUST** accompany a body heading INSERT with matching anchor_block_id
- TOC DELETE **MUST** accompany a body heading DELETE
- TOC REPLACE is allowed only when the corresponding body heading text also changes
- NEVER add/remove TOC entries without corresponding body heading changes

Page numbers are placeholders. Word auto-updates on open (Ctrl+A, F9).

---

## D. Creating New DOCX

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

## E. Final Checklist

Before writing edits.json, verify:
1. Every edit has `semantic_tag`
2. Paragraph REPLACE/INSERT has `style_alias`
3. Table edits have `edit_unit`
4. Table row INSERT has `row_style_aliases` + `cell_style_aliases`
5. Table INSERT has `table_style_alias`
6. No `\n` in paragraph `new_text`
7. `new_text` is COMPLETE content
8. TOC INSERT/REPLACE has `anchor_block_id`
9. TOC entry `target_id` uses `bN:pN` format
10. Run `validate_edits.py` before `apply_edits.py`
11. Run `generate_run_prompts.py` and process prompts before `apply_edits.py`
12. Run `validate_docx.py` after `repack_docx.py`
