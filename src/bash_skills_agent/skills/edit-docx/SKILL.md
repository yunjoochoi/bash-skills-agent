---
name: edit-docx
description: Edit existing DOCX documents (preserving styles) or create new ones. Use when the user asks to modify, update, or create Word documents.
---

# DOCX Document Editing Skill

## When to Use
- User asks to edit, modify, or update a DOCX/Word document
- User asks to create a new DOCX document
- User uploads a DOCX and wants content changes applied

---

## A. Editing Workflow

**Execute steps IN ORDER. Each step depends on the previous step's output.**

### Step 1: Analyze Document

```
bash(command="python3 /skills/edit-docx/scripts/analyze_docx.py /workspace/<input.docx> /workspace/docx_work")
```

Outputs:
- **stdout**: `text_merge` — document content with block IDs and style aliases (see Section B for format)
- **`/workspace/docx_work/analysis.json`**: full analysis (blocks, templates, aliases)
- **`/workspace/docx_work/extracted/`**: extracted XML files

### Step 2: Plan Edits

Read the text_merge from Step 1 and write edits to `/workspace/docx_work/edits.json`. See Section B for JSON format and Section C for rules.

### Step 3: Validate

```
bash(command="python3 /skills/edit-docx/scripts/validation/validate_edits.py /workspace/docx_work")
```

- `"valid": false` → fix ONLY the failing edits (use `edit_file`, do NOT rewrite the entire file), re-validate
- `"valid": true` → proceed

### Step 4: Run Distribution (if needed)

```
bash(command="python3 /skills/edit-docx/scripts/generate_run_prompts.py /workspace/docx_work")
```

**Prerequisite:** edits.json must pass validation (Step 3).

If `"prompts": [...]` is non-empty, process each prompt:
1. Read the `prompt` — shows run styles (RS0, RS1...) and original text distribution
2. Distribute the new text across the styles following the pattern
3. Update edits.json: add `"runs"` AND `"run_style_templates"` from the prompt output:
```json
{
  "runs": [{"text": "...", "run_style": "RS0"}, {"text": "...", "run_style": "RS1"}],
  "run_style_templates": {"RS0": {"rpr_xml": "..."}, "RS1": {"rpr_xml": "..."}}
}
```
- Copy `run_style_templates` from prompt output as-is (XML templates)
- Concatenation of all `runs[].text` must equal `new_text`
- Empty prompts → skip to Step 5

### Step 5: Apply Edits

```
bash(command="python3 /skills/edit-docx/scripts/apply_edits.py /workspace/docx_work")
```

Reads `analysis.json` + `edits.json`, modifies `document.xml`.

If heading RENAME only (no INSERT/DELETE) AND document has TOC → edit TOC XML now (Section D), before repacking.

### Step 6: Repack DOCX

```
bash(command="python3 /skills/edit-docx/scripts/repack_docx.py /workspace/<input.docx> /workspace/docx_work/extracted /workspace/<output.docx>")
```

**Do NOT use `--update-fields`** if TOC was manually edited — it rebuilds the entire TOC.

### Step 7: Validate & Verify

```
bash(command="python3 /skills/edit-docx/scripts/validation/validate_docx.py /workspace/<output.docx> /workspace/docx_work")
bash(command="python3 /skills/edit-docx/scripts/analyze_docx.py /workspace/<output.docx> /workspace/docx_work_verify")
```

Run both: structural validation + re-analysis. Check the verification text_merge against every edit:

- **INSERT**: New block exists as separate block AND original target still present
- **REPLACE**: Target block contains new_text, block count unchanged
- **DELETE**: Target block absent
- **Numbering**: If numPr heading inserted/deleted, compare before/after prefixes for ALL subsequent headings
- **TOC**: Each entry text + bookmark must match corresponding body heading with correct numPr values

If corrections needed → use OUTPUT as new input, repeat from Step 1.

### Step 8: Second Pass (heading INSERT/DELETE + TOC)

**Only if heading INSERT/DELETE AND document has TOC.**

Heading INSERT/DELETE shifts auto-numbering for ALL subsequent headings. You cannot know final numPr values until after re-analysis.

1. Steps 1–7 complete with body edits only (TOC skipped)
2. Use the output DOCX as new input → re-analyze (read NEW numPr values)
3. Edit TOC XML using re-analyzed values (Section D)
4. Update ALL TOC entries whose numbering shifted — not just new entries
5. Repack (without `--update-fields`) and re-verify

---

## B. Edit Plan Reference

### Reading text_merge

```
[b0:H1|S1] Chapter 1: Introduction
[b1:BODY|S2] This is the first paragraph of body text.
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

- `|numPr:"1-1."` — auto-numbering: Word renders "1-1." before the text
- `|numPr:"•"` — auto bullet
- `|numPr` (no value) — auto-numbering detected but prefix could not be computed

### Action Types

| Action | Value | Description |
|--------|-------|-------------|
| Replace | `"replace"` | Modify existing block content |
| Insert After | `"insert_after"` | Add new block AFTER target |
| Insert Before | `"insert_before"` | Add new block BEFORE target |
| Delete | `"delete"` | Remove existing block |

### Paragraph Edit

semantic_tag: H1, H2, H3, BODY, LIST, TITLE, SUBTITLE, OTHER

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

Multi-paragraph insert (heading + body after b10):
```json
{
  "edits": [
    {"action": "insert_after", "target_id": "b10", "semantic_tag": "H1", "new_text": "New Section", "style_alias": "S1"},
    {"action": "insert_after", "target_id": "b10", "semantic_tag": "BODY", "new_text": "Body paragraph text.", "style_alias": "S2"}
  ]
}
```

### Table Edit

semantic_tag: TBL

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

#### edit_unit Reference

| edit_unit | target_id | Description | Required fields |
|-----------|-----------|-------------|-----------------|
| `"cell"` | `b13:r0c0p0` | Edit single cell paragraph | `cell_style_aliases` |
| `"row"` | `b13:r2` | Insert/delete entire row | `row_style_aliases`, `cell_style_aliases` |
| `"column"` | `b13:c1` | Insert/delete entire column | `cell_style_aliases` (for INSERT) |
| `"table"` | `b13` | Insert new table after target | `table_style_alias`, `row_style_aliases`, `cell_style_aliases` |

#### Examples

**Cell REPLACE:**
```json
{"action": "replace", "target_id": "b13:r0c0p0", "semantic_tag": "TBL", "edit_unit": "cell", "new_text": "New cell text", "cell_style_aliases": [["CS2"]]}
```

**Row INSERT:**
```json
{"action": "insert_after", "target_id": "b13:r2", "semantic_tag": "TBL", "edit_unit": "row", "new_text": "col1 | col2 | col3", "row_style_aliases": ["RS0"], "cell_style_aliases": [["CS2", "CS3", "CS2"]]}
```
- `|` separates cells. Match column count from existing rows.
- Copy RS alias from similar row, CS aliases from similar cells.

**Column INSERT:**
```json
{"action": "insert_after", "target_id": "b13:c0", "semantic_tag": "TBL", "edit_unit": "column", "new_text": "No\n1\n2\n3", "cell_style_aliases": ["CS2", "CS2", "CS2", "CS2"]}
```
- `\n` separates cell values (one per row, top to bottom)
- `cell_style_aliases`: **flat list** (one CS alias per row), NOT nested

**Row/Column DELETE:**
```json
{"action": "delete", "target_id": "b13:r5", "semantic_tag": "TBL", "edit_unit": "row"}
{"action": "delete", "target_id": "b13:c1", "semantic_tag": "TBL", "edit_unit": "column"}
```

---

## C. Rules

### General
- **semantic_tag is REQUIRED** for all edits. Copy TAG from text_merge: `[b5:BODY|S2]` → `"semantic_tag": "BODY"`
  - TBL → Table edit, everything else → Paragraph edit
  - Valid: H1, H2, H3, BODY, LIST, TBL, TITLE, SUBTITLE, OTHER
  - TOC blocks are NOT edited via edits.json (see Section D)
- **style_alias**: REQUIRED for REPLACE and INSERT. REPLACE uses same alias as original. INSERT uses appropriate alias from text_merge.
- **One block = one edit**: Never combine multiple styled elements. Never use `\n` in paragraph `new_text` (split into separate INSERT edits).
- **new_text** must be COMPLETE content (not a summary)
- **Prefer `edit_file` over scripts** for XML modifications. Only write a script when `edit_file` cannot handle the change.
- **`grep_search` in XML**: Tags split text — use short keywords, not full phrases. Then `read_file` with offset/limit for context.

### Auto-Numbering
- `|numPr:"1-5-1."` = Word auto-generates "1-5-1." before the text
- **DO NOT include number prefixes in `new_text`** — Word adds them automatically
- Use the computed prefix when creating TOC entries for these blocks
- **Numbering cascade**: Inserting/deleting headings with numPr shifts auto-numbering for ALL subsequent same-level headings

### Table-Specific
- Row `cell_style_aliases`: **nested** list `[["CS0", "CS1"]]`
- Column `cell_style_aliases`: **flat** list `["CS0", "CS0", "CS0"]`
- Row cell count (pipe-separated in `new_text`) must match table column count

---

## D. TOC Editing

### When Required
If ANY heading (H1/H2/H3) was added, removed, or renamed, you MUST update the TOC. Skipping leaves stale/broken entries. validate_edits.py emits a `toc_impact` warning as a reminder.

TOC editing is done via **direct XML manipulation** using bash, NOT through edits.json.

### Same-pass vs Second-pass

| Scenario | When to Edit TOC |
|----------|------------------|
| Heading RENAME only | Same pass — between Step 5 (apply) and Step 6 (repack) |
| Heading INSERT or DELETE | Second pass — Step 8 workflow (re-analyze first for correct numPr) |

### XML Structure

**TOC entry** (inside `<w:sdtContent>`):
```xml
<w:hyperlink w:anchor="_Toc12345">
  <w:r><w:t>1. Introduction</w:t></w:r>
</w:hyperlink>
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

### How to Edit

Use `lxml` for TOC XML manipulation (preserves namespace prefixes, handles OOXML reliably).

1. Parse extracted `document.xml` with `lxml.etree`
2. Find the TOC `<w:sdt>` block and target heading's bookmark
3. Edit operations:
   - **Add entry**: Copy existing TOC paragraph, change anchor + text + PAGEREF
   - **Modify entry**: Update `<w:t>` text and `<w:hyperlink>` anchor + PAGEREF
   - **Delete entry**: Remove `<w:p>` from `<w:sdtContent>`
   - **Add bookmark**: Insert `<w:bookmarkStart>` / `<w:bookmarkEnd>` around body heading
4. Save with `tree.write()`

### Rules
- `<w:hyperlink w:anchor="name">` must match `<w:bookmarkStart w:name="name">`
- PAGEREF instrText must reference the same bookmark name
- Bookmark IDs must be unique across the document
- **Do NOT use `--update-fields`** when TOC was manually edited
- Page numbers in manually added entries are static placeholders
- **Number prefixes in TOC entries MUST match re-analyzed numPr values** — never use pre-edit numbering

---

## E. Creating New DOCX

### Step 1: Write Content JSON

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

Supported types: `heading` (level 1-6), `paragraph`, `bullet_list`, `numbered_list`, `table`

---

## F. Pre-flight Checklist

Before writing edits.json:
1. Every edit has `semantic_tag` (→ C)
2. Paragraph REPLACE/INSERT has `style_alias` (→ C)
3. Table edits have `edit_unit` + correct style aliases (→ B)
4. `new_text` is complete, no `\n` in paragraph edits (→ C)
5. `new_text` does NOT include numPr prefixes (→ C)
6. If heading edits + TOC exists → plan TOC update (→ D)
