# Patch types

## `replace_text`

Replace a substring in a paragraph. Formatting of the original text is preserved.

```ts
{
  type: "replace_text",
  target: "p1",           // paragraph ID
  find: "valid",          // exact substring to find
  replace: "enforceable", // replacement text
  occurrence: 1,          // which occurrence (1-based, default: 1)
}
```

**Inline formatting markers in `find` and `replace`:** You can include GFM/HTML markers (`**bold**`, `_italic_`, `~~strikethrough~~`, `<u>underline</u>`) in these fields. Markers in `find` are stripped before matching against the plain paragraph text. Markers in `replace` are parsed and create properly formatted OOXML runs:

```ts
{
  type: "replace_text",
  target: "p1",
  find: "old text",                        // matches plain text "old text"
  replace: "**bold** and _italic_ text",   // creates bold and italic runs
}
```

This also works for `insert_text` (`anchor` and `text`), `delete_text` (`find`), `insert_paragraph`, `append_text`, `replace_cell_text`, `add_list` / `add_list_item`, `add_header` / `add_footer` and `add_comment`.

Rules that keep ordinary document text intact:

- A `_` is an italic marker only when it stands alone at a word boundary. Runs of underscores (`Name: ______`, `a(n) __________`) and intra-word underscores (`snake_case`) are literal.
- A single `*` is literal (`5 * 3`); only `**` toggles bold.
- Backslash escapes write the character itself: `\_`, `\*`, `\~`, `\<` and `\\` give `_`, `*`, `~`, `<` and `\`. Any other backslash is kept.
- To skip marker parsing altogether, set `markers: false` on the patch or `inlineMarkers: false` in the options; the text (and `find`) is then used verbatim:

```ts
{ type: "replace_text", target: "p1", find: "[AMOUNT]", replace: "**TBD** _____", markers: false }
// paragraph text now contains the literal "**TBD** _____"
```

**Line breaks.** A soft line break (`w:br`) is one `\n` in paragraph text, both in projections and in the text that `find` / `occurrence` are resolved against, so offsets computed from a projection always line up. `find` may span a break (`"By:\nName:"`), and `\n` in `replace` / `text` becomes a `w:br` (`insert_paragraph` and `append_text` are the exception: there `\n` starts a new paragraph).

## `insert_text`

Insert text before or after an anchor string.

```ts
{
  type: "insert_text",
  target: "p2",
  position: "after",      // "before" or "after"
  anchor: "terms:",       // anchor text to locate
  text: " and conditions",
  occurrence: 1,
}
```

## `delete_text`

Delete a substring from a paragraph.

```ts
{
  type: "delete_text",
  target: "p5",
  find: "DRAFT - ",
  occurrence: 1,
}
```

## `append_text`

Append a new paragraph to the end of the document body (or a header/footer). No `target` needed — this is a document-level operation. Particularly useful with `Docxtor.create()` to build documents from scratch.

```ts
{
  type: "append_text",
  text: "New paragraph at the end",
}
```

To append to a header or footer instead of the body, use `container`:

```ts
{
  type: "append_text",
  text: "Updated: {{DATE}}",
  container: { type: "footer", variant: "default" },
}
```

When more than one section has a header or footer of that variant, say which with `section` (1-based, in document order): `container: { type: "header", variant: "first", section: 2 }`. Without it the patch fails validation with a message that lists the sections. The same applies to `container` on `add_table` and `add_list`.

**Field placeholders:** Text supports `{{FIELD}}` placeholders that are converted to proper OOXML field codes. Supported fields: `{{PAGE}}` (current page number), `{{NUMPAGES}}` (total page count), `{{DATE}}`, `{{TIME}}`, `{{FILENAME}}`. Field placeholders and inline formatting can be combined in the same text:

```ts
{
  type: "append_text",
  text: "**Page** {{PAGE}} of {{NUMPAGES}}",
}
```

The new paragraph gets the next available ID (e.g., `p3` if the last paragraph is `p2`) and is immediately available for subsequent patches in the same batch:

```ts
doc.applyPatches([
  { type: "append_text", text: "Conclusion paragraph" },
  { type: "apply_formatting", target: "p3", find: "Conclusion", formatting: { bold: true } },
]);
```

**Newline handling:** If the text contains `\n` characters, each line becomes a separate `w:p` paragraph in the OOXML output. For example, `"Line one\nLine two\nLine three"` creates three paragraphs, each with its own sequential ID. This matches how Word represents distinct paragraphs and ensures proper rendering.

```ts
// Creates 3 separate paragraphs: p1, p2, p3
doc.applyPatches([
  { type: "append_text", text: "First line\nSecond line\nThird line" },
]);
// doc.getParagraphs() → [..., { id: "p1", text: "First line" },
//                              { id: "p2", text: "Second line" },
//                              { id: "p3", text: "Third line" }]
```

**Markdown table support:** If the text contains a markdown table (pipe-delimited rows with a `| --- |` separator line), it is automatically converted to a real OOXML table (`w:tbl`) with TableGrid styling and visible borders. Plain text before and after the table becomes normal paragraphs.

```ts
// Creates an actual table in the document
doc.applyPatches([
  { type: "append_text", text: "| Name | Age |\n| --- | --- |\n| Alice | 30 |\n| Bob | 25 |" },
]);
// doc.getTables() → [{ id: "tbl0", rows: [{ cells: ["Name", "Age"] }, ...] }]
```

**Markdown list support:** If the text contains markdown list syntax (`- ` for bullets, `1. ` for ordered), items are converted to real OOXML list paragraphs with proper numbering definitions in `word/numbering.xml`. Nesting uses 2-space indentation.

```ts
// Creates a bullet list with nesting
doc.applyPatches([
  { type: "append_text", text: "- First item\n- Second item\n  - Sub-item\n- Third item" },
]);

// Creates an ordered list
doc.applyPatches([
  { type: "append_text", text: "1. Step one\n2. Step two\n3. Step three" },
]);
```

Mixed content works naturally — text lines become paragraphs, table blocks become tables, list blocks become list paragraphs:

```ts
doc.applyPatches([
  {
    type: "append_text",
    text: "Summary:\n| Metric | Value |\n| --- | --- |\n| Revenue | $1M |\nEnd of report.",
  },
]);
// Creates: paragraph "Summary:", then a table, then paragraph "End of report."
```

With revision tracking enabled, each appended paragraph's run is wrapped in `w:ins`, and each table row gets `w:ins` in its `w:trPr`, so all content appears as tracked insertions in Word.

## `insert_paragraph`

Insert one or more new paragraphs immediately before or after an existing paragraph — in the body, inside a table cell (a sibling `w:p` in the same `w:tc`), or in a header/footer.

```ts
{
  type: "insert_paragraph",
  target: "p12",                 // or "para:4F2A1C7B"
  position: "after",             // "before" | "after"
  text: "Landlord shall indemnify, defend and hold harmless Tenant ...",
  inheritStyle: true,            // default true
}
```

- `text` supports inline formatting markers (`markers: false` writes it verbatim) and `{{FIELD}}` placeholders; `\n` creates multiple paragraphs, inserted contiguously in the given order. Markdown tables and lists are **not** supported here (the validator rejects them) — use `add_table` / `add_list` with `afterId`.
- `inheritStyle` copies the target's `w:pPr` (paragraph style, numbering `numId`/`ilvl`, indentation, spacing — not section breaks or revision state) and gives the new text the formatting of the target's last run, which is what Word does when you press Enter at the end of a paragraph.
- New paragraphs get the next free positional id in the target's container (e.g. `p13`, `tbl0.r1.c0.p1`), a fresh `paraId` when the part declares `w14`, and are addressable by later patches in the same batch.
- With `trackRevisions`, the runs are wrapped in `w:ins` **and** the paragraph mark carries `w:pPr/w:rPr/w:ins`, so accepting the change leaves exactly one new paragraph and no stray blank one.

## `delete_paragraph`

Delete a whole paragraph.

```ts
{ type: "delete_paragraph", target: "p12" }   // or "para:4F2A1C7B"
```

- With `trackRevisions`: every run is wrapped in `w:del` (`w:t` → `w:delText`, `w:instrText` → `w:delInstrText`) and the paragraph mark gets `w:pPr/w:rPr/w:del` — Word's convention for a deleted paragraph. Runs already inside a `w:ins` by the same author are removed outright rather than double-marked; another author's insertion gets the `w:del` nested inside it. Accepting removes the paragraph with no empty paragraph left behind.
- Without tracking: the `w:p` is removed. The only paragraph of a table cell is emptied instead, since a cell must keep one paragraph.

## `add_comment`

Attach a Word comment to a paragraph or to a substring of it. Creates `word/comments.xml` (with its relationship and content type) on demand.

```ts
{
  type: "add_comment",
  target: "p12",                 // or "para:4F2A1C7B"
  find: "recapture",             // optional: anchor the comment to this text (occurrence? default 1)
  text: "Delete per firm rule R-12.\nSecond paragraph of the comment.",
  author: "USI Legal AI",        // defaults to PatchOptions.revisionAuthor, then "docxtor"
  initials: "ULA",               // optional; derived from author if omitted
}
```

The comment date comes from `PatchOptions.revisionDate` (default: now). Comments are left alone by `acceptAllRevisions` / `rejectAllRevisions`.

## `replace_cell_text`

Replace text in a table cell paragraph (same as `replace_text` but uses the cell's paragraph ID).

```ts
{
  type: "replace_cell_text",
  target: "tbl0.r1.c0.p0",
  find: "TBD",
  replace: "Confirmed",
}
```

## `add_table_row`

Add a row to a table. The new row copies the structure (cell properties, widths) from the last existing row.

```ts
{
  type: "add_table_row",
  target: "tbl0",
  afterRow: 1,            // insert after row index 1; use -1 to insert at the beginning
  cells: ["Col A", "Col B", "Col C"],
}
```

## `delete_table_row`

Delete a row from a table.

```ts
{
  type: "delete_table_row",
  target: "tbl0",
  row: 2,                 // 0-based row index
}
```

## `add_table`

Add a new table to the document body (or a header/footer). This is a document-level operation — no existing table is required. The table is created with TableGrid styling (visible single borders).

```ts
{
  type: "add_table",
  headers: ["Name", "Age", "Role"],           // optional header row
  rows: [["Alice", "30", "Engineer"], ["Bob", "25", "Designer"]],
}
```

Use `afterId` to insert the table after a specific paragraph or table. If omitted, the table is appended to the end of the document (before `w:sectPr`).

```ts
{
  type: "add_table",
  afterId: "p2",          // insert after paragraph p2
  headers: ["Item", "Qty"],
  rows: [["Widget", "10"]],
}
```

To add a table inside a header or footer, use `container`:

```ts
{
  type: "add_table",
  headers: ["A", "B"],
  rows: [["1", "2"]],
  container: { type: "header", variant: "default" },
}
```

> **Note:** `afterId` is not supported when `container` is set. Container-targeted tables are always appended to the end of the header/footer. The validator will reject patches that specify both.

Headers are optional. If provided, they become the first row of the table. All rows (and headers, if present) must have the same number of cells.

```ts
// Table without headers
{ type: "add_table", rows: [["A", "B"], ["C", "D"]] }
```

With revision tracking, every row in the new table gets `w:ins` in its `w:trPr`.

## `apply_formatting`

Apply inline formatting to a substring or an entire paragraph. Only touches `w:rPr` on targeted runs -- never modifies global styles, paragraph properties, or theme definitions.

```ts
{
  type: "apply_formatting",
  target: "p1",
  find: "important",       // exact substring; omit to format the entire paragraph
  occurrence: 1,           // which occurrence (1-based, default: 1)
  formatting: {
    bold: true,            // true = add, false/null = remove, undefined = leave unchanged
    italic: true,
    underline: "single",   // "single", "double", "wave", "dash", etc. or null to remove
    strikethrough: false,
    fontSize: 28,          // in half-points (28 = 14pt); null to remove
    fontFamily: "Arial",   // sets w:rFonts w:ascii + w:hAnsi; null to remove
    color: "FF0000",       // hex color; null to remove
    highlight: "yellow",   // "yellow", "green", "cyan", etc.; null to remove
  },
}
```

Properties not mentioned in `formatting` are left unchanged on the run. This means you can add bold without affecting an existing color, or remove underline without touching font size.

When `find` is provided, runs are split to isolate exactly the matched text, and formatting is applied only to those runs. When `find` is omitted, formatting is applied to every run in the paragraph.

## `apply_style`

Re-point a paragraph or substring to a pre-existing style defined in the document's `styles.xml`. This is the equivalent of clicking a style in Word's ribbon -- it never creates or modifies style definitions.

**Paragraph style** (no `find`): sets `w:pPr > w:pStyle`. The style must be type `"paragraph"`.

```ts
{
  type: "apply_style",
  target: "p3",
  style: "Heading1",      // style ID from styles.xml
}
```

**Character style** (with `find`): isolates matched runs and sets `w:rPr > w:rStyle`. The style must be type `"character"`.

```ts
{
  type: "apply_style",
  target: "p1",
  find: "important term",
  style: "Strong",        // character style ID
  occurrence: 1,
}
```

Available styles are listed in the markdown projection under "Available Styles" and in the JSON projection's `styles` array. Use `doc.getStyles()` to inspect them programmatically.

## `add_header`

Add a header to the document's last section. Creates the header XML part, relationship, and content type entry automatically.

```ts
{
  type: "add_header",
  variant: "default",     // "default" (all/odd pages), "first" (first page only), "even" (even pages)
  paragraphs: ["CONFIDENTIAL", "Page {{PAGE}}"],
}
```

Paragraph text supports `{{FIELD}}` placeholders: `{{PAGE}}`, `{{NUMPAGES}}`, `{{DATE}}`, `{{TIME}}`, `{{FILENAME}}`.

**Variant behavior:**
- `"default"` — applies to all pages (or odd pages when even/odd mode is active).
- `"first"` — applies to the first page only. Automatically adds `w:titlePg` to the section properties.
- `"even"` — applies to even pages. Automatically adds `<w:evenAndOddHeaders/>` to `word/settings.xml` so Word honours the even-page reference.

Headers and footers can be added to documents created via `Docxtor.create()` — the required `document.xml.rels` is initialized on-demand.

## `add_footer`

Add a footer to the document's last section. Same options as `add_header`.

```ts
{
  type: "add_footer",
  variant: "default",
  paragraphs: ["Page {{PAGE}} of {{NUMPAGES}}"],
}
```

## `remove_header`

Remove a header from the document's last section. Cleans up the relationship, content type entry, and the XML part from the archive.

```ts
{
  type: "remove_header",
  variant: "default",     // "default", "first", or "even"
}
```

## `remove_footer`

Remove a footer from the document's last section. Cleans up the relationship, content type entry, and the XML part from the archive.

```ts
{
  type: "remove_footer",
  variant: "default",
}
```

## `add_list`

Create a new list (bullet, ordered, or multilevel) in the document body or a header/footer. Creates the numbering definition in `word/numbering.xml` automatically.

```ts
{
  type: "add_list",
  listType: "bullet",             // "bullet", "ordered", or "multilevel"
  items: ["First item", "Second item", "Third item"],
}
```

Items can specify individual nesting levels:

```ts
{
  type: "add_list",
  listType: "ordered",
  items: [
    { text: "Top-level item", level: 0 },
    { text: "Sub-item", level: 1 },
    { text: "Another top-level", level: 0 },
  ],
}
```

Use `afterId` to insert the list after a specific element, or `container` to add it to a header/footer:

```ts
{
  type: "add_list",
  listType: "bullet",
  items: ["Item A", "Item B"],
  afterId: "p2",                 // insert after paragraph p2
}

{
  type: "add_list",
  listType: "ordered",
  items: ["Step 1", "Step 2"],
  container: { type: "footer", variant: "default" },
}
```

To reuse an existing numbering definition (e.g., to continue a list sequence), pass `numId`:

```ts
{
  type: "add_list",
  listType: "bullet",
  numId: "3",                    // reuse existing numId from numbering.xml
  items: ["Continuation item"],
}
```

The `numId` is validated against `numbering.xml` — if the ID doesn't exist, the patch is rejected.

## `add_list_item`

Add a new item to an existing list, positioned before or after a target list paragraph.

```ts
{
  type: "add_list_item",
  target: "p3",                  // must be an existing list paragraph
  position: "after",             // "before" or "after"
  text: "New item text",
  level: 1,                     // optional; defaults to the target's level
}
```

The new item inherits the `numId` from the target paragraph. Works correctly for paragraphs in table cells and headers/footers.

## `remove_list_item`

Remove a list item paragraph entirely (or mark it as deleted with revision tracking).

```ts
{
  type: "remove_list_item",
  target: "p3",                  // must be a list paragraph
}
```

With revision tracking, the paragraph's runs are wrapped in `w:del` (using proper `w:delText`) and the paragraph properties get a `w:pPrChange` record. Runs that were inside `w:ins` (unaccepted insertions) are simply removed — not re-wrapped in `w:del`.

## `convert_to_list`

Convert one or more existing plain paragraphs into a list by adding `w:numPr` to their `w:pPr`.

```ts
{
  type: "convert_to_list",
  targets: ["p2", "p3", "p4"],   // paragraph IDs to convert
  listType: "bullet",            // "bullet", "ordered", or "multilevel"
  level: 0,                     // optional; default 0
}
```

All targeted paragraphs share the same numbering definition.

## `remove_list`

Remove list formatting from one or more paragraphs, converting them back to plain paragraphs.

```ts
{
  type: "remove_list",
  targets: ["p2", "p3", "p4"],
}
```

## `set_list_properties`

Change the level or list type of an existing list paragraph.

```ts
{
  type: "set_list_properties",
  target: "p3",
  level: 2,                     // optional; change indentation level
  listType: "ordered",          // optional; switch list type (creates new numbering definition)
}
```
