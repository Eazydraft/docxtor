<img src="logo.png" style="margin-bottom:30px" alt="Docxtor logo"/>

Surgical, LLM-friendly DOCX editing. Load a `.docx`, get a normalized view an LLM can read, apply patches the LLM produces, and write out a valid `.docx` -- all without destroying formatting, styles, headers, or tables.

## How it works

```
DOCX --> XML --> Normalized AST (read-only, with xmlRef pointers) --> LLM Projection
                       |                                                    |
                       |                                              LLM produces
                       |                                              patch objects
                       |                                                    |
                       v                                                    v
              Mutable FXP objects <------- Patch Engine (mutates via xmlRef pointers)
                       |
                       v
                  Serialize --> DOCX
```

The AST is never roundtripped. It's a read-only lens with `xmlRef` pointers back to the mutable [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) (FXP) objects. Patches follow these pointers to mutate FXP nodes directly. Serializing the mutated FXP objects produces the output XML. Formatting, styles, and all non-text OOXML elements survive untouched.

## Install

```bash
pnpm install
```

Requires Node.js >= 20.

## Quick start

```ts
import { Docxtor } from "docxtor";

// Load from file path or buffer
const doc = await Docxtor.load("contract.docx");

// Or create a blank document from scratch
// const doc = await Docxtor.create();

// Get an LLM-friendly projection
const markdown = doc.project("markdown");
// [p0] (Heading1) Service Agreement
// [p1] This agreement is valid from January 1, 2025...
// [p2] The parties agree to the following terms:
// ...

// Apply patches
const patches = [
  { type: "replace_text", target: "p1", find: "valid", replace: "enforceable" },
  { type: "insert_text", target: "p2", position: "after", anchor: "terms:", text: " and conditions:" },
  { type: "delete_text", target: "p5", find: "DRAFT - " },
  { type: "add_table", afterId: "p2", headers: ["Term", "Value"], rows: [["Duration", "12 months"]] },
  { type: "apply_formatting", target: "p1", find: "enforceable", formatting: { bold: true, color: "FF0000" } },
  { type: "apply_style", target: "p0", style: "Heading1" },
  { type: "add_footer", variant: "default", paragraphs: ["Page {{PAGE}} of {{NUMPAGES}}"] },
];
const results = doc.applyPatches(patches);

// Or with tracked changes (produces w:ins/w:del markup that Word shows as redlines)
const trackedResults = doc.applyPatches(patches, {
  trackRevisions: true,
  revisionAuthor: "AI Editor",
});

// Check results
for (const r of results) {
  if (!r.success) console.error(`Patch failed: ${r.error}`);
}

// Save
await doc.toFile("contract-edited.docx");
```

## API

### `Docxtor.load(input, options?)`

Load a DOCX document. `input` can be a file path (`string`), `Buffer`, `ArrayBuffer`, or `Uint8Array`.

```ts
const doc = await Docxtor.load("file.docx");
const doc = await Docxtor.load(buffer);
const doc = await Docxtor.load(buffer, { includeHeadersFooters: false });
```

### `Docxtor.create(input?, options?)`

Create a new Docxtor instance. If `input` is provided, delegates to `Docxtor.load()`. Otherwise, creates a valid empty DOCX document in memory with a single empty paragraph.

```ts
// Create a blank document from scratch
const doc = await Docxtor.create();

// Equivalent to Docxtor.load(buffer)
const doc = await Docxtor.create(buffer);
const doc = await Docxtor.create("file.docx");
```

The empty document is a fully valid `.docx` that can be patched, projected, and exported:

```ts
const doc = await Docxtor.create();
doc.applyPatches([
  { type: "append_text", text: "Hello World" },
]);
await doc.toFile("new-document.docx");
```

### `doc.project(format?, options?)`

Generate an LLM-friendly projection of the document.

**Markdown format** (default for LLM consumption):

```ts
const md = doc.project("markdown");
```

```
[p0] (Heading1) Master Services Agreement
[p1] This Agreement is entered into as of the Effective Date...

| [tbl0.r0.c0.p0] Term | [tbl0.r0.c1.p0] Value |
| [tbl0.r1.c0.p0] Duration | [tbl0.r1.c1.p0] 12 months |

--- Header (default) ---
[hdr.default.p0] CONFIDENTIAL

--- Footer (default) ---
[ftr.default.p0] Page 1

--- Available Styles ---
[paragraph] Normal — Normal
[paragraph] Heading1 — heading 1
[character] Strong — Strong
```

Every element has a stable ID (`p0`, `tbl0.r1.c0.p0`, `hdr.default.p0`, etc.) that patches use as targets.

**Inline formatting markers:** When `includeFormatting: true` is set, the markdown projection annotates text with inline markers showing existing formatting:

```ts
const md = doc.project("markdown", { includeFormatting: true });
```

```
[p0] This has **bold text** and _italic text_ and ~~struck text~~ and <u>underlined text</u>.
[p1] Combined: **_bold italic_** and ~~**bold strikethrough**~~
```

These markers are annotations — they are not part of the actual document text. The supported markers are:

| Marker | Formatting |
|---|---|
| `**text**` | Bold |
| `_text_` | Italic |
| `~~text~~` | Strikethrough |
| `<u>text</u>` | Underline |

Markers can be nested for combined formatting (e.g., `**_bold italic_**`).

**JSON format** (structured data):

```ts
const json = doc.project("json");
// or with options
const json = doc.project("json", {
  includeFormatting: true,      // include bold/italic/underline/etc per run
  includeHeadersFooters: true,  // include headers and footers (default: true)
  includeTables: true,          // include tables (default: true)
  includeRevisions: true,       // include revision marks from tracked changes (default: false)
  includePageBreaks: true,      // include page break elements and hints (default: false)
});
```

When `includeFormatting` is enabled, each paragraph gains a `formatting` array of `JsonProjectionRunFormatting` objects. Only runs that have at least one formatting property set (bold, italic, underline, strikethrough, fontSize, fontFamily, color, highlight) are included — plain/unformatted runs are omitted to keep the output lean. Each entry includes `text`, `startOffset`, and `endOffset` for precise positioning within the paragraph text.

```json
{
  "id": "p0",
  "text": "Hello bold world",
  "formatting": [
    { "text": "bold", "startOffset": 6, "endOffset": 10, "bold": true }
  ]
}
```

When `includePageBreaks` is enabled, docxtor detects all three OOXML page break mechanisms:

| Mechanism | Source | Typical use |
|---|---|---|
| `w:br w:type="page"` | User-inserted explicit break | "Insert Page Break" in Word |
| `w:lastRenderedPageBreak` | Word rendering hint inside a run | Most common in real-world `.docx` files |
| `w:pageBreakBefore` | Paragraph property in `w:pPr` | "Page break before" paragraph setting |

There are two categories of page breaks:

1. **Page-break-only paragraphs** — a `w:p` that contains only a break and no text. These become standalone `NormalizedPageBreak` elements with IDs like `pb0`, `pb1`, ... (they don't consume paragraph indices).
2. **Inline page breaks** — a paragraph that has both text and a break. The paragraph keeps its normal ID (`p0`, `p1`, ...) and gains `hasPageBreak: true` plus `pageBreakOffset`, the exact character index where the break falls in the paragraph text.

The offset is positionally accurate: it reflects where the break actually occurs — at the start, middle, or end of the text — not always at the beginning. This matters for ingestion pipelines that need to split text at page boundaries.

**Markdown** — `[PAGE BREAK]` is spliced into the text at the exact character position:
```
[p0] First page content
[pb0] --- PAGE BREAK ---
[p1] Some text [PAGE BREAK] continues on next page
[p2] [PAGE BREAK] Starts on new page
[p3] All the text up to here [PAGE BREAK]
```

**JSON** — standalone page breaks are `{ id: "pb0", type: "page_break" }`. Inline breaks add `hasPageBreak` and `pageBreakOffset` (character index within `.text`):
```json
{ "id": "p1", "text": "Some text continues on next page", "hasPageBreak": true, "pageBreakOffset": 10 }
```

### `doc.applyPatches(patches, options?)`

Apply an array of patches sequentially. After each patch, the affected paragraph is re-normalized so subsequent patches see updated text. Returns a `PatchResult[]`.

```ts
const results = doc.applyPatches(patches);
// [{ success: true, patch: ... }, { success: false, patch: ..., error: "..." }]
```

With revision tracking enabled, patches produce OOXML tracked-change markup (`w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange`) instead of destructive edits:

```ts
const results = doc.applyPatches(patches, {
  trackRevisions: true,          // enable tracked changes (default: false)
  revisionAuthor: "AI Editor",   // author name shown in Word's review pane (default: "docxtor")
});
```

See [Revision Tracking](#revision-tracking-tracked-changes) below for details.

### `doc.validatePatches(patches)`

Validate patches without applying them.

```ts
const validations = doc.validatePatches(patches);
// [{ valid: true }, { valid: false, error: "Paragraph 'p99' not found" }]
```

### `doc.toBuffer()` / `doc.toFile(path)`

Export the modified document.

```ts
const buffer = await doc.toBuffer();
await doc.toFile("output.docx");
```

### `doc.toString()`

Export the current document as a raw OOXML string (`word/document.xml` content). Reflects any patches that have been applied. Useful for inspecting the XML directly or forwarding it to another system without producing a full DOCX file.

```ts
const xml = doc.toString();
// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
//   <w:body>
//     <w:p><w:r><w:t>Hello World</w:t></w:r></w:p>
//   </w:body>
// </w:document>
```

### Inspection methods

```ts
doc.getAST();                    // Full NormalizedDocument
doc.getBody();                   // Body elements (paragraphs + tables)
doc.getParagraphs();             // Body-level paragraphs only
doc.getTables();                 // Body-level tables only
doc.getHeaders();                // Map<string, HeaderFooter>
doc.getFooters();                // Map<string, HeaderFooter>
doc.findParagraphById("p5");     // Search everywhere (body, tables, headers, footers)
doc.findTableById("tbl0");       // Search body, headers, footers + nested tables
doc.getStyles();                 // StyleInfo[] from styles.xml
```

## Patch types

### `replace_text`

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

This also works for `insert_text` (`anchor` and `text` fields) and `append_text` (`text` field).

### `insert_text`

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

### `delete_text`

Delete a substring from a paragraph.

```ts
{
  type: "delete_text",
  target: "p5",
  find: "DRAFT - ",
  occurrence: 1,
}
```

### `append_text`

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

Mixed content works naturally — text lines become paragraphs, table blocks become tables:

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

### `replace_cell_text`

Replace text in a table cell paragraph (same as `replace_text` but uses the cell's paragraph ID).

```ts
{
  type: "replace_cell_text",
  target: "tbl0.r1.c0.p0",
  find: "TBD",
  replace: "Confirmed",
}
```

### `add_table_row`

Add a row to a table. The new row copies the structure (cell properties, widths) from the last existing row.

```ts
{
  type: "add_table_row",
  target: "tbl0",
  afterRow: 1,            // insert after row index 1; use -1 to insert at the beginning
  cells: ["Col A", "Col B", "Col C"],
}
```

### `delete_table_row`

Delete a row from a table.

```ts
{
  type: "delete_table_row",
  target: "tbl0",
  row: 2,                 // 0-based row index
}
```

### `add_table`

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

### `apply_formatting`

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

### `apply_style`

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

### `add_header`

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

### `add_footer`

Add a footer to the document's last section. Same options as `add_header`.

```ts
{
  type: "add_footer",
  variant: "default",
  paragraphs: ["Page {{PAGE}} of {{NUMPAGES}}"],
}
```

### `remove_header`

Remove a header from the document's last section. Cleans up the relationship, content type entry, and the XML part from the archive.

```ts
{
  type: "remove_header",
  variant: "default",     // "default", "first", or "even"
}
```

### `remove_footer`

Remove a footer from the document's last section. Cleans up the relationship, content type entry, and the XML part from the archive.

```ts
{
  type: "remove_footer",
  variant: "default",
}
```

## ID scheme

| Location | Pattern | Example |
|---|---|---|
| Body paragraph | `p{n}` | `p0`, `p1`, `p12` |
| Body table | `tbl{n}` | `tbl0`, `tbl1` |
| Page break | `pb{n}` | `pb0`, `pb1` |
| Table cell paragraph | `tbl{t}.r{r}.c{c}.p{p}` | `tbl0.r0.c0.p0` |
| Nested table | `tbl{t}.r{r}.c{c}.tbl{n}` | `tbl0.r2.c1.tbl0` |
| Nested cell paragraph | `tbl0.r0.c0.tbl0.r0.c0.p0` | (recursive) |
| Header paragraph | `hdr.{variant}.p{n}` | `hdr.default.p0` |
| Header table | `hdr.{variant}.tbl{n}` | `hdr.default.tbl0` |
| Footer paragraph | `ftr.{variant}.p{n}` | `ftr.default.p0` |
| Footer table | `ftr.{variant}.tbl{n}` | `ftr.default.tbl0` |

IDs are positional and stable within a single editing session. They are reassigned if you re-load the document.

## Revision Tracking (Tracked Changes)

When `trackRevisions: true` is passed to `applyPatches()`, every edit produces OOXML tracked-change markup instead of destructive modifications. The resulting `.docx` opens in Word with redline marks that a human reviewer can accept or reject.

```ts
const results = doc.applyPatches(
  [
    { type: "replace_text", target: "p1", find: "valid", replace: "enforceable" },
    { type: "delete_text", target: "p3", find: "DRAFT - " },
    { type: "insert_text", target: "p2", position: "after", anchor: "terms:", text: " and conditions" },
    { type: "apply_formatting", target: "p1", find: "enforceable", formatting: { bold: true } },
    { type: "apply_style", target: "p0", style: "Heading2" },
    { type: "add_table", headers: ["Col1", "Col2"], rows: [["A", "B"]] },
    { type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["A", "B"] },
    { type: "delete_table_row", target: "tbl0", row: 2 },
  ],
  { trackRevisions: true, revisionAuthor: "AI Editor" },
);
```

### What each patch type produces

| Patch type | OOXML markup |
|---|---|
| `replace_text` | `w:del` wrapping the original runs + `w:ins` wrapping the replacement runs |
| `insert_text` | `w:ins` wrapping the new runs |
| `delete_text` | `w:del` wrapping the deleted runs (text preserved in XML) |
| `append_text` | `w:ins` wrapping each new paragraph's run; markdown tables get `w:ins` in each row's `w:trPr` |
| `apply_formatting` | `w:rPrChange` inside each affected run's `w:rPr` (captures old formatting) |
| `apply_style` (paragraph) | `w:pPrChange` inside `w:pPr` (captures old paragraph style) |
| `apply_style` (character) | `w:rPrChange` inside each affected run's `w:rPr` |
| `add_table` | `w:ins` element inside every row's `w:trPr` |
| `add_table_row` | `w:ins` element inside the new row's `w:trPr` |
| `delete_table_row` | Row is kept; `w:del` element added to `w:trPr` |
| `add_header` | `w:ins` wrapping each paragraph's runs |
| `add_footer` | `w:ins` wrapping each paragraph's runs |
| `remove_header` | Removes header reference, relationship, and content type (no revision markup) |
| `remove_footer` | Removes footer reference, relationship, and content type (no revision markup) |

### Revision IDs

Each revision element gets a unique `w:id`. IDs start from `max(existing w:id values in the document) + 1` so they never collide with pre-existing tracked changes.

### Reading revision marks

The normalizer captures revision marks when present. Use `includeRevisions` in projections to surface them:

```ts
// JSON -- each paragraph gains a `revisions` array
const json = doc.project("json", { includeRevisions: true });

// Markdown -- inline annotations like {+inserted+} and {-deleted-}
const md = doc.project("markdown", { includeRevisions: true });
```

### Behavior notes

- Without `trackRevisions`, all patches work exactly as before (destructive edits). The option is fully backward-compatible.
- Sequential patches within the same `applyPatches()` call share a single revision context: IDs auto-increment, and re-normalization between patches correctly handles the `w:ins`/`w:del` wrappers.
- The normalizer's `collectRuns()` recurses into `w:ins` (treats the content as current text) and skips `w:del` (excludes deleted text from `.text`). This means subsequent patches in the same batch correctly see the post-edit text.

## LLM integration pattern

docxtor is LLM-agnostic. The typical integration:

```ts
import { Docxtor } from "docxtor";

const doc = await Docxtor.load("contract.docx");
const projection = doc.project("markdown", { includeFormatting: true });

// Send projection to your LLM with instructions to produce patches
const response = await llm.chat([
  {
    role: "system",
    content: `You are a document editor. The user will give you a document 
and editing instructions. Respond with a JSON array of patches.

Patch types:
- replace_text: { type: "replace_text", target: "<id>", find: "<text>", replace: "<text>" }
- insert_text: { type: "insert_text", target: "<id>", position: "before"|"after", anchor: "<text>", text: "<text>" }
- delete_text: { type: "delete_text", target: "<id>", find: "<text>" }
- append_text: { type: "append_text", text: "<text>" } (supports markdown tables, {{FIELD}} placeholders, and inline formatting)
- add_table: { type: "add_table", headers: ["col1", ...], rows: [["val1", ...], ...], afterId: "<id>" }
- apply_formatting: { type: "apply_formatting", target: "<id>", find: "<text>", formatting: { bold: true, color: "FF0000" } }
- apply_style: { type: "apply_style", target: "<id>", style: "Heading1" }
- add_header: { type: "add_header", variant: "default"|"first"|"even", paragraphs: ["text", ...] }
- add_footer: { type: "add_footer", variant: "default"|"first"|"even", paragraphs: ["text", ...] }
- remove_header: { type: "remove_header", variant: "default"|"first"|"even" }
- remove_footer: { type: "remove_footer", variant: "default"|"first"|"even" }

Document:
${projection}`,
  },
  { role: "user", content: "Change all mentions of 'Provider' to 'Vendor' and make them bold" },
]);

const patches = JSON.parse(response);

// Validate before applying
const validations = doc.validatePatches(patches);
const invalid = validations.filter((v) => !v.valid);
if (invalid.length > 0) {
  console.error("Invalid patches:", invalid);
}

const results = doc.applyPatches(patches);
await doc.toFile("contract-edited.docx");
```

## Low-level exports

For advanced use cases, the internal modules are exported:

```ts
import {
  readDocx,             // DOCX buffer -> DocxZipEntries
  writeDocx,            // Modified entries -> DOCX buffer
  parseXml,             // XML string -> FXP ordered array
  serializeXml,         // FXP ordered array -> XML string
  normalizeDocument,    // DocxZipEntries -> NormalizedDocument
  applyPatches,         // Direct patch application (accepts PatchOptions)
  project,              // Direct projection
  projectToJson,
  projectToMarkdown,
  validatePatch,
  findParagraphById,
  findTableById,
} from "docxtor";

// Types
import type { PatchOptions, RevisionMark, NormalizedPageBreak, JsonProjectionPageBreak, JsonProjectionRunFormatting, HeaderFooterVariant, HeaderFooterRef } from "docxtor";
```

## What it handles

- Document creation from scratch (`Docxtor.create()`)
- Paragraph appending (`append_text` — document-level, no target needed)
- Table creation (`add_table` — inserts a new table with TableGrid style)
- Markdown table conversion in `append_text` (pipe-delimited syntax → real OOXML tables)
- Paragraphs, runs, and text across split runs
- Tables (add/delete rows, cell text replacement)
- Nested tables
- Headers and footers (including tables within headers/footers)
- Header/footer creation and removal (`add_header`, `add_footer`, `remove_header`, `remove_footer`)
- Field code placeholders (`{{PAGE}}`, `{{NUMPAGES}}`, `{{DATE}}`, `{{TIME}}`, `{{FILENAME}}`) — combinable with inline formatting
- Container targeting for `append_text` and `add_table` (add content to headers/footers)
- Hyperlinks (text inside `w:hyperlink` elements)
- Run splitting with formatting preservation
- Inline formatting patches (bold, italic, underline, strikethrough, fontSize, fontFamily, color, highlight)
- Inline formatting markers in projections and text ops (`**bold**`, `_italic_`, `~~strikethrough~~`, `<u>underline</u>`)
- Style application (paragraph styles like Heading1, character styles like Strong)
- Style validation against the document's `styles.xml`
- Tracked changes / revision tracking (`w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange`)
- Page breaks — surfaced via `includePageBreaks` option with positional `pageBreakOffset` (see [projection options](#docprojectformat-options) for details)
- XML entity escaping (`&amp;`, `&lt;`, `&gt;`)
- Newlines in `append_text` (`\n` creates separate `w:p` paragraphs; markdown tables create `w:tbl`)
- Newlines in run text (`\n` maps to `w:br` elements)
- List/numbering info in projections
- Styles, fonts, colors, highlighting in projections

## What it does not handle (yet)

- Merged cells (`gridSpan` / `vMerge`)
- Numbered list renumbering
- Image manipulation
- Style creation (can only apply pre-existing styles)
- Comments and annotations
- Math equations

## Development

```bash
pnpm install
pnpm test          # run tests (vitest)
pnpm test:watch    # watch mode
pnpm build         # build ESM + CJS + DTS (tsup)
pnpm lint          # type-check
```

414 tests covering ZIP I/O, XML parsing/serialization, normalization, mapping, patch operations (text, table, formatting, style, append, add_table), document creation, header/footer lifecycle (add/remove with round-trip, field placeholders, container targeting, stale zip entry cleanup, even/odd settings), revision tracking (tracked changes), page breaks (explicit, lastRenderedPageBreak, pageBreakBefore, positional offset, revision interaction), integration flows, hyperlinks, entity escaping, newlines, nested tables, markdown table conversion, projection, and API surface.

## License

Proprietary. All rights reserved.
