# API reference

## `Docxtor.load(input, options?)`

Load a DOCX document. `input` can be a file path (`string`), `Buffer`, `ArrayBuffer`, or `Uint8Array`.

```ts
const doc = await Docxtor.load("file.docx");
const doc = await Docxtor.load(buffer);
const doc = await Docxtor.load(buffer, { includeHeadersFooters: false });
const doc = await Docxtor.load(buffer, { assignParaIds: true });  // stamp w14:paraId on every paragraph
```

Options:

| Option | Default | Description |
|---|---|---|
| `includeHeadersFooters` | `true` | Normalize headers and footers. |
| `assignParaIds` | `false` | Stamp a native `w14:paraId` (and `w14:textId`) on every paragraph in the body, headers and footers that lacks one, declaring the `w14` namespace and adding it to `mc:Ignorable`. See [Native paragraph identity](ids.md#native-paragraph-identity-w14paraid). |

## `Docxtor.create(input?, options?)`

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

## `doc.project(format?, options?)`

Generate an LLM-friendly projection of the document.

**Markdown format** (default for LLM consumption):

```ts
const md = doc.project("markdown");
```

```
[p0] (Heading1) Master Services Agreement
[p1] This Agreement is entered into as of the Effective Date...
[p2] {list:1 L0} - First item
[p3] {list:1 L0} - Second item
[p4] {list:1 L1}   - Sub-item

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

List paragraphs show a `{list:N LX}` hint where `N` is the numbering definition ID (`numId`) and `X` is the nesting level. Bullet items render with `- ` prefix, ordered items with `1. `, `2. `, etc., and multilevel items with compound prefixes like `1.2.1. `. To continue an existing list sequence when adding items, reuse the same `numId` from the hint.

Every element has a stable ID (`p0`, `tbl0.r1.c0.p0`, `hdr.default.p0`, etc.) that patches use as targets. In a document with more than one section the header and footer ids carry the section, `hdr.s2.first.p0`, and the divider names it: `--- Header (first, section 2) ---`. See [ids.md](ids.md).

**Content controls.** The text inside a Word content control (a form field, `w:sdt`) is part of the paragraph text, whether the control sits inside a paragraph, wraps whole paragraphs or wraps a table cell. The markdown projection marks the controls a reader must not mistake for ordinary text:

```
[p7] {field: Insert Title for Contract} {field: Add Contract reference if applicable}
| [tbl5.r3.c0.p0] Type of contract | [tbl5.r3.c1.p0] {field dropdown "One-off purchase" OR "Ongoing supply": Choose an item.} |
[p119] {field date: Select date}
[p120] {field date= 1 May 2026}
[p40] ☐ Agency has implemented the plan
```

| Marker | Meaning |
|---|---|
| `{field: prompt}` | An empty control. The text is its prompt, not an answer. |
| `{field date: prompt}`, `{field dropdown "A" OR "B": prompt}`, `{field combo "A" OR "B": prompt}` | An empty date picker, drop-down list or combo box, with the list items. |
| `{field date= value}`, `{field dropdown "A" OR "B"= value}` | A filled date picker, drop-down list or combo box. |
| `{field:}` | A control with no text at all. |
| `{field picture}` | A picture control. |

A filled rich-text or plain-text control prints as ordinary text, a checkbox prints its glyph (`☐`, `☒`), and a building-block gallery such as a table of contents prints as its content. The markers are annotations, not document text: `find` matches the text inside them. [Text patches inside a content control](patches.md#text-patches-inside-a-content-control) says what an edit there does.

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
  runFilter: { color: "FF0000" }, // JSON only: report exactly the runs matching this filter (see below)
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

A paragraph that holds content controls, or sits inside one, has a `controls` array. A control wrapped around the paragraph (or around its table cell) comes first and covers all of the text; the controls inside the paragraph follow in document order, a nested control after the one that holds it. Each entry is a `ControlInfo` without its XML reference, plus the range of `text` it covers:

```json
{
  "id": "tbl5.r3.c1.p0",
  "text": "Choose an item.",
  "controls": [
    {
      "id": "-564876532", "kind": "dropDownList", "alias": "Type of contract",
      "placeholder": true, "options": ["One-off purchase", "Ongoing supply"],
      "startOffset": 0, "endOffset": 15
    }
  ]
}
```

| Field | |
|---|---|
| `id` | the control's `w:id` |
| `kind` | `richText`, `text`, `date`, `dropDownList`, `comboBox`, `checkbox`, `picture`, `docPart` (a building-block gallery) or `other` |
| `alias`, `tag` | the control's title and developer tag, when set |
| `placeholder` | `true` while the control shows its prompt: the text in the range is not an answer |
| `options` | display texts of a drop-down list or combo box |
| `checked` | state of a checkbox |
| `locked` | `content` (the content cannot be edited), `control` (the control cannot be deleted) or `both` |
| `bound` | `true` when the control mirrors a value in a custom XML part |

In the AST the same information is on `RunMapping.control` (the innermost control holding a run), `NormalizedParagraph.controls` (the ranges), `NormalizedParagraph.control` and `NormalizedTableCell.control` (a control wrapped around the paragraph or the cell). A table cell's JSON entry carries its `paragraphs` whenever one of them has controls.

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

## `doc.applyPatches(patches, options?)`

Apply an array of patches sequentially. After each patch, the affected paragraph is re-normalized so subsequent patches see updated text. Returns a `PatchResult[]`.

```ts
const results = doc.applyPatches(patches);
// [{ success: true, patch: ... }, { success: false, patch: ..., error: "..." }]
```

With revision tracking enabled, patches produce OOXML tracked-change markup (`w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange`) instead of destructive edits:

```ts
const results = doc.applyPatches(patches, {
  trackRevisions: true,                  // enable tracked changes (default: false)
  revisionAuthor: "AI Editor",           // author name shown in Word's review pane (default: "docxtor")
  revisionDate: "2026-10-13T00:00:00Z",  // ISO 8601 date on every revision mark (default: now)
  atomic: true,                          // all-or-nothing batch (default: false)
  inlineMarkers: false,                  // write patch text verbatim, no GFM markers (default: true)
});
```

**Verbatim text.** By default the text fields of patches (`find`, `replace`, `anchor`, `text`, list items, header/footer paragraphs, comment text) are parsed for inline formatting markers (see [`replace_text`](patches.md#replace_text)). `inlineMarkers: false` turns that off for the whole batch, and a patch's own `markers: true | false` overrides the batch setting, so approved language can be written exactly as given.

`revisionDate` is stamped on every `w:ins` / `w:del` / `w:*Change` element the batch creates. Pin it when you need byte-stable output (e.g., fixed-point tests that compare `document.xml`).

**Atomic batches.** With `atomic: true` the patches run against a deep clone of the document. If every patch succeeds the clone replaces the document; if any patch fails the document is left exactly as it was, and `toBuffer()` produces the same file as before the batch. The returned `PatchResult[]` still has one entry per patch so you can see what would have applied and what failed — in atomic mode a `success: true` entry means "would have applied", not "did apply"; check `results.every(r => r.success)` for the batch outcome. Because the batch runs on the clone, a patch may target a paragraph created by an earlier patch in the same batch.

See [Revision tracking](revisions.md) for details.

## `doc.acceptAllRevisions(options?)` / `doc.rejectAllRevisions(options?)`

Resolve every tracked change in the document — body, headers, footers, footnotes, endnotes and comments.

```ts
const { accepted } = doc.acceptAllRevisions();                        // everything
const { accepted } = doc.acceptAllRevisions({ authors: ["AI Editor"] }); // only this author's marks
const { rejected } = doc.rejectAllRevisions({ authors: ["AI Editor"] });
const { rejected } = doc.rejectAllRevisions({ ids: ["14", "15"] });     // only these w:ids
```

| Element | Accept | Reject |
|---|---|---|
| `w:ins` / `w:moveTo` (runs) | unwrapped | removed |
| `w:del` / `w:moveFrom` (runs) | removed | unwrapped, `w:delText` → `w:t` |
| `w:rPrChange`, `w:pPrChange`, `w:trPrChange`, `w:tblPrChange`, `w:tcPrChange`, `w:sectPrChange` | dropped (new properties kept) | old properties restored |
| Row `w:trPr/w:ins` / `w:trPr/w:del` | marker dropped / row removed | row removed / marker dropped |
| Paragraph mark `w:pPr/w:rPr/w:ins` | marker dropped | paragraph merged into the next |
| Paragraph mark `w:pPr/w:rPr/w:del` | paragraph merged into the next | marker dropped |

With `authors`, only revisions by those authors are touched; a reviewer's hand edits survive while the tool's own marks are accepted. With `ids`, only the revisions whose `w:id` is listed are touched (both filters apply when both are given). Each insertion and deletion in a projection with `includeRevisions` carries its `id`; a replacement is a `w:del` and a `w:ins`, so pass both ids to resolve it. Move ranges are matched by their own ids. The document is re-normalized afterwards: positional ids are reassigned, `paraId`s are not. Use these to produce the clean final document from a redline (accept the tool's author, then strip any remaining bracket tokens with `delete_text`).

## Metadata: `doc.getCoreProperties()` / `doc.setCoreProperties(props)` / `doc.scrubMetadata()`

```ts
doc.getCoreProperties();
// { creator, lastModifiedBy, title, subject, description, keywords, category, contentStatus, revision, created, modified }

doc.setCoreProperties({ creator: "Your Firm", title: "Lease (Tenant Draft)", modified: "2026-10-13T00:00:00Z" });
// creates docProps/core.xml (plus its relationship and content type) if the file has none

doc.scrubMetadata();
// core.xml: removes creator, lastModifiedBy, description, keywords, category, contentStatus; sets revision = "1"
// app.xml:  removes Company, Manager, HyperlinkBase
// removes docProps/custom.xml and its relationship / content-type entries
// leaves title, created and modified untouched
```

## `doc.validatePatches(patches)`

Validate patches without applying them.

```ts
const validations = doc.validatePatches(patches);
// [{ valid: true }, { valid: false, error: "Paragraph 'p99' not found" }]
```

## `doc.toBuffer()` / `doc.toFile(path)`

Export the modified document.

```ts
const buffer = await doc.toBuffer();
await doc.toFile("output.docx");
```

## `doc.toString()`

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

## Inspection methods

`doc.findParagraphByParaId("4F2A1C7B")` resolves a paragraph by its native `w14:paraId` (case-insensitive; body including table cells, then headers, then footers).

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
  findParagraphByParaId,
  findTableById,
  reNormalizeDocument,  // rebuild body/headers/footers from the XML parts in place
  cloneDocument,        // deep clone (used by atomic batches)
  acceptAllRevisions,   // operate on a NormalizedDocument directly (call reNormalizeDocument after)
  rejectAllRevisions,
  getCoreProperties,
  setCoreProperties,
  scrubMetadata,
} from "@lucio-ai/docxtor";

// Types
import type { PatchOptions, RevisionMark, RevisionFilterOptions, CoreProperties, RunFilter, NormalizedPageBreak, JsonProjectionPageBreak, JsonProjectionRunFormatting, HeaderFooterVariant, HeaderFooterRef, InsertParagraphPatch, DeleteParagraphPatch, AddCommentPatch, AddListPatch, AddListItemPatch, RemoveListItemPatch, ConvertToListPatch, RemoveListPatch, SetListPropertiesPatch } from "@lucio-ai/docxtor";
```
