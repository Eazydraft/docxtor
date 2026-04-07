# docxtor

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

// Get an LLM-friendly projection
const markdown = doc.project("markdown");
// [p0] (Heading1) Service Agreement
// [p1] This agreement is valid from January 1, 2025...
// [p2] The parties agree to the following terms:
// ...

// Apply patches
const results = doc.applyPatches([
  { type: "replace_text", target: "p1", find: "valid", replace: "enforceable" },
  { type: "insert_text", target: "p2", position: "after", anchor: "terms:", text: " and conditions:" },
  { type: "delete_text", target: "p5", find: "DRAFT - " },
]);

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
```

Every element has a stable ID (`p0`, `tbl0.r1.c0.p0`, `hdr.default.p0`, etc.) that patches use as targets.

**JSON format** (structured data):

```ts
const json = doc.project("json");
// or with options
const json = doc.project("json", {
  includeFormatting: true,      // include bold/italic/underline/etc per run
  includeHeadersFooters: true,  // include headers and footers (default: true)
  includeTables: true,          // include tables (default: true)
});
```

### `doc.applyPatches(patches)`

Apply an array of patches sequentially. After each patch, the affected paragraph is re-normalized so subsequent patches see updated text. Returns a `PatchResult[]`.

```ts
const results = doc.applyPatches(patches);
// [{ success: true, patch: ... }, { success: false, patch: ..., error: "..." }]
```

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

### Inspection methods

```ts
doc.getAST();                    // Full NormalizedDocument
doc.getBody();                   // Body elements (paragraphs + tables)
doc.getParagraphs();             // Body-level paragraphs only
doc.getTables();                 // Body-level tables only
doc.getHeaders();                // Map<string, HeaderFooter>
doc.getFooters();                // Map<string, HeaderFooter>
doc.findParagraphById("p5");     // Search everywhere (body, tables, headers, footers)
doc.findTableById("tbl0");       // Search body + nested tables
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

## ID scheme

| Location | Pattern | Example |
|---|---|---|
| Body paragraph | `p{n}` | `p0`, `p1`, `p12` |
| Body table | `tbl{n}` | `tbl0`, `tbl1` |
| Table cell paragraph | `tbl{t}.r{r}.c{c}.p{p}` | `tbl0.r0.c0.p0` |
| Nested table | `tbl{t}.r{r}.c{c}.tbl{n}` | `tbl0.r2.c1.tbl0` |
| Nested cell paragraph | `tbl0.r0.c0.tbl0.r0.c0.p0` | (recursive) |
| Header paragraph | `hdr.{variant}.p{n}` | `hdr.default.p0` |
| Footer paragraph | `ftr.{variant}.p{n}` | `ftr.default.p0` |

IDs are positional and stable within a single editing session. They are reassigned if you re-load the document.

## LLM integration pattern

docxtor is LLM-agnostic. The typical integration:

```ts
import { Docxtor } from "docxtor";

const doc = await Docxtor.load("contract.docx");
const projection = doc.project("markdown");

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

Document:
${projection}`,
  },
  { role: "user", content: "Change all mentions of 'Provider' to 'Vendor'" },
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
  applyPatches,         // Direct patch application
  project,              // Direct projection
  projectToJson,
  projectToMarkdown,
  validatePatch,
  findParagraphById,
  findTableById,
} from "docxtor";
```

## What it handles

- Paragraphs, runs, and text across split runs
- Tables (add/delete rows, cell text replacement)
- Nested tables
- Headers and footers
- Hyperlinks (text inside `w:hyperlink` elements)
- Run splitting with formatting preservation
- XML entity escaping (`&amp;`, `&lt;`, `&gt;`)
- Newlines (`\n` maps to `w:br` elements)
- List/numbering info in projections
- Styles, fonts, colors, highlighting in projections

## What it does not handle (yet)

- Tracked changes (revisions)
- Merged cells (`gridSpan` / `vMerge`)
- Numbered list renumbering
- Image manipulation
- Style creation
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

166 tests covering ZIP I/O, XML parsing/serialization, normalization, mapping, patch operations, integration flows, hyperlinks, entity escaping, newlines, nested tables, projection, and API surface.

## License

Proprietary. All rights reserved.
