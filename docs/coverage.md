# Feature coverage

- Document creation from scratch (`Docxtor.create()`)
- Paragraph appending (`append_text` — document-level, no target needed)
- Paragraph insertion at a location and whole-paragraph deletion (`insert_paragraph`, `delete_paragraph`), tracked or not
- Native paragraph identity (`w14:paraId` read, `para:` targets, `assignParaIds` on load, fresh ids on new paragraphs)
- Accept / reject all revisions, optionally per author (`acceptAllRevisions`, `rejectAllRevisions`)
- Atomic (all-or-nothing) patch batches (`atomic: true`)
- Document metadata (`getCoreProperties`, `setCoreProperties`, `scrubMetadata`)
- Comments (`add_comment`, creating `word/comments.xml` on demand)
- Reading runs by formatting, e.g. red-text answers (`runFilter` in the JSON projection)
- Footnotes and endnotes parts (preserved, and walked by accept/reject)
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
- Page breaks — surfaced via `includePageBreaks` option with positional `pageBreakOffset` (see [projection options](api.md#docprojectformat-options) for details)
- XML entity escaping (`&amp;`, `&lt;`, `&gt;`)
- Newlines in `append_text` (`\n` creates separate `w:p` paragraphs; markdown tables create `w:tbl`)
- Soft line breaks (`w:br` ↔ `\n`, one character each, in projections and patch offsets alike)
- Verbatim text (`inlineMarkers: false` / per-patch `markers: false`, plus backslash escapes)
- List/numbering info in projections
- List operations (`add_list`, `add_list_item`, `remove_list_item`, `convert_to_list`, `remove_list`, `set_list_properties`)
- Markdown list conversion in `append_text` (bullet `- ` and ordered `1. ` syntax with nesting → real OOXML lists)
- Numbering definition management (`word/numbering.xml` auto-created with content types and relationships)
- Styles, fonts, colors, highlighting in projections

## What it does not handle (yet)

- Merged cells (`gridSpan` / `vMerge`)
- Image manipulation
- Style creation (can only apply pre-existing styles)
- Reading or resolving existing comments (only `add_comment` is supported)
- Document compare (clean vs original is a different algorithm — use an external engine)
- Math equations
