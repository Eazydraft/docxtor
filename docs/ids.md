# ID scheme

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

In a document where more than one section has headers or footers, their ids carry the section: `hdr.s2.first.p0`, `ftr.s1.default.tbl0`. Sections are numbered from 1 in document order, and every header and footer id in such a document has the `s{n}` part, section 1 included. `HeaderFooter.sectionIndex` gives the section and `HeaderFooter.idPrefix` the start of its ids (`hdr` or `hdr.s2`). A part that several sections share appears once, under the first section that references it; `HeaderFooter.references` lists every reference.

IDs are positional and stable within a single editing session. They are reassigned if you re-load the document.

## Native paragraph identity (`w14:paraId`)

Word 2010+ files stamp every paragraph with a `w14:paraId` (8 uppercase hex chars) that survives round trips through Word, hand edits and re-saves — unlike positional ids, which shift whenever a paragraph is inserted or deleted. docxtor exposes it everywhere a paragraph appears:

```ts
doc.findParagraphById("p12").paraId;         // "4F2A1C7B" (undefined if the file has none)
doc.findParagraphByParaId("4F2A1C7B");       // → NormalizedParagraph
doc.project("markdown");                     // "[p12 ¶4F2A1C7B] text..."
doc.project("json").body[12].paraId;         // "4F2A1C7B"
```

Any patch `target` (and `afterId`) accepts the `para:` form interchangeably with a positional id:

```ts
{ type: "replace_text", target: "para:4F2A1C7B", find: "...", replace: "..." }
{ type: "replace_text", target: "p12", ... }   // unchanged
```

Resolution order for `para:` targets is body (paragraphs and table cells in document order), headers, footers.

Files that lack paraIds (older Word versions, generated files) can be stamped on load with `Docxtor.load(buffer, { assignParaIds: true })`. Ids are unique within the document and in the range Word accepts (`< 0x80000000`); `w14:textId` is set to the same value if absent. Paragraphs created by patches (`insert_paragraph`, `append_text`, `add_list`, `add_list_item`, table cells of `add_table`, rows cloned by `add_table_row`) receive fresh ids whenever the part they land in declares the `w14` namespace.
