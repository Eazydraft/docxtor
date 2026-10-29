# Design

This page explains how docxtor is put together and the rules the code and tests hold to. Read it before changing the normalizer, the patch engine or the revision code.

## Why it exists

docxtor was built for legal document review: a language model proposes edits to a contract, and a lawyer reviews them in Word as tracked changes. That setting forces three things on a DOCX layer.

1. **The model never edits the file.** It reads a projection and produces patch objects. Application code decides which patches to apply. docxtor applies them. It never decides *what* changes, only how a known change lands in the file.
2. **Everything else must survive.** Styles, numbering, section properties, images, fields, another author's tracked changes. A reviewer who finds collateral damage stops trusting the tool.
3. **Edits must be reviewable.** Every change should be able to appear as a tracked change with an author and a date, and accepting those changes must give the same document as applying the edits directly.

## Pipeline

```
DOCX --> XML --> Normalized AST (read-only, with xmlRef pointers) --> projection
                        |                                                 |
                        |                                           patch objects
                        v                                                 v
               Mutable XML objects <-------- Patch engine (mutates via xmlRef pointers)
                        |
                        v
                   Serialize --> DOCX
```

- `src/zip` reads the archive into a map of parts and writes it back, preserving entries it does not understand.
- `src/xml` parses each part with fast-xml-parser in ordered mode (`preserveOrder`) and serializes it again. Element order, attributes and unknown elements are kept.
- `src/normalizer` walks the parsed XML and builds the normalized AST: paragraphs, runs, tables, headers, footers, page breaks, revision marks. Every node carries an `xmlRef` back to the XML object it came from.
- `src/projection` renders the AST as markdown or JSON for a model.
- `src/patches` validates and applies patches by following `xmlRef` pointers and mutating the XML objects. `src/mapping` turns character offsets in paragraph text back into runs, and `run-splitter` splits runs so a patch touches exactly the matched text.
- `src/revisions` accepts or rejects tracked changes across every part. `src/metadata` reads and writes `docProps`.

## Invariants

The test suite enforces each of these. A change that breaks one is a bug, however small the diff.

### The AST is a read-only lens

The AST is never serialized. Patches mutate the XML objects through `xmlRef`, then the engine re-normalizes the affected paragraph (or the whole document for structural patches) so the next patch in the batch sees current text and ids. Writing back from the AST would lose everything the normalizer does not model.

### Untouched content survives

A patch may only change the XML it targets. Run splitting copies `w:rPr` to every piece. New paragraphs copy the target's `w:pPr` when `inheritStyle` is on. Nothing rewrites styles, numbering definitions, section properties or theme parts unless a patch exists for exactly that purpose (`add_list` creates numbering definitions, `add_header` creates header parts).

### Text is one string per paragraph, and offsets line up

Paragraph text is the concatenation of its run text, including runs inside `w:ins` and excluding runs inside `w:del`, with hyperlinks and fields flattened to their visible text. A soft line break is one `\n`. Projections print that same string, so an offset or `find` computed from a projection resolves to the same characters in the patch engine.

### Two kinds of identity

Positional ids (`p12`, `tbl0.r1.c0.p0`, `hdr.default.p0`) are assigned in document order and are stable within one editing session. Structural patches re-normalize and may shift them, and a patch later in the same batch may target an id created by an earlier patch.

Native `w14:paraId` values survive everything: patches, accept and reject, a round trip through Word. docxtor reads them, exposes them in projections, accepts `para:XXXXXXXX` as a target anywhere a positional id is accepted, can stamp them on load, and gives new paragraphs fresh ids when the part declares the `w14` namespace. It never rewrites an existing one.

### Tracked changes are a fixed point under accept

With `trackRevisions`, a patch emits `w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange` or the table-row equivalents instead of editing destructively. Applying a batch tracked and then calling `acceptAllRevisions` must yield the same text as applying it untracked. For `replace_text`, `insert_paragraph`, `delete_paragraph` and `replace_cell_text` the tests compare `document.xml` byte for byte, with `revisionDate` pinned. Revision ids start above the highest id already in the document, so they never collide with an existing round of review.

Editing inside an unaccepted insertion follows Word's behaviour: the batch author's own insertion is edited in place, another author's insertion is split around the edit so both attributions survive.

### Atomic batches cannot half-apply

With `atomic: true` the batch runs on a deep clone. Only if every patch succeeds does the clone replace the document. Otherwise `toBuffer()` produces the same file as before.

### Validation runs before mutation

`validatePatches` and the first phase of `applyPatches` check targets, occurrences, style ids, `numId` references and option combinations without touching the document. A patch that fails validation reports an error and leaves the document as it was.

## Acceptance criteria

The integration test in `tests/integration/lease-acceptance.test.ts` checks the workflow end to end on a synthetic lease built in code:

- **Round trip.** Load and save with no patches. Paragraph count, text and `paraId`s are unchanged.
- **Identity stability.** Ten tracked patches including two paragraph inserts, then save and reload. Every `paraId` resolves to the same text; positional ids shift as expected.
- **Fixed point.** A tracked batch followed by accept equals the untracked batch byte for byte, with `revisionDate` pinned and fresh `paraId`s normalised.
- **Correspondence.** Insert N bracket tokens at N distinct anchors. The projection finds each exactly once, at its paragraph.
- **Paragraph insert in a list.** The new paragraph shares the target's `numId` and `ilvl`; after accept there is exactly one new paragraph and no blank.
- **Paragraph delete.** Tracked delete plus accept drops the count by one and leaves no empty paragraph.
- **Author filter.** Accepting only the tool's author leaves another author's marks untouched.
- **Red-text read.** The colour filter returns only the red runs, with correct ids and offsets.
- **Atomic failure.** A failing atomic batch leaves the exported file identical to a no-patch export.
- **Metadata scrub.** No author or company fields remain and `custom.xml` is gone.

## Non-goals

- Document compare (clean versus original). A different algorithm; use an external engine.
- PDF conversion and OCR.
- Deriving patches from a diff of text. Application code owns that, because it knows about anchors and occurrence counting.
- Rendering tracked changes in a browser. Reviewers open redlines in Word.

## Testing approach

Fixtures are generated in code (`tests/fixtures/create-docx.ts`) so no binary `.docx` files live in the repository and every fixture is reviewable in a diff. Unit tests cover each layer; integration tests cover whole flows. Tests that compare XML pin `revisionDate` and use fixed `paraId`s so output is deterministic.
