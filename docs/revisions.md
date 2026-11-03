# Revision tracking (tracked changes)

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

## What each patch type produces

| Patch type | OOXML markup |
|---|---|
| `replace_text` | `w:del` wrapping the original runs + `w:ins` wrapping the replacement runs |
| `insert_text` | `w:ins` wrapping the new runs |
| `delete_text` | `w:del` wrapping the deleted runs (text preserved in XML) |
| `append_text` | `w:ins` wrapping each new paragraph's run; markdown tables get `w:ins` in each row's `w:trPr` |
| `insert_paragraph` | `w:ins` wrapping the new runs **and** `w:ins` on the paragraph mark (`w:pPr/w:rPr/w:ins`) |
| `delete_paragraph` | All runs wrapped in `w:del` (`w:delText`) plus `w:del` on the paragraph mark (`w:pPr/w:rPr/w:del`); own unaccepted insertions removed outright |
| `add_comment` | No revision markup — comments are annotations, not changes |
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
| `add_list` | `w:ins` wrapping each new list paragraph's runs |
| `add_list_item` | `w:ins` wrapping the new item's runs |
| `remove_list_item` | Runs cloned and wrapped in `w:del` (with `w:delText`); `w:pPrChange` on paragraph properties. Runs inside `w:ins` are removed entirely. |
| `convert_to_list` | `w:pPrChange` capturing old paragraph properties |
| `remove_list` | `w:pPrChange` capturing old paragraph properties |
| `set_list_properties` | `w:pPrChange` capturing old paragraph properties |

## Revision IDs

Each revision element gets a unique `w:id`. IDs start from `max(existing w:id values in the document) + 1` so they never collide with pre-existing tracked changes.

## Reading revision marks

The normalizer captures revision marks when present. Use `includeRevisions` in projections to surface them:

```ts
// JSON -- each paragraph gains a `revisions` array
const json = doc.project("json", { includeRevisions: true });

// Markdown -- inline annotations like {+inserted+} and {-deleted-}
const md = doc.project("markdown", { includeRevisions: true });
```

## Reading coloured runs (e.g. answers typed in red)

The JSON projection can report exactly the runs matching a formatting filter, with their paragraph id, `paraId` and character offsets. Black text in the same paragraph is excluded; paragraphs with no matching run have no `formatting` array; table cells get a `paragraphs` array so their matches are reported too.

```ts
const json = doc.project("json", { runFilter: { color: "FF0000" } });
// body[9].formatting → [{ text: "two months", startOffset: 25, endOffset: 35, color: "FF0000" },
//                       { text: "Parent Co",  startOffset: 48, endOffset: 57, color: "FF0000", themeColor: "accent2" }]
```

Colour comparison is case-insensitive on the raw `w:color w:val`, which Word also writes for theme-mapped colours (the resolved RGB), so theme reds and literal reds match alike; `themeColor` is exposed when present. The filter may also include `bold`, `italic`, `underline`, `strikethrough` and `highlight`; every property given must match.

## Behavior notes

- Without `trackRevisions`, all patches work exactly as before (destructive edits). The option is fully backward-compatible.
- `acceptAllRevisions` after a tracked batch yields the same text (and, for `replace_text`, `insert_paragraph`, `delete_paragraph` and `replace_cell_text`, the same `document.xml`) as applying the batch untracked — the engine is a fixed point under accept.
- Sequential patches within the same `applyPatches()` call share a single revision context: IDs auto-increment, and re-normalization between patches correctly handles the `w:ins`/`w:del` wrappers.
- The normalizer's `collectRuns()` recurses into `w:ins` (treats the content as current text) and skips `w:del` (excludes deleted text from `.text`). This means subsequent patches in the same batch correctly see the post-edit text.
- Replacing or deleting text that is itself an unaccepted insertion produces no `w:del` — the inserted text is simply removed. If the insertion belongs to another author, their `w:ins` is split around the edit and the replacement sits between the halves in its own `w:ins`, so order and attribution are both kept. If it is the batch author's own insertion, the replacement joins it in place.
