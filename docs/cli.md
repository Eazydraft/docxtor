# Command reference

`docxtor` is the command the agent skill drives. You can run it yourself too, after `npm install -g @lucio-ai/docxtor`, or with `npx @lucio-ai/docxtor`. The agent skill runs its own bundled copy, `scripts/docxtor.mjs`. It needs Node.js 20 or newer.

```
docxtor <command> [options]
```

`docxtor help` prints a summary. Errors go to stderr with exit code 1.

## Reading

### `view FILE`

Prints the document as one line per paragraph, each with the id that `edit` patches and `--in` options target: `[p12]` for body paragraphs, `[tbl0.r1.c0.p0]` for a paragraph in a table cell, `[hdr.default.p0]` and `[ftr.default.p0]` for headers and footers, `[fn1.p0]` and `[en1.p0]` for footnotes and endnotes. In a document with more than one section the ids carry the section, `[hdr.s2.first.p0]`. Paragraphs that carry a native Word id show it as `¶4F2A1C7B`, addressable as `para:4F2A1C7B`. Tracked table rows end with `{+row inserted by …+}` or `{-row deleted by …-}`. Style names and list levels are shown, and list numbering appears at the start of the text as Word renders it. Form fields (content controls) show as `{field: prompt}` while empty, with the kind and the list items for a date picker or a drop-down list (`{field dropdown "A" OR "B": Choose an item.}`), and as `{field date= 1 May 2026}` once a date or a list item is chosen; see the [API reference](api.md#docprojectformat-options).

A tab in the text prints as a tab character, so a clause typed as a number, a tab and its text reads `17.2⇥For goods…`. A footnote reference prints as `[^1]` and an endnote reference as `[^e1]`, numbered as Word displays them; the mark is not part of the text. The notes are listed after the footers under `--- Footnotes ---` and `--- Endnotes ---`, and `[^1]` is the note whose first paragraph is `[fn1.p0]`.

| Option | |
|---|---|
| `--grep REGEX` | only paragraphs matching REGEX (case-insensitive) |
| `--context N` | with `--grep`, also N paragraphs before and after each match, groups separated by `--` |
| `--no-revisions` | show the current text without `{+inserted+}` / `{-deleted-}` marks |
| `--formatting` | mark bold, italic, strikethrough and underline inline (`**bold**`, `_italic_`, `~~strike~~`, `<u>underline</u>`) |
| `--styles` | append the document's available styles |
| `--json` | the JSON projection instead (see [API reference](api.md#docprojectformat-options)) |

### `revisions FILE`

Lists every tracked change in document order:

```
39 changes by "Counterparty"
#24 [p65] ids=25,26 by Counterparty: replace "any one or more" -> "at least two (2)"
    …downgraded below any of the following thresholds by [-any one or more-] [+at least two (2)+] of the following…
```

A deletion followed by an insertion is one replacement with both ids, even when a space separates them. An inserted or deleted whole paragraph is one change, with its paragraph mark's id included. Formatting and paragraph-property changes are listed too, along with any remaining marked elements, such as table rows, under `[-]`. A change inside a footnote or endnote is listed under the note's paragraph, `[fn1.p0]`. `--json` gives the same list as JSON.

### `comments FILE`

Lists the comments in the document: the paragraph where each comment's passage starts, author, date, text, and the passage itself. `--json` for JSON.

## Writing

Every write command (`replace`, `comment`, `edit`, `accept`, `reject`, `metadata --set/--scrub`):

- takes `-o OUT`, the file to write. It may be the input file, to update it in place;
- writes nothing if any part fails, and says why.

`replace` and `edit` write tracked changes unless you pass `--untracked`, and take `--author NAME` for the revision author (default `docxtor`). `comment` takes `--author` for the comment author and changes no text. `accept`, `reject` and `metadata` take neither option: they resolve existing tracked changes or change document properties, and create no tracked changes.

Text is written verbatim: `*`, `_` and `~` are not read as formatting.

### `replace FILE -o OUT --find TEXT --replace TEXT`

Finds TEXT anywhere: body, tables, headers, footers, footnotes and endnotes.

- Repeat `--find`/`--replace` pairs to make several edits in one call. They apply in order, and each sees the results of the ones before it.
- Only what differs is marked. `--find "within thirty (30) days" --replace "within sixty (60) days"` tracks "thirty (30)" → "sixty (60)"; a pure deletion or insertion is marked exactly, so `--find "including, without limitation," --replace "including"` deletes ", without limitation,".
- If the text isn't found exactly, straight and curly quote variants are tried.
- A find on the prompt of an empty form field fills the field: the whole prompt is replaced by the whole `--replace` text, even when the find covers only part of the prompt. A drop-down list takes one of its items only. See [text patches inside a content control](patches.md#text-patches-inside-a-content-control).
- A space and a tab are then treated as the same character, so `--find "17.2 For goods"` matches a clause that Word holds as `17.2⇥For goods` and no literal tab is needed in a shell argument. The tab is kept: where `--replace` repeats the start or end of `--find`, the document's own characters are used there, and only the words that differ are marked. The report shows the text that matched. (Patches given to `edit` match exactly; write the tab as `\t` in the JSON.)
- A find that matches more than once is refused with a numbered list of the matches, unless one of these says which:

| Option | |
|---|---|
| `--all` | replace every match |
| `--occurrence N` | the Nth match in the refusal's list |
| `--in ID` | only inside this paragraph (`p12`, `para:4F2A1C7B`) |
| `--whole-word` | skip matches inside longer words ("Seller" but not "Sellers" or "Reseller") |

### `comment FILE -o OUT --text TEXT --find PASSAGE`

Adds a Word comment anchored to PASSAGE, which must match exactly once; `--in ID` and `--occurrence N` narrow it as for `replace`, and a space in PASSAGE matches a tab in the same way. Pass `--in ID` without `--find` to comment on a whole paragraph. The document's text is not changed and no tracked change is added. `\n` in TEXT starts a new comment paragraph.

### `edit FILE -o OUT [PATCHES.json | -]`

Applies a JSON array of patches, read from a file or, with `-` or no file, from stdin. They apply all or nothing, in order. When one fails, the error shows the target paragraph's text as it stood at that point in the batch. `--dry-run` (instead of `-o`) checks every patch and reports each as `ok` or `FAILS` without writing. Every [patch type](patches.md) works, for example:

```json
[
  { "type": "insert_paragraph", "target": "p12", "position": "after", "text": "12.5 New clause text." },
  { "type": "delete_paragraph", "target": "p14" },
  { "type": "insert_text", "target": "p3", "position": "after", "anchor": "Agreement", "text": " (as amended)" }
]
```

### `accept FILE -o OUT` / `reject FILE -o OUT`

Accept or reject tracked changes: all of them by default, or only the ones selected.

| Option | |
|---|---|
| `--ids 14,15` | only revisions with these ids, from `revisions`; pass every id of a change |
| `--by AUTHOR` | only revisions by this author (repeatable) |
| `--others reject` / `--others accept` | resolve every other change the opposite way (default: leave them tracked) |

`reject FILE -o OUT --ids 14,15,22,23 --others accept` resolves a whole review in one step. The output says how many changes are still tracked.

### `metadata FILE`

Prints the document's properties (author, last modified by, title, subject, company, manager, dates, …); `--json` for JSON.

`metadata FILE -o OUT --set key=value [--set …]` sets properties (`creator`, `lastModifiedBy`, `title`, `subject`, `description`, `keywords`, `category`, `contentStatus`, `revision`, `created`, `modified`). `--scrub` removes identifying ones (creator, last modified by, description, keywords, category, content status, company, manager and custom properties) and resets the revision count; title and dates stay. `--set` applies after `--scrub`.
