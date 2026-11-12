# Command reference

All commands: `docxtor <command> ...` (see SKILL.md for how to run it). Errors go to stderr with exit code 1; `docxtor help` prints a summary.

## Contents

- Reading: [view](#view), [revisions](#revisions), [comments](#comments), [metadata (read)](#metadata)
- Writing: [common options](#common-write-options), [replace](#replace), [comment](#comment), [edit](#edit), [accept / reject](#accept--reject), [metadata (write)](#metadata)

## view

```
docxtor view FILE [--grep REGEX [--context N]] [--no-revisions] [--formatting] [--styles] [--json]
```

One line per paragraph, prefixed with its id. Table rows print as `| [tbl0.r1.c0.p0] text | [tbl0.r1.c1.p0] text |`, with an id before each paragraph of a cell that has several (`| [tbl0.r1.c0.p0] Heading [tbl0.r1.c0.p1] Body |`) and `{+row inserted by …+}` or `{-row deleted by …-}` after a tracked row; the rows of a table nested in a cell follow the row that holds it, with ids under the cell's (`| [tbl0.r1.c0.tbl0.r0.c0.p0] text |`); headers and footers follow the body under `--- Header (default) ---` lines, then footnotes and endnotes under `--- Footnotes ---` and `--- Endnotes ---` with ids `fn1.p0` and `en1.p0`. In the text, a tab prints as a tab character, and a footnote or endnote reference as `[^1]` or `[^e1]`. The number is the one Word displays, the position of the reference in the text, not the `w:id` stored in the file; `[^1]` is the note listed as `fn1.p0` and `[^e1]` the one listed as `en1.p0`. The mark is not part of the text. An empty form field prints as `{field: prompt}` (`{field date: …}`, `{field dropdown "A" OR "B": …}` for typed ones) and a chosen date or list item as `{field date= 1 May 2026}`.

| Option | Effect |
|---|---|
| `--grep REGEX` | Only lines matching REGEX (case-insensitive, JavaScript syntax: `a\|b`, `\d+`, `\(`). |
| `--context N` | With `--grep`, also N lines before and after each match; groups are separated by `--`. |
| `--no-revisions` | Current text only, without `{+ +}` / `{- -}` marks. |
| `--formatting` | Shows bold, italic, strikethrough and underline as `**b**`, `_i_`, `~~s~~`, `<u>u</u>`. |
| `--styles` | Appends the document's styles, for `apply_style`. |
| `--json` | Structured output: ids, text, styles, list info, revisions (with ids) and, with `--formatting`, run formatting. |

For long documents, `--grep` for headings first, then read a range: `docxtor view FILE | sed -n '/^\[p199\]/,/^\[p281\]/p'`. Piping into `head`, `grep` or `sed` is fine.

## revisions

```
docxtor revisions FILE [--json]
```

```
39 changes by "Author"
#24 [p65] ids=25,26 by Author: replace "any one or more" -> "at least two (2)"
    …thresholds by [-any one or more-] [+at least two (2)+] of the following…
#38 [p73] ids=37 by Author: insert "Drafting note (Rebecca Tsao, …): …"
    whole paragraph inserted: [+Drafting note …+]
```

- One numbered entry per change, in document order: replacement (two ids), insertion, deletion, whole paragraph inserted or deleted (text and paragraph-mark ids together), formatting change, or `other` (table rows, cells, sections) under `[-]`.
- Pass **all** ids of a change to `accept` / `reject`.
- `--json` gives `n`, `paragraph`, `kind`, `ids`, `authors`, `date`, `deleted`, `inserted`, `context` per change, which is useful for filtering many changes with a script.

## comments

```
docxtor comments FILE [--json]
```

```
1 comment
#1 [p228] by Acme 2026-11-06: The cap is too low for us. …
    on: "shall not exceed an amount equal to the total value of Products actually delivered"
```

Every comment in document order: the paragraph where its passage starts, author, date, text (paragraphs joined with ` / `), and the passage it is on (empty for a comment on a single point). `--json` adds `id` and `initials`. Read-only: comments are added with `comment` and are not edited or deleted by any command.

## metadata

```
docxtor metadata FILE [--json]                              # read
docxtor metadata FILE -o OUT [--set key=value ...] [--scrub] # write
```

The read form also shows `company` and `manager` from the application properties. Keys for `--set`: `creator`, `lastModifiedBy`, `title`, `subject`, `description`, `keywords`, `category`, `contentStatus`, `revision`, `created`, `modified`. `--scrub` removes `creator`, `lastModifiedBy`, `description`, `keywords`, `category`, `contentStatus`, the company and manager, and custom properties, and resets `revision` to 1; title and dates stay. `--set` applies after `--scrub`, so `--scrub --set title="Final"` works.

## Common write options

| Option | Effect |
|---|---|
| `-o OUT` | Required. The file to write; may be the input path to update it in place. |
| `--author NAME` | Author of tracked changes and comments. Default `docxtor`. |
| `--untracked` | Edit directly instead of as tracked changes (`replace`, `edit`). |

Nothing is written when any part fails.

## replace

```
docxtor replace FILE -o OUT --find TEXT --replace TEXT [--find TEXT --replace TEXT ...]
                [--all] [--occurrence N] [--in ID] [--whole-word]
```

```
"due within thirty (30) days" -> "due within sixty (60) days" (marked: "thirty (30)" -> "sixty (60)"): 1 replacement in 1 paragraph (p144)
wrote redline.docx (tracked, author "Acme")
```

- Searches the body, table cells, headers, footers, footnotes and endnotes.
- Marks only the words that differ between `--find` and `--replace`.
- Tries straight/curly quote variants when the literal text isn't found, then lets a space match a tab: `--find "17.2 For goods"` finds `17.2⇥For goods`. The tab is kept, and the report shows the text that matched.
- An empty `--replace ""` deletes the text.
- A find on the prompt of an empty form field fills the field with the whole `--replace` text.
- A find stays inside one paragraph. When nothing matches, the error says so if the text is in another paragraph than `--in` (`"X" is not in p12; it is in p13. Use --in p13.`) or runs across a paragraph break (`"X" runs across two paragraphs (tbl0.r3.c1.p0 and tbl0.r3.c1.p1)`): split such a find into one `--find`/`--replace` pair per paragraph. `comment` reports the same.

| Option | Effect |
|---|---|
| `--all` | Replace every match. |
| `--occurrence N` | Only the Nth match, numbered as in the refusal list. |
| `--in ID` | Only inside paragraph ID (`p12` or `para:4F2A1C7B`). |
| `--whole-word` | Skip matches inside longer words. |

Ambiguous find, refused:

```
error: "Seller" matches 83 times; nothing was written. Use --all, --in ID or --occurrence N (counted in this list):
  1. [p5] CASPIAN ENERGY TRADING FZE (as Seller)
  2. [p10] …(hereinafter referred to as the “Seller” or “CET”);
  …
```

## comment

```
docxtor comment FILE -o OUT --text TEXT (--find PASSAGE [--in ID] [--occurrence N] | --in ID)
```

Anchors a Word comment to PASSAGE, which must match once (narrow it as for `replace`; a space matches a tab here too), or to the whole paragraph with `--in ID` alone. `\n` in TEXT starts a new comment paragraph. `--author` sets the comment author.

## edit

```
docxtor edit FILE -o OUT [PATCHES.json | -]
docxtor edit FILE --dry-run [PATCHES.json | -]
```

Applies a JSON array of patches from a file or stdin; see [patches.md](patches.md). `--dry-run` checks every patch against a copy and prints:

```
1. ok append_text
2. FAILS replace_cell_text tbl0.r1.c0.p0
   -> Text "LSFO" not found in paragraph "tbl0.r1.c0.p0" (occurrence 1)
       tbl0.r1.c0.p0 currently reads: "Jet A-1"
1 of 2 would fail; fix them and run again.
```

"Currently reads" is the paragraph as it stood when that patch ran, after the earlier patches in the batch.

## accept / reject

```
docxtor accept FILE -o OUT [--ids 1,2,...] [--by AUTHOR ...] [--others reject|keep]
docxtor reject FILE -o OUT [--ids 1,2,...] [--by AUTHOR ...] [--others accept|keep]
```

```
rejected 13 revisions, accepted the other 34; 0 changes still tracked
wrote resolved.docx
```

| Option | Effect |
|---|---|
| (none) | Every tracked change. |
| `--ids` | Only revisions with these w:ids from `revisions` (comma or space separated). Unknown ids are refused. |
| `--by AUTHOR` | Only this author's revisions (repeatable). Combines with `--ids`. |
| `--others` | Resolve every remaining change the opposite way. Default `keep`: they stay tracked. |
