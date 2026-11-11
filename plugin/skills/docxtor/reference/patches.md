# Patch reference for `docxtor edit`

`edit` applies a JSON array of patches. Each patch has a `type` and, usually, a `target`: a paragraph id from `view` (`p12`, `tbl0.r1.c0.p0`, `hdr.default.p0`, `fn1.p0`) or a native id (`para:4F2A1C7B`). In a document with more than one section the header and footer ids carry the section, `hdr.s2.first.p0`.

## Contents

- [How a batch behaves](#how-a-batch-behaves)
- Text in a paragraph: [replace_text](#replace_text), [insert_text](#insert_text), [delete_text](#delete_text), [replace_cell_text](#replace_cell_text)
- Paragraphs: [insert_paragraph](#insert_paragraph), [delete_paragraph](#delete_paragraph), [append_text](#append_text)
- Form fields: [text in a field](#text-in-a-form-field), [set_checkbox](#set_checkbox)
- Comments: [add_comment](#add_comment)
- Formatting and styles: [apply_formatting](#apply_formatting), [apply_style](#apply_style)
- Tables: [add_table](#add_table), [add_table_row](#add_table_row), [delete_table_row](#delete_table_row)
- Lists: [add_list](#add_list), [add_list_item](#add_list_item), [remove_list_item](#remove_list_item), [convert_to_list](#convert_to_list), [remove_list](#remove_list), [set_list_properties](#set_list_properties)
- Headers and footers: [add_header / add_footer](#add_header--add_footer), [remove_header / remove_footer](#remove_header--remove_footer)
- [Footnotes and endnotes](#footnotes-and-endnotes)
- [Text options: markers, line breaks, tabs, fields](#text-options)

## How a batch behaves

- **Tracked by default.** Every content change is a tracked change by `--author`. `--untracked` edits directly. Comments and removing a header or footer are never tracked (Word doesn't track them).
- **All or nothing.** If any patch fails, nothing is written. The error names the patch and, for a text that isn't found, prints what the paragraph says at that point.
- **In order.** Patches apply one after another, and each sees the results of the earlier ones. A new paragraph gets the next free id (an insert after `p12` in a 40-paragraph document becomes `p40`) and can be targeted by later patches. Adding a table row shifts the rows below it within the batch: after inserting at row 1, the old row 1 is row 2.
- **Check first:** `docxtor edit FILE --dry-run patches.json` reports each patch as `ok` or `FAILS` with the reason, and writes nothing.
- `find`, `anchor` and `occurrence`: `find` must match the paragraph's current text exactly (without the `{-deleted-}` parts and the `[^1]` footnote marks shown by `view`). A tab in the text is `\t` in the JSON: `"find": "17.2\tFor goods"`. Unlike `replace --find`, a space here does not match a tab. `occurrence` picks the Nth match (default 1).

## Text in a paragraph

### replace_text

```json
{ "type": "replace_text", "target": "p12", "find": "thirty (30) days", "replace": "sixty (60) days" }
```

Keeps the formatting of the replaced text. Tracked: the old text becomes a deletion and the new text an insertion. Replacing text that is still an unaccepted insertion by the same author rewrites that insertion instead of stacking a deletion on it. For the whole document at once, the `replace` command is simpler.

### insert_text

```json
{ "type": "insert_text", "target": "p12", "position": "after", "anchor": "the Products", "text": " and the Services" }
```

`position` is `before` or `after` the `anchor`.

### delete_text

```json
{ "type": "delete_text", "target": "p12", "find": ", without limitation,", "occurrence": 1 }
```

### replace_cell_text

```json
{ "type": "replace_cell_text", "target": "tbl0.r2.c1.p0", "find": "TBD", "replace": "USD 10.00" }
```

The same as `replace_text`, for a paragraph inside a table cell. `replace_text` also accepts cell ids.

## Paragraphs

### insert_paragraph

```json
{ "type": "insert_paragraph", "target": "p12", "position": "after", "text": "12.5 The Supplier shall maintain insurance." }
```

- `position`: `before` or `after`. Works in the body, inside a table cell, and in headers and footers.
- Copies the target's style, list settings, indentation and spacing, and the formatting of its last run, as pressing Enter in Word does. `"inheritStyle": false` turns that off.
- In a list Word numbers itself, the new paragraph becomes the next item. Typed numbers ("12.5") are part of the text.
- `\n` in `text` creates several consecutive paragraphs. Use `add_table` or `add_list` for tables and lists.
- Tracked: inserted text and its paragraph mark, so accepting leaves exactly the new paragraph and rejecting leaves nothing behind.

### delete_paragraph

```json
{ "type": "delete_paragraph", "target": "p14" }
```

Tracked: the whole paragraph, including its mark, is a deletion. Untracked: the paragraph is removed (a table cell's only paragraph is emptied instead).

### append_text

```json
{ "type": "append_text", "text": "Schedule 9\n| Product | Price |\n| --- | --- |\n| VLSFO | USD 900 |\n- first point\n- second point" }
```

Adds content at the end of the body: each line becomes a paragraph, a markdown table (with a `| --- |` line) becomes a Word table, and `- ` or `1. ` lines become a real list (indent two spaces per level). Add `"container": { "type": "footer", "variant": "default" }` to append to a header or footer instead. If several sections have that header or footer, add `"section": 2` (the number shown in the `view` divider, `--- Footer (default, section 2) ---`); without it the patch fails and lists the sections.

## Form fields

### Text in a form field

`view` shows an empty form field (a Word content control) as `{field: prompt}`, and a date picker or drop-down list with its kind and items. The text inside the braces is paragraph text, so the text patches above reach it:

```json
{ "type": "replace_text", "target": "p7", "find": "Insert Title for Contract", "replace": "Supply of Widgets" }
```

- On a prompt, the patch fills the field: the whole prompt is removed and the replacement becomes the field's text, even when `find` is only part of the prompt. `insert_text` anchored on a prompt fills it too. Tracked: one insertion, no deletion.
- In a filled text field, patches edit the text as anywhere else.
- A drop-down list must end up holding one of its items; otherwise the patch fails and names them. A date field takes any text and stores the date when the text is one (`2026-11-11`, `11 November 2026`).
- Refused with a reason: emptying a field, a `find` that crosses the edge of a field, a locked field, a field bound to document data, a picture field.

### set_checkbox

```json
{ "type": "set_checkbox", "target": "p40", "checked": true, "occurrence": 1 }
```

Ticks or unticks a checkbox field in the paragraph (`occurrence` picks among several, default 1). Replacing `☐` with `☒` through `replace_text` does the same. Only Word checkbox content controls count; a plain `☐` character is ordinary text.

## Comments

### add_comment

```json
{ "type": "add_comment", "target": "p228", "find": "shall not exceed", "text": "Too low.\nPlease propose twice the annual fees.", "author": "Acme", "initials": "AC" }
```

`find` (and `occurrence`) anchor the comment to a passage; without it the comment covers the whole paragraph. `\n` starts a new comment paragraph. `author` defaults to `--author`. Comments add no tracked change.

## Formatting and styles

### apply_formatting

```json
{ "type": "apply_formatting", "target": "p12", "find": "Effective Date", "formatting": { "bold": true, "highlight": "yellow" } }
```

Without `find`, formats the whole paragraph. Properties: `bold`, `italic`, `strikethrough` (true or false; `null` removes), `underline` (`"single"`, `"double"`, `"wave"`, `"dash"`, … or `null`), `fontSize` in half-points (24 = 12 pt), `fontFamily`, `color` (hex such as `"C00000"`), `highlight` (`"yellow"`, `"green"`, `"cyan"`, …). Unlisted properties are left alone. Tracked as a formatting change.

### apply_style

```json
{ "type": "apply_style", "target": "p14", "style": "Heading2" }
{ "type": "apply_style", "target": "p12", "find": "Seller", "style": "Strong" }
```

Without `find`: a paragraph style. With `find`: a character style for that text. The style must already exist; `view FILE --styles` lists the document's styles. Tracked as a formatting change.

## Tables

### add_table

```json
{ "type": "add_table", "afterId": "p40", "headers": ["Product", "Price"], "rows": [["LSFO", "USD 900"], ["GO", "USD 950"]] }
```

A bordered table after `afterId` (a paragraph or table id), or at the end of the body without it. `headers` is optional; every row needs the same number of cells. Tracked: each row is an inserted row.

### add_table_row

```json
{ "type": "add_table_row", "target": "tbl0", "afterRow": 1, "cells": ["Jet A-1", "60,000 MT", "~5,000 MT", "USD 10.00"] }
```

`afterRow` is 0-based (`-1` inserts at the top). The cell count must match the table's; the new row copies the last row's cell properties, as plain cells: a form field in the copied row is not repeated. Rows below shift down for later patches.

### delete_table_row

```json
{ "type": "delete_table_row", "target": "tbl0", "row": 2 }
```

Tracked: the row stays, marked deleted.

## Lists

In `view`, list paragraphs show `{list:N Lx}`: `N` is the numbering id, `Lx` the level.

### add_list

```json
{ "type": "add_list", "afterId": "p40", "listType": "ordered", "items": ["First", { "text": "Sub-point", "level": 1 }, "Second"] }
```

`listType`: `bullet`, `ordered` or `multilevel`. Add `"numId": "1017"` to continue an existing list's numbering.

### add_list_item

```json
{ "type": "add_list_item", "target": "p59", "position": "after", "text": "any other matter agreed in writing", "level": 0 }
```

The target must be a list paragraph; the new item joins its list, and Word numbers it (don't type the number). `level` defaults to the target's.

### remove_list_item

```json
{ "type": "remove_list_item", "target": "p61" }
```

### convert_to_list

```json
{ "type": "convert_to_list", "targets": ["p20", "p21", "p22"], "listType": "bullet", "level": 0 }
```

### remove_list

```json
{ "type": "remove_list", "targets": ["p20", "p21", "p22"] }
```

### set_list_properties

```json
{ "type": "set_list_properties", "target": "p59", "level": 1, "listType": "ordered" }
```

Changes an existing list paragraph's level and/or list type.

## Headers and footers

### add_header / add_footer

```json
{ "type": "add_footer", "variant": "default", "paragraphs": ["Page {{PAGE}} of {{NUMPAGES}}"] }
```

`variant`: `default` (every page), `first` (first page only) or `even`. Applies to the last section. To edit an existing header or footer, use its paragraph ids (`hdr.default.p0`) with the text patches above.

### remove_header / remove_footer

```json
{ "type": "remove_header", "variant": "first" }
```

## Footnotes and endnotes

```json
{ "type": "replace_text", "target": "fn1.p0", "find": "World Bank", "replace": "Bank Group" }
```

The paragraphs of a note are `fn1.p0`, `fn1.p1` (footnote 1) and `en1.p0` (endnote 1), listed at the end of `view`; the number is the one in the note's mark in the text, `[^1]` or `[^e1]`. `replace_text`, `insert_text`, `delete_text`, `apply_formatting`, `apply_style`, `add_comment` and `insert_paragraph` work on them as on any paragraph, tracked or not.

Adding a note and deleting a note are not supported. `delete_paragraph` on the paragraph that holds a note's number, or on a note's only paragraph, fails; use `delete_text` to remove its text. The reference in the body is not text and cannot be matched, moved or removed with a patch; edits beside it leave it in place.

## Text options

- **Text is written exactly as given.** `*`, `_` and `~` are ordinary characters. To write formatted text, add `"markers": true` to that patch and use `**bold**`, `_italic_`, `~~strikethrough~~` or `<u>underline</u>`:
  ```json
  { "type": "insert_paragraph", "target": "p12", "position": "after", "text": "**Note:** prices exclude VAT.", "markers": true }
  ```
- **Line breaks.** In `replace_text`, `insert_text` and similar, `\n` is a line break inside the paragraph. In `insert_paragraph` and `append_text`, `\n` starts a new paragraph.
- **Tabs.** `\t` is a tab, in every text field: `find` and `anchor` match one, `replace` and `text` write one. To change a tabbed clause and keep its tab, put `\t` on both sides: `"find": "17.2\tFor goods", "replace": "17.2\tFor Goods"`.
- **Fields.** `{{PAGE}}`, `{{NUMPAGES}}`, `{{DATE}}`, `{{TIME}}` and `{{FILENAME}}` become live Word fields in `append_text`, `insert_paragraph`, `add_header` and `add_footer`.
