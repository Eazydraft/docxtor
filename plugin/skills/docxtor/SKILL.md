---
name: docxtor
description: Reads, edits and redlines Word .docx files with the docxtor command-line tool. Handles tracked changes with an author, comments anchored to text, accepting or rejecting tracked changes one by one, find-and-replace across the whole document, inserting and deleting paragraphs, tables, lists, formatting, headers and footers, and document properties. Leaves everything it does not touch byte for byte. It does not compare two versions of a document. Use when a task involves a .docx or Word document, including answering questions about what it says and citing sections, editing it as tracked changes or clean edits, reviewing or resolving someone else's tracked changes, adding comments, and removing author metadata before sending.
license: MIT
---

# docxtor

## Running it

The tool is `scripts/docxtor.mjs`, next to this file. It is self-contained and needs only Node.js 20 or newer:

```bash
node ${CLAUDE_SKILL_DIR}/scripts/docxtor.mjs <command> ...
```

If `${CLAUDE_SKILL_DIR}` is still literal text, use the absolute path of the folder that contains this SKILL.md. If `node --version` fails, Node.js must be installed first. If `docxtor` is installed globally (`npm install -g @lucio-ai/docxtor`), plain `docxtor <command>` works too. Below, `docxtor` means whichever of these works.

Use it for every read and write of a .docx. Do not unzip the file, edit its XML, or convert it with pandoc or python-docx: that loses formatting, numbering and tracked-change markup.

## Commands

| Task | Command |
|---|---|
| Read the document, find a passage | `docxtor view FILE [--grep REGEX [--context N]]` |
| List tracked changes | `docxtor revisions FILE` |
| List comments and what they're on | `docxtor comments FILE` |
| Change wording (tracked) | `docxtor replace FILE -o OUT --find TEXT --replace TEXT` |
| Add a comment | `docxtor comment FILE -o OUT --find TEXT --text COMMENT` |
| Insert or delete paragraphs; tables, lists, formatting, styles, headers, footers | `docxtor edit FILE -o OUT patches.json` |
| Check patches before writing | `docxtor edit FILE --dry-run patches.json` |
| Accept or reject tracked changes | `docxtor accept` / `reject FILE -o OUT [--ids ...] [--by AUTHOR] [--others ...]` |
| Read, set or scrub author, company and other properties | `docxtor metadata FILE [-o OUT --set key=value \| --scrub]` |

Every write command:

- writes **tracked changes** by default (`--untracked` edits directly);
- takes `--author NAME` (default `docxtor`). Use the user's name or organisation when known;
- writes to `-o OUT`. Use a new file unless the user asks to change the file in place; then `-o` is the input path;
- writes **nothing** if any part fails, and says why, including the paragraph's actual text when a find doesn't match. Fix the input and run the same command again.

Full options: [reference/commands.md](reference/commands.md). Every patch type for `edit`: [reference/patches.md](reference/patches.md). Worked sessions with real output: [reference/examples.md](reference/examples.md).

## Reading the output of `view`

```
[p144] (FirstParagraph) 11.1 Payment for each Shipment of Products shall be due within thirty (30) days …
[p210] {list:1017 L0} c. the following Clauses shall survive the termination …
| [tbl0.r1.c0.p0] Diesel | [tbl0.r1.c1.p0] $900 |
[hdr.default.p0] CONFIDENTIAL
[p9] (BodyText) Payment is due in {-30-}{+45+} days.
```

- `[p144]` is the paragraph id that patches target. Table cells are `tbl0.r1.c0.p0` (table, row, cell, paragraph). Headers and footers are `hdr.default.p0` and `ftr.default.p0`.
- `(FirstParagraph)` is the style; `{list:1017 L0}` is a list paragraph (numbering id, level).
- Numbering that Word generates, such as `c.` or `11.1`, is shown at the start of the text but is not part of it: don't put it in a `--find`.
- `{+…+}` is a tracked insertion and `{-…-}` a tracked deletion. The current text is the text without the `{-…-}` parts. A tracked table row ends with `{+row inserted by …+}` or `{-row deleted by …-}`.
- Ids are positions. After an edit that adds or removes paragraphs, run `view` again before using ids.

## Core workflow

Copy this checklist and tick it off for any edit:

```
- [ ] 1. Locate: view FILE --grep '<distinctive words>' (add --context 3 to see the clause around it)
- [ ] 2. Copy the exact text to change from the view output
- [ ] 3. Edit with the narrowest command (replace / comment / edit / accept / reject)
- [ ] 4. If the command refuses or fails, read its message, fix the input, run it again
- [ ] 5. Verify: revisions OUT (or view OUT --grep ...) shows exactly the intended changes
- [ ] 6. Report what changed, by its text, and the output file path
```

Step 2 matters: .docx text often has curly quotes (’ “ ”), non-breaking spaces, en dashes or section signs that differ from what you would type. `replace` tries straight and curly quote variants itself.

## Recipes

### Change wording

```bash
docxtor replace contract.docx -o contract-redline.docx --author "Acme" \
  --find "due within thirty (30) days" --replace "due within sixty (60) days" \
  --find "Acme Corp." --replace "Acme Holdings Inc."
```

- Make each `--find` long enough to be unambiguous. Only the words that differ are marked: the example tracks "thirty (30)" → "sixty (60)".
- A find that matches more than once is refused with a numbered list of the matches. Choose with `--occurrence N` (a number from that list), limit to one paragraph with `--in p12`, lengthen the find, or use `--all`.
- Several pairs apply in order, and each sees the earlier results. Nothing is written unless all of them apply.

### Rename a term throughout

```bash
docxtor replace agreement.docx -o agreement-redline.docx --find Seller --replace Vendor --all --whole-word
```

`--whole-word` skips words that merely contain the term (Sellers, Reseller). Possessives such as "Seller's" do match. The report gives the count and the paragraphs; `view` lists headers and footers, so check those too.

### Add a comment

```bash
docxtor comment msa.docx -o msa-commented.docx --author "Acme" \
  --find "shall not exceed the fees paid" --text "This cap is too low. We need twice the fees for the preceding 12 months."
```

A comment changes no text and adds no tracked change. Anchor it on a unique phrase from the clause, or use `--in p12` alone to comment on a whole paragraph. Check it with `docxtor comments msa-commented.docx`, which lists each comment with the passage it's on.

### Work through a reviewer's comments

```bash
docxtor comments draft.docx
```
```
2 comments
#1 [p228] by Jane 2026-11-04: Cap too low. Use twice annual fees.
    on: "shall not exceed an amount equal to the total value of Products"
#2 [p81] by Jane 2026-11-04: Three business days is enough.
    on: "five (5) Business Days"
```

Each entry gives the paragraph id and the exact passage, so it maps straight onto a `replace --in ID` edit. Comments stay in the file; mention which ones you addressed.

### Insert, delete or restructure paragraphs

Write the patches to a file, check them, then apply:

```bash
cat > /tmp/patches.json <<'EOF'
[
  { "type": "insert_paragraph", "target": "p12", "position": "after", "text": "12.5 The Supplier shall maintain insurance of at least USD 5,000,000." },
  { "type": "delete_paragraph", "target": "p14" }
]
EOF
docxtor edit msa.docx --dry-run /tmp/patches.json
docxtor edit msa.docx -o msa-redline.docx /tmp/patches.json
```

- `insert_paragraph` copies the target's style and list settings. In a list that Word numbers (`{list:…}` in `view`), the new paragraph becomes the next item automatically. Where numbers are typed as text, as in "12.5 …" above, write the number yourself.
- All patches in one `edit` apply together or not at all, and can target paragraphs created earlier in the same batch.
- Text is written exactly as given. To make part of it bold or italic, add `"markers": true` to that patch and use `**bold**` / `_italic_`, or use `apply_formatting`.

Tables, lists, formatting, styles, headers and footers: see [reference/patches.md](reference/patches.md).

### Resolve someone else's tracked changes

1. List them:
   ```bash
   docxtor revisions counterparty.docx
   ```
   ```
   #24 [p65] ids=25,26 by Counterparty: replace "any one or more" -> "at least two (2)"
       …downgraded below any of the following thresholds by [-any one or more-] [+at least two (2)+] of the following…
   ```
2. Decide each change against the user's rule. Read both the deleted and the inserted text: a change "affects an amount" if either side contains one.
3. Resolve everything in one call, passing **every id** of each chosen change:
   ```bash
   docxtor reject counterparty.docx -o resolved.docx --ids 25,26,27,28 --others accept
   ```
   Use `accept ... --others reject` for the opposite rule. Without `--others`, unlisted changes stay tracked.
4. Verify with `revisions resolved.docx`: with `--others`, it should report no tracked changes. Then report the rejected and accepted changes by their text.

Flag consequences you notice, such as a rejected definition whose term is still used elsewhere.

### Undo one of your own changes later

"We've decided to keep the original payment terms" means **reject** that tracked change, not edit over it:

```bash
docxtor revisions redline.docx                       # find your change and its ids
docxtor reject redline.docx -o redline.docx --ids 1,2
```

The file then reads as if that change was never made, and your other changes stay tracked. Someone may have edited the file since your last turn, so always run `revisions` or `view` again on an existing file before changing it.

### Comparing two versions: not supported

docxtor cannot compare two documents or build a redline from two versions. Don't try to approximate one with `replace` or `edit` patches: the result is unreliable. Tell the user it isn't supported, and suggest Word's Review > Compare (or another compare tool). docxtor can then work with the redline it produces: `revisions`, `accept`, `reject`.

### Produce a clean copy

```bash
docxtor accept redline.docx -o final.docx                       # all changes
docxtor accept redline.docx -o final.docx --by "Acme"           # only Acme's
docxtor metadata final.docx -o final.docx --scrub               # drop author, company, custom properties
docxtor revisions final.docx && docxtor metadata final.docx     # check: no changes left, no names
```

### Answer questions about a document

```bash
docxtor view credit-agreement.docx --grep 'conditions precedent' --context 2
docxtor view credit-agreement.docx | sed -n '/^\[p199\]/,/^\[p281\]/p'   # read from p199 to p281
```

Cite section numbers as `view` prints them. For long documents, grep for headings first, then read the range between them.

## When something goes wrong

| Message | What to do |
|---|---|
| `"X" matches N times` | Pick with `--occurrence N` or `--in ID`, or lengthen the find |
| `"X" not found` | `view FILE --grep '<one distinctive word>'` and copy the exact text; it may sit inside a tracked deletion, which is no longer current text |
| `not found in paragraph "p12"` (from `edit`) | The error prints the paragraph's current text: copy from it |
| `Paragraph "p99" not found` | Ids shifted or are wrong: run `view` again |
| `no tracked change has w:id ...` | Take the ids from `revisions` on the same file |
