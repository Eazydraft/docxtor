# Worked examples

Real sessions on real agreements, with each request, the commands run, and the output. Paths are shortened.

## Contents

1. [Two edits as tracked changes](#1-two-edits-as-tracked-changes)
2. [Rename a defined term throughout](#2-rename-a-defined-term-throughout)
3. [Comment on a clause](#3-comment-on-a-clause)
4. [Resolve a counterparty's changes by a rule](#4-resolve-a-counterpartys-changes-by-a-rule)
5. [Undo one of your own changes later](#5-undo-one-of-your-own-changes-later)
6. [Add a clause and a table row, fixing a failed patch](#6-add-a-clause-and-a-table-row-fixing-a-failed-patch)
7. [Answer a question with citations](#7-answer-a-question-with-citations)
8. [Final clean copy](#8-final-clean-copy)

## 1. Two edits as tracked changes

> Change the payment terms to net 45 days and the order acceptance period to three business days, as tracked changes, in a new file.

```bash
docxtor view msa.docx --grep 'days (of|from)|payment' | cut -c1-160
```
```
[p81] (BodyText) 3.4 The Seller shall confirm acceptance or rejection of each Purchase Order in writing within five (5) Business Days of receipt. …
[p144] (FirstParagraph) 11.1 Payment for each Shipment of Products shall be due within thirty (30) days from the date of the Bill of Lading …
```
```bash
docxtor replace msa.docx -o msa-redline.docx --author "Acme" \
  --find "due within thirty (30) days" --replace "due within forty-five (45) days" \
  --find "within five (5) Business Days of receipt" --replace "within three (3) Business Days of receipt"
docxtor revisions msa-redline.docx
```
```
2 changes by "Acme"
#1 [p81] ids=3,4 by Acme: replace "five (5)" -> "three (3)"
    …in writing within [-five (5)-][+three (3)+] Business Days of receipt. …
#2 [p144] ids=1,2 by Acme: replace "thirty (30)" -> "forty-five (45)"
    …shall be due within [-thirty (30)-][+forty-five (45)+] days from the date of the Bill of Lading …
```

The finds are long enough to be unique; only the differing words are marked.

## 2. Rename a defined term throughout

> Change 'Seller' to 'Vendor' throughout, as tracked changes.

```bash
docxtor replace msa.docx -o msa-vendor.docx --find Seller --replace Vendor --all --whole-word
```
```
"Seller" -> "Vendor": 83 replacements in 50 paragraphs (p5, p10, p13, p15, p17, p32, p45, p49, p73, p80, p81, p84, …)
wrote msa-vendor.docx (tracked, author "docxtor")
```

Check: `docxtor view msa-vendor.docx --no-revisions --grep seller` should find nothing left except words you meant to keep.

## 3. Comment on a clause

> Add a comment to the limitation of liability clause saying the cap is too low.

```bash
docxtor view msa.docx --grep 'LIMITATION OF LIABILITY' --context 3 | cut -c1-140
```
```
[p226] (Heading2) CLAUSE 18 — LIMITATION OF LIABILITY
[p227] (FirstParagraph) 18.1 EXCLUSION OF INDIRECT DAMAGES. NEITHER PARTY SHALL BE LIABLE …
[p228] (BodyText) 18.2 Subject to Clause 18.3, the aggregate liability of either Party … shall not exceed an amount equal to the total value of Products …
```
```bash
docxtor comment msa.docx -o msa-commented.docx --author "Acme" --in p228 \
  --find "shall not exceed an amount equal to the total value of Products" \
  --text "The cap is too low for us. We want it raised to twice the value of Products delivered in the preceding 12 months."
```
```
comment on [p228] anchored to "shall not exceed an amount equal to the total value of Products" by "Acme"
```

## 4. Resolve a counterparty's changes by a rule

> Reject the changes to the credit rating thresholds and accept everything else.

```bash
docxtor revisions confirmation.docx | cut -c1-140
```
```
39 changes by "Author"
…
#24 [p65] ids=25,26 by Author: replace "any one or more" -> "at least two (2)"
#25 [p66] ids=27,28 by Author: replace "BBB−" -> "BB"
#26 [p67] ids=29,30 by Author: replace "Baa3" -> "Ba2"
#27 [p68] ids=31,32 by Author: replace "BBB−" -> "BB"
#28 [p69] ids=33,34 by Author: replace "For the avoidance of doubt, a downgrade of … by any single Rating Agency …"
#29 [p72] ids=35,36 by Author: replace "Accordingly, the trigger thresholds of BBB−/Baa3 are one notch below …"
…
```

#25–#27 change the rating levels; #29 restates them. #24 and #28 change how many agencies must downgrade, which is part of the threshold rule too: decide such borderline cases explicitly and say how you decided.

```bash
docxtor reject confirmation.docx -o confirmation-resolved.docx --ids 27,28,29,30,31,32,35,36 --others accept
```
```
rejected 8 revisions, accepted the other 39; 0 changes still tracked
```

With many changes, filter the JSON instead of reading every line: `docxtor revisions FILE --json` gives `n`, `ids`, `deleted` and `inserted` per change.

Report back by text ("Rejected: S&P BBB− → BB, Moody's Baa3 → Ba2, …"), and flag knock-on effects, for example a drafting note that still argues for the rejected thresholds.

## 5. Undo one of your own changes later

After example 1, the user says:

> We'll keep the original payment terms after all.

```bash
docxtor revisions msa-redline.docx            # payment change is #2, ids=1,2
docxtor reject msa-redline.docx -o msa-redline.docx --ids 1,2
docxtor revisions msa-redline.docx
```
```
rejected 2 revisions; 1 change still tracked
1 change by "Acme"
#1 [p81] ids=3,4 by Acme: replace "five (5)" -> "three (3)"
```

Rejecting restores the original text with no trace; editing "forty-five" back to "thirty" would leave a confusing double change.

## 6. Add a clause and a table row, fixing a failed patch

> Add Jet A-1 to the product table and change LSFO to VLSFO in it.

```bash
docxtor view msa.docx --grep '^\| \[tbl0\.r[01]\.' | cut -c1-120
```
```
| [tbl0.r0.c0.p0] Product | [tbl0.r0.c1.p0] Minimum Annual Quantity (MT) | [tbl0.r0.c2.p0] Monthly Target (approx.) | …
| [tbl0.r1.c0.p0] Low-Sulfur Fuel Oil (LSFO) | [tbl0.r1.c1.p0] 480,000 MT | [tbl0.r1.c2.p0] ~40,000 MT | …
```
```json
[
  { "type": "add_table_row", "target": "tbl0", "afterRow": 0, "cells": ["Jet A-1", "60,000 MT", "~5,000 MT", "USD 10.00"] },
  { "type": "replace_cell_text", "target": "tbl0.r1.c0.p0", "find": "LSFO", "replace": "VLSFO" }
]
```
```bash
docxtor edit msa.docx --dry-run patches.json
```
```
1. ok add_table_row tbl0
2. FAILS replace_cell_text tbl0.r1.c0.p0
   -> Text "LSFO" not found in paragraph "tbl0.r1.c0.p0" (occurrence 1)
       tbl0.r1.c0.p0 currently reads: "Jet A-1"
1 of 2 would fail; fix them and run again.
```

The new row became row 1 and pushed LSFO to row 2. Change the target to `tbl0.r2.c0.p0`, dry-run again, then write with `-o`.

## 7. Answer a question with citations

> What are the conditions precedent to the initial loan? Cite the sections.

```bash
docxtor view credit.docx --grep 'conditions precedent|^\[p[0-9]+\] \(Heading' | cut -c1-100
```
```
…
[p199] (Heading1) ARTICLE IV — CONDITIONS PRECEDENT
[p200] (Heading2) Section 4.01 — Conditions to All Credit Extensions
…
[p282] (Heading1) ARTICLE V — AFFIRMATIVE COVENANTS
```
```bash
docxtor view credit.docx | sed -n '/^\[p199\]/,/^\[p281\]/p'
```

Answer from that text, citing "Section 4.01(a)" and so on as the document numbers them. Nothing is written.

## 8. Final clean copy

> Accept our changes, keep theirs tracked, and remove our names from the file properties.

```bash
docxtor accept redline.docx -o final.docx --by "Acme"
docxtor metadata final.docx -o final.docx --scrub
docxtor revisions final.docx | head -1
```
