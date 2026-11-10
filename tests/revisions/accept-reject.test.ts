/**
 * Accept and reject all revisions.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor } from "../../src/index.js";
import type { Patch } from "../../src/index.js";
import {
  createLeaseDocx,
  createSimpleDocx,
  createTableDocx,
  createStyledDocx,
  createMultiParagraphDocx,
} from "../fixtures/create-docx.js";

const AUTHOR = "USI Legal AI";
const TRACKED = {
  trackRevisions: true,
  revisionAuthor: AUTHOR,
  revisionDate: "2026-10-13T00:00:00Z",
};

/** Plain text of every body paragraph, including table cells, in order. */
function bodyText(doc: Docxtor): string[] {
  const out: string[] = [];
  for (const el of doc.getBody()) {
    if (el.kind === "paragraph") out.push(el.text);
    else if (el.kind === "table") {
      for (const row of el.rows) for (const cell of row.cells) for (const p of cell.paragraphs) out.push(p.text);
    }
  }
  return out;
}

// =============================================================================
// Engine fixed point: tracked + accept == untracked
// =============================================================================

describe("engine fixed point", () => {
  const BATCH: Patch[] = [
    { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "shall have no right to recapture" },
    { type: "insert_text", target: "para:0A000003", position: "after", anchor: "Exhibit A", text: " attached hereto" },
    { type: "delete_text", target: "para:0A000005", find: ", defend" },
    { type: "insert_paragraph", target: "para:0A000005", position: "after", text: "Landlord shall indemnify Tenant." },
    { type: "delete_paragraph", target: "para:0A000006" },
    { type: "delete_paragraph", target: "para:0A000007" },
    { type: "replace_cell_text", target: "para:0A000012", find: "Operating", replace: "Building Operating" },
    { type: "apply_formatting", target: "para:0A000014", formatting: { italic: true } },
    { type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["Taxes", "Pro rata"] },
    { type: "delete_table_row", target: "tbl0", row: 0 },
    { type: "insert_paragraph", target: "para:0A000014", position: "before", text: "IN WITNESS WHEREOF" },
  ];

  it("text(apply(B, tracked)).acceptAll() equals text(apply(B, untracked))", async () => {
    const tracked = await Docxtor.load(await createLeaseDocx());
    const untracked = await Docxtor.load(await createLeaseDocx());

    const r1 = tracked.applyPatches(BATCH, TRACKED);
    const r2 = untracked.applyPatches(BATCH);
    expect(r1.every((r) => r.success)).toBe(true);
    expect(r2.every((r) => r.success)).toBe(true);

    tracked.acceptAllRevisions({ authors: [AUTHOR] });

    expect(bodyText(tracked)).toEqual(bodyText(untracked));
    // Compare projections without id tokens: accept re-normalises (contiguous
    // positional ids) while the untracked path keeps next-free numbering, and
    // new paragraphs receive random paraIds.
    const stripIds = (md: string): string => md.replace(/^\[[^\]]+\] ?/gm, "").replace(/ ¶[0-9A-F]{8}/g, "");
    expect(stripIds(tracked.project("markdown"))).toBe(stripIds(untracked.project("markdown")));
  });

  it("document.xml is byte-identical for replace/insert_paragraph/delete_paragraph batches with a pinned date", async () => {
    const batch: Patch[] = [
      { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "shall have no right to recapture" },
      { type: "insert_paragraph", target: "para:0A000005", position: "after", text: "Landlord shall indemnify Tenant." },
      { type: "delete_paragraph", target: "para:0A000006" },
      { type: "replace_text", target: "para:0A000014", find: "follows", replace: "is attached" },
      { type: "insert_paragraph", target: "para:0A000014", position: "before", text: "IN WITNESS WHEREOF" },
    ];
    const tracked = await Docxtor.load(await createLeaseDocx());
    const untracked = await Docxtor.load(await createLeaseDocx());

    // paraIds for new paragraphs are random; pin them by copying from the untracked run
    tracked.applyPatches(batch, TRACKED);
    untracked.applyPatches(batch);
    tracked.acceptAllRevisions({ authors: [AUTHOR] });

    const strip = (xml: string): string => xml.replace(/ w14:paraId="[0-9A-F]{8}" w14:textId="[0-9A-F]{8}"/g, "");
    expect(strip(tracked.toString())).toBe(strip(untracked.toString()));
  });

  it("re-applying the same batch after accept produces no edits for text patches it already made", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "shall have no right to recapture" },
    ], TRACKED);
    doc.acceptAllRevisions();
    const reloaded = await Docxtor.load(await doc.toBuffer());
    const [r] = reloaded.applyPatches([
      { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "shall have no right to recapture" },
    ], TRACKED);
    // The find string no longer exists → the rule is a no-op on the second pass
    expect(r!.success).toBe(false);
    expect(reloaded.toString()).not.toContain("w:ins");
  });
});

// =============================================================================
// Accept: element coverage
// =============================================================================

describe("acceptAllRevisions", () => {
  it("accepts pre-existing revisions from another author in body and footnotes", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    expect(doc.findParagraphById("p8")!.text).toBe("Base Rent shall be $42.00 per rentable square foot per year.");

    const { accepted } = doc.acceptAllRevisions();
    expect(accepted).toBe(3); // del + ins in p8, ins in footnote

    expect(doc.findParagraphById("p8")!.text).toBe("Base Rent shall be $42.00 per rentable square foot per year.");
    expect(doc.findParagraphById("p8")!.revisions).toBeUndefined();
    const xml = doc.toString();
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
    expect(xml).not.toContain("$40.00");

    const out = await doc.toBuffer();
    const zip = await JSZip.loadAsync(out);
    const footnotes = await zip.file("word/footnotes.xml")!.async("string");
    expect(footnotes).not.toContain("w:ins");
    expect(footnotes).toContain("revised ");
  });

  it("returns 0 and leaves the document alone when there is nothing to accept", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    const before = doc.toString();
    expect(doc.acceptAllRevisions()).toEqual({ accepted: 0 });
    expect(doc.toString()).toBe(before);
  });

  it("drops w:rPrChange and w:pPrChange, keeping the new formatting", async () => {
    const doc = await Docxtor.load(await createStyledDocx());
    doc.applyPatches([
      { type: "apply_formatting", target: "p2", find: "Plain", formatting: { bold: true } },
      { type: "apply_style", target: "p2", style: "Heading2" },
    ], TRACKED);
    expect(doc.toString()).toContain("w:rPrChange");
    expect(doc.toString()).toContain("w:pPrChange");

    const { accepted } = doc.acceptAllRevisions();
    expect(accepted).toBe(2);
    const xml = doc.toString();
    expect(xml).not.toContain("w:rPrChange");
    expect(xml).not.toContain("w:pPrChange");
    const p2 = doc.findParagraphById("p2")!;
    expect(p2.style).toBe("Heading2");
    expect(p2.runs.find((r) => r.text === "Plain")!.formatting.bold).toBe(true);
  });

  it("keeps inserted rows (marker dropped) and removes deleted rows", async () => {
    const doc = await Docxtor.load(await createTableDocx());
    doc.applyPatches([
      { type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["New A", "New B"] },
      { type: "delete_table_row", target: "tbl0", row: 0 },
    ], TRACKED);
    expect(doc.getTables()[0]!.rows.length).toBe(3);

    doc.acceptAllRevisions();
    const table = doc.getTables()[0]!;
    expect(table.rows.length).toBe(2);
    expect(table.rows.map((r) => r.cells[0]!.paragraphs[0]!.text)).toEqual(["Cell 1", "New A"]);
    const xml = doc.toString();
    expect(xml).not.toContain("Header A");
    expect(xml).not.toMatch(/<w:trPr><w:ins/);
  });

  it("walks headers and footers", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "append_text", text: "Draft", container: { type: "header", variant: "default" } },
      { type: "replace_text", target: "ftr.default.p0", find: "Page ", replace: "Pg. " },
    ], TRACKED);
    doc.acceptAllRevisions();
    const out = await doc.toBuffer();
    const zip = await JSZip.loadAsync(out);
    const hdr = await zip.file("word/header1.xml")!.async("string");
    const ftr = await zip.file("word/footer1.xml")!.async("string");
    expect(hdr).not.toContain("w:ins");
    expect(hdr).toContain("Draft");
    expect(ftr).not.toContain("w:del");
    expect(ftr).toContain("Pg. ");
    expect(ftr).not.toContain("Page ");
  });

  it("re-normalises: positional ids are contiguous again, paraIds unchanged", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([{ type: "delete_paragraph", target: "p3" }], TRACKED);
    doc.acceptAllRevisions();
    const ids = doc.getParagraphs().map((p) => p.id);
    expect(ids).toEqual(["p0", "p1", "p2", "p3", "p4", "p5", "p6", "p7", "p8", "p9"]);
    expect(doc.findParagraphById("p3")!.paraId).toBe("0A000005");
    expect(doc.findParagraphByParaId("0A000004")).toBeUndefined();
  });

  it("keeps a deleted paragraph mark at the end of a container as an empty paragraph (Word semantics)", async () => {
    const doc = await Docxtor.load(await createMultiParagraphDocx(["One", "Two"]));
    doc.applyPatches([{ type: "delete_paragraph", target: "p1" }], TRACKED);
    doc.acceptAllRevisions();
    // No following paragraph to merge into → an empty paragraph remains, as in Word
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["One", ""]);
  });

  it("merges a split paragraph when the inserted mark is accepted... and keeps content on reject", async () => {
    // Simulate Word splitting "Hello World" into two paragraphs with tracking on:
    // first half has an inserted paragraph mark.
    const doc = await Docxtor.load(await createMultiParagraphDocx(["Hello ", "World"]));
    const ast = doc.getAST();
    const first = ast.body[0]!.xmlRef;
    (first["w:p"] as unknown[]).unshift({
      "w:pPr": [{ "w:rPr": [{ "w:ins": [], ":@": { "@_w:id": "9", "@_w:author": "Someone", "@_w:date": "2026-01-01T00:00:00Z" } }] }],
    });

    const accept = await Docxtor.load(await doc.toBuffer());
    accept.acceptAllRevisions();
    expect(accept.getParagraphs().map((p) => p.text)).toEqual(["Hello ", "World"]);

    const reject = await Docxtor.load(await doc.toBuffer());
    reject.rejectAllRevisions();
    expect(reject.getParagraphs().map((p) => p.text)).toEqual(["Hello World"]);
  });
});

// =============================================================================
// Reject: element coverage
// =============================================================================

describe("rejectAllRevisions", () => {
  it("restores original text for a mixed tracked batch", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const original = bodyText(doc);
    const results = doc.applyPatches([
      { type: "replace_text", target: "p6", find: "may recapture", replace: "may not recapture" },
      { type: "insert_text", target: "p2", position: "after", anchor: "Exhibit A", text: " attached" },
      { type: "delete_text", target: "p4", find: ", defend" },
      { type: "insert_paragraph", target: "p4", position: "after", text: "Mirror." },
      { type: "delete_paragraph", target: "p5" },
      { type: "replace_cell_text", target: "tbl0.r1.c0.p0", find: "Operating", replace: "Building" },
      { type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["x", "y"] },
      { type: "delete_table_row", target: "tbl0", row: 0 },
      { type: "apply_formatting", target: "p0", formatting: { italic: true } },
      { type: "apply_style", target: "p10", style: "Heading2" },
    ], TRACKED);
    expect(results.every((r) => r.success)).toBe(true);

    const { rejected } = doc.rejectAllRevisions({ authors: [AUTHOR] });
    expect(rejected).toBeGreaterThanOrEqual(10);

    expect(bodyText(doc)).toEqual(original);
    expect(doc.findParagraphById("p0")!.runs[0]!.formatting.italic).toBeUndefined();
    expect(doc.findParagraphById("p10")!.style).toBeUndefined();
    expect(doc.getTables()[0]!.rows.length).toBe(2);
    const xml = doc.toString();
    expect(xml).not.toContain(AUTHOR);
    // Opposing Counsel's revisions are untouched by the author filter: their
    // one deletion ($40.00) is the only w:delText left
    expect(xml).toContain('w:author="Opposing Counsel"');
    expect((xml.match(/<w:delText/g) ?? []).length).toBe(1);
    expect(xml).toContain("<w:delText>$40.00</w:delText>");
  });

  it("rejecting another author's pre-existing changes restores their original text", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.rejectAllRevisions();
    expect(doc.findParagraphById("p8")!.text).toBe("Base Rent shall be $40.00 per rentable square foot per year.");
    expect(doc.toString()).not.toContain("$42.00");
    expect(doc.toString()).not.toContain("w:delText");
  });

  it("restores old run properties from w:rPrChange", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Bold this"));
    doc.applyPatches([{ type: "apply_formatting", target: "p0", formatting: { bold: true, color: "FF0000" } }], TRACKED);
    doc.rejectAllRevisions();
    const run = doc.findParagraphById("p0")!.runs[0]!;
    expect(run.formatting.bold).toBeUndefined();
    expect(run.formatting.color).toBeUndefined();
    expect(doc.toString()).not.toContain("w:rPrChange");
  });
});

// =============================================================================
// Author filter
// =============================================================================

describe("author filter", () => {
  it("accepts only the tool's revisions, leaving another author's marks untouched", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "replace_text", target: "p6", find: "may recapture", replace: "may not recapture" },
      { type: "insert_paragraph", target: "p4", position: "after", text: "Mirror." },
      { type: "delete_paragraph", target: "p5" },
    ], TRACKED);

    const { accepted } = doc.acceptAllRevisions({ authors: [AUTHOR] });
    expect(accepted).toBeGreaterThanOrEqual(5); // del+ins (replace), ins runs + mark (insert), del runs + mark (delete)

    const xml = doc.toString();
    expect(xml).not.toContain(`w:author="${AUTHOR}"`);
    expect(xml).toContain('w:author="Opposing Counsel"');
    expect((xml.match(/w:author="Opposing Counsel"/g) ?? []).length).toBe(2);

    // Their paragraph still reads with their insertion and excludes their deletion
    const rent = doc.getParagraphs().find((p) => p.text.startsWith("Base Rent"))!;
    expect(rent.text).toBe("Base Rent shall be $42.00 per rentable square foot per year.");
    expect(rent.revisions!.length).toBe(2);

    // And the footnote insertion by Opposing Counsel survives
    const out = await doc.toBuffer();
    const zip = await JSZip.loadAsync(out);
    expect(await zip.file("word/footnotes.xml")!.async("string")).toContain('w:author="Opposing Counsel"');
  });

  it("an author filter that matches nothing is a no-op", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const before = doc.toString();
    expect(doc.acceptAllRevisions({ authors: ["Nobody"] })).toEqual({ accepted: 0 });
    expect(doc.rejectAllRevisions({ authors: ["Nobody"] })).toEqual({ rejected: 0 });
    expect(doc.toString()).toBe(before);
  });
});

describe("id filter", () => {
  async function twoReplacements(): Promise<Docxtor> {
    const doc = await Docxtor.load(await createMultiParagraphDocx(["Pay within 30 days.", "Governed by Texas law."]));
    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "30", replace: "45" },
      { type: "replace_text", target: "p1", find: "Texas", replace: "New York" },
    ], TRACKED);
    return doc;
  }

  it("exposes each mark's w:id in the projection", async () => {
    const doc = await twoReplacements();
    const marks = doc.getParagraphs().flatMap((p) => p.revisions ?? []);
    expect(marks.map((m) => [m.type, m.text])).toEqual([
      ["insert", "45"], ["delete", "30"], ["insert", "New York"], ["delete", "Texas"],
    ]);
    for (const m of marks) expect(m.id).toMatch(/^\d+$/);
    expect(new Set(marks.map((m) => m.id)).size).toBe(4);
  });

  it("rejects only the listed revisions and leaves the rest tracked", async () => {
    const doc = await twoReplacements();
    const ids = doc.getParagraphs()[0]!.revisions!.map((m) => m.id!);
    expect(doc.rejectAllRevisions({ ids })).toEqual({ rejected: 2 });
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["Pay within 30 days.", "Governed by New York law."]);
    expect(doc.getParagraphs()[0]!.revisions).toBeUndefined();
    expect(doc.getParagraphs()[1]!.revisions!.length).toBe(2);

    expect(doc.acceptAllRevisions()).toEqual({ accepted: 2 });
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["Pay within 30 days.", "Governed by New York law."]);
  });

  it("combines with an author filter", async () => {
    const doc = await twoReplacements();
    const ids = doc.getParagraphs().flatMap((p) => p.revisions ?? []).map((m) => m.id!);
    expect(doc.acceptAllRevisions({ ids, authors: ["Someone Else"] })).toEqual({ accepted: 0 });
    expect(doc.acceptAllRevisions({ ids: [ids[0]!], authors: [AUTHOR] })).toEqual({ accepted: 1 });
  });
});

// =============================================================================
// Moves and tracked cell merges (review follow-ups)
// =============================================================================

describe("moves with an author filter", () => {
  const MOVE_BODY = `
<w:p><w:moveFromRangeStart w:id="10" w:author="Mover" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:moveFrom w:id="11" w:author="Mover" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Moved text.</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="10"/></w:p>
<w:p><w:r><w:t>Stays.</w:t></w:r></w:p>
<w:p><w:moveToRangeStart w:id="12" w:author="Mover" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:moveTo w:id="13" w:author="Mover" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Moved text.</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="12"/></w:p>`;

  it("accepting the mover's changes removes both range start and end markers", async () => {
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    const doc = await Docxtor.load(await createBodyDocx(MOVE_BODY));
    doc.acceptAllRevisions({ authors: ["Mover"] });
    const xml = doc.toString();
    expect(xml).not.toContain("moveFromRange");
    expect(xml).not.toContain("moveToRange");
    expect(xml).not.toContain("w:moveFrom");
    expect(xml).not.toContain("w:moveTo");
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["", "Stays.", "Moved text."]);
  });

  it("rejecting the mover's changes restores the source and removes all markers", async () => {
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    const doc = await Docxtor.load(await createBodyDocx(MOVE_BODY));
    doc.rejectAllRevisions({ authors: ["Mover"] });
    const xml = doc.toString();
    expect(xml).not.toContain("RangeStart");
    expect(xml).not.toContain("RangeEnd");
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["Moved text.", "Stays.", ""]);
  });

  it("an ids filter naming any part of a move resolves the whole move", async () => {
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    // 11 is the moveFrom content, 12 the moveTo range marker, 13 the moveTo content.
    for (const ids of [["11"], ["12"], ["13"]]) {
      const accepted = await Docxtor.load(await createBodyDocx(MOVE_BODY));
      accepted.acceptAllRevisions({ ids });
      const a = accepted.toString();
      expect(a).not.toMatch(/w:move(From|To)/);
      expect(accepted.getParagraphs().map((p) => p.text)).toEqual(["", "Stays.", "Moved text."]);

      const rejected = await Docxtor.load(await createBodyDocx(MOVE_BODY));
      rejected.rejectAllRevisions({ ids });
      expect(rejected.toString()).not.toMatch(/w:move(From|To)/);
      expect(rejected.getParagraphs().map((p) => p.text)).toEqual(["Moved text.", "Stays.", ""]);
    }
  });

  it("another author's filter leaves the move and its markers untouched", async () => {
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    const doc = await Docxtor.load(await createBodyDocx(MOVE_BODY));
    const before = doc.toString();
    expect(doc.acceptAllRevisions({ authors: ["Someone Else"] })).toEqual({ accepted: 0 });
    expect(doc.toString()).toBe(before);
    expect(doc.toString()).toContain('<w:moveFromRangeEnd w:id="10"/>');
  });
});

describe("tracked cell merges", () => {
  const cellMergeBody = (vMergeOrig: string | undefined, currentVMerge: string): string => `
<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="4788"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="4788" w:type="dxa"/>${currentVMerge}<w:cellMerge w:id="20" w:author="Merger" w:date="2026-01-01T00:00:00Z" w:vMerge="rest"${vMergeOrig ? ` w:vMergeOrig="${vMergeOrig}"` : ""}/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="4788" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>
</w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>`;

  it("accept keeps the merged state and drops the marker", async () => {
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    const doc = await Docxtor.load(await createBodyDocx(cellMergeBody(undefined, '<w:vMerge w:val="restart"/>')));
    expect(doc.acceptAllRevisions()).toEqual({ accepted: 1 });
    const xml = doc.toString();
    expect(xml).not.toContain("w:cellMerge");
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
  });

  it("reject with no vMergeOrig removes w:vMerge (the cell was not merged before)", async () => {
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    const doc = await Docxtor.load(await createBodyDocx(cellMergeBody(undefined, '<w:vMerge w:val="restart"/>')));
    expect(doc.rejectAllRevisions()).toEqual({ rejected: 1 });
    const xml = doc.toString();
    expect(xml).not.toContain("w:cellMerge");
    expect(xml).not.toContain('w:val="restart"');
    // The second row's own (untracked) vMerge is untouched
    expect((xml.match(/<w:vMerge\/>/g) ?? []).length).toBe(1);
  });

  it("reject restores the recorded original merge state", async () => {
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    const cont = await Docxtor.load(await createBodyDocx(cellMergeBody("cont", '<w:vMerge w:val="restart"/>')));
    cont.rejectAllRevisions();
    expect(cont.toString()).toMatch(/<w:tcPr><w:tcW[^>]*\/><w:vMerge\/><\/w:tcPr>/);

    const rest = await Docxtor.load(await createBodyDocx(cellMergeBody("rest", "")));
    rest.rejectAllRevisions();
    expect(rest.toString()).toMatch(/<w:tcPr><w:tcW[^>]*\/><w:vMerge w:val="restart"\/><\/w:tcPr>/);
  });
});

describe("revision id allocation", () => {
  it("skips ids used in footnotes, endnotes and comments", async () => {
    // The lease fixture's body uses ids 1-2 and its footnote uses id 3
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "replace_text", target: "p6", find: "may recapture", replace: "may not recapture" },
    ], TRACKED);
    const ours = [...doc.toString().matchAll(/<w:(?:ins|del) w:id="(\d+)" w:author="USI Legal AI"/g)].map((m) => Number(m[1]));
    expect(ours.length).toBe(2);
    expect(Math.min(...ours)).toBeGreaterThanOrEqual(4);
  });
});
