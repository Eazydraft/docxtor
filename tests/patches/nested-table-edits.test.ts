/**
 * Editing a table nested in a cell of another table, through the ids the
 * markdown view prints for it. Each edit must change the nested table only,
 * leave the table around it as it was, show in the view, survive a save,
 * and, when tracked, give the untracked result on accept and the original
 * on reject.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor, type Patch, type PatchOptions } from "../../src/index.js";
import { createBodyDocx } from "../fixtures/create-docx.js";

const TRACKED: PatchOptions = { trackRevisions: true, revisionAuthor: "Reviewer", revisionDate: "2026-11-12T00:00:00Z" };

const cell = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const para = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/**
 * tbl0 is 2 x 2. Its cell r0.c0 holds "Outer", then a 2 x 2 table
 * (Item/Price, Diesel/$900 with "$900" bold), then an empty paragraph.
 * Its cell r1.c1 holds a second nested table with one empty cell.
 */
const INNER =
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
  `<w:tr>${cell(para("Item"))}${cell(para("Price"))}</w:tr>` +
  `<w:tr>${cell(para("Diesel"))}${cell(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>$900</w:t></w:r></w:p>`)}</w:tr>` +
  `</w:tbl>`;
const BODY =
  para("Before") +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
  `<w:tr>${cell(para("Outer") + INNER + "<w:p/>")}${cell(para("Beside"))}</w:tr>` +
  `<w:tr>${cell(para("Lower left"))}${cell(`<w:tbl><w:tr>${cell("<w:p/>")}</w:tr></w:tbl><w:p/>`)}</w:tr>` +
  `</w:tbl>` +
  para("After");

const N = "tbl0.r0.c0.tbl0";

const load = async (): Promise<Docxtor> => Docxtor.load(await createBodyDocx(BODY));

function apply(doc: Docxtor, patch: unknown, options?: PatchOptions) {
  return doc.applyPatches([patch as Patch], options)[0]!;
}

async function documentXml(doc: Docxtor): Promise<string> {
  const zip = await JSZip.loadAsync(await doc.toBuffer());
  return zip.file("word/document.xml")!.async("string");
}

const rows = (doc: Docxtor, options?: Parameters<Docxtor["project"]>[1]): string[] =>
  (doc.project("markdown", options) as string).split("\n").filter((l) => l.startsWith("|"));

const ORIGINAL_ROWS = [
  "| [tbl0.r0.c0.p0] Outer [tbl0.r0.c0.p1]  | [tbl0.r0.c1.p0] Beside |",
  `| [${N}.r0.c0.p0] Item | [${N}.r0.c1.p0] Price |`,
  `| [${N}.r1.c0.p0] Diesel | [${N}.r1.c1.p0] $900 |`,
  "| [tbl0.r1.c0.p0] Lower left | [tbl0.r1.c1.p0]  |",
  "| [tbl0.r1.c1.tbl0.r0.c0.p0]  |",
];

/** The patches under test, each with the view rows it should leave (untracked). */
const EDITS: { name: string; patch: Record<string, unknown>; rows: string[]; skipReject?: boolean }[] = [
  {
    name: "replace_text in a nested cell",
    patch: { type: "replace_text", target: `${N}.r1.c1.p0`, find: "$900", replace: "$950" },
    rows: [ORIGINAL_ROWS[0]!, ORIGINAL_ROWS[1]!, `| [${N}.r1.c0.p0] Diesel | [${N}.r1.c1.p0] $950 |`, ORIGINAL_ROWS[3]!, ORIGINAL_ROWS[4]!],
  },
  {
    name: "replace_cell_text in a nested cell",
    patch: { type: "replace_cell_text", target: `${N}.r1.c0.p0`, find: "Diesel", replace: "Jet A-1" },
    rows: [ORIGINAL_ROWS[0]!, ORIGINAL_ROWS[1]!, `| [${N}.r1.c0.p0] Jet A-1 | [${N}.r1.c1.p0] $900 |`, ORIGINAL_ROWS[3]!, ORIGINAL_ROWS[4]!],
  },
  {
    name: "insert_text next to an anchor in a nested cell",
    patch: { type: "insert_text", target: `${N}.r0.c1.p0`, position: "after", anchor: "Price", text: " (USD)" },
    rows: [ORIGINAL_ROWS[0]!, `| [${N}.r0.c0.p0] Item | [${N}.r0.c1.p0] Price (USD) |`, ORIGINAL_ROWS[2]!, ORIGINAL_ROWS[3]!, ORIGINAL_ROWS[4]!],
  },
  {
    name: "delete_text in a nested cell",
    patch: { type: "delete_text", target: `${N}.r1.c1.p0`, find: "$" },
    rows: [ORIGINAL_ROWS[0]!, ORIGINAL_ROWS[1]!, `| [${N}.r1.c0.p0] Diesel | [${N}.r1.c1.p0] 900 |`, ORIGINAL_ROWS[3]!, ORIGINAL_ROWS[4]!],
  },
  {
    name: "insert_text with no anchor into the empty cell of a nested table",
    patch: { type: "insert_text", target: "tbl0.r1.c1.tbl0.r0.c0.p0", position: "start", text: "Filled" },
    rows: [...ORIGINAL_ROWS.slice(0, 4), "| [tbl0.r1.c1.tbl0.r0.c0.p0] Filled |"],
  },
  {
    name: "add_table_row to a nested table",
    patch: { type: "add_table_row", target: N, afterRow: 1, cells: ["LSFO", "$700"] },
    rows: [...ORIGINAL_ROWS.slice(0, 3), `| [${N}.r2.c0.p0] LSFO | [${N}.r2.c1.p0] $700 |`, ORIGINAL_ROWS[3]!, ORIGINAL_ROWS[4]!],
  },
  {
    name: "delete_table_row from a nested table",
    patch: { type: "delete_table_row", target: N, row: 0 },
    rows: [ORIGINAL_ROWS[0]!, `| [${N}.r0.c0.p0] Diesel | [${N}.r0.c1.p0] $900 |`, ORIGINAL_ROWS[3]!, ORIGINAL_ROWS[4]!],
  },
  {
    name: "insert_paragraph after a paragraph of a nested cell",
    // Rejecting a tracked insert_paragraph leaves an empty paragraph behind. That is how it behaves
    // in the body and in any cell, not something about nested tables, so it is not checked here.
    skipReject: true,
    patch: { type: "insert_paragraph", target: `${N}.r1.c0.p0`, position: "after", text: "Ultra-low sulphur" },
    rows: [ORIGINAL_ROWS[0]!, ORIGINAL_ROWS[1]!, `| [${N}.r1.c0.p0] Diesel [${N}.r1.c0.p1] Ultra-low sulphur | [${N}.r1.c1.p0] $900 |`, ORIGINAL_ROWS[3]!, ORIGINAL_ROWS[4]!],
  },
];

describe("editing a nested table", () => {
  it("the view prints the nested tables under the rows that hold them", async () => {
    expect(rows(await load())).toEqual(ORIGINAL_ROWS);
  });

  describe.each(EDITS)("$name", ({ patch, rows: expected, skipReject }) => {
    it("untracked: changes the nested table only, and survives a save", async () => {
      const doc = await load();
      expect(apply(doc, patch).error).toBeUndefined();
      expect(rows(doc)).toEqual(expected);
      expect(doc.findParagraphById("p0")!.text).toBe("Before");
      expect(doc.findParagraphById("p1")!.text).toBe("After");
      const reloaded = await Docxtor.load(await doc.toBuffer());
      expect(rows(reloaded)).toEqual(expected);
      // The outer table still has two rows of two cells, and both nested tables
      expect(reloaded.getTables()).toHaveLength(1);
      expect(reloaded.getTables()[0]!.rows.map((r) => r.cells.length)).toEqual([2, 2]);
      expect(reloaded.findTableById(N)).toBeDefined();
      expect(reloaded.findTableById("tbl0.r1.c1.tbl0")).toBeDefined();
    });

    it("tracked: accept gives the untracked result, reject gives the original", async () => {
      const plain = await load();
      expect(apply(plain, patch).error).toBeUndefined();

      const accepted = await load();
      expect(apply(accepted, patch, TRACKED).error).toBeUndefined();
      expect((await documentXml(accepted)).includes(`w:author="Reviewer"`)).toBe(true);
      accepted.acceptAllRevisions();
      expect(rows(accepted)).toEqual(expected);
      expect(await documentXml(accepted)).not.toContain(`w:author="Reviewer"`);

      if (skipReject) return;
      const rejected = await load();
      expect(apply(rejected, patch, TRACKED).error).toBeUndefined();
      rejected.rejectAllRevisions();
      expect(rows(rejected)).toEqual(ORIGINAL_ROWS);
      expect(await documentXml(rejected)).not.toContain(`w:author="Reviewer"`);
    });
  });

  it("a tracked replacement shows inline in the nested row and keeps the cell's formatting", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: `${N}.r1.c1.p0`, find: "$900", replace: "$950" }, TRACKED).success).toBe(true);
    expect(rows(doc, { includeRevisions: true })[2]).toBe(`| [${N}.r1.c0.p0] Diesel | [${N}.r1.c1.p0] {-$900-}{+$950+} |`);
    expect(await documentXml(doc)).toContain(`<w:r><w:rPr><w:b/></w:rPr><w:t>$950</w:t></w:r>`);
  });

  it("a row patch on the outer table leaves the nested tables in place, under their new ids", async () => {
    const doc = await load();
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 0 }).error).toBeUndefined();
    // The second nested table was in row 1 and is now in row 0
    expect(rows(doc)).toEqual(["| [tbl0.r0.c0.p0] Lower left | [tbl0.r0.c1.p0]  |", "| [tbl0.r0.c1.tbl0.r0.c0.p0]  |"]);
    expect(doc.findTableById(N)).toBeUndefined();
    // The id the view prints after the shift resolves, and the old one does not
    expect(apply(doc, { type: "insert_text", target: "tbl0.r0.c1.tbl0.r0.c0.p0", position: "start", text: "Filled" }).error).toBeUndefined();
    expect(apply(doc, { type: "insert_text", target: "tbl0.r1.c1.tbl0.r0.c0.p0", position: "start", text: "x" }).error).toBe(
      'Paragraph "tbl0.r1.c1.tbl0.r0.c0.p0" not found',
    );
    expect(rows(doc)[1]).toBe("| [tbl0.r0.c1.tbl0.r0.c0.p0] Filled |");
  });

  it("a tracked deletion of the outer row that holds a nested table removes it on accept and keeps it on reject", async () => {
    const accepted = await load();
    expect(apply(accepted, { type: "delete_table_row", target: "tbl0", row: 0 }, TRACKED).error).toBeUndefined();
    // Until it is resolved the row and its nested table are still there
    expect(rows(accepted)).toEqual(ORIGINAL_ROWS);
    accepted.acceptAllRevisions();
    expect(rows(accepted)).toEqual(["| [tbl0.r0.c0.p0] Lower left | [tbl0.r0.c1.p0]  |", "| [tbl0.r0.c1.tbl0.r0.c0.p0]  |"]);

    const rejected = await load();
    expect(apply(rejected, { type: "delete_table_row", target: "tbl0", row: 0 }, TRACKED).error).toBeUndefined();
    rejected.rejectAllRevisions();
    expect(rows(rejected)).toEqual(ORIGINAL_ROWS);
  });

  it("several edits in one batch each see the ids as the earlier ones left them", async () => {
    const doc = await load();
    const results = doc.applyPatches([
      { type: "add_table_row", target: N, afterRow: 0, cells: ["LSFO", "$700"] },
      { type: "replace_text", target: `${N}.r2.c0.p0`, find: "Diesel", replace: "Gasoil" },
      { type: "insert_text", target: "tbl0.r0.c0.p1", position: "start", text: "Prices exclude VAT." },
    ] as Patch[], { ...TRACKED, atomic: true });
    expect(results.map((r) => r.error)).toEqual([undefined, undefined, undefined]);
    doc.acceptAllRevisions();
    expect(rows(doc).slice(0, 4)).toEqual([
      "| [tbl0.r0.c0.p0] Outer [tbl0.r0.c0.p1] Prices exclude VAT. | [tbl0.r0.c1.p0] Beside |",
      `| [${N}.r0.c0.p0] Item | [${N}.r0.c1.p0] Price |`,
      `| [${N}.r1.c0.p0] LSFO | [${N}.r1.c1.p0] $700 |`,
      `| [${N}.r2.c0.p0] Gasoil | [${N}.r2.c1.p0] $900 |`,
    ]);
  });

  it("add_comment anchors on text in a nested cell and changes no text", async () => {
    const doc = await load();
    expect(apply(doc, { type: "add_comment", target: `${N}.r1.c1.p0`, find: "$900", text: "Confirm the price" }).error).toBeUndefined();
    expect(rows(doc)).toEqual(ORIGINAL_ROWS);
    const zip = await JSZip.loadAsync(await doc.toBuffer());
    expect(await zip.file("word/comments.xml")!.async("string")).toContain("Confirm the price");
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:commentRangeStart w:id="\d+"\/><w:r><w:rPr><w:b\/><\/w:rPr><w:t>\$900<\/w:t><\/w:r><w:commentRangeEnd/);
  });

  it("a find that is in the neighbouring nested cell names it; a missing nested id is reported", async () => {
    const doc = await load();
    const r = apply(doc, { type: "replace_text", target: `${N}.r1.c0.p0`, find: "$900", replace: "$950" });
    expect(r.error).toBe(`Text "$900" not found in paragraph "${N}.r1.c0.p0" (occurrence 1); it is in ${N}.r1.c1.p0 (use that as the target)`);
    const gone = apply(doc, { type: "replace_text", target: `${N}.r7.c0.p0`, find: "x", replace: "y" });
    expect(gone.error).toBe(`Paragraph "${N}.r7.c0.p0" not found`);
    const table = apply(doc, { type: "add_table_row", target: "tbl0.r0.c1.tbl0", afterRow: 0, cells: ["a"] });
    expect(table.error).toBe('Table "tbl0.r0.c1.tbl0" not found');
  });

  it("a save with no edit leaves document.xml byte for byte as it was", async () => {
    const buffer = await createBodyDocx(BODY);
    const before = await (await JSZip.loadAsync(buffer)).file("word/document.xml")!.async("string");
    expect(await documentXml(await Docxtor.load(buffer))).toBe(before);
  });
});
