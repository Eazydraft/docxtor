/**
 * §3.2 insert_paragraph and §3.4 delete_paragraph.
 */

import { describe, it, expect } from "vitest";
import { Docxtor, serializeXml } from "../../src/index.js";
import {
  createLeaseDocx,
  createSimpleDocx,
  createTableDocx,
  createFormattedDocx,
  createMultiParagraphDocx,
} from "../fixtures/create-docx.js";

const TRACKED = {
  trackRevisions: true,
  revisionAuthor: "USI Legal AI",
  revisionDate: "2026-10-13T00:00:00Z",
};

/** Paragraph-level XML for a given body paragraph, for assertions. */
function paragraphXml(doc: Docxtor, id: string): string {
  return serializeXml([doc.findParagraphById(id)!.xmlRef]);
}

// =============================================================================
// insert_paragraph
// =============================================================================

describe("insert_paragraph", () => {
  it("inserts after a level-2 numbered item, inheriting numId and ilvl (tracked)", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const originalCount = doc.getParagraphs().length;

    const [r] = doc.applyPatches([
      { type: "insert_paragraph", target: "p4", position: "after", text: "Landlord shall indemnify, defend and hold harmless Tenant." },
    ], TRACKED);
    expect(r!.success).toBe(true);

    const inserted = doc.getParagraphs()[5]!;
    expect(inserted.text).toBe("Landlord shall indemnify, defend and hold harmless Tenant.");
    expect(inserted.listInfo?.numId).toBe("1");
    expect(inserted.listInfo?.level).toBe(1);
    expect(inserted.id).toBe("p11"); // next free positional id

    const xml = paragraphXml(doc, inserted.id);
    // Runs wrapped in w:ins, and the paragraph mark carries w:ins inside w:pPr/w:rPr
    expect(xml).toMatch(/<w:pPr>.*<w:rPr><w:ins w:id="\d+" w:author="USI Legal AI" w:date="2026-10-13T00:00:00Z"><\/w:ins><\/w:rPr><\/w:pPr>/);
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="USI Legal AI"[^>]*><w:r>/);
    expect(xml).toContain('<w:numPr><w:ilvl w:val="1"></w:ilvl><w:numId w:val="1"></w:numId></w:numPr>');

    // Accept all: exactly one new paragraph, no blank paragraph
    doc.acceptAllRevisions();
    const paras = doc.getParagraphs();
    expect(paras.length).toBe(originalCount + 1);
    expect(paras[5]!.text).toBe("Landlord shall indemnify, defend and hold harmless Tenant.");
    expect(paras.some((p) => p.text.trim() === "")).toBe(false);
    expect(doc.toString()).not.toContain("w:ins");
  });

  it("inserts before the target", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([{ type: "insert_paragraph", target: "p1", position: "before", text: "RECITALS" }]);
    const paras = doc.getParagraphs();
    expect(paras[1]!.text).toBe("RECITALS");
    expect(paras[2]!.text).toBe("PREMISES");
    // Inherited the list properties of PREMISES (level 0)
    expect(paras[1]!.listInfo?.level).toBe(0);
  });

  it("splits on \\n into several paragraphs addressable later in the same batch", async () => {
    const doc = await Docxtor.load(await createMultiParagraphDocx(["One", "Two", "Three"]));
    const results = doc.applyPatches([
      { type: "insert_paragraph", target: "p1", position: "after", text: "A\nB\nC" },
      // p3, p4, p5 are the new paragraphs (next free ids after p2)
      { type: "replace_text", target: "p4", find: "B", replace: "Bee" },
    ]);
    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["One", "Two", "A", "Bee", "C", "Three"]);
    expect(doc.getParagraphs().map((p) => p.id)).toEqual(["p0", "p1", "p3", "p4", "p5", "p2"]);
  });

  it("supports inline formatting markers and field placeholders", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    doc.applyPatches([
      { type: "insert_paragraph", target: "p0", position: "after", text: "This is **bold** and page {{PAGE}}." },
    ]);
    const p = doc.getParagraphs()[1]!;
    expect(p.text).toBe("This is bold and page .");
    expect(p.runs.find((r) => r.text === "bold")!.formatting.bold).toBe(true);
    expect(doc.toString()).toContain("w:instrText");
  });

  it("inheritStyle: false produces a bare paragraph", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "insert_paragraph", target: "p4", position: "after", text: "Plain.", inheritStyle: false },
    ]);
    const p = doc.getParagraphs()[5]!;
    expect(p.listInfo).toBeUndefined();
    expect(p.style).toBeUndefined();
    expect(paragraphXml(doc, p.id)).not.toContain("w:pPr");
  });

  it("inherits the paragraph style and the last run's formatting", async () => {
    // createFormattedDocx p1: "This is " + bold + " and " + italic + " text." (last run plain)
    // p0 has Heading1 style and a bold run
    const doc = await Docxtor.load(await createFormattedDocx());
    doc.applyPatches([
      { type: "insert_paragraph", target: "p0", position: "after", text: "Subtitle" },
    ]);
    const p = doc.getParagraphs()[1]!;
    expect(p.style).toBe("Heading1");
    expect(p.runs[0]!.formatting.bold).toBe(true);
  });

  it("does not copy a section break or tracked-change state from the target", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "apply_style", target: "p10", style: "Heading2" },
    ], TRACKED);
    // p10 now has a w:pPrChange; the inserted paragraph must not
    doc.applyPatches([
      { type: "insert_paragraph", target: "p10", position: "after", text: "Tail" },
    ], TRACKED);
    const tail = doc.getParagraphs().find((p) => p.text === "Tail")!;
    const xml = paragraphXml(doc, tail.id);
    expect(xml).not.toContain("w:pPrChange");
    expect(xml).not.toContain("w:sectPr");
    expect(tail.style).toBe("Heading2");
  });

  it("works inside a table cell (sibling w:p in the same w:tc)", async () => {
    const doc = await Docxtor.load(await createTableDocx());
    const [r] = doc.applyPatches([
      { type: "insert_paragraph", target: "tbl0.r1.c0.p0", position: "after", text: "Second line" },
    ]);
    expect(r!.success).toBe(true);
    const cell = doc.getTables()[0]!.rows[1]!.cells[0]!;
    expect(cell.paragraphs.map((p) => p.text)).toEqual(["Cell 1", "Second line"]);
    expect(cell.paragraphs[1]!.id).toBe("tbl0.r1.c0.p1");
    const xml = doc.toString();
    expect(xml).toMatch(/<w:tc>(?:(?!<\/w:tc>)[\s\S])*Cell 1(?:(?!<\/w:tc>)[\s\S])*Second line[\s\S]*?<\/w:tc>/);
    expect(doc.findParagraphById("tbl0.r1.c0.p1")!.text).toBe("Second line");
  });

  it("multi-line insert in a table cell keeps every line, in order, directly after the target (issue #36)", async () => {
    const u = "_".repeat(29);
    const signature = `<w:p>
      <w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:b/></w:rPr></w:pPr>
      <w:r><w:t>[CLIENT]</w:t></w:r><w:r><w:t xml:space="preserve"> LLC,</w:t></w:r>
      <w:r><w:br/></w:r><w:r><w:br/></w:r><w:r><w:br/></w:r>
      <w:r><w:t>By:</w:t></w:r><w:r><w:br/></w:r>
      <w:r><w:t>Name:</w:t></w:r><w:r><w:t xml:space="preserve">  _</w:t></w:r><w:r><w:t>${"_".repeat(28)}</w:t></w:r><w:r><w:br/></w:r>
      <w:r><w:t>Title:</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">  ${u}</w:t></w:r>
    </w:p>`;
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    const buffer = await createBodyDocx(
      `<w:p><w:r><w:t>Intro</w:t></w:r></w:p>
       <w:tbl><w:tr>
         <w:tc><w:p><w:r><w:t>LANDLORD:</w:t></w:r></w:p></w:tc>
         <w:tc>${signature}<w:p><w:r><w:t>By:</w:t></w:r></w:p></w:tc>
       </w:tr></w:tbl>`,
    );
    const doc = await Docxtor.load(buffer, { assignParaIds: true });
    const target = `para:${doc.findParagraphById("tbl0.r0.c1.p0")!.paraId}`;
    const text = `[CLIENT] LLC, REVISED\nBy:\nName:  ${u}\nTitle:  ${u}`;

    const results = doc.applyPatches(
      [
        { type: "insert_paragraph", target, position: "after", text, inheritStyle: true },
        { type: "delete_paragraph", target },
      ],
      TRACKED,
    );
    expect(results.every((r) => r.success)).toBe(true);

    const cellTexts = (d: Docxtor): string[] => {
      const json = d.project("json") as { body: Array<{ rows?: Array<Array<{ paragraphs?: Array<{ text: string }> }>> }> };
      return json.body[1]!.rows![0]![1]!.paragraphs!.map((p) => p.text);
    };
    const expected = ["", "[CLIENT] LLC, REVISED", "By:", `Name:  ${u}`, `Title:  ${u}`, "By:"];
    expect(cellTexts(doc)).toEqual(expected);

    // Same after a round trip, and the new paragraphs carry the inherited style
    const reloaded = await Docxtor.load(await doc.toBuffer());
    expect(cellTexts(reloaded)).toEqual(expected);
    expect(reloaded.findParagraphById("tbl0.r0.c1.p3")!.style).toBe("Normal");
    expect(reloaded.findParagraphById("tbl0.r0.c1.p3")!.runs[0]!.formatting.bold).toBe(true);

    // Accepting leaves the four new lines plus the pre-existing one
    reloaded.acceptAllRevisions();
    expect(cellTexts(reloaded)).toEqual(expected.slice(1));
  });

  it("works inside a header", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "insert_paragraph", target: "hdr.default.p0", position: "after", text: "Draft of 13 Oct 2026" },
    ]);
    const hdr = [...doc.getHeaders().values()][0]!;
    expect(hdr.paragraphs.map((p) => p.text)).toEqual(["Landlord Form Lease - CONFIDENTIAL", "Draft of 13 Oct 2026"]);
    expect(hdr.paragraphs[1]!.id).toBe("hdr.default.p1");
  });

  it("untracked insert has no revision markup", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    doc.applyPatches([{ type: "insert_paragraph", target: "p0", position: "after", text: "World" }]);
    expect(doc.toString()).not.toContain("w:ins");
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["Hello", "World"]);
  });

  it("rejects markdown tables and lists in text", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    const validations = doc.validatePatches([
      { type: "insert_paragraph", target: "p0", position: "after", text: "| a | b |\n| --- | --- |\n| 1 | 2 |" },
      { type: "insert_paragraph", target: "p0", position: "after", text: "- item one\n- item two" },
      { type: "insert_paragraph", target: "p0", position: "after", text: "" },
      { type: "insert_paragraph", target: "p99", position: "after", text: "x" },
    ]);
    expect(validations[0]!.valid).toBe(false);
    expect(validations[0]!.error).toContain("add_table");
    expect(validations[1]!.valid).toBe(false);
    expect(validations[2]!.valid).toBe(false);
    expect(validations[3]!.valid).toBe(false);
  });

  it("survives export and reload", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "insert_paragraph", target: "p4", position: "after", text: "Mirror." },
    ], TRACKED);
    const reloaded = await Docxtor.load(await doc.toBuffer());
    expect(reloaded.getParagraphs()[5]!.text).toBe("Mirror.");
    expect(reloaded.getParagraphs()[5]!.revisions?.[0]?.type).toBe("insert");
  });
});

// =============================================================================
// delete_paragraph
// =============================================================================

describe("delete_paragraph", () => {
  it("tracked: wraps runs in w:del, marks the paragraph mark, and accept removes it cleanly", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const originalCount = doc.getParagraphs().length;

    const [r] = doc.applyPatches([{ type: "delete_paragraph", target: "p6" }], TRACKED);
    expect(r!.success).toBe(true);

    const p6 = doc.findParagraphById("p6")!;
    expect(p6.text).toBe(""); // deleted text is excluded from normalized text
    expect(p6.revisions?.some((rev) => rev.type === "delete" && rev.text.includes("recapture"))).toBe(true);

    const xml = paragraphXml(doc, "p6");
    expect(xml).toContain("<w:delText");
    expect(xml).toMatch(/<w:pPr>.*<w:rPr><w:del w:id="\d+" w:author="USI Legal AI"/);
    expect(xml).not.toContain("<w:t ");
    expect(doc.getParagraphs().length).toBe(originalCount); // still present until accepted

    doc.acceptAllRevisions();
    const paras = doc.getParagraphs();
    expect(paras.length).toBe(originalCount - 1);
    expect(paras.some((p) => p.text.trim() === "")).toBe(false);
    expect(paras.some((p) => p.text.includes("recapture"))).toBe(false);
    // RECAPTURE heading (p5) is now followed by RENT (previously p7)
    expect(paras[5]!.text).toBe("RECAPTURE");
    expect(paras[6]!.text).toBe("RENT");
  });

  it("tracked: reject restores the paragraph", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([{ type: "delete_paragraph", target: "p6" }], TRACKED);
    doc.rejectAllRevisions({ authors: ["USI Legal AI"] });
    expect(doc.findParagraphById("p6")!.text).toBe("Landlord may recapture the Premises upon any request to assign or sublet.");
    expect(doc.toString()).not.toContain("USI Legal AI");
  });

  it("untracked: removes the w:p from XML and from the body", async () => {
    const doc = await Docxtor.load(await createMultiParagraphDocx(["One", "Two", "Three"]));
    const [r] = doc.applyPatches([{ type: "delete_paragraph", target: "p1" }]);
    expect(r!.success).toBe(true);
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["One", "Three"]);
    expect(doc.toString()).not.toContain("Two");
    expect(doc.findParagraphById("p1")).toBeUndefined();
  });

  it("handles another author's insertion inside the paragraph (nested w:del in w:ins)", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    // p8 contains Opposing Counsel's w:del($40.00) and w:ins($42.00)
    doc.applyPatches([{ type: "delete_paragraph", target: "p8" }], TRACKED);
    const xml = paragraphXml(doc, "p8");
    // Their insertion is kept, with our deletion nested inside
    expect(xml).toMatch(/<w:ins w:id="2" w:author="Opposing Counsel"[^>]*><w:del w:id="\d+" w:author="USI Legal AI"/);
    // Their deletion is untouched (not double wrapped)
    expect(xml).toMatch(/<w:del w:id="1" w:author="Opposing Counsel"[^>]*><w:r><w:delText>\$40.00/);

    doc.acceptAllRevisions();
    expect(doc.getParagraphs().some((p) => p.text.includes("Base Rent"))).toBe(false);
  });

  it("removes outright a paragraph that is wholly the author's own unaccepted insertion", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const originalCount = doc.getParagraphs().length;
    const results = doc.applyPatches([
      { type: "insert_paragraph", target: "p4", position: "after", text: "Oops." },
      { type: "delete_paragraph", target: "p11" },
    ], TRACKED);
    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.getParagraphs().length).toBe(originalCount);
    expect(doc.toString()).not.toContain("Oops.");
  });

  it("keeps a table cell valid when deleting its only paragraph (untracked)", async () => {
    const doc = await Docxtor.load(await createTableDocx());
    doc.applyPatches([{ type: "delete_paragraph", target: "tbl0.r1.c0.p0" }]);
    const cell = doc.getTables()[0]!.rows[1]!.cells[0]!;
    expect(cell.paragraphs.length).toBe(1);
    expect(cell.paragraphs[0]!.text).toBe("");
    expect(serializeXml([cell.xmlRef]).replace(/>\s+</g, "><")).toMatch(/<w:tc><w:tcPr>.*?<\/w:tcPr><w:p><\/w:p><\/w:tc>/);
  });

  it("accepts para: targets", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const [r] = doc.applyPatches([{ type: "delete_paragraph", target: "para:0A000007" }]);
    expect(r!.success).toBe(true);
    expect(doc.findParagraphByParaId("0A000007")).toBeUndefined();
  });

  it("converts field instructions to w:delInstrText", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([{ type: "delete_paragraph", target: "ftr.default.p0" }], TRACKED);
    const footer = [...doc.getFooters().values()][0]!;
    const { serializeXml } = await import("../../src/xml/serializer.js");
    const xml = serializeXml(footer.xmlRef);
    expect(xml).toContain("w:delInstrText");
    expect(xml).not.toContain("<w:instrText");
  });

  it("validation reports a missing target", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    const [v] = doc.validatePatches([{ type: "delete_paragraph", target: "p9" }]);
    expect(v!.valid).toBe(false);
  });
});

describe("delete_paragraph keeps every story valid (untracked)", () => {
  it("keeps an empty paragraph when deleting the only header paragraph", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const [r] = doc.applyPatches([{ type: "delete_paragraph", target: "hdr.default.p0" }]);
    expect(r!.success).toBe(true);
    const hdr = [...doc.getHeaders().values()][0]!;
    expect(hdr.paragraphs.length).toBe(1);
    expect(hdr.paragraphs[0]!.text).toBe("");
    expect(serializeXml(hdr.xmlRef)).toMatch(/<w:hdr[^>]*>\s*<w:p[^>]*><\/w:p>\s*<\/w:hdr>/);
  });

  it("keeps an empty paragraph when deleting the only body paragraph", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Only one"));
    doc.applyPatches([{ type: "delete_paragraph", target: "p0" }]);
    expect(doc.getParagraphs().length).toBe(1);
    expect(doc.getParagraphs()[0]!.text).toBe("");
    expect(doc.toString()).toContain("<w:p>");
  });

  it("removes a body paragraph outright when a table remains", async () => {
    const doc = await Docxtor.load(await createTableDocx());
    doc.applyPatches([{ type: "delete_paragraph", target: "p0" }, { type: "delete_paragraph", target: "p1" }]);
    expect(doc.getParagraphs().length).toBe(0);
    expect(doc.getTables().length).toBe(1);
  });
});
