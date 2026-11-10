/**
 * Issue #33: every `w:br` is one `\n` in paragraph text — in the projection
 * and in the offsets the patch engine resolves `find` / `occurrence` against.
 */

import { describe, it, expect } from "vitest";
import { Docxtor } from "../../src/index.js";
import { createBodyDocx } from "../fixtures/create-docx.js";

const U29 = "_".repeat(29);

/** The signature block from the issue: break-only runs between the lines. */
const SIGNATURE_PARAGRAPH = `<w:p>
  <w:r><w:t>[CLIENT]</w:t></w:r><w:r><w:t xml:space="preserve"> LLC,</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>
  <w:r><w:t xml:space="preserve">a(n) __________ </w:t></w:r><w:r><w:t>limited liability company</w:t></w:r>
  <w:r><w:br/></w:r><w:r><w:br/></w:r><w:r><w:br/></w:r>
  <w:r><w:t>By:</w:t></w:r><w:r><w:br/></w:r>
  <w:r><w:t>Name:</w:t></w:r><w:r><w:t xml:space="preserve">  _</w:t></w:r><w:r><w:t>${"_".repeat(28)}</w:t></w:r><w:r><w:br/></w:r>
  <w:r><w:t>Title:</w:t></w:r><w:r><w:t xml:space="preserve">  ${U29}</w:t></w:r>
</w:p>`;

const EXPECTED_TEXT =
  `[CLIENT] LLC, a(n) __________ limited liability company\n\n\nBy:\nName:  ${U29}\nTitle:  ${U29}`;

async function signatureCellDoc(): Promise<Docxtor> {
  return Docxtor.load(
    await createBodyDocx(`<w:tbl><w:tr><w:tc>${SIGNATURE_PARAGRAPH}</w:tc></w:tr></w:tbl>`),
  );
}

type Cell = { id: string; text: string };

describe("w:br in paragraph text (issue #33)", () => {
  it("projects one \\n per w:br, including consecutive break-only runs", async () => {
    const doc = await signatureCellDoc();
    const json = doc.project("json");
    const cell = (json.body[0] as { rows: Cell[][] }).rows[0]![0]!;
    expect(cell.text).toBe(EXPECTED_TEXT);
    expect(doc.findParagraphById("tbl0.r0.c0.p0")!.text).toBe(EXPECTED_TEXT);
  });

  it("markdown projection carries the same text", async () => {
    const doc = await signatureCellDoc();
    expect(doc.project("markdown")).toContain(EXPECTED_TEXT);
  });

  it("replace_text occurrence counted on the projected text lands on the Title line (tracked)", async () => {
    const doc = await signatureCellDoc();
    const [r] = doc.applyPatches(
      [{ type: "replace_text", target: "tbl0.r0.c0.p0", find: U29, replace: `${U29} X.`, occurrence: 2 }],
      { trackRevisions: true, revisionAuthor: "T", revisionDate: "2026-10-17T00:00:00Z" },
    );
    expect(r!.success).toBe(true);
    const text = doc.findParagraphById("tbl0.r0.c0.p0")!.text;
    expect(text).toBe(`${EXPECTED_TEXT} X.`);
    expect(text.split("\n")[3]).toBe("By:");
    expect(text.split("\n")[4]).toBe(`Name:  ${U29}`);
  });

  it("replace_text occurrence counted on the projected text lands on the Title line (untracked)", async () => {
    const doc = await signatureCellDoc();
    const [r] = doc.applyPatches([
      { type: "replace_text", target: "tbl0.r0.c0.p0", find: U29, replace: `${U29} X.`, occurrence: 2 },
    ]);
    expect(r!.success).toBe(true);
    expect(doc.findParagraphById("tbl0.r0.c0.p0")!.text).toBe(`${EXPECTED_TEXT} X.`);
  });

  it("find may span a line break; the break is deleted with the match and re-created from \\n in the replacement", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:t>By:</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Name:</w:t></w:r></w:p>`),
    );
    expect(doc.getParagraphs()[0]!.text).toBe("By:\nName:");

    const [r] = doc.applyPatches([{ type: "replace_text", target: "p0", find: "By:\nName:", replace: "Signed by:\nPrinted name:" }]);
    expect(r!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("Signed by:\nPrinted name:");

    const xml = doc.toString();
    expect(xml).toMatch(/<w:t>Signed by:<\/w:t><w:br\/><w:t>Printed name:<\/w:t>/);
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/); // never a raw newline inside w:t
  });

  it("tracked deletion across a break keeps the w:br inside w:del", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:t>By:</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Name:</w:t></w:r></w:p>`),
    );
    const [r] = doc.applyPatches([{ type: "delete_text", target: "p0", find: "\nName:" }], { trackRevisions: true });
    expect(r!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("By:");
    expect(doc.toString()).toMatch(/<w:del[^>]*><w:r><w:br\/><\/w:r><w:r><w:delText>Name:<\/w:delText><\/w:r><\/w:del>/);
    expect(doc.getParagraphs()[0]!.revisions).toEqual([
      expect.objectContaining({ type: "delete", text: "\nName:", startOffset: 3, endOffset: 3 }),
    ]);
    doc.rejectAllRevisions();
    expect(doc.getParagraphs()[0]!.text).toBe("By:\nName:");
  });

  it("insert_text anchored on either side of a break goes to the right line", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:t>By:</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Name:</w:t></w:r></w:p>`),
    );
    doc.applyPatches([
      { type: "insert_text", target: "p0", anchor: "By:", position: "after", text: " ____" },
      { type: "insert_text", target: "p0", anchor: "Name:", position: "before", text: "Printed " },
    ]);
    expect(doc.getParagraphs()[0]!.text).toBe("By: ____\nPrinted Name:");
  });

  it("a break-only run splitting a formatted-run match is handled by the run splitter without a raw newline", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t><w:br/><w:t>World</w:t></w:r></w:p>`),
    );
    expect(doc.getParagraphs()[0]!.text).toBe("Hello\nWorld");
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Hello", replace: "Hi" }]);
    expect(doc.getParagraphs()[0]!.text).toBe("Hi\nWorld");
    const xml = doc.toString();
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/);
    expect((xml.match(/<w:br\/>/g) ?? []).length).toBe(1);
  });

  it("untracked: deleting a break from a run that holds text on both sides removes the w:br", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t><w:br/><w:t>World</w:t></w:r></w:p>`),
    );
    const [r] = doc.applyPatches([{ type: "delete_text", target: "p0", find: "\n" }]);
    expect(r!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("HelloWorld");
    const xml = doc.toString();
    expect(xml).not.toMatch(/<w:br[ \/>]/);
    expect(xml).toContain("<w:b/>"); // formatting kept on the rebuilt runs
    expect(await Docxtor.load(await doc.toBuffer()).then((d) => d.getParagraphs()[0]!.text)).toBe("HelloWorld");
  });

  it("untracked: replacing a break-only run's newline with text drops the w:br", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:t>A</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>B</w:t></w:r></w:p>`),
    );
    const [r] = doc.applyPatches([{ type: "replace_text", target: "p0", find: "\n", replace: "-" }]);
    expect(r!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("A-B");
    expect(doc.toString()).not.toMatch(/<w:br[ \/>]/);
  });

  it("untracked: replacing text that spans a break within one run, with a single-line replacement", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:t>By:</w:t><w:br/><w:t>Name:</w:t></w:r></w:p>`),
    );
    doc.applyPatches([{ type: "replace_text", target: "p0", find: ":\nName", replace: " / Name" }]);
    expect(doc.getParagraphs()[0]!.text).toBe("By / Name:");
    expect(doc.toString()).not.toMatch(/<w:br[ \/>]/);
  });

  it("untracked: a trailing page break in the matched run stays after the replacement", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:t>Hello</w:t><w:br/><w:t>World</w:t><w:br w:type="page"/></w:r></w:p>`),
    );
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Hello\nWorld", replace: "Hi" }]);
    const para = doc.getParagraphs()[0]!;
    expect(para.text).toBe("Hi");
    expect(para.pageBreakOffset).toBe(2);
    expect(doc.toString().replace(/>\s+</g, "><")).toMatch(/<w:t>Hi<\/w:t><w:br w:type="page"\/><\/w:r>/);
  });

  it("untracked: a leading page break in the matched run stays before the replacement", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:br w:type="page"/><w:t>Hello</w:t><w:br/><w:t>World</w:t></w:r></w:p>`),
    );
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Hello\nWorld", replace: "Hi" }]);
    const para = doc.getParagraphs()[0]!;
    expect(para.text).toBe("Hi");
    expect(para.pageBreakOffset).toBe(0);
    expect(doc.toString().replace(/>\s+</g, "><")).toMatch(/<w:r><w:rPr><w:b\/><\/w:rPr><w:br w:type="page"\/><w:t>Hi<\/w:t><\/w:r>/);
  });

  it("page and column breaks still contribute no text", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:t>A</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:br w:type="column"/></w:r><w:r><w:t>B</w:t></w:r></w:p>`),
    );
    expect(doc.getParagraphs()[0]!.text).toBe("AB");
  });

  it("round-trips: save, reload, same text and offsets", async () => {
    const doc = await signatureCellDoc();
    const reloaded = await Docxtor.load(await doc.toBuffer());
    expect(reloaded.findParagraphById("tbl0.r0.c0.p0")!.text).toBe(EXPECTED_TEXT);
  });
});
