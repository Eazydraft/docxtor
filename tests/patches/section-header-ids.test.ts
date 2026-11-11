/**
 * Header and footer ids in a document with more than one section: every id
 * carries its section (`hdr.s2.first.p0`), and a patch against one section's
 * header never lands in another's.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor, projectToJson, projectToMarkdown } from "../../src/index.js";
import type { Patch } from "../../src/index.js";
import { createTwoSectionDocx, createHeaderFooterDocx } from "../fixtures/create-docx.js";

async function part(buffer: Buffer, path: string): Promise<string> {
  return (await JSZip.loadAsync(buffer)).file(path)!.async("string");
}

const TRACKED = { trackRevisions: true, revisionAuthor: "Reviewer", revisionDate: "2026-11-10T00:00:00Z" };

describe("header and footer ids across sections", () => {
  it("number the sections in document order and put the section into every id", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());

    const headers = [...doc.getHeaders().values()];
    expect(headers.map((h) => [h.sectionIndex, h.variant, h.idPrefix, h.paragraphs[0]!.id, h.paragraphs[0]!.text])).toEqual([
      [1, "first", "hdr.s1", "hdr.s1.first.p0", "Section one first-page header"],
      [2, "default", "hdr.s2", "hdr.s2.default.p0", "Running header"],
      [2, "first", "hdr.s2", "hdr.s2.first.p0", "Section two first-page header"],
    ]);
    const footers = [...doc.getFooters().values()];
    expect(footers.map((f) => [f.sectionIndex, f.paragraphs[0]!.id, f.paragraphs[0]!.text])).toEqual([
      [1, "ftr.s1.first.p0", ""],
      [2, "ftr.s2.first.p0", "Prepared by Acme Legal"],
    ]);

    // No id is used twice anywhere in the document
    const ids = [...headers, ...footers].flatMap((hf) => hf.paragraphs.map((p) => p.id));
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("the view names the section in each divider, section 1 first", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());
    const md = projectToMarkdown(doc.getAST());
    const lines = md.split("\n").filter((l) => /^--- |^\[(hdr|ftr)\./.test(l));
    expect(lines).toEqual([
      "--- Header (first, section 1) ---",
      "[hdr.s1.first.p0] Section one first-page header",
      "--- Header (default, section 2) ---",
      "[hdr.s2.default.p0] Running header",
      "--- Header (first, section 2) ---",
      "[hdr.s2.first.p0] Section two first-page header",
      "--- Footer (first, section 1) ---",
      "[ftr.s1.first.p0] ",
      "--- Footer (first, section 2) ---",
      "[ftr.s2.first.p0] Prepared by Acme Legal",
    ]);

    const json = projectToJson(doc.getAST());
    expect(json.headers!.map((h) => [h.variant, h.section])).toEqual([["first", 1], ["default", 2], ["first", 2]]);
    expect(json.footers!.map((f) => [f.variant, f.section])).toEqual([["first", 1], ["first", 2]]);
  });

  it("a single-section document keeps the short ids and the plain divider", async () => {
    const doc = await Docxtor.load(await createHeaderFooterDocx());
    const header = [...doc.getHeaders().values()][0]!;
    expect([header.sectionIndex, header.idPrefix, header.paragraphs[0]!.id]).toEqual([1, "hdr", "hdr.default.p0"]);
    const md = projectToMarkdown(doc.getAST());
    expect(md).toContain("--- Header (default) ---");
    expect(md).toContain("--- Footer (default) ---");
    expect(projectToJson(doc.getAST()).headers![0]).not.toHaveProperty("section");
  });

  it("replace_text against hdr.s2.first.p0 changes section 2's header and leaves section 1's untouched", async () => {
    const input = await createTwoSectionDocx();
    const doc = await Docxtor.load(input);
    const results = doc.applyPatches(
      [{ type: "replace_text", target: "hdr.s2.first.p0", find: "Section two", replace: "Part B" }],
      TRACKED,
    );
    expect(results[0]!.success).toBe(true);
    expect(doc.findParagraphById("hdr.s2.first.p0")!.text).toBe("Part B first-page header");
    expect(doc.findParagraphById("hdr.s1.first.p0")!.text).toBe("Section one first-page header");

    const output = await doc.toBuffer();
    expect(await part(output, "word/header2.xml")).toContain("Part B");
    expect(await part(output, "word/header1.xml")).toBe(await part(input, "word/header1.xml"));
    expect(await part(output, "word/header3.xml")).toBe(await part(input, "word/header3.xml"));
    expect(await part(output, "word/document.xml")).toBe(await part(input, "word/document.xml"));
  });

  it("an id without the section does not resolve in a multi-section document", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());
    expect(doc.findParagraphById("hdr.first.p0")).toBeUndefined();
    const results = doc.applyPatches([{ type: "replace_text", target: "hdr.first.p0", find: "Section", replace: "Part" }]);
    expect(results[0]!.success).toBe(false);
  });

  it("insert_paragraph after a section-2 footer paragraph lands in that footer with a section-2 id", async () => {
    const input = await createTwoSectionDocx();
    const doc = await Docxtor.load(input);
    const results = doc.applyPatches([
      { type: "insert_paragraph", target: "ftr.s2.first.p0", position: "after", text: "Confidential" },
    ]);
    expect(results[0]!.success).toBe(true);
    expect(doc.findParagraphById("ftr.s2.first.p1")!.text).toBe("Confidential");

    const output = await doc.toBuffer();
    expect(await part(output, "word/footer2.xml")).toContain("Confidential");
    expect(await part(output, "word/footer1.xml")).toBe(await part(input, "word/footer1.xml"));
  });
});

describe("container with section", () => {
  const append = (section?: number): Patch => ({
    type: "append_text",
    text: "Draft",
    container: { type: "header", variant: "first", ...(section ? { section } : {}) },
  });

  it("fails validation when the variant exists in more than one section and none is given", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());
    expect(doc.validatePatches([append()])[0]).toEqual({
      valid: false,
      error: "header 'first' exists in sections 1 and 2; say which with `section`",
    });
    const results = doc.applyPatches([append()]);
    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("say which with `section`");
  });

  it("appends to the named section's part", async () => {
    const input = await createTwoSectionDocx();
    const doc = await Docxtor.load(input);
    const results = doc.applyPatches([append(2)]);
    expect(results[0]!.success).toBe(true);
    expect(doc.findParagraphById("hdr.s2.first.p1")!.text).toBe("Draft");

    const output = await doc.toBuffer();
    expect(await part(output, "word/header2.xml")).toContain("Draft");
    expect(await part(output, "word/header1.xml")).toBe(await part(input, "word/header1.xml"));

    const reloaded = await Docxtor.load(output);
    expect(reloaded.findParagraphById("hdr.s2.first.p1")!.text).toBe("Draft");
    expect(reloaded.findParagraphById("hdr.s1.first.p1")).toBeUndefined();
  });

  it("needs no section when only one section has that variant", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());
    const results = doc.applyPatches([{ type: "append_text", text: "Draft", container: { type: "header", variant: "default" } }]);
    expect(results[0]!.success).toBe(true);
    expect(doc.findParagraphById("hdr.s2.default.p1")!.text).toBe("Draft");
  });

  it("add_table and add_list honour the section", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());
    const results = doc.applyPatches([
      { type: "add_table", headers: ["A", "B"], rows: [["1", "2"]], container: { type: "footer", variant: "first", section: 1 } },
      { type: "add_list", listType: "bullet", items: ["one"], container: { type: "footer", variant: "first", section: 2 } },
    ]);
    expect(results.map((r) => r.success)).toEqual([true, true]);
    const [f1, f2] = [...doc.getFooters().values()];
    expect(f1!.tables.map((t) => t.id)).toEqual(["ftr.s1.first.tbl0"]);
    expect(f2!.paragraphs.map((p) => p.id)).toEqual(["ftr.s2.first.p0", "ftr.s2.first.p1"]);
    expect(f2!.tables).toHaveLength(0);
  });

  it("rejects a section that has no such part, and a section that is not a positive integer", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());
    expect(doc.validatePatches([{ type: "append_text", text: "x", container: { type: "header", variant: "default", section: 1 } }])[0]).toEqual({
      valid: false,
      error: 'No header with variant "default" in section 1; it exists in section 2',
    });
    expect(doc.validatePatches([append(0)])[0]!.valid).toBe(false);
    expect(doc.validatePatches([append(3)])[0]!.valid).toBe(false);
  });

  it("section 1 is accepted on a single-section document, and omitting it still works", async () => {
    const doc = await Docxtor.load(await createHeaderFooterDocx());
    const results = doc.applyPatches([
      { type: "append_text", text: "A", container: { type: "header", variant: "default" } },
      { type: "append_text", text: "B", container: { type: "header", variant: "default", section: 1 } },
    ]);
    expect(results.map((r) => r.success)).toEqual([true, true]);
    expect([...doc.getHeaders().values()][0]!.paragraphs.map((p) => p.id)).toEqual(["hdr.default.p0", "hdr.default.p1", "hdr.default.p2"]);
  });
});
