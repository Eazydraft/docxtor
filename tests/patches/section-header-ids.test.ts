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

describe("a part referenced more than once", () => {
  /** Rewrite the fixture so section 2's first-page header is reached through a second relationship to header1.xml. */
  async function sharedPartDocx(sameSection: boolean): Promise<Buffer> {
    const zip = await JSZip.loadAsync(await createTwoSectionDocx());
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    zip.file("word/_rels/document.xml.rels", rels.replace('Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"', 'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"'));
    if (sameSection) {
      // Both relationship ids referenced from section 1: as its first-page and its even-page header
      const doc = await zip.file("word/document.xml")!.async("string");
      zip.file("word/document.xml", doc
        .replace('<w:headerReference w:type="first" r:id="rId1"/>', '<w:headerReference w:type="first" r:id="rId1"/><w:headerReference w:type="even" r:id="rId3"/>')
        .replace('<w:headerReference w:type="first" r:id="rId3"/>', ""));
    }
    return zip.generateAsync({ type: "nodebuffer" }) as Promise<Buffer>;
  }

  it.each([
    ["from two sections", false],
    ["twice from one section", true],
  ])("under two relationship ids %s appears once, with its first reference", async (_name, sameSection) => {
    const doc = await Docxtor.load(await sharedPartDocx(sameSection));
    const headers = [...doc.getHeaders().values()];
    const shared = headers.filter((h) => h.paragraphs[0]!.text === "Section one first-page header");
    expect(shared).toHaveLength(1);
    expect([shared[0]!.rId, shared[0]!.sectionIndex, shared[0]!.variant]).toEqual(["rId1", 1, "first"]);

    // No id names two paragraphs, anywhere
    const ids = [...headers, ...doc.getFooters().values()].flatMap((hf) => hf.paragraphs.map((p) => p.id));
    expect(new Set(ids).size).toBe(ids.length);
    const md = projectToMarkdown(doc.getAST());
    expect(md.match(/Section one first-page header/g)).toHaveLength(1);
  });

  it("an edit to the shared part is written once", async () => {
    const input = await sharedPartDocx(false);
    const doc = await Docxtor.load(input);
    const id = [...doc.getHeaders().values()][0]!.paragraphs[0]!.id;
    const results = doc.applyPatches([{ type: "replace_text", target: id, find: "Section one", replace: "Shared" }]);
    expect(results[0]!.success).toBe(true);
    const out = await part(await doc.toBuffer(), "word/header1.xml");
    expect(out.match(/Shared/g)).toHaveLength(1);
  });
});

describe("one header part shared by every section", () => {
  /** Two sections whose only header is header1.xml: by one relationship id, or by two. */
  async function onlySharedDocx(twoRelationships: boolean): Promise<Buffer> {
    const zip = await JSZip.loadAsync(await createTwoSectionDocx());
    const second = twoRelationships ? "rId3" : "rId1";
    let doc = await zip.file("word/document.xml")!.async("string");
    doc = doc
      .replace(/<w:footerReference [^>]*\/>/g, "")
      .replace('<w:headerReference w:type="default" r:id="rId5"/>', "")
      .replace('<w:headerReference w:type="first" r:id="rId3"/>', `<w:headerReference w:type="first" r:id="${second}"/>`);
    zip.file("word/document.xml", doc);
    if (twoRelationships) {
      const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
      zip.file("word/_rels/document.xml.rels", rels.replace('Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"', 'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"'));
    }
    return zip.generateAsync({ type: "nodebuffer" }) as Promise<Buffer>;
  }

  it.each([
    ["one relationship id", false],
    ["two relationship ids", true],
  ])("still gets section ids, because two sections reference it (%s)", async (_name, two) => {
    const doc = await Docxtor.load(await onlySharedDocx(two));
    const headers = [...doc.getHeaders().values()];
    expect(headers).toHaveLength(1);
    expect([headers[0]!.sectionIndex, headers[0]!.idPrefix, headers[0]!.paragraphs[0]!.id]).toEqual([1, "hdr.s1", "hdr.s1.first.p0"]);
    expect(projectToMarkdown(doc.getAST())).toContain("--- Header (first, section 1) ---");
  });

  it.each([
    ["one relationship id", false],
    ["two relationship ids", true],
  ])("remove_header on the last section keeps the part section 1 still uses (%s)", async (_name, two) => {
    const doc = await Docxtor.load(await onlySharedDocx(two));
    const results = doc.applyPatches([{ type: "remove_header", variant: "first" }]);
    expect(results[0]!.success).toBe(true);

    // Section 1 still has its header, now the only section with one
    const headers = [...doc.getHeaders().values()];
    expect(headers.map((h) => [h.rId, h.paragraphs[0]!.id, h.paragraphs[0]!.text])).toEqual([
      ["rId1", "hdr.first.p0", "Section one first-page header"],
    ]);

    const output = await doc.toBuffer();
    const zip = await JSZip.loadAsync(output);
    expect(zip.file("word/header1.xml")).not.toBeNull();
    expect(await part(output, "[Content_Types].xml")).toContain('PartName="/word/header1.xml"');
    const rels = await part(output, "word/_rels/document.xml.rels");
    expect(rels).toContain('Id="rId1"');
    // The last section's own relationship to the part is gone; section 1's stays
    expect(rels.match(/Target="header1\.xml"/g)).toHaveLength(1);
    // Every header reference left in the document resolves to a relationship
    const refs = [...(await part(output, "word/document.xml")).matchAll(/<w:headerReference [^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
    expect(refs).toEqual(["rId1"]);

    const reloaded = await Docxtor.load(output);
    expect([...reloaded.getHeaders().values()].map((h) => h.paragraphs[0]!.text)).toEqual(["Section one first-page header"]);
  });

  it("remove_header still deletes a part nothing else references", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());
    const results = doc.applyPatches([{ type: "remove_header", variant: "first" }]);
    expect(results[0]!.success).toBe(true);
    const output = await doc.toBuffer();
    const zip = await JSZip.loadAsync(output);
    expect(zip.file("word/header2.xml")).toBeNull();
    expect(zip.file("word/header1.xml")).not.toBeNull();
    expect(await part(output, "[Content_Types].xml")).not.toContain("/word/header2.xml");
    expect(await part(output, "word/_rels/document.xml.rels")).not.toContain("header2.xml");
  });
});

describe("container on a part shared by sections", () => {
  /** Section 2's first-page header is header1.xml too, through its own relationship id. */
  async function sharedFirstHeaderDocx(): Promise<Buffer> {
    const zip = await JSZip.loadAsync(await createTwoSectionDocx());
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    zip.file("word/_rels/document.xml.rels", rels.replace('Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"', 'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"'));
    return zip.generateAsync({ type: "nodebuffer" }) as Promise<Buffer>;
  }
  const append = (section?: number): Patch => ({
    type: "append_text",
    text: "Draft",
    container: { type: "header", variant: "first", ...(section ? { section } : {}) },
  });

  it("the part lists every reference to it", async () => {
    const doc = await Docxtor.load(await sharedFirstHeaderDocx());
    const shared = [...doc.getHeaders().values()].find((h) => h.variant === "first")!;
    expect(shared.references).toEqual([
      { sectionIndex: 1, variant: "first", rId: "rId1" },
      { sectionIndex: 2, variant: "first", rId: "rId3" },
    ]);
    const own = [...doc.getHeaders().values()].find((h) => h.variant === "default")!;
    expect(own.references).toEqual([{ sectionIndex: 2, variant: "default", rId: "rId5" }]);
  });

  it.each([[1], [2]])("section %i resolves to the shared part", async (section) => {
    const doc = await Docxtor.load(await sharedFirstHeaderDocx());
    expect(doc.validatePatches([append(section)])[0]).toEqual({ valid: true });
    const results = doc.applyPatches([append(section)]);
    expect(results[0]!.success).toBe(true);
    // The new paragraph takes the part's own ids, whichever section named it
    expect(doc.findParagraphById("hdr.s1.first.p1")!.text).toBe("Draft");
    const out = await part(await doc.toBuffer(), "word/header1.xml");
    expect(out.match(/Draft/g)).toHaveLength(1);
  });

  it("needs no section: both sections name the same part", async () => {
    const doc = await Docxtor.load(await sharedFirstHeaderDocx());
    const results = doc.applyPatches([append()]);
    expect(results[0]!.success).toBe(true);
    expect(doc.findParagraphById("hdr.s1.first.p1")!.text).toBe("Draft");
  });

  it("a section that does not reference the part is still refused, and the message lists both that do", async () => {
    const doc = await Docxtor.load(await sharedFirstHeaderDocx());
    expect(doc.validatePatches([append(3)])[0]).toEqual({
      valid: false,
      error: 'No header with variant "first" in section 3; it exists in sections 1 and 2',
    });
  });

  it("a part used as two variants is found by either, and new paragraphs keep its ids", async () => {
    const zip = await JSZip.loadAsync(await sharedFirstHeaderDocx());
    const xml = await zip.file("word/document.xml")!.async("string");
    zip.file("word/document.xml", xml.replace('<w:headerReference w:type="first" r:id="rId3"/>', '<w:headerReference w:type="even" r:id="rId3"/>'));
    const doc = await Docxtor.load(await zip.generateAsync({ type: "nodebuffer" }) as Buffer);
    const results = doc.applyPatches([{ type: "append_text", text: "Draft", container: { type: "header", variant: "even" } }]);
    expect(results[0]!.success).toBe(true);
    expect(doc.findParagraphById("hdr.s1.first.p1")!.text).toBe("Draft");
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
