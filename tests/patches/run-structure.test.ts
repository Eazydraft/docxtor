/**
 * Edits must keep a run's zero-width children (page and column breaks, tabs,
 * rendering hints) and must handle a match that spans runs in different
 * parents, such as text that starts inside another author's insertion.
 */

import { describe, it, expect } from "vitest";
import { Docxtor } from "../../src/index.js";
import { createBodyDocx } from "../fixtures/create-docx.js";

const body = (xml: string): string =>
  xml.replace(/^[\s\S]*<w:body>/, "").replace(/<w:sectPr[\s\S]*$/, "").replace(/<\/w:body>[\s\S]*$/, "");

async function load(paragraphXml: string): Promise<Docxtor> {
  return Docxtor.load(await createBodyDocx(paragraphXml));
}

function first(doc: Docxtor): { text: string; hasPageBreak?: boolean; pageBreakOffset?: number } {
  return doc.project("json", { includePageBreaks: true }).body[0] as never;
}

describe("zero-width children survive text edits", () => {
  it("keeps a leading page break when the word after it is replaced", async () => {
    const doc = await load(`<w:p><w:r><w:br w:type="page"/><w:t>Hello World</w:t></w:r></w:p>`);
    const r = doc.applyPatches([{ type: "replace_text", target: "p0", find: "Hello", replace: "Hi" }]);
    expect(r[0]!.success).toBe(true);
    expect(first(doc)).toMatchObject({ text: "Hi World", hasPageBreak: true, pageBreakOffset: 0 });
    expect(body(doc.toString())).toBe(
      `<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Hi</w:t></w:r><w:r><w:t xml:space="preserve"> World</w:t></w:r></w:p>`,
    );
  });

  it("keeps a page break between two w:t nodes at its character position", async () => {
    const doc = await load(`<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "A", replace: "XY" }]);
    expect(first(doc)).toMatchObject({ text: "XYB", hasPageBreak: true, pageBreakOffset: 2 });
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "B", replace: "Z" }]);
    expect(first(doc)).toMatchObject({ text: "XYZ", hasPageBreak: true, pageBreakOffset: 2 });
  });

  it("keeps a trailing lastRenderedPageBreak when the last word is replaced", async () => {
    const doc = await load(`<w:p><w:r><w:t>Hello World</w:t><w:lastRenderedPageBreak/></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "World", replace: "There" }]);
    expect(first(doc)).toMatchObject({ text: "Hello There", hasPageBreak: true, pageBreakOffset: 11 });
    expect(body(doc.toString())).toBe(
      `<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>There</w:t></w:r><w:r><w:lastRenderedPageBreak/></w:r></w:p>`,
    );
  });

  it("keeps a tab between two text nodes", async () => {
    const doc = await load(`<w:p><w:r><w:t>Name:</w:t><w:tab/><w:t>Alice</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Alice", replace: "Bob" }]);
    expect(body(doc.toString())).toBe(`<w:p><w:r><w:t>Name:</w:t><w:tab/></w:r><w:r><w:t>Bob</w:t></w:r></w:p>`);
  });

  it("keeps a page break inside a match that is replaced whole", async () => {
    const doc = await load(`<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "AB", replace: "XY" }]);
    expect(first(doc)).toMatchObject({ text: "XY", hasPageBreak: true });
  });

  it("keeps a page break inside deleted text", async () => {
    const doc = await load(`<w:p><w:r><w:t>Before </w:t></w:r><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t></w:r><w:r><w:t> after</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "delete_text", target: "p0", find: "AB" }]);
    expect(first(doc)).toMatchObject({ text: "Before  after", hasPageBreak: true });
  });

  it("keeps a trailing page break through a formatted replacement", async () => {
    const doc = await load(`<w:p><w:r><w:t>Hello World</w:t><w:br w:type="page"/></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Hello World", replace: "**Hi**" }]);
    expect(first(doc)).toMatchObject({ text: "Hi", hasPageBreak: true, pageBreakOffset: 2 });
    expect(body(doc.toString())).toBe(
      `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hi</w:t><w:br w:type="page"/></w:r></w:p>`,
    );
  });

  it("keeps a leading page break through a formatted partial replacement", async () => {
    const doc = await load(`<w:p><w:r><w:br w:type="page"/><w:t>Hello World</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "World", replace: "_There_" }]);
    expect(first(doc)).toMatchObject({ text: "Hello There", hasPageBreak: true, pageBreakOffset: 0 });
  });

  it("keeps a page break in place when text is inserted on either side of it", async () => {
    const doc = await load(`<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "insert_text", target: "p0", position: "after", anchor: "A", text: "12" }]);
    expect(first(doc)).toMatchObject({ text: "A12B", hasPageBreak: true, pageBreakOffset: 3 });
    doc.applyPatches([{ type: "insert_text", target: "p0", position: "before", anchor: "B", text: "34" }]);
    expect(first(doc)).toMatchObject({ text: "A1234B", hasPageBreak: true, pageBreakOffset: 3 });
  });

  it("keeps a page break when formatted text is inserted into the run", async () => {
    const doc = await load(`<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "insert_text", target: "p0", position: "before", anchor: "B", text: "**x**" }]);
    expect(first(doc)).toMatchObject({ text: "AxB", hasPageBreak: true, pageBreakOffset: 1 });
  });

  it("keeps a page break when a tracked insertion splits the run", async () => {
    const doc = await load(`<w:p><w:r><w:t>Hello</w:t><w:br w:type="page"/><w:t>World</w:t></w:r></w:p>`);
    doc.applyPatches(
      [{ type: "insert_text", target: "p0", position: "after", anchor: "Hel", text: "p" }],
      { trackRevisions: true, revisionAuthor: "Me" },
    );
    expect(first(doc)).toMatchObject({ text: "HelploWorld", hasPageBreak: true, pageBreakOffset: 6 });
    const clean = await Docxtor.load(await doc.toBuffer());
    clean.rejectAllRevisions();
    expect(first(clean)).toMatchObject({ text: "HelloWorld", hasPageBreak: true, pageBreakOffset: 5 });
  });

  it("keeps a page break next to a tracked replacement", async () => {
    const doc = await load(`<w:p><w:r><w:t>Hello World</w:t><w:br w:type="page"/></w:r></w:p>`);
    doc.applyPatches(
      [{ type: "replace_text", target: "p0", find: "World", replace: "There" }],
      { trackRevisions: true, revisionAuthor: "Me" },
    );
    expect(first(doc)).toMatchObject({ text: "Hello There", hasPageBreak: true, pageBreakOffset: 11 });
    const clean = await Docxtor.load(await doc.toBuffer());
    clean.acceptAllRevisions();
    expect(first(clean)).toMatchObject({ text: "Hello There", hasPageBreak: true, pageBreakOffset: 11 });
  });
});

describe("a match that spans another author's insertion and plain text", () => {
  const xml = `<w:p><w:r><w:t xml:space="preserve">Pay within </w:t></w:r><w:ins w:id="1" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:t>thirty</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> (30) days.</w:t></w:r></w:p>`;

  it("tracked replace: the plain part is deleted, the inserted part removed", async () => {
    const doc = await load(xml);
    const r = doc.applyPatches(
      [{ type: "replace_text", target: "p0", find: "thirty (30)", replace: "sixty (60)" }],
      { trackRevisions: true, revisionAuthor: "Me" },
    );
    expect(r[0]!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("Pay within sixty (60) days.");
    const out = body(doc.toString());
    expect(out).toContain(`<w:delText xml:space="preserve"> (30)</w:delText>`);
    expect(out).not.toContain(`w:author="Bob"`); // Bob's insertion held only the matched text
    expect(out).toMatch(/<w:ins [^>]*w:author="Me"[^>]*><w:r><w:t>sixty \(60\)<\/w:t><\/w:r><\/w:ins><w:del /);

    const accepted = await Docxtor.load(await doc.toBuffer());
    accepted.acceptAllRevisions();
    expect(accepted.getParagraphs()[0]!.text).toBe("Pay within sixty (60) days.");

    const rejected = await Docxtor.load(await doc.toBuffer());
    rejected.rejectAllRevisions();
    expect(rejected.getParagraphs()[0]!.text).toBe("Pay within  (30) days.");
  });

  it("tracked replace keeps the rest of a longer foreign insertion", async () => {
    const doc = await load(
      `<w:p><w:r><w:t xml:space="preserve">Pay within </w:t></w:r><w:ins w:id="1" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">net thirty</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> (30) days.</w:t></w:r></w:p>`,
    );
    doc.applyPatches(
      [{ type: "replace_text", target: "p0", find: "thirty (30)", replace: "sixty (60)" }],
      { trackRevisions: true, revisionAuthor: "Me" },
    );
    expect(doc.getParagraphs()[0]!.text).toBe("Pay within net sixty (60) days.");
    const out = body(doc.toString());
    expect(out).toMatch(/<w:ins [^>]*w:author="Bob"[^>]*><w:r><w:t xml:space="preserve">net <\/w:t><\/w:r><\/w:ins><w:ins [^>]*w:author="Me"/);
    const rejected = await Docxtor.load(await doc.toBuffer());
    rejected.rejectAllRevisions();
    expect(rejected.getParagraphs()[0]!.text).toBe("Pay within  (30) days.");
  });

  it("tracked delete: the plain part is deleted, the inserted part removed", async () => {
    const doc = await load(xml);
    doc.applyPatches(
      [{ type: "delete_text", target: "p0", find: "thirty (30)" }],
      { trackRevisions: true, revisionAuthor: "Me" },
    );
    expect(doc.getParagraphs()[0]!.text).toBe("Pay within  days.");
    const out = body(doc.toString());
    expect(out).toContain(`<w:delText xml:space="preserve"> (30)</w:delText>`);
    expect(out).not.toContain(`<w:t>thirty</w:t>`);
  });

  it("tracked replace of a match that starts in plain text and ends in an insertion", async () => {
    const doc = await load(
      `<w:p><w:r><w:t xml:space="preserve">Pay within </w:t></w:r><w:ins w:id="1" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:r><w:t>days.</w:t></w:r></w:ins></w:p>`,
    );
    const r = doc.applyPatches(
      [{ type: "replace_text", target: "p0", find: "within days", replace: "in 30 days" }],
      { trackRevisions: true, revisionAuthor: "Me" },
    );
    expect(r[0]!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("Pay in 30 days.");
    const out = body(doc.toString());
    expect(out).toMatch(/<w:del [^>]*><w:r><w:delText xml:space="preserve">within <\/w:delText><\/w:r><\/w:del><w:ins [^>]*w:author="Me"[^>]*><w:r><w:t>in 30 days<\/w:t><\/w:r><\/w:ins><w:ins [^>]*w:author="Bob"[^>]*><w:r><w:t>\.<\/w:t><\/w:r><\/w:ins>/);
    const accepted = await Docxtor.load(await doc.toBuffer());
    accepted.acceptAllRevisions();
    expect(accepted.getParagraphs()[0]!.text).toBe("Pay in 30 days.");
  });

  it("untracked replace gets the text right", async () => {
    const doc = await load(xml);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "thirty (30)", replace: "sixty (60)" }]);
    expect(doc.getParagraphs()[0]!.text).toBe("Pay within sixty (60) days.");
  });
});
