/**
 * Byte-level fidelity of a save: a document saved with no patches returns
 * every part unchanged, and an edit changes only the parts it touches.
 *
 * This holds for XML written the way Word, LibreOffice and Google Docs write
 * it. The parser does not keep formatting inside a tag, so a part with line
 * breaks between attributes or single-quoted attributes comes back as
 * equivalent XML, not the same bytes. Empty elements keep the form the part
 * uses, self-closed or an open-close pair, but a part that mixes the two is
 * written self-closed.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor } from "../../src/index.js";
import { createLeaseDocx, createBodyDocx, createContentControlDocx, createTwoSectionDocx } from "../fixtures/create-docx.js";

/** Every file in a .docx (directory entries left out), as path → bytes. */
async function parts(buffer: Buffer): Promise<Map<string, Buffer>> {
  const zip = await JSZip.loadAsync(buffer);
  const out = new Map<string, Buffer>();
  for (const [path, file] of Object.entries(zip.files)) {
    if (!file.dir) out.set(path, await file.async("nodebuffer"));
  }
  return out;
}

/** Paths whose bytes differ between two packages, plus any added or removed. */
async function changedParts(before: Buffer, after: Buffer): Promise<string[]> {
  const [a, b] = [await parts(before), await parts(after)];
  const changed: string[] = [];
  for (const path of new Set([...a.keys(), ...b.keys()])) {
    const [x, y] = [a.get(path), b.get(path)];
    if (!x || !y || !x.equals(y)) changed.push(path);
  }
  return changed.sort();
}

/** A minimal file in the shape Word writes: CRLF after the declaration, one line, self-closed tags. */
function wordStyleDocx(): Promise<Buffer> {
  return createBodyDocx(
    `<w:p><w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr><w:bookmarkStart w:id="0" w:name="_GoBack"/>` +
    `<w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Payment is due </w:t></w:r>` +
    `<w:r><w:t>within thirty days.</w:t><w:tab/></w:r><w:bookmarkEnd w:id="0"/></w:p>` +
    `<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`,
  ).then(async (buffer) => {
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    zip.file("word/document.xml", xml.replace("?>\n", "?>\r\n"));
    return zip.generateAsync({ type: "nodebuffer" }) as Promise<Buffer>;
  });
}

describe("a save with no patches", () => {
  it.each([
    ["lease (numbering, table, header, footer, footnotes, tracked changes, properties)", () => createLeaseDocx()],
    ["Word-style single-line XML with CRLF", () => wordStyleDocx()],
    ["content controls (inline, around a cell, around paragraphs, nested, in a header)", () => createContentControlDocx()],
    ["two sections, empty docProps properties as open-close pairs, an apostrophe in an attribute", () => createTwoSectionDocx()],
    ["two sections with native paragraph ids", () => createTwoSectionDocx({ withParaIds: true })],
  ])("returns every part byte for byte: %s", async (_name, make) => {
    const input = await make();
    const output = await (await Docxtor.load(input)).toBuffer();
    expect(await changedParts(input, output)).toEqual([]);
  });

  it("is still byte for byte after a failed atomic batch", async () => {
    const input = await createLeaseDocx();
    const doc = await Docxtor.load(input);
    const results = doc.applyPatches(
      [
        { type: "replace_text", target: "p2", find: "Landlord", replace: "Lessor" },
        { type: "replace_text", target: "p2", find: "no such text anywhere", replace: "x" },
      ],
      { trackRevisions: true, atomic: true },
    );
    expect(results.some((r) => !r.success)).toBe(true);
    expect(await changedParts(input, await doc.toBuffer())).toEqual([]);
  });
});

describe("docProps parts", () => {
  it("come back with their empty properties as open-close pairs", async () => {
    const input = await createTwoSectionDocx();
    const output = await (await Docxtor.load(input)).toBuffer();
    const [before, after] = [await parts(input), await parts(output)];
    for (const path of ["docProps/core.xml", "docProps/app.xml"]) {
      expect(after.get(path)!.toString()).toBe(before.get(path)!.toString());
    }
    expect(after.get("docProps/core.xml")!.toString()).toContain("<dc:title></dc:title>");
    expect(after.get("docProps/app.xml")!.toString()).toContain("<Company></Company>");
    // word/*.xml in the same package still self-closes
    expect(after.get("word/footer1.xml")!.toString()).toMatch(/<w:p\/>/);
  });

  it("keep the pair form when a property is set", async () => {
    const doc = await Docxtor.load(await createTwoSectionDocx());
    doc.setCoreProperties({ title: "Master Agreement" });
    const core = (await parts(await doc.toBuffer())).get("docProps/core.xml")!.toString();
    expect(core).toContain("<dc:title>Master Agreement</dc:title>");
    expect(core).toContain("<dc:subject></dc:subject>");
    expect(core).not.toContain("/>");
  });
});

describe("a save after one edit", () => {
  it("changes document.xml only, and only inside the edited paragraph", async () => {
    const input = await wordStyleDocx();
    const doc = await Docxtor.load(input);
    const results = doc.applyPatches(
      [{ type: "replace_text", target: "p0", find: "thirty", replace: "forty-five" }],
      { trackRevisions: true, revisionAuthor: "Reviewer" },
    );
    expect(results[0]!.success).toBe(true);

    const output = await doc.toBuffer();
    expect(await changedParts(input, output)).toEqual(["word/document.xml"]);

    const before = (await parts(input)).get("word/document.xml")!.toString();
    const after = (await parts(output)).get("word/document.xml")!.toString();
    // Everything up to the edited run, and everything from the end of the paragraph on, is unchanged
    const head = before.slice(0, before.indexOf("<w:r><w:t>within"));
    const tail = before.slice(before.indexOf("<w:bookmarkEnd"));
    expect(after.startsWith(head)).toBe(true);
    expect(after.endsWith(tail)).toBe(true);

    // The one bookmark has id 0, so the new revisions take 1 and 2, dated without milliseconds
    const revisions = [...after.matchAll(/<w:(del|ins) w:id="(\d+)" w:author="Reviewer" w:date="([^"]+)"/g)];
    expect(revisions.map((m) => `${m[1]}:${m[2]}`)).toEqual(["del:1", "ins:2"]);
    for (const m of revisions) expect(m[3]).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/);
  });

  it("an edit to the body leaves the header, footer and footnotes parts untouched", async () => {
    const input = await createLeaseDocx();
    const doc = await Docxtor.load(input);
    doc.applyPatches(
      [{ type: "replace_text", target: "p2", find: "Landlord", replace: "Lessor" }],
      { trackRevisions: true, revisionDate: "2026-11-10T00:00:00Z" },
    );
    expect(await changedParts(input, await doc.toBuffer())).toEqual(["word/document.xml"]);
  });
});
