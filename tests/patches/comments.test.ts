/**
 * add_comment.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor } from "../../src/index.js";
import { createLeaseDocx, createSimpleDocx, createStyledDocx } from "../fixtures/create-docx.js";

const OPTIONS = { revisionAuthor: "USI Legal AI", revisionDate: "2026-10-13T00:00:00Z" };

async function entry(buffer: Buffer, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(buffer);
  const f = zip.file(path);
  return f ? f.async("string") : undefined;
}

describe("add_comment", () => {
  it("creates comments.xml on demand and anchors the comment to a substring", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const [r] = doc.applyPatches([
      { type: "add_comment", target: "p6", find: "recapture the Premises", text: "Delete per firm rule R-12.", author: "USI Legal AI" },
    ], OPTIONS);
    expect(r!.success).toBe(true);
    expect(doc.findParagraphById("p6")!.text).toBe("Landlord may recapture the Premises upon any request to assign or sublet.");

    const xml = doc.toString();
    expect(xml).toMatch(/<w:commentRangeStart w:id="0"><\/w:commentRangeStart><w:r><w:t[^>]*>recapture the Premises<\/w:t><\/w:r><w:commentRangeEnd w:id="0"><\/w:commentRangeEnd><w:r><w:commentReference w:id="0">/);

    const out = await doc.toBuffer();
    const comments = (await entry(out, "word/comments.xml"))!;
    expect(comments).toContain('<w:comment w:id="0" w:author="USI Legal AI" w:date="2026-10-13T00:00:00Z" w:initials="ULA">');
    expect(comments).toContain("<w:annotationRef>");
    expect(comments).toContain("Delete per firm rule R-12.");
    expect((await entry(out, "word/_rels/document.xml.rels"))!).toContain('Target="comments.xml"');
    expect((await entry(out, "[Content_Types].xml"))!).toContain('PartName="/word/comments.xml"');

    // Reloads cleanly and the comment part is kept
    const reloaded = await Docxtor.load(out);
    expect(reloaded.findParagraphById("p6")!.text).toContain("recapture the Premises");
    expect(await entry(await reloaded.toBuffer(), "word/comments.xml")).toContain("R-12");
  });

  it("comments the whole paragraph when find is omitted; ids increment; author falls back to options", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello World"));
    const results = doc.applyPatches([
      { type: "add_comment", target: "p0", text: "First" },
      { type: "add_comment", target: "p0", find: "World", text: "Second\nwith two paragraphs", initials: "ZZ" },
    ], OPTIONS);
    expect(results.every((r) => r.success)).toBe(true);

    const xml = doc.toString().replace(/>\s+</g, "><");
    expect(xml).toMatch(/<w:p><w:commentRangeStart w:id="0">.*<w:commentRangeEnd w:id="0"><\/w:commentRangeEnd><w:r><w:commentReference w:id="0">/);
    expect(xml).toContain('<w:commentRangeStart w:id="1">');

    const comments = (await entry(await doc.toBuffer(), "word/comments.xml"))!;
    expect(comments).toContain('w:id="0" w:author="USI Legal AI"');
    expect(comments).toContain('w:id="1" w:author="USI Legal AI" w:date="2026-10-13T00:00:00Z" w:initials="ZZ"');
    expect((comments.match(/<w:p>/g) ?? []).length).toBe(3);
  });

  it("continues numbering from an existing comments part after reload", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello World"));
    doc.applyPatches([{ type: "add_comment", target: "p0", text: "One" }], OPTIONS);
    const reloaded = await Docxtor.load(await doc.toBuffer());
    reloaded.applyPatches([{ type: "add_comment", target: "p0", find: "World", text: "Two" }], OPTIONS);
    const comments = (await entry(await reloaded.toBuffer(), "word/comments.xml"))!;
    expect(comments).toContain('w:id="0"');
    expect(comments).toContain('w:id="1"');
  });

  it("uses CommentReference / CommentText styles when the document defines them", async () => {
    const doc = await Docxtor.load(await createStyledDocx());
    // createStyledDocx has no comment styles → no rStyle
    doc.applyPatches([{ type: "add_comment", target: "p2", text: "x" }], OPTIONS);
    expect(doc.toString()).not.toContain("CommentReference");
  });

  it("validates target, text and find", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    const v = doc.validatePatches([
      { type: "add_comment", target: "p9", text: "x" },
      { type: "add_comment", target: "p0", text: "" },
      { type: "add_comment", target: "p0", find: "", text: "x" },
      { type: "add_comment", target: "p0", find: "nope", text: "x" },
    ]);
    expect(v.map((x) => x.valid)).toEqual([false, false, false, true]);
    const [r] = doc.applyPatches([{ type: "add_comment", target: "p0", find: "nope", text: "x" }]);
    expect(r!.success).toBe(false);
    expect(r!.error).toContain("not found");
  });

  it("is untouched by accept/reject of revisions", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello World"));
    doc.applyPatches([
      { type: "add_comment", target: "p0", find: "World", text: "c" },
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi" },
    ], { ...OPTIONS, trackRevisions: true });
    doc.acceptAllRevisions();
    expect(doc.toString()).toContain("w:commentReference");
    expect(doc.getParagraphs()[0]!.text).toBe("Hi World");
  });
});

describe("add_comment failure has no side effects", () => {
  it("does not create comments.xml, a relationship or a content type when the anchor is missing", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello World"));
    const [r] = doc.applyPatches([{ type: "add_comment", target: "p0", find: "nope", text: "x" }]);
    expect(r!.success).toBe(false);
    const out = await doc.toBuffer();
    expect(await entry(out, "word/comments.xml")).toBeUndefined();
    expect(await entry(out, "word/_rels/document.xml.rels") ?? "").not.toContain("comments.xml");
    expect((await entry(out, "[Content_Types].xml"))!).not.toContain("comments.xml");
    expect(doc.toString()).not.toContain("commentRange");
  });
});
