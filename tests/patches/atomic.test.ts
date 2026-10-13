/**
 * §3.6 Atomic batch.
 */

import { describe, it, expect } from "vitest";
import { Docxtor } from "../../src/index.js";
import type { Patch } from "../../src/index.js";
import { createLeaseDocx, createMultiParagraphDocx, createSimpleDocx } from "../fixtures/create-docx.js";

const TRACKED = {
  trackRevisions: true,
  revisionAuthor: "USI Legal AI",
  revisionDate: "2026-10-13T00:00:00Z",
};

describe("atomic batches", () => {
  it("a 50-patch batch whose last patch fails leaves the document untouched (49 successes, 1 failure)", async () => {
    const texts = Array.from({ length: 49 }, (_, i) => `Paragraph ${i} alpha beta`);
    const doc = await Docxtor.load(await createMultiParagraphDocx(texts));
    const before = doc.toString();
    const beforeMd = doc.project("markdown");

    const patches: Patch[] = texts.map((_, i) => ({
      type: "replace_text",
      target: `p${i}`,
      find: "alpha",
      replace: "gamma",
    }));
    patches.push({ type: "replace_text", target: "p0", find: "does-not-exist", replace: "x" });

    const results = doc.applyPatches(patches, { ...TRACKED, atomic: true });
    expect(results.length).toBe(50);
    expect(results.filter((r) => r.success).length).toBe(49);
    expect(results[49]!.success).toBe(false);
    expect(results[49]!.error).toContain("not found");

    expect(doc.toString()).toBe(before);
    expect(doc.project("markdown")).toBe(beforeMd);
    expect(doc.getParagraphs()[0]!.text).toBe("Paragraph 0 alpha beta");
  });

  it("applies everything when every patch succeeds", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const results = doc.applyPatches([
      { type: "replace_text", target: "p6", find: "may recapture", replace: "may not recapture" },
      { type: "insert_paragraph", target: "p4", position: "after", text: "Mirror." },
      { type: "delete_paragraph", target: "p5" },
      { type: "append_text", text: "Tail" },
    ], { ...TRACKED, atomic: true });
    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.findParagraphById("p6")!.text).toContain("may not recapture");
    expect(doc.getParagraphs()[5]!.text).toBe("Mirror.");
    expect(doc.getParagraphs().at(-1)!.text).toBe("Tail");
    expect(doc.toString()).toContain('w:author="USI Legal AI"');
  });

  it("later patches may target paragraphs created earlier in the same atomic batch", async () => {
    const doc = await Docxtor.load(await createMultiParagraphDocx(["One", "Two"]));
    const results = doc.applyPatches([
      { type: "insert_paragraph", target: "p0", position: "after", text: "New" },
      { type: "replace_text", target: "p2", find: "New", replace: "Newer" },
    ], { atomic: true });
    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["One", "Newer", "Two"]);
  });

  it("a mid-batch failure rolls back earlier patches and still reports every patch", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const before = doc.toString();
    const results = doc.applyPatches([
      { type: "replace_text", target: "p6", find: "may recapture", replace: "may not recapture" },
      { type: "replace_text", target: "p99", find: "x", replace: "y" },
      { type: "append_text", text: "Tail" },
      { type: "remove_header", variant: "default" },
    ], { ...TRACKED, atomic: true });
    expect(results.map((r) => r.success)).toEqual([true, false, true, true]);
    expect(doc.toString()).toBe(before);
    expect(doc.getHeaders().size).toBe(1);
    expect(doc.getParagraphs().some((p) => p.text === "Tail")).toBe(false);
  });

  it("the exported file after a failed atomic batch equals a no-patch export", async () => {
    const buffer = await createLeaseDocx();
    const pristine = await Docxtor.load(buffer);
    const pristineOut = await pristine.toBuffer();

    const doc = await Docxtor.load(buffer);
    doc.applyPatches([
      { type: "replace_text", target: "p6", find: "may recapture", replace: "may not recapture" },
      { type: "replace_text", target: "p6", find: "nope", replace: "y" },
    ], { atomic: true });
    const out = await doc.toBuffer();

    const JSZip = (await import("jszip")).default;
    const a = await JSZip.loadAsync(pristineOut);
    const b = await JSZip.loadAsync(out);
    for (const path of Object.keys(a.files)) {
      if (a.files[path]!.dir) continue;
      expect(await b.file(path)!.async("string"), path).toBe(await a.file(path)!.async("string"));
    }
  });

  it("the document keeps working after a rolled-back batch", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello World"));
    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi" },
      { type: "replace_text", target: "p0", find: "nope", replace: "y" },
    ], { atomic: true });
    expect(doc.getParagraphs()[0]!.text).toBe("Hello World");
    const [r] = doc.applyPatches([{ type: "replace_text", target: "p0", find: "World", replace: "Earth" }]);
    expect(r!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("Hello Earth");
  });

  it("atomic: false is the default sequential behaviour", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello World"));
    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi" },
      { type: "replace_text", target: "p0", find: "nope", replace: "y" },
    ]);
    expect(doc.getParagraphs()[0]!.text).toBe("Hi World");
  });
});
