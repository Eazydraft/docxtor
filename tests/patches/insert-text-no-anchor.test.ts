/**
 * insert_text without an anchor: text at the start or end of a paragraph,
 * which is how an empty paragraph or table cell is filled. Also the errors
 * of a text patch whose text is in another paragraph than its target.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor, type Patch, type PatchOptions } from "../../src/index.js";
import { createBodyDocx } from "../fixtures/create-docx.js";

const TRACKED: PatchOptions = { trackRevisions: true, revisionAuthor: "Reviewer", revisionDate: "2026-11-12T00:00:00Z" };

const cell = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const para = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/**
 * - p0: "Signed by the parties." in two runs, the second bold
 * - tbl0 r0: a cell with two paragraphs ("Heading", "Body text"), and an empty cell
 * - tbl0 r1: "Date:", and an empty cell whose paragraph mark is bold, 14pt and
 *   carries a tracked paragraph-mark insertion
 * - p1: an empty paragraph with no properties
 */
const BODY =
  `<w:p><w:r><w:t xml:space="preserve">Signed by </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>the parties.</w:t></w:r></w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
  `<w:tr>${cell(para("Heading") + para("Body text"))}${cell("<w:p/>")}</w:tr>` +
  `<w:tr>${cell(para("Date:"))}${cell(
    `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:ins w:id="90" w:author="Other" w:date="2026-01-01T00:00:00Z"/><w:b/><w:sz w:val="28"/></w:rPr></w:pPr></w:p>`,
  )}</w:tr>` +
  `</w:tbl>` +
  `<w:p/>`;

const load = async (): Promise<Docxtor> => Docxtor.load(await createBodyDocx(BODY));

function apply(doc: Docxtor, patch: unknown, options?: PatchOptions) {
  return doc.applyPatches([patch as Patch], options)[0]!;
}

async function documentXml(doc: Docxtor): Promise<string> {
  const zip = await JSZip.loadAsync(await doc.toBuffer());
  return zip.file("word/document.xml")!.async("string");
}

const text = (doc: Docxtor, id: string): string => doc.findParagraphById(id)!.text;

describe("view of a table cell with several paragraphs", () => {
  it("prints an id before each paragraph of the cell", async () => {
    const md = (await load()).project("markdown");
    expect(md).toContain("| [tbl0.r0.c0.p0] Heading [tbl0.r0.c0.p1] Body text | [tbl0.r0.c1.p0]  |");
  });

  it("leaves a one-paragraph cell and an empty cell as they were", async () => {
    const md = (await load()).project("markdown");
    expect(md).toContain("| [tbl0.r1.c0.p0] Date: | [tbl0.r1.c1.p0]  |");
  });

  it("prints the synthesized id for a cell with no paragraph", async () => {
    const doc = await Docxtor.load(await createBodyDocx(
      `<w:tbl><w:tr>${cell(para("A"))}<w:tc><w:tcPr/></w:tc></w:tr></w:tbl><w:p/>`,
    ));
    expect(doc.project("markdown")).toContain("| [tbl0.r0.c0.p0] A | [tbl0.r0.c1.p0]  |");
  });

  it("shows revisions and formatting per paragraph", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "tbl0.r0.c0.p1", find: "Body", replace: "Main" }, TRACKED).success).toBe(true);
    expect(doc.project("markdown", { includeRevisions: true })).toContain(
      "[tbl0.r0.c0.p0] Heading [tbl0.r0.c0.p1] {-Body-}{+Main+} text",
    );
  });
});

describe("insert_text without an anchor", () => {
  it("inserts at the start of a paragraph that has text", async () => {
    const doc = await load();
    const r = apply(doc, { type: "insert_text", target: "p0", position: "after", text: "DRAFT. " });
    expect(r.error).toBeUndefined();
    expect(text(doc, "p0")).toBe("DRAFT. Signed by the parties.");
  });

  it('inserts at the start with position "start", "before" or none', async () => {
    for (const position of ["start", "before", undefined]) {
      const doc = await load();
      const r = apply(doc, { type: "insert_text", target: "p0", position, text: "DRAFT. " });
      expect(r.error, String(position)).toBeUndefined();
      expect(text(doc, "p0")).toBe("DRAFT. Signed by the parties.");
    }
  });

  it('appends after the existing text with position "end", in the last run\'s formatting', async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_text", target: "p0", position: "end", text: " Witnessed." }).success).toBe(true);
    expect(text(doc, "p0")).toBe("Signed by the parties. Witnessed.");
    expect(await documentXml(doc)).toContain(`<w:r><w:rPr><w:b/></w:rPr><w:t>the parties. Witnessed.</w:t></w:r>`);
  });

  it("tracks the insertion at either end", async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_text", target: "p0", position: "start", text: "DRAFT. " }, TRACKED).success).toBe(true);
    expect(apply(doc, { type: "insert_text", target: "p0", position: "end", text: " Witnessed." }, TRACKED).success).toBe(true);
    expect(doc.project("markdown", { includeRevisions: true })).toContain("[p0] {+DRAFT. +}Signed by the parties.{+ Witnessed.+}");
    doc.rejectAllRevisions();
    expect(text(doc, "p0")).toBe("Signed by the parties.");
  });

  it('treats anchor "" as no anchor', async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_text", target: "tbl0.r0.c1.p0", position: "after", anchor: "", text: "June 22, 2026" }).success).toBe(true);
    expect(text(doc, "tbl0.r0.c1.p0")).toBe("June 22, 2026");
  });

  it("fills an empty table-cell paragraph, tracked, with the paragraph mark's run properties", async () => {
    const doc = await load();
    const r = apply(doc, { type: "insert_text", target: "tbl0.r1.c1.p0", position: "after", text: "June 22, 2026" }, TRACKED);
    expect(r.error).toBeUndefined();
    expect(text(doc, "tbl0.r1.c1.p0")).toBe("June 22, 2026");
    const xml = await documentXml(doc);
    // The mark's formatting is copied; its tracked-insertion marker is not.
    expect(xml).toContain(
      `</w:pPr><w:ins w:id="91" w:author="Reviewer" w:date="2026-11-12T00:00:00Z"><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>June 22, 2026</w:t></w:r></w:ins></w:p>`,
    );
    expect(doc.project("markdown", { includeRevisions: true })).toContain("[tbl0.r1.c1.p0] {+June 22, 2026+}");
  });

  it("fills an empty paragraph untracked as a plain run", async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_text", target: "tbl0.r1.c1.p0", position: "start", text: "June 22, 2026" }).success).toBe(true);
    const xml = await documentXml(doc);
    expect(xml).toContain(`</w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>June 22, 2026</w:t></w:r></w:p>`);
    expect(xml).not.toContain(`w:author="Reviewer"`);
  });

  it("fills a paragraph that has no properties at all", async () => {
    for (const position of ["start", "end"]) {
      const doc = await load();
      expect(apply(doc, { type: "insert_text", target: "p1", position, text: "Filled" }).success).toBe(true);
      expect(await documentXml(doc)).toContain("<w:p><w:r><w:t>Filled</w:t></w:r></w:p>");
    }
  });

  it("applies inline formatting markers in the filled text", async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_text", target: "p1", position: "start", text: "a **b**" }).success).toBe(true);
    expect(text(doc, "p1")).toBe("a b");
    expect(await documentXml(doc)).toContain("<w:r><w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r>");
  });

  it("reaches a fixed point: tracked fill then accept-all equals the untracked fill, byte for byte", async () => {
    // Empty paragraphs only: in a paragraph with text, an untracked insertion joins the run next to it.
    for (const target of ["tbl0.r1.c1.p0", "tbl0.r0.c1.p0", "p1"]) {
      for (const position of ["start", "end"]) {
        const patch = { type: "insert_text", target, position, text: "June 22, 2026" };
        const tracked = await load();
        expect(apply(tracked, patch, TRACKED).success).toBe(true);
        tracked.acceptAllRevisions({ authors: ["Reviewer"] });
        const plain = await load();
        expect(apply(plain, patch).success).toBe(true);
        expect(await documentXml(tracked), `${target} ${position}`).toBe(await documentXml(plain));
      }
    }
  });

  it("fills an empty paragraph inside a block content control, and refuses a locked one", async () => {
    const control = (lock: string): string =>
      `<w:sdt><w:sdtPr><w:id w:val="7"/>${lock}</w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt><w:p/>`;
    const open = await Docxtor.load(await createBodyDocx(control("")));
    expect(apply(open, { type: "insert_text", target: "p0", position: "start", text: "Filled" }).error).toBeUndefined();
    expect(text(open, "p0")).toBe("Filled");
    const locked = await Docxtor.load(await createBodyDocx(control(`<w:lock w:val="contentLocked"/>`)));
    const r = apply(locked, { type: "insert_text", target: "p0", position: "start", text: "Filled" });
    expect(r.success).toBe(false);
    expect(r.error).toMatch(/locked/);
  });
});

describe("insert_text validation", () => {
  it("never surfaces a TypeError for a malformed patch", async () => {
    const doc = await load();
    for (const patch of [
      { type: "insert_text", target: "p0", position: "after" },
      { type: "insert_text", target: "p0", position: "after", anchor: 5, text: "x" },
      { type: "insert_text", target: "p0", position: "middle", text: "x" },
      { type: "replace_text", target: "p0", replace: "x" },
      { type: "delete_text", target: "p0" },
    ]) {
      const r = apply(doc, patch);
      expect(r.success).toBe(false);
      expect(r.error).not.toMatch(/Cannot read|undefined|is not a function/);
    }
  });

  it('refuses "start" and "end" with an anchor, saying what to do', async () => {
    const doc = await load();
    const r = apply(doc, { type: "insert_text", target: "p0", position: "end", anchor: "Signed", text: "x" });
    expect(r.success).toBe(false);
    expect(r.error).toBe(
      'position "end" takes no anchor: drop the anchor to insert at the end of the paragraph, or use "before" or "after" with it',
    );
  });

  it("still inserts next to an anchor, and still reports a missing anchor", async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_text", target: "p0", position: "after", anchor: "Signed", text: " and sealed" }).success).toBe(true);
    expect(text(doc, "p0")).toBe("Signed and sealed by the parties.");
    const r = apply(doc, { type: "insert_text", target: "p0", position: "after", anchor: "Nowhere", text: "x" });
    expect(r.error).toBe('Anchor text "Nowhere" not found in paragraph "p0" (occurrence 1)');
  });
});

describe("text that is in another paragraph than the target", () => {
  it("replace_text names the paragraph that holds the text", async () => {
    const doc = await load();
    const r = apply(doc, { type: "replace_text", target: "tbl0.r0.c0.p0", find: "Body", replace: "Main" });
    expect(r.error).toBe(
      'Text "Body" not found in paragraph "tbl0.r0.c0.p0" (occurrence 1); it is in tbl0.r0.c0.p1 (use that as the target)',
    );
  });

  it("add_comment and insert_text do too", async () => {
    const doc = await load();
    const c = apply(doc, { type: "add_comment", target: "tbl0.r0.c0.p0", find: "Body", text: "Check" });
    expect(c.error).toContain("; it is in tbl0.r0.c0.p1 (use that as the target)");
    const i = apply(doc, { type: "insert_text", target: "p1", position: "after", anchor: "Date:", text: " today" });
    expect(i.error).toContain("; it is in tbl0.r1.c0.p0 (use that as the target)");
  });

  it("lists every paragraph when several hold the text", async () => {
    const doc = await Docxtor.load(await createBodyDocx(para("alpha") + para("beta one") + para("gamma") + para("beta two")));
    const r = apply(doc, { type: "replace_text", target: "p0", find: "beta", replace: "delta" });
    expect(r.error).toContain("; it is in p1, p3 (use one of those as the target)");
  });

  it("leaves the error unchanged when the text is nowhere", async () => {
    const doc = await load();
    const r = apply(doc, { type: "replace_text", target: "p0", find: "Nowhere", replace: "x" });
    expect(r.error).toBe('Text "Nowhere" not found in paragraph "p0" (occurrence 1)');
  });
});
