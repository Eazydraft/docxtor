/**
 * A run-level tab is one character of text, `\t`, as a soft line break is
 * `\n`: the view shows it, `find` matches across it, and a patch can keep,
 * add or delete one.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor } from "../../src/index.js";
import { parseXml } from "../../src/xml/parser.js";
import { getChildren, getRunText, getRunPageBreakOffset, createRun, textToRunChildren } from "../../src/xml/helpers.js";
import { serializeXml } from "../../src/xml/serializer.js";
import { createBodyDocx } from "../fixtures/create-docx.js";

const body = (xml: string): string =>
  xml.replace(/^[\s\S]*<w:body>/, "").replace(/<w:sectPr[\s\S]*$/, "").replace(/<\/w:body>[\s\S]*$/, "");

async function load(paragraphXml: string): Promise<Docxtor> {
  return Docxtor.load(await createBodyDocx(paragraphXml));
}

const run = (xml: string) => parseXml(xml)[0]!;
const TRACK = { trackRevisions: true, revisionAuthor: "Reviewer", revisionDate: "2026-11-11T00:00:00Z" } as const;

describe("reading a tab", () => {
  it("getRunText reads w:tab as a tab character", () => {
    expect(getRunText(run(`<w:r><w:t>Name:</w:t><w:tab/><w:t>Alice</w:t></w:r>`))).toBe("Name:\tAlice");
  });

  it("reads the positional tab w:ptab the same way", () => {
    const xml = `<w:r><w:t>Left</w:t><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/><w:t>Right</w:t></w:r>`;
    expect(getRunText(run(xml))).toBe("Left\tRight");
  });

  it("reads w:cr as a line break, like w:br", () => {
    expect(getRunText(run(`<w:r><w:t>one</w:t><w:cr/><w:t>two</w:t><w:br/><w:t>three</w:t></w:r>`))).toBe("one\ntwo\nthree");
  });

  it("counts a tab when it locates a page break in the run", () => {
    expect(getRunPageBreakOffset(run(`<w:r><w:t>A</w:t><w:tab/><w:br w:type="page"/><w:t>B</w:t></w:r>`))).toBe(2);
  });

  it("puts the tab into the paragraph text, the run offsets and both projections", async () => {
    const doc = await load(`<w:p><w:r><w:t>17.2</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>For goods</w:t></w:r></w:p>`);
    const para = doc.getParagraphs()[0]!;
    expect(para.text).toBe("17.2\tFor goods");
    expect(para.runs.map((r) => [r.text, r.startOffset, r.endOffset])).toEqual([
      ["17.2", 0, 4],
      ["\t", 4, 5],
      ["For goods", 5, 14],
    ]);
    expect(doc.project("markdown")).toBe("[p0] 17.2\tFor goods");
    expect(doc.project("json").body[0]).toMatchObject({ text: "17.2\tFor goods" });
  });

  it("does not read the tab stops in w:pPr/w:tabs as text", async () => {
    const doc = await load(
      `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="9360"/></w:tabs></w:pPr>` +
      `<w:r><w:t>No tab here</w:t></w:r></w:p>`,
    );
    expect(doc.getParagraphs()[0]!.text).toBe("No tab here");
  });

  it("shows a tab inside a table cell and inside a tracked deletion", async () => {
    const doc = await load(
      `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>By:</w:t><w:tab/><w:t>Alice</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
      `<w:p><w:r><w:t>Kept</w:t></w:r><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:tab/><w:delText>gone</w:delText></w:r></w:del></w:p>`,
    );
    const md = doc.project("markdown", { includeRevisions: true });
    expect(md).toContain("| [tbl0.r0.c0.p0] By:\tAlice |");
    expect(md).toContain("[p0] Kept{-\tgone-}");
  });
});

describe("writing a tab", () => {
  it("createRun turns a tab character into w:tab and a newline into w:br", () => {
    expect(serializeXml([createRun("a\tb\nc")])).toBe(`<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r>`);
    expect(serializeXml([createRun("\tlead")])).toBe(`<w:r><w:tab/><w:t>lead</w:t></w:r>`);
    expect(serializeXml([createRun("")])).toBe(`<w:r><w:t/></w:r>`);
  });

  it("textToRunChildren reuses an earlier node of the same character, in order", () => {
    const ptab = run(`<w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/>`);
    const children = textToRunChildren("a\tb\tc", [ptab]);
    expect(children[1]).toBe(ptab);
    expect(Object.keys(children[3]!)).toEqual(["w:tab"]);
  });

  it.each([
    ["untracked", undefined],
    ["tracked, then accepted", TRACK],
  ])("replace_text across a tab keeps one w:tab, in place (%s)", async (_name, options) => {
    const doc = await load(`<w:p><w:r><w:t>Name:</w:t><w:tab/><w:t>Alice</w:t></w:r></w:p>`);
    const r = doc.applyPatches([{ type: "replace_text", target: "p0", find: "Name:\tAlice", replace: "Name:\tBob" }], options);
    expect(r[0]).toMatchObject({ success: true });
    if (options) {
      expect(body(doc.toString())).toContain(`<w:r><w:delText>Name:</w:delText><w:tab/><w:delText>Alice</w:delText></w:r>`);
      doc.acceptAllRevisions();
    }
    expect(body(doc.toString())).toBe(`<w:p><w:r><w:t>Name:</w:t><w:tab/><w:t>Bob</w:t></w:r></w:p>`);
    expect(doc.getParagraphs()[0]!.text).toBe("Name:\tBob");
  });

  it("replace_text across a number, a tab and the text in three runs keeps one tab", async () => {
    const xml = `<w:p><w:r><w:t>17.2</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>For goods</w:t></w:r></w:p>`;
    const untracked = await load(xml);
    untracked.applyPatches([{ type: "replace_text", target: "p0", find: "17.2\tFor", replace: "17.3\tFor" }]);
    expect(body(untracked.toString())).toBe(
      `<w:p><w:r><w:t>17.3</w:t><w:tab/><w:t>For</w:t></w:r><w:r><w:t xml:space="preserve"> goods</w:t></w:r></w:p>`,
    );

    const tracked = await load(xml);
    tracked.applyPatches([{ type: "replace_text", target: "p0", find: "17.2\tFor", replace: "17.3\tFor" }], TRACK);
    tracked.acceptAllRevisions();
    expect(body(tracked.toString())).toBe(body(untracked.toString()));
  });

  it.each([
    ["untracked", undefined],
    ["tracked", TRACK],
  ])("a tab added by a replacement is a w:tab, not a raw character in w:t (%s)", async (_name, options) => {
    const doc = await load(`<w:p><w:r><w:t>Name: Alice</w:t></w:r></w:p>`);
    const r = doc.applyPatches([{ type: "replace_text", target: "p0", find: ": ", replace: ":\t" }], options);
    expect(r[0]).toMatchObject({ success: true });
    const xml = body(doc.toString());
    expect(xml).not.toContain("\t");
    expect(xml).toContain(`<w:t>:</w:t><w:tab/>`);
    expect(xml.match(/<w:tab\/>/g)).toHaveLength(1);
    expect(doc.getParagraphs()[0]!.text).toBe("Name:\tAlice");
  });

  it("insert_text, insert_paragraph, append_text and a table cell write tabs as w:tab", async () => {
    const doc = await load(`<w:p><w:r><w:t>By:</w:t></w:r></w:p>`);
    const results = doc.applyPatches([
      { type: "insert_text", target: "p0", anchor: "By:", position: "after", text: "\t____" },
      { type: "insert_paragraph", target: "p0", position: "after", text: "Name:\tAlice" },
      { type: "append_text", text: "Date:\t**today**" },
      { type: "add_table", headers: ["A\tB"], rows: [] },
    ]);
    expect(results.every((r) => r.success)).toBe(true);
    const xml = body(doc.toString());
    expect(xml).not.toContain("\t");
    expect(xml.match(/<w:tab\/>/g)).toHaveLength(4);
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["By:\t____", "Name:\tAlice", "Date:\ttoday"]);
  });

  it("a tab in a list item's text stays a tab; a tab in its indentation is a level", async () => {
    const doc = await load(`<w:p><w:r><w:t>Intro</w:t></w:r></w:p>`);
    const r = doc.applyPatches([{ type: "append_text", text: "- Name:\tAlice\n\t- nested" }]);
    expect(r[0]).toMatchObject({ success: true });
    const items = doc.getParagraphs().slice(1);
    expect(items.map((p) => [p.text, p.listInfo?.level])).toEqual([["Name:\tAlice", 0], ["nested", 1]]);
  });

  it("insert_text anchored on a tab puts the text on the right side of it", async () => {
    const doc = await load(`<w:p><w:r><w:t>15.1</w:t><w:tab/><w:t>Prices</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "insert_text", target: "p0", anchor: "\t", position: "after", text: "The " }]);
    expect(body(doc.toString())).toBe(`<w:p><w:r><w:t>15.1</w:t><w:tab/><w:t>The Prices</w:t></w:r></w:p>`);
  });
});

describe("deleting across a tab", () => {
  it("a tracked delete_text puts the w:tab inside the w:del run, and reject-all restores the input bytes", async () => {
    const input = await createBodyDocx(
      `<w:p><w:r><w:t xml:space="preserve">Signed </w:t></w:r><w:r><w:t>By:</w:t><w:tab/><w:t>Alice</w:t></w:r></w:p>`,
    );
    const before = await (await JSZip.loadAsync(input)).file("word/document.xml")!.async("string");
    const doc = await Docxtor.load(input);
    const r = doc.applyPatches([{ type: "delete_text", target: "p0", find: "By:\tAlice" }], TRACK);
    expect(r[0]).toMatchObject({ success: true });
    expect(body(doc.toString())).toBe(
      `<w:p><w:r><w:t xml:space="preserve">Signed </w:t></w:r>` +
      `<w:del w:id="1" w:author="Reviewer" w:date="2026-11-11T00:00:00Z"><w:r><w:delText>By:</w:delText><w:tab/><w:delText>Alice</w:delText></w:r></w:del></w:p>`,
    );
    expect(doc.getParagraphs()[0]!.text).toBe("Signed ");
    expect(doc.project("markdown", { includeRevisions: true })).toBe("[p0] Signed {-By:\tAlice-}");

    doc.rejectAllRevisions();
    const after = await (await JSZip.loadAsync(await doc.toBuffer())).file("word/document.xml")!.async("string");
    expect(after).toBe(before);
  });

  it("an untracked delete_text of the tab alone removes the element", async () => {
    const doc = await load(`<w:p><w:r><w:t>Name:</w:t><w:tab/><w:t>Alice</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "delete_text", target: "p0", find: "\t" }]);
    expect(doc.getParagraphs()[0]!.text).toBe("Name:Alice");
    expect(body(doc.toString())).not.toContain("w:tab");
  });
});

describe("a positional tab that an edit keeps", () => {
  const PTAB = `<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/>`;

  it("is still a w:ptab after an edit beside it", async () => {
    const doc = await load(`<w:p><w:r><w:t>Draft</w:t>${PTAB}<w:t>Page one</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "one", replace: "two" }]);
    expect(body(doc.toString())).toBe(
      `<w:p><w:r><w:t>Draft</w:t>${PTAB}<w:t xml:space="preserve">Page </w:t></w:r><w:r><w:t>two</w:t></w:r></w:p>`,
    );
  });

  it.each([
    ["untracked", undefined],
    ["tracked, then accepted", TRACK],
  ])("is still a w:ptab when the replacement spans it (%s)", async (_name, options) => {
    const doc = await load(`<w:p><w:r><w:t>Draft</w:t>${PTAB}<w:t>Page one</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Draft\tPage", replace: "Final\tSheet" }], options);
    if (options) doc.acceptAllRevisions();
    const xml = body(doc.toString());
    expect(xml).toContain(`<w:t>Final</w:t>${PTAB}<w:t>Sheet</w:t>`);
    expect(xml).not.toContain("<w:tab/>");
  });

  it("a second tab added next to it is a plain w:tab", async () => {
    const doc = await load(`<w:p><w:r><w:t>Draft</w:t>${PTAB}<w:t>Page</w:t></w:r></w:p>`);
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Draft\tPage", replace: "Draft\t\tPage" }]);
    const children = getChildren(doc.getParagraphs()[0]!.runs[0]!.xmlRef).map((c) => Object.keys(c)[0]);
    expect(children).toEqual(["w:t", "w:ptab", "w:tab", "w:t"]);
  });
});
