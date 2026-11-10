/**
 * Issue #35: inserting text verbatim. Underscore blanks and lone asterisks
 * are literal even with markers on; `markers: false` / `inlineMarkers: false`
 * turn parsing off entirely; backslash escapes write marker characters.
 */

import { describe, it, expect } from "vitest";
import { Docxtor, serializeXml } from "../../src/index.js";
import {
  parseFormattedText,
  stripFormattingMarkers,
  hasFormattingMarkers,
} from "../../src/patches/markdown-parser.js";
import { createSimpleDocx, createTableDocx, createListDocx } from "../fixtures/create-docx.js";

const BLANK = "Name:  " + "_".repeat(29);

describe("markdown parser (issue #35)", () => {
  it("runs of underscores are literal", () => {
    expect(stripFormattingMarkers(BLANK)).toBe(BLANK);
    expect(stripFormattingMarkers("a(n) __________ limited")).toBe("a(n) __________ limited");
    expect(stripFormattingMarkers("____________ end")).toBe("____________ end");
    expect(stripFormattingMarkers("__init__")).toBe("__init__");
    expect(hasFormattingMarkers(BLANK)).toBe(false);
  });

  it("single underscores at word boundaries are still italic", () => {
    expect(parseFormattedText("a _b_ c")).toEqual([{ text: "a " }, { text: "b", italic: true }, { text: " c" }]);
    expect(stripFormattingMarkers("_italic_")).toBe("italic");
    expect(stripFormattingMarkers("snake_case stays")).toBe("snake_case stays");
    expect(stripFormattingMarkers("**_bold italic_**")).toBe("bold italic");
  });

  it("an opening underscore followed by a space is literal", () => {
    expect(stripFormattingMarkers("_ not italic _")).toBe("_ not italic _");
  });

  it("backslash escapes write the marker characters", () => {
    expect(stripFormattingMarkers("Name:  \\_\\_\\_\\_\\_")).toBe("Name:  _____");
    expect(stripFormattingMarkers("\\*\\*not bold\\*\\*")).toBe("**not bold**");
    expect(stripFormattingMarkers("\\~\\~kept\\~\\~")).toBe("~~kept~~");
    expect(stripFormattingMarkers("\\<u>literal\\</u>")).toBe("<u>literal</u>");
    expect(stripFormattingMarkers("C:\\\\path")).toBe("C:\\path");
    // A backslash before anything else is kept
    expect(stripFormattingMarkers("a\\nb")).toBe("a\\nb");
    expect(parseFormattedText("**\\*starred\\***")).toEqual([{ text: "*starred*", bold: true }]);
  });

  it("markers: false returns the input untouched", () => {
    const raw = "**bold** _i_ ~~s~~ <u>u</u> \\_ \\\\";
    expect(parseFormattedText(raw, false)).toEqual([{ text: raw }]);
    expect(stripFormattingMarkers(raw, false)).toBe(raw);
    expect(parseFormattedText("", false)).toEqual([]);
  });
});

describe("verbatim patch text (issue #35)", () => {
  async function blank(): Promise<Docxtor> {
    const doc = await Docxtor.create();
    doc.applyPatches([{ type: "append_text", text: "anchor" }]);
    return doc;
  }

  it("the issue's table, with markers on", async () => {
    const doc = await blank();
    const target = doc.getParagraphs().at(-1)!.id;
    const inputs = [BLANK, "Price: 5 * 3 * 2 and **bold**", "a ~~x~~ b _c_ d", "Name:  \\_\\_\\_\\_\\_"];
    doc.applyPatches(inputs.map((text) => ({ type: "insert_paragraph" as const, target, position: "after" as const, text })));
    const texts = doc.getParagraphs().slice(-4).map((p) => p.text).reverse();
    expect(texts).toEqual([BLANK, "Price: 5 * 3 * 2 and bold", "a x b c d", "Name:  _____"]);
  });

  it("insert_paragraph with markers: false writes the text verbatim", async () => {
    const doc = await blank();
    const target = doc.getParagraphs().at(-1)!.id;
    const raw = "Price: 5 * 3 * 2 and **bold** _c_ ~~x~~ \\_";
    doc.applyPatches([{ type: "insert_paragraph", target, position: "after", text: raw, markers: false }]);
    expect(doc.getParagraphs().at(-1)!.text).toBe(raw);
    expect(doc.toString()).not.toMatch(/<w:b[ \/>]/);
  });

  it("PatchOptions.inlineMarkers: false applies to every patch in the batch; a patch can opt back in", async () => {
    const doc = await blank();
    const target = doc.getParagraphs().at(-1)!.id;
    doc.applyPatches(
      [
        { type: "insert_paragraph", target, position: "after", text: "**raw**" },
        { type: "insert_paragraph", target, position: "after", text: "**bold**", markers: true },
      ],
      { inlineMarkers: false },
    );
    const [bold, raw] = doc.getParagraphs().slice(-2).map((p) => p.text);
    expect(bold).toBe("bold");
    expect(raw).toBe("**raw**");
  });

  it("replace_text: find and replace verbatim", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Sign: **here** then ____________ end"));
    doc.applyPatches(
      [{ type: "replace_text", target: "p0", find: "**here**", replace: "_there_", markers: false }],
    );
    expect(doc.getParagraphs()[0]!.text).toBe("Sign: _there_ then ____________ end");

    // With markers on, the same replace yields an italic run and the blank stays intact
    const doc2 = await Docxtor.load(await createSimpleDocx("Sign: here then ____________ end"));
    doc2.applyPatches([{ type: "replace_text", target: "p0", find: "here", replace: "_there_ ____________" }]);
    expect(doc2.getParagraphs()[0]!.text).toBe("Sign: there ____________ then ____________ end");
    expect(doc2.toString()).toContain("<w:i/>");
  });

  it("replace_text tracked: verbatim replacement", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Amount: [AMOUNT]"));
    doc.applyPatches(
      [{ type: "replace_text", target: "p0", find: "[AMOUNT]", replace: "**TBD** _____", markers: false }],
      { trackRevisions: true },
    );
    expect(doc.getParagraphs()[0]!.text).toBe("Amount: **TBD** _____");
    doc.acceptAllRevisions();
    expect(doc.getParagraphs()[0]!.text).toBe("Amount: **TBD** _____");
  });

  it("insert_text and delete_text honour markers: false for anchor, text and find", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("a **b** c"));
    doc.applyPatches([
      { type: "insert_text", target: "p0", anchor: "**b**", position: "after", text: " ~~d~~", markers: false },
    ]);
    expect(doc.getParagraphs()[0]!.text).toBe("a **b** ~~d~~ c");
    doc.applyPatches([{ type: "delete_text", target: "p0", find: "**b** ", markers: false }]);
    expect(doc.getParagraphs()[0]!.text).toBe("a ~~d~~ c");
    // Tracked variant
    doc.applyPatches([{ type: "delete_text", target: "p0", find: "~~d~~ ", markers: false }], { trackRevisions: true });
    expect(doc.getParagraphs()[0]!.text).toBe("a c");
  });

  it("append_text, add_list, add_list_item, replace_cell_text, add_header, add_comment", async () => {
    const doc = await Docxtor.load(await createTableDocx());
    const n = doc.getParagraphs().length;
    const results = doc.applyPatches(
      [
        { type: "append_text", text: "**verbatim** one\n_two_" },
        { type: "add_list", listType: "bullet", items: ["**item** _x_", { text: "~~y~~", level: 1 }] },
        { type: "replace_cell_text", target: "tbl0.r1.c0.p0", find: "Cell 1", replace: "**Term**" },
        { type: "add_header", variant: "default", paragraphs: ["**CONF** {{PAGE}}"] },
        { type: "add_comment", target: "p0", text: "**note**" },
      ],
      { inlineMarkers: false },
    );
    expect(results.every((r) => r.success)).toBe(true);
    const texts = doc.getParagraphs().slice(n).map((p) => p.text);
    expect(texts).toEqual(["**verbatim** one", "_two_", "**item** _x_", "~~y~~"]);
    expect(doc.findParagraphById("tbl0.r1.c0.p0")!.text).toBe("**Term**");
    const header = [...doc.getAST().headers.values()].find((h) => h.paragraphs.some((p) => p.text.includes("CONF")));
    expect(header!.paragraphs[0]!.text).toBe("**CONF** ");
    const headerXml = serializeXml(header!.xmlRef);
    expect(headerXml).toContain("PAGE"); // the field still works
    expect(headerXml).not.toMatch(/<w:b[ \/>]/);
    const comments = serializeXml(doc.getAST().xmlParts.comments!);
    expect(comments).toContain("**note**");
    expect(comments).not.toMatch(/<w:b[ \/>]/);

    const listDoc = await Docxtor.load(await createListDocx());
    const item = listDoc.getParagraphs().find((p) => p.listInfo)!;
    listDoc.applyPatches([{ type: "add_list_item", target: item.id, position: "after", text: "**raw**", markers: false }]);
    expect(listDoc.getParagraphs().some((p) => p.text === "**raw**")).toBe(true);
  });

  it("a literal U+FFFC next to a field placeholder is kept, with markers on and off", async () => {
    for (const markers of [true, false]) {
      const doc = await Docxtor.create();
      doc.applyPatches([{ type: "append_text", text: "\uFFFC {{PAGE}} \uFFFC", markers }]);
      const para = doc.getParagraphs().at(-1)!;
      expect(para.text).toBe("\uFFFC  \uFFFC");
      const xml = doc.toString();
      expect((xml.match(/w:instrText[^>]*> PAGE /g) ?? []).length).toBe(1);
      expect(xml).not.toContain("undefined");
    }
  });

  it("validatePatches accepts the markers field", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    const v = doc.validatePatches([{ type: "replace_text", target: "p0", find: "Hello", replace: "x", markers: false }]);
    expect(v[0]!.valid).toBe(true);
  });
});
