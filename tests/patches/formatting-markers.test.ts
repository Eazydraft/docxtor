/**
 * Tests for GFM markdown formatting in projections and text ops.
 */
import { describe, it, expect } from "vitest";
import {
  parseFormattedText,
  stripFormattingMarkers,
  hasFormattingMarkers,
} from "../../src/patches/markdown-parser.js";
import { Docxtor } from "../../src/index.js";
import { readDocx } from "../../src/zip/reader.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { projectToJson, projectToMarkdown } from "../../src/projection/llm.js";
import { serializeXml } from "../../src/xml/serializer.js";

// ============================================================================
// Markdown parser tests
// ============================================================================

describe("parseFormattedText", () => {
  it("returns plain text as single segment", () => {
    const result = parseFormattedText("hello world");
    expect(result).toEqual([{ text: "hello world" }]);
  });

  it("parses bold", () => {
    const result = parseFormattedText("**bold**");
    expect(result).toEqual([{ text: "bold", bold: true }]);
  });

  it("parses italic", () => {
    const result = parseFormattedText("_italic_");
    expect(result).toEqual([{ text: "italic", italic: true }]);
  });

  it("parses strikethrough", () => {
    const result = parseFormattedText("~~strike~~");
    expect(result).toEqual([{ text: "strike", strikethrough: true }]);
  });

  it("parses bold + italic combined", () => {
    const result = parseFormattedText("**_bold italic_**");
    expect(result).toEqual([{ text: "bold italic", bold: true, italic: true }]);
  });

  it("parses mixed formatting in one string", () => {
    const result = parseFormattedText("plain **bold** and _italic_");
    expect(result).toHaveLength(4);
    expect(result[0]).toEqual({ text: "plain " });
    expect(result[1]).toEqual({ text: "bold", bold: true });
    expect(result[2]).toEqual({ text: " and " });
    expect(result[3]).toEqual({ text: "italic", italic: true });
  });

  it("handles empty text", () => {
    const result = parseFormattedText("");
    expect(result).toEqual([]);
  });

  it("handles no markers", () => {
    const result = parseFormattedText("no formatting here");
    expect(result).toEqual([{ text: "no formatting here" }]);
  });
});

describe("stripFormattingMarkers", () => {
  it("strips bold markers", () => {
    expect(stripFormattingMarkers("**bold**")).toBe("bold");
  });

  it("strips italic markers", () => {
    expect(stripFormattingMarkers("_italic_")).toBe("italic");
  });

  it("strips strikethrough markers", () => {
    expect(stripFormattingMarkers("~~strike~~")).toBe("strike");
  });

  it("strips mixed markers", () => {
    expect(stripFormattingMarkers("**bold** and _italic_")).toBe(
      "bold and italic",
    );
  });

  it("strips nested markers", () => {
    expect(stripFormattingMarkers("**_bold italic_**")).toBe("bold italic");
  });

  it("returns plain text unchanged", () => {
    expect(stripFormattingMarkers("plain text")).toBe("plain text");
  });
});

describe("hasFormattingMarkers", () => {
  it("detects bold markers", () => {
    expect(hasFormattingMarkers("**bold**")).toBe(true);
  });

  it("detects strikethrough markers", () => {
    expect(hasFormattingMarkers("~~strike~~")).toBe(true);
  });

  it("returns false for plain text without underscores", () => {
    expect(hasFormattingMarkers("plain text")).toBe(false);
  });

  it("returns false for intra-word underscores like snake_case", () => {
    expect(hasFormattingMarkers("snake_case")).toBe(false);
    expect(hasFormattingMarkers("foo_bar_baz")).toBe(false);
    expect(hasFormattingMarkers("my_var")).toBe(false);
  });

  it("detects underscore italic at word boundary", () => {
    expect(hasFormattingMarkers("_italic_")).toBe(true);
    expect(hasFormattingMarkers("hello _world_")).toBe(true);
    expect(hasFormattingMarkers("_start")).toBe(true);
    expect(hasFormattingMarkers("end_")).toBe(true);
  });

  it("detects underline markers", () => {
    expect(hasFormattingMarkers("<u>text</u>")).toBe(true);
  });
});

// ============================================================================
// Projection tests
// ============================================================================

async function createBoldDocx(): Promise<Buffer> {
  const doc = await Docxtor.create();
  // We need a doc with bold text — build one from scratch with bold run
  const JSZip = (await import("jszip")).default;
  const zip = new JSZip();

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">Hello </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>world</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> and </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>everyone</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file("word/document.xml", documentXml);

  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  }) as Promise<Buffer>;
}

describe("JSON projection with formatting offsets", () => {
  it("includes text, startOffset, endOffset in formatting entries", async () => {
    const buf = await createBoldDocx();
    const doc = await Docxtor.load(buf);

    const proj = doc.project("json", { includeFormatting: true });
    const para = proj.body[0] as any;

    expect(para.formatting).toBeDefined();
    expect(para.formatting).toHaveLength(2); // bold + italic, plain runs filtered

    // Bold run: "world" at offset 6..11
    expect(para.formatting[0]).toEqual({
      text: "world",
      startOffset: 6,
      endOffset: 11,
      bold: true,
    });

    // Italic run: "everyone" at offset 16..24
    expect(para.formatting[1]).toEqual({
      text: "everyone",
      startOffset: 16,
      endOffset: 24,
      italic: true,
    });
  });

  it("omits formatting array when no runs have formatting", async () => {
    const doc = await Docxtor.create();
    const proj = doc.project("json", { includeFormatting: true });
    const para = proj.body[0] as any;
    expect(para.formatting).toBeUndefined();
  });
});

describe("Markdown projection with formatting", () => {
  it("renders bold and italic markers in paragraph text", async () => {
    const buf = await createBoldDocx();
    const doc = await Docxtor.load(buf);

    const md = doc.project("markdown", { includeFormatting: true });

    // Should contain: Hello **world** and _everyone_
    expect(md).toContain("**world**");
    expect(md).toContain("_everyone_");
    expect(md).toContain("Hello **world** and _everyone_");
  });

  it("does not render markers when includeFormatting is false", async () => {
    const buf = await createBoldDocx();
    const doc = await Docxtor.load(buf);

    const md = doc.project("markdown");

    expect(md).not.toContain("**");
    expect(md).toContain("Hello world and everyone");
  });
});

// ============================================================================
// Text ops with markers
// ============================================================================

describe("Text ops with GFM markers", () => {
  it("replace_text strips markers from find", async () => {
    const buf = await createBoldDocx();
    const doc = await Docxtor.load(buf);

    // The bold text "world" appears as "**world**" in the markdown projection
    // but the find should work with markers stripped
    const results = doc.applyPatches([
      {
        type: "replace_text",
        target: "p0",
        find: "**world**",
        replace: "earth",
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const md = doc.project("markdown");
    expect(md).toContain("earth");
    expect(md).not.toContain("world");
  });

  it("replace_text creates formatted runs from replace markers", async () => {
    const doc = await Docxtor.create();

    // append_text creates p1, so target that paragraph for replace
    doc.applyPatches([
      { type: "append_text", text: "Hello bold world" },
    ]);

    const md = doc.project("markdown");
    // Find which paragraph has "world"
    const targetP = md.includes("[p1]") && md.includes("world") ? "p1" : "p0";

    const results = doc.applyPatches([
      { type: "replace_text", target: targetP, find: "world", replace: "**bold** _earth_" },
    ]);

    expect(results[0]!.success).toBe(true);

    const xml = doc.toString();
    expect(xml).toContain("w:b");
    expect(xml).toContain("w:i");

    const md2 = doc.project("markdown", { includeFormatting: true });
    expect(md2).toContain("**bold**");
    expect(md2).toContain("_earth_");
  });

  it("replace_text with bold replacement creates w:b in OOXML", async () => {
    const buf = await createBoldDocx();
    const doc = await Docxtor.load(buf);

    // Replace "Hello " with "**Greetings** "
    const results = doc.applyPatches([
      {
        type: "replace_text",
        target: "p0",
        find: "Hello ",
        replace: "**Greetings** ",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    // Verify the output XML has w:b
    const xml = doc.toString();
    expect(xml).toContain("w:b");

    // Re-project and verify
    const proj = doc.project("json", { includeFormatting: true });
    const para = proj.body[0] as any;
    const boldRun = para.formatting?.find((f: any) => f.text === "Greetings");
    expect(boldRun).toBeDefined();
    expect(boldRun.bold).toBe(true);
  });

  it("insert_text strips markers from anchor", async () => {
    const buf = await createBoldDocx();
    const doc = await Docxtor.load(buf);

    const results = doc.applyPatches([
      {
        type: "insert_text",
        target: "p0",
        anchor: "**world**",
        position: "after",
        text: "!",
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const md = doc.project("markdown");
    expect(md).toContain("world!");
  });

  it("insert_text with formatted text creates formatted runs", async () => {
    const buf = await createBoldDocx();
    const doc = await Docxtor.load(buf);

    const results = doc.applyPatches([
      {
        type: "insert_text",
        target: "p0",
        anchor: "Hello ",
        position: "after",
        text: "**beautiful** ",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const xml = doc.toString();
    // Should have a run with w:b for "beautiful"
    expect(xml).toContain("w:b");

    const proj = doc.project("json", { includeFormatting: true });
    const para = proj.body[0] as any;
    const beautifulRun = para.formatting?.find(
      (f: any) => f.text === "beautiful",
    );
    expect(beautifulRun).toBeDefined();
    expect(beautifulRun.bold).toBe(true);
  });

  it("delete_text strips markers from find", async () => {
    const buf = await createBoldDocx();
    const doc = await Docxtor.load(buf);

    const results = doc.applyPatches([
      { type: "delete_text", target: "p0", find: "**world**" },
    ]);

    expect(results[0]!.success).toBe(true);
    const md = doc.project("markdown");
    expect(md).not.toContain("world");
  });

  it("append_text with formatted text creates formatted runs", async () => {
    const doc = await Docxtor.create();

    doc.applyPatches([
      { type: "append_text", text: "**Important:** This is _critical_" },
    ]);

    const xml = doc.toString();
    expect(xml).toContain("w:b");
    expect(xml).toContain("w:i");

    const proj = doc.project("json", { includeFormatting: true });
    // Find the appended paragraph (last one)
    const paras = proj.body.filter((e: any) => e.text);
    const last = paras[paras.length - 1] as any;
    expect(last.formatting).toBeDefined();

    const boldRun = last.formatting.find((f: any) => f.bold);
    expect(boldRun).toBeDefined();
    expect(boldRun.text).toBe("Important:");

    const italicRun = last.formatting.find((f: any) => f.italic);
    expect(italicRun).toBeDefined();
    expect(italicRun.text).toBe("critical");
  });

  it("append_text with both field placeholders and formatting", async () => {
    const doc = await Docxtor.create();

    doc.applyPatches([
      { type: "append_text", text: "**Page** {{PAGE}} of {{NUMPAGES}}" },
    ]);

    const xml = doc.toString();
    // Should have bold formatting on "Page"
    expect(xml).toContain("w:b");
    // Should have field codes for PAGE and NUMPAGES
    expect(xml).toContain("w:fldChar");
    expect(xml).toContain("PAGE");
    expect(xml).toContain("NUMPAGES");
  });

  it("append_text with _italic_ text and {{PAGE}} field", async () => {
    const doc = await Docxtor.create();

    doc.applyPatches([{ type: "append_text", text: "_Page_ {{PAGE}}" }]);

    const xml = doc.toString();
    expect(xml).toContain("w:i");
    expect(xml).toContain("w:fldChar");
    expect(xml).toContain("PAGE");
  });
});

// ============================================================================
// Formatted runs + field codes integration
// ============================================================================

describe("formatted runs and field codes together", () => {
  it("bold text before a field code produces bold run then field runs", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([{ type: "append_text", text: "**Total:** {{NUMPAGES}}" }]);

    const xml = doc.toString();
    // Bold formatting present
    expect(xml).toContain("w:b");
    // Field code present
    expect(xml).toContain("w:fldChar");
    expect(xml).toContain("NUMPAGES");

    // Structural: bold run for "Total:" should appear before the field begin
    const boldPos = xml.indexOf("w:b");
    const fldBeginPos = xml.indexOf('w:fldCharType="begin"');
    expect(boldPos).toBeLessThan(fldBeginPos);
  });

  it("field code between two formatted segments", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      { type: "append_text", text: "**Page** {{PAGE}} of **{{NUMPAGES}}**" },
    ]);

    const xml = doc.toString();

    // Both field codes present
    expect(xml).toContain("PAGE");
    expect(xml).toContain("NUMPAGES");

    // fldChar pattern for both fields
    const fldCharMatches = xml.match(/w:fldCharType="begin"/g);
    expect(fldCharMatches).not.toBeNull();
    // 2 fields → 2 begin markers
    expect(fldCharMatches!.length).toBeGreaterThanOrEqual(2);
  });

  it("strikethrough + field code", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      { type: "append_text", text: "~~old~~ {{DATE}}" },
    ]);

    const xml = doc.toString();
    expect(xml).toContain("w:strike");
    expect(xml).toContain("DATE");
    expect(xml).toContain("w:fldChar");
  });

  it("multiple formatting styles around multiple fields", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      {
        type: "append_text",
        text: "_File:_ {{FILENAME}} | **Page** {{PAGE}} of {{NUMPAGES}}",
      },
    ]);

    const xml = doc.toString();
    // Italic for "File:"
    expect(xml).toContain("w:i");
    // Bold for "Page"
    expect(xml).toContain("w:b");
    // All three field codes
    expect(xml).toContain("FILENAME");
    expect(xml).toContain("PAGE");
    expect(xml).toContain("NUMPAGES");

    // 3 fields → 3 begin markers
    const begins = xml.match(/w:fldCharType="begin"/g);
    expect(begins!.length).toBe(3);
  });

  it("field code with no surrounding text produces only field runs", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([{ type: "append_text", text: "{{PAGE}}" }]);

    const xml = doc.toString();
    expect(xml).toContain("w:fldChar");
    expect(xml).toContain("PAGE");
    // No bold/italic artifacts
    expect(xml).not.toContain("w:b/>");
    expect(xml).not.toContain("w:i/>");
  });

  it("round-trips formatted text + field codes through write/read", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      { type: "append_text", text: "**Page** {{PAGE}} of {{NUMPAGES}}" },
    ]);

    const buffer = await doc.toBuffer();
    const doc2 = await Docxtor.load(buffer);
    const xml2 = doc2.toString();

    expect(xml2).toContain("w:b");
    expect(xml2).toContain("w:fldChar");
    expect(xml2).toContain("PAGE");
    expect(xml2).toContain("NUMPAGES");
  });

  it("header with formatted text and field codes", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      {
        type: "add_header",
        variant: "default",
        paragraphs: ["**Document Title** — Page {{PAGE}} of {{NUMPAGES}}"],
      },
    ]);

    const ast = doc.getAST();
    const headerXmls = [...ast.xmlParts.headers.values()].map((fxp) =>
      serializeXml(fxp),
    );
    const headerXml = headerXmls.find(
      (x) => x.includes("PAGE") && x.includes("w:b"),
    );
    expect(headerXml).toBeDefined();
    expect(headerXml).toContain("NUMPAGES");
  });

  it("footer with italic label and {{PAGE}} field", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      {
        type: "add_footer",
        variant: "default",
        paragraphs: ["_Page_ {{PAGE}}"],
      },
    ]);

    const ast = doc.getAST();
    const footerXmls = [...ast.xmlParts.footers.values()].map((fxp) =>
      serializeXml(fxp),
    );
    const footerXml = footerXmls.find(
      (x) => x.includes("PAGE") && x.includes("w:i"),
    );
    expect(footerXml).toBeDefined();
  });

  it("unsupported placeholder is treated as plain text with formatting", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      { type: "append_text", text: "**{{UNSUPPORTED}}** value" },
    ]);

    const xml = doc.toString();
    // Should NOT produce field codes
    expect(xml).not.toContain("w:fldChar");
    // The text including braces should be bold
    expect(xml).toContain("w:b");
    expect(xml).toContain("{{UNSUPPORTED}}");
  });

  it("bold wrapping a field code applies bold to field runs, no literal **", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      { type: "append_text", text: "**{{PAGE}}**" },
    ]);

    const xml = doc.toString();
    // Field code should be present
    expect(xml).toContain("w:fldChar");
    expect(xml).toContain("PAGE");
    // Bold should be applied to the field runs
    expect(xml).toContain("w:b");
    // Literal ** should NOT appear in the output
    expect(xml).not.toContain(">**<");
    expect(xml).not.toMatch(/w:t[^>]*>\*\*/);
  });

  it("italic wrapping a field code applies italic to field runs", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      { type: "append_text", text: "_{{PAGE}}_" },
    ]);

    const xml = doc.toString();
    expect(xml).toContain("w:fldChar");
    expect(xml).toContain("w:i");
    // No literal underscores leaked
    expect(xml).not.toMatch(/w:t[^>]*>_</);
  });

  it("formatting carries across multiple field codes", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      { type: "append_text", text: "**{{PAGE}} of {{NUMPAGES}}**" },
    ]);

    const xml = doc.toString();
    // Both fields present
    expect(xml).toContain("PAGE");
    expect(xml).toContain("NUMPAGES");
    // Bold on both field groups and the " of " text
    expect(xml).toContain("w:b");
    // No literal **
    expect(xml).not.toMatch(/w:t[^>]*>\*\*/);

    // The " of " text should also be bold
    const ofMatch = xml.match(/<w:r>(.*?)<\/w:r>/gs);
    const ofRun = ofMatch?.find((r) => r.includes(" of "));
    expect(ofRun).toBeDefined();
    expect(ofRun).toContain("w:b");
  });

  it("markers that span only text, not field, do not leak to field runs", async () => {
    const doc = await Docxtor.create();
    doc.applyPatches([
      { type: "append_text", text: "**Bold** {{PAGE}} plain" },
    ]);

    const xml = doc.toString();
    expect(xml).toContain("w:b");
    expect(xml).toContain("w:fldChar");

    // "plain" should not be bold
    const runs = xml.match(/<w:r>(.*?)<\/w:r>/gs) ?? [];
    const plainRun = runs.find((r) => r.includes("plain"));
    expect(plainRun).toBeDefined();
    expect(plainRun).not.toContain("w:b");
  });
});
