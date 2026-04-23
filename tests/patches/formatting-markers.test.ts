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

    // Replace the empty text with formatted text
    let results = doc.applyPatches([
      {
        type: "replace_text",
        target: "p0",
        find: "",
        replace: "Hello **bold** world",
      },
    ]);

    results = doc.applyPatches([
      { type: "replace_text", target: "p1", find: "world", replace: "**bold** _earth_" },
    ]);

    expect(results[0]!.success).toBe(true);

    const xml = doc.toString();
    expect(xml).toContain("w:b");
    expect(xml).toContain("w:i");

    const md = doc.project("markdown", { includeFormatting: true });
    expect(md).toContain("**bold**");
    expect(md).toContain("_earth_");
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
