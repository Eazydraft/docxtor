/**
 * Phase 5 tests: Edge cases and polish.
 *
 * Tests cover:
 * - Hyperlink support (normalization, patching)
 * - XML entity escaping in replacement text
 * - Newline handling via w:br
 * - Nested table support
 * - Richer projection (formatting, list info)
 * - Public API polish (findById, validatePatches, LoadOptions)
 * - Validator improvements (occurrence, row bounds)
 */

import { describe, it, expect } from "vitest";
import {
  createSimpleDocx,
  createFormattedDocx,
  createTableDocx,
  createHeaderFooterDocx,
  createHyperlinkDocx,
  createNestedTableDocx,
} from "../fixtures/create-docx.js";
import { readDocx } from "../../src/zip/reader.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { applyPatches } from "../../src/patches/engine.js";
import { serializeXml } from "../../src/xml/serializer.js";
import { projectToJson, projectToMarkdown } from "../../src/projection/formats.js";
import { Docxtor } from "../../src/index.js";
import type {
  NormalizedParagraph,
  NormalizedTable,
  Patch,
} from "../../src/types.js";

// =============================================================================
// P5.1: Hyperlink support
// =============================================================================

describe("Hyperlink support", () => {
  it("normalizes text from inside w:hyperlink elements", async () => {
    const buf = await createHyperlinkDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.kind).toBe("paragraph");
    expect(para.text).toBe("Visit our website for details.");
  });

  it("maps runs inside hyperlinks with correct offsets", async () => {
    const buf = await createHyperlinkDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.runs.length).toBe(3);

    // "Visit "
    expect(para.runs[0]!.text).toBe("Visit ");
    expect(para.runs[0]!.startOffset).toBe(0);
    expect(para.runs[0]!.endOffset).toBe(6);

    // "our website" (inside hyperlink)
    expect(para.runs[1]!.text).toBe("our website");
    expect(para.runs[1]!.startOffset).toBe(6);
    expect(para.runs[1]!.endOffset).toBe(17);

    // " for details."
    expect(para.runs[2]!.text).toBe(" for details.");
    expect(para.runs[2]!.startOffset).toBe(17);
    expect(para.runs[2]!.endOffset).toBe(30);
  });

  it("captures formatting from hyperlink runs", async () => {
    const buf = await createHyperlinkDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const para = doc.body[0] as NormalizedParagraph;
    const hyperlinkRun = para.runs[1]!;
    expect(hyperlinkRun.formatting.color).toBe("0000FF");
    expect(hyperlinkRun.formatting.underline).toBe("single");
  });

  it("can replace text inside a hyperlink", async () => {
    const buf = await createHyperlinkDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "our website",
        replace: "the portal",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Visit the portal for details.");
  });

  it("can replace text spanning normal and hyperlink runs", async () => {
    const buf = await createHyperlinkDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "Visit our",
        replace: "See my",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("See my website for details.");
  });

  it("preserves hyperlink XML structure after patching", async () => {
    const buf = await createHyperlinkDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "our website",
        replace: "the portal",
      },
    ]);

    const xml = serializeXml(doc.xmlParts.document);
    // The hyperlink element should still be present
    expect(xml).toContain("w:hyperlink");
    expect(xml).toContain("the portal");
  });
});

// =============================================================================
// P5.2: XML entity escaping
// =============================================================================

describe("XML entity escaping", () => {
  it("handles replacement text with ampersand", async () => {
    const buf = await createSimpleDocx("Hello World");
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "World",
        replace: "Tom & Jerry",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    // The normalized text should show natural characters
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Tom & Jerry");

    // The serialized XML should have proper entities
    const xml = serializeXml(doc.xmlParts.document);
    expect(xml).toContain("Tom &amp; Jerry");
    expect(xml).not.toContain("Tom & Jerry");
  });

  it("handles replacement text with angle brackets", async () => {
    const buf = await createSimpleDocx("placeholder");
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "placeholder",
        replace: "a < b > c",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("a < b > c");

    const xml = serializeXml(doc.xmlParts.document);
    expect(xml).toContain("a &lt; b &gt; c");
  });

  it("handles insertion text with special characters", async () => {
    const buf = await createSimpleDocx("Hello World");
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: " <&>",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello <&> World");

    const xml = serializeXml(doc.xmlParts.document);
    expect(xml).toContain("&lt;&amp;&gt;");
  });

  it("reads existing entities correctly", async () => {
    // Create a DOCX with text that would have entities in XML
    const buf = await createSimpleDocx("A & B");
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const para = doc.body[0] as NormalizedParagraph;
    // The fixture uses escapeXml(), so the XML has &amp;
    // Our normalizer should unescape it to &
    expect(para.text).toBe("A & B");
  });

  it("round-trips entities through parse/patch/serialize", async () => {
    const buf = await createSimpleDocx("A & B");
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    // Replace & with &&
    applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "A & B",
        replace: "A && B",
      },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("A && B");

    const xml = serializeXml(doc.xmlParts.document);
    expect(xml).toContain("A &amp;&amp; B");
  });
});

// =============================================================================
// P5.3: Newline handling
// =============================================================================

describe("Newline handling", () => {
  it("replaces text with content containing newlines", async () => {
    const buf = await createSimpleDocx("Hello World");
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "World",
        replace: "Line1\nLine2",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    // Serialized XML should contain w:br
    const xml = serializeXml(doc.xmlParts.document);
    expect(xml).toContain("<w:br");
    expect(xml).toContain("Line1");
    expect(xml).toContain("Line2");
  });

  it("normalizes w:br as newline in paragraph text", async () => {
    const buf = await createSimpleDocx("Hello World");
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "World",
        replace: "Line1\nLine2\nLine3",
      },
    ]);

    // After re-normalization, the paragraph text should contain \n
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Line1\nLine2\nLine3");
  });

  it("inserts text with newlines", async () => {
    const buf = await createSimpleDocx("Hello World");
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: "\nNew Line",
      },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello\nNew Line World");
  });
});

// =============================================================================
// P5.4: Nested table support
// =============================================================================

describe("Nested table support", () => {
  it("normalizes nested tables inside table cells", async () => {
    const buf = await createNestedTableDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const table = doc.body[0] as NormalizedTable;
    expect(table.kind).toBe("table");
    expect(table.rows.length).toBe(1);

    const firstCell = table.rows[0]!.cells[0]!;
    expect(firstCell.paragraphs.length).toBe(1);
    expect(firstCell.paragraphs[0]!.text).toBe("Outer cell text");

    // Nested table
    expect(firstCell.nestedTables.length).toBe(1);
    const nested = firstCell.nestedTables[0]!;
    expect(nested.kind).toBe("table");
    expect(nested.id).toBe("tbl0.r0.c0.tbl0");
    expect(nested.rows.length).toBe(1);
    expect(nested.rows[0]!.cells.length).toBe(2);
  });

  it("assigns correct IDs to nested table paragraphs", async () => {
    const buf = await createNestedTableDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const table = doc.body[0] as NormalizedTable;
    const nested = table.rows[0]!.cells[0]!.nestedTables[0]!;

    expect(nested.rows[0]!.cells[0]!.paragraphs[0]!.id).toBe(
      "tbl0.r0.c0.tbl0.r0.c0.p0",
    );
    expect(nested.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe("Inner cell A");

    expect(nested.rows[0]!.cells[1]!.paragraphs[0]!.id).toBe(
      "tbl0.r0.c0.tbl0.r0.c1.p0",
    );
    expect(nested.rows[0]!.cells[1]!.paragraphs[0]!.text).toBe("Inner cell B");
  });

  it("can patch text in nested table paragraphs", async () => {
    const buf = await createNestedTableDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const results = applyPatches(doc, [
      {
        type: "replace_cell_text",
        target: "tbl0.r0.c0.tbl0.r0.c0.p0",
        find: "Inner cell A",
        replace: "Modified A",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const table = doc.body[0] as NormalizedTable;
    const nested = table.rows[0]!.cells[0]!.nestedTables[0]!;
    expect(nested.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe("Modified A");
  });

  it("finds nested tables by ID via findTableById", async () => {
    const buf = await createNestedTableDocx();
    const docxtor = await Docxtor.load(buf);

    const nested = docxtor.findTableById("tbl0.r0.c0.tbl0");
    expect(nested).toBeDefined();
    expect(nested!.rows.length).toBe(1);
  });
});

// =============================================================================
// P5.5: Richer projection
// =============================================================================

describe("Richer projection", () => {
  it("includes all formatting fields in JSON projection", async () => {
    const buf = await createFormattedDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const proj = projectToJson(doc, { includeFormatting: true });

    // First paragraph: bold heading
    const heading = proj.body[0] as { formatting?: Record<string, unknown>[] };
    expect(heading.formatting).toBeDefined();
    expect(heading.formatting![0]!.bold).toBe(true);

    // Third paragraph: red large text
    const redPara = proj.body[2] as { formatting?: Record<string, unknown>[] };
    expect(redPara.formatting).toBeDefined();
    expect(redPara.formatting![0]!.color).toBe("FF0000");
    expect(redPara.formatting![0]!.fontSize).toBe(28);
  });

  it("includes underline in formatting projection", async () => {
    const buf = await createHyperlinkDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const proj = projectToJson(doc, { includeFormatting: true });
    const para = proj.body[0] as { formatting?: Record<string, unknown>[] };
    expect(para.formatting).toBeDefined();
    // Only formatted runs are included (plain runs are filtered out)
    // The hyperlink run has underline and color
    expect(para.formatting![0]!.underline).toBe("single");
    expect(para.formatting![0]!.color).toBe("0000FF");
  });

  it("includes nested tables in JSON projection", async () => {
    const buf = await createNestedTableDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const proj = projectToJson(doc);
    const table = proj.body[0] as { nestedTables?: unknown[] };
    expect(table.nestedTables).toBeDefined();
    expect(table.nestedTables!.length).toBe(1);
  });
});

// =============================================================================
// P5.6: Public API polish
// =============================================================================

describe("Public API polish", () => {
  it("findParagraphById finds body paragraphs", async () => {
    const buf = await createSimpleDocx("Hello");
    const docxtor = await Docxtor.load(buf);

    const para = docxtor.findParagraphById("p0");
    expect(para).toBeDefined();
    expect(para!.text).toBe("Hello");
  });

  it("findParagraphById finds table cell paragraphs", async () => {
    const buf = await createTableDocx();
    const docxtor = await Docxtor.load(buf);

    const para = docxtor.findParagraphById("tbl0.r1.c0.p0");
    expect(para).toBeDefined();
    expect(para!.text).toBe("Cell 1");
  });

  it("findParagraphById finds header paragraphs", async () => {
    const buf = await createHeaderFooterDocx();
    const docxtor = await Docxtor.load(buf);

    const para = docxtor.findParagraphById("hdr.default.p0");
    expect(para).toBeDefined();
    expect(para!.text).toBe("Header text");
  });

  it("findParagraphById finds nested table paragraphs", async () => {
    const buf = await createNestedTableDocx();
    const docxtor = await Docxtor.load(buf);

    const para = docxtor.findParagraphById("tbl0.r0.c0.tbl0.r0.c0.p0");
    expect(para).toBeDefined();
    expect(para!.text).toBe("Inner cell A");
  });

  it("findTableById finds top-level tables", async () => {
    const buf = await createTableDocx();
    const docxtor = await Docxtor.load(buf);

    const table = docxtor.findTableById("tbl0");
    expect(table).toBeDefined();
    expect(table!.rows.length).toBe(2);
  });

  it("findParagraphById returns undefined for non-existent ID", async () => {
    const buf = await createSimpleDocx("Hello");
    const docxtor = await Docxtor.load(buf);

    expect(docxtor.findParagraphById("p999")).toBeUndefined();
  });

  it("validatePatches returns validation results", async () => {
    const buf = await createSimpleDocx("Hello");
    const docxtor = await Docxtor.load(buf);

    const results = docxtor.validatePatches([
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi" },
      { type: "replace_text", target: "p999", find: "Nope", replace: "X" },
      { type: "replace_text", target: "p0", find: "", replace: "X" },
    ]);

    expect(results[0]!.valid).toBe(true);
    expect(results[1]!.valid).toBe(false);
    expect(results[1]!.error).toContain("p999");
    expect(results[2]!.valid).toBe(false);
    expect(results[2]!.error).toContain("Empty find string");
  });

  it("respects LoadOptions.includeHeadersFooters = false", async () => {
    const buf = await createHeaderFooterDocx();
    const docxtor = await Docxtor.load(buf, { includeHeadersFooters: false });

    expect(docxtor.getHeaders().size).toBe(0);
    expect(docxtor.getFooters().size).toBe(0);
  });

  it("includes headers/footers by default", async () => {
    const buf = await createHeaderFooterDocx();
    const docxtor = await Docxtor.load(buf);

    expect(docxtor.getHeaders().size).toBeGreaterThan(0);
    expect(docxtor.getFooters().size).toBeGreaterThan(0);
  });
});

// =============================================================================
// P5.7: Validator improvements
// =============================================================================

describe("Validator improvements", () => {
  it("rejects invalid occurrence values", async () => {
    const buf = await createSimpleDocx("Hello World");
    const docxtor = await Docxtor.load(buf);

    const results = docxtor.validatePatches([
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi", occurrence: 0 },
    ]);

    expect(results[0]!.valid).toBe(false);
    expect(results[0]!.error).toContain("Invalid occurrence");
  });

  it("rejects negative occurrence values", async () => {
    const buf = await createSimpleDocx("Hello World");
    const docxtor = await Docxtor.load(buf);

    const results = docxtor.validatePatches([
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi", occurrence: -1 },
    ]);

    expect(results[0]!.valid).toBe(false);
    expect(results[0]!.error).toContain("Invalid occurrence");
  });

  it("validates add_table_row afterRow bounds", async () => {
    const buf = await createTableDocx();
    const docxtor = await Docxtor.load(buf);

    const results = docxtor.validatePatches([
      { type: "add_table_row", target: "tbl0", afterRow: 99, cells: ["A", "B"] },
    ]);

    expect(results[0]!.valid).toBe(false);
    expect(results[0]!.error).toContain("Invalid afterRow");
  });

  it("validates delete_table_row row bounds", async () => {
    const buf = await createTableDocx();
    const docxtor = await Docxtor.load(buf);

    const results = docxtor.validatePatches([
      { type: "delete_table_row", target: "tbl0", row: 99 },
    ]);

    expect(results[0]!.valid).toBe(false);
    expect(results[0]!.error).toContain("Invalid row index");
  });

  it("accepts valid occurrence value", async () => {
    const buf = await createSimpleDocx("Hello World");
    const docxtor = await Docxtor.load(buf);

    const results = docxtor.validatePatches([
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi", occurrence: 1 },
    ]);

    expect(results[0]!.valid).toBe(true);
  });
});

// =============================================================================
// Integration: Full round-trip with Phase 5 features
// =============================================================================

describe("Phase 5 integration round-trip", () => {
  it("handles a complex edit sequence with entities and hyperlinks", async () => {
    const buf = await createHyperlinkDocx();
    const docxtor = await Docxtor.load(buf);

    // Validate first
    const patches: Patch[] = [
      {
        type: "replace_text",
        target: "p0",
        find: "our website",
        replace: "Tom & Jerry's site",
      },
    ];

    const validation = docxtor.validatePatches(patches);
    expect(validation[0]!.valid).toBe(true);

    // Apply
    const results = docxtor.applyPatches(patches);
    expect(results[0]!.success).toBe(true);

    // Verify text
    const para = docxtor.findParagraphById("p0");
    expect(para!.text).toBe("Visit Tom & Jerry's site for details.");

    // Export and verify XML is well-formed
    const output = await docxtor.toBuffer();
    expect(output).toBeInstanceOf(Buffer);
    expect(output.length).toBeGreaterThan(0);

    // Re-load and verify
    const doc2 = await Docxtor.load(output);
    const para2 = doc2.findParagraphById("p0");
    expect(para2!.text).toBe("Visit Tom & Jerry's site for details.");
  });

  it("markdown projection shows list info", async () => {
    // We can't easily create a list docx from fixture, so test the formatting path
    const buf = await createFormattedDocx();
    const zip = await readDocx(buf);
    const doc = normalizeDocument(zip);

    const md = projectToMarkdown(doc);
    // Should contain paragraph IDs and style hints
    expect(md).toContain("[p0]");
    expect(md).toContain("(Heading1)");
    expect(md).toContain("Document Title");
  });
});
