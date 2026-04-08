import { describe, it, expect } from "vitest";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { readDocx } from "../../src/zip/reader.js";
import { writeDocx } from "../../src/zip/writer.js";
import { serializeXml } from "../../src/xml/serializer.js";
import { applyPatches } from "../../src/patches/engine.js";
import {
  createStyledDocx,
  createFormattedDocx,
  createSimpleDocx,
  createSplitRunDocx,
  createHyperlinkDocx,
  createTableDocx,
} from "../fixtures/create-docx.js";
import type {
  NormalizedParagraph,
  NormalizedTable,
  Patch,
} from "../../src/types.js";
import { validatePatch } from "../../src/patches/validator.js";
import { findChild, findChildren, getChildren, getAttr, getAttrs, getTagName } from "../../src/xml/helpers.js";

// =============================================================================
// Helpers
// =============================================================================

async function loadAndPatch(buffer: Buffer, patches: Patch[]) {
  const entries = await readDocx(buffer);
  const doc = normalizeDocument(entries);
  const results = applyPatches(doc, patches);
  return { doc, results, entries };
}

/** Get the raw w:rPr child elements from a run's xmlRef. */
function getRPrChildren(para: NormalizedParagraph, runIndex: number) {
  const run = para.runs[runIndex]!;
  const rPr = findChild(run.xmlRef, "w:rPr");
  return rPr ? getChildren(rPr) : [];
}

/** Check if a specific tag exists in w:rPr of a run. */
function hasRPrTag(para: NormalizedParagraph, runIndex: number, tagName: string): boolean {
  return getRPrChildren(para, runIndex).some((c) => tagName in c);
}

/** Get attribute from a specific tag in w:rPr of a run. */
function getRPrTagAttr(
  para: NormalizedParagraph,
  runIndex: number,
  tagName: string,
  attrName: string,
): string | undefined {
  const rPrChildren = getRPrChildren(para, runIndex);
  const node = rPrChildren.find((c) => tagName in c);
  return node ? getAttr(node, attrName) : undefined;
}

// =============================================================================
// apply_formatting — bold on/off
// =============================================================================

describe("apply_formatting — bold", () => {
  it("should add bold to a plain run", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Hello", formatting: { bold: true } },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    // After re-normalization, the "Hello" run should be bold
    expect(para.runs[0]!.formatting.bold).toBe(true);
    // The rest should NOT be bold
    expect(para.runs[1]!.formatting.bold).toBeUndefined();
  });

  it("should remove bold from a bold run", async () => {
    const buffer = await createFormattedDocx();
    const { doc, results } = await loadAndPatch(buffer, [
      // p1 has "This is bold and italic text." — "bold" run is bold
      { type: "apply_formatting", target: "p1", find: "bold", formatting: { bold: false } },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[1] as NormalizedParagraph;
    // Find the run containing "bold" — it should no longer have bold
    const boldRun = para.runs.find((r) => r.text === "bold");
    expect(boldRun).toBeDefined();
    expect(boldRun!.formatting.bold).toBeUndefined();
  });

  it("should remove bold with null", async () => {
    const buffer = await createFormattedDocx();
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p1", find: "bold", formatting: { bold: null } },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[1] as NormalizedParagraph;
    const boldRun = para.runs.find((r) => r.text === "bold");
    expect(boldRun!.formatting.bold).toBeUndefined();
  });
});

// =============================================================================
// apply_formatting — italic
// =============================================================================

describe("apply_formatting — italic", () => {
  it("should add italic to a plain run", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "World", formatting: { italic: true } },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    const worldRun = para.runs.find((r) => r.text === "World");
    expect(worldRun!.formatting.italic).toBe(true);
  });

  it("should remove italic from an italic run", async () => {
    const buffer = await createFormattedDocx();
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p1", find: "italic", formatting: { italic: false } },
    ]);

    const para = doc.body[1] as NormalizedParagraph;
    const italicRun = para.runs.find((r) => r.text === "italic");
    expect(italicRun!.formatting.italic).toBeUndefined();
  });
});

// =============================================================================
// apply_formatting — multiple properties at once
// =============================================================================

describe("apply_formatting — multiple properties", () => {
  it("should apply bold + italic + color in one patch", async () => {
    const buffer = await createSimpleDocx("Important text here");
    const { doc, results } = await loadAndPatch(buffer, [
      {
        type: "apply_formatting",
        target: "p0",
        find: "Important",
        formatting: { bold: true, italic: true, color: "FF0000" },
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    const run = para.runs.find((r) => r.text === "Important");
    expect(run!.formatting.bold).toBe(true);
    expect(run!.formatting.italic).toBe(true);
    expect(run!.formatting.color).toBe("FF0000");
  });

  it("should set some properties and remove others", async () => {
    const buffer = await createFormattedDocx();
    // p0 has "Document Title" with bold. Let's remove bold and add italic.
    const { doc } = await loadAndPatch(buffer, [
      {
        type: "apply_formatting",
        target: "p0",
        find: "Document Title",
        formatting: { bold: false, italic: true },
      },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.runs[0]!.formatting.bold).toBeUndefined();
    expect(para.runs[0]!.formatting.italic).toBe(true);
  });
});

// =============================================================================
// apply_formatting — underline, strikethrough
// =============================================================================

describe("apply_formatting — underline & strikethrough", () => {
  it("should add underline with style", async () => {
    const buffer = await createSimpleDocx("Underline me");
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Underline", formatting: { underline: "single" } },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    const run = para.runs.find((r) => r.text === "Underline");
    expect(run!.formatting.underline).toBe("single");
  });

  it("should add strikethrough", async () => {
    const buffer = await createSimpleDocx("Strike this out");
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "this", formatting: { strikethrough: true } },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    const run = para.runs.find((r) => r.text === "this");
    expect(run!.formatting.strikethrough).toBe(true);
  });

  it("should remove underline with null", async () => {
    const buffer = await createSimpleDocx("Has underline");
    // First add underline, then remove it
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "underline", formatting: { underline: "single" } },
      { type: "apply_formatting", target: "p0", find: "underline", formatting: { underline: null } },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    const run = para.runs.find((r) => r.text === "underline");
    expect(run!.formatting.underline).toBeUndefined();
  });
});

// =============================================================================
// apply_formatting — fontSize, fontFamily, color, highlight
// =============================================================================

describe("apply_formatting — value properties", () => {
  it("should set fontSize in half-points", async () => {
    const buffer = await createSimpleDocx("Big text");
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Big", formatting: { fontSize: 48 } },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    const run = para.runs.find((r) => r.text === "Big");
    expect(run!.formatting.fontSize).toBe(48);
  });

  it("should set fontFamily", async () => {
    const buffer = await createSimpleDocx("Custom font text");
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Custom", formatting: { fontFamily: "Arial" } },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    const run = para.runs.find((r) => r.text === "Custom");
    expect(run!.formatting.fontFamily).toBe("Arial");
  });

  it("should set highlight color", async () => {
    const buffer = await createSimpleDocx("Highlighted text");
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Highlighted", formatting: { highlight: "yellow" } },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    const run = para.runs.find((r) => r.text === "Highlighted");
    expect(run!.formatting.highlight).toBe("yellow");
  });

  it("should remove fontSize with null", async () => {
    const buffer = await createFormattedDocx();
    // p2 has "Red large text" with sz=28
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p2", find: "Red large text", formatting: { fontSize: null } },
    ]);

    const para = doc.body[2] as NormalizedParagraph;
    expect(para.runs[0]!.formatting.fontSize).toBeUndefined();
  });

  it("should remove color with null", async () => {
    const buffer = await createFormattedDocx();
    // p2 has "Red large text" with color=FF0000
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p2", find: "Red large text", formatting: { color: null } },
    ]);

    const para = doc.body[2] as NormalizedParagraph;
    expect(para.runs[0]!.formatting.color).toBeUndefined();
  });

  it("should remove fontFamily with null", async () => {
    const buffer = await createSimpleDocx("Font test");
    // Add then remove
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Font", formatting: { fontFamily: "Arial" } },
      { type: "apply_formatting", target: "p0", find: "Font", formatting: { fontFamily: null } },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    const run = para.runs.find((r) => r.text === "Font");
    expect(run!.formatting.fontFamily).toBeUndefined();
  });
});

// =============================================================================
// apply_formatting — whole paragraph (no find)
// =============================================================================

describe("apply_formatting — whole paragraph", () => {
  it("should apply formatting to all runs when find is omitted", async () => {
    const buffer = await createFormattedDocx();
    // p1 has "This is bold and italic text." with multiple runs
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p1", formatting: { underline: "single" } },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[1] as NormalizedParagraph;
    // Every run should now have underline
    for (const run of para.runs) {
      expect(run.formatting.underline).toBe("single");
    }
  });

  it("should preserve existing formatting when adding new properties to all runs", async () => {
    const buffer = await createFormattedDocx();
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p1", formatting: { color: "00FF00" } },
    ]);

    const para = doc.body[1] as NormalizedParagraph;
    // The bold run should still be bold
    const boldRun = para.runs.find((r) => r.text === "bold");
    expect(boldRun!.formatting.bold).toBe(true);
    expect(boldRun!.formatting.color).toBe("00FF00");

    // The italic run should still be italic
    const italicRun = para.runs.find((r) => r.text === "italic");
    expect(italicRun!.formatting.italic).toBe(true);
    expect(italicRun!.formatting.color).toBe("00FF00");
  });
});

// =============================================================================
// apply_formatting — occurrence targeting
// =============================================================================

describe("apply_formatting — occurrence", () => {
  it("should target the second occurrence of a word", async () => {
    const buffer = await createSimpleDocx("foo bar foo baz foo");
    const { doc, results } = await loadAndPatch(buffer, [
      {
        type: "apply_formatting",
        target: "p0",
        find: "foo",
        occurrence: 2,
        formatting: { bold: true },
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("foo bar foo baz foo");

    // After splitting, the second "foo" (at offset 8) should be in its own bold run.
    // The first "foo" (offset 0) should not be bold, the third "foo" (offset 16) also not.
    // Find the run that starts at offset 8 (the second "foo")
    const secondFooRun = para.runs.find((r) => r.startOffset === 8);
    expect(secondFooRun).toBeDefined();
    expect(secondFooRun!.text).toBe("foo");
    expect(secondFooRun!.formatting.bold).toBe(true);

    // The runs before and after should NOT be bold
    const firstFooRun = para.runs.find((r) => r.startOffset === 0);
    expect(firstFooRun).toBeDefined();
    expect(firstFooRun!.formatting.bold).toBeUndefined();
  });
});

// =============================================================================
// apply_formatting — run splitting (text spans multiple runs)
// =============================================================================

describe("apply_formatting — run splitting", () => {
  it("should format text that spans two runs", async () => {
    const buffer = await createSplitRunDocx();
    // "valid" spans "va" (run 1) + "lid" (run 2)
    const { doc, results } = await loadAndPatch(buffer, [
      {
        type: "apply_formatting",
        target: "p0",
        find: "valid",
        formatting: { italic: true },
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    // All text should be intact
    expect(para.text).toBe("This agreement is valid and enforceable");

    // The runs containing "va" and "lid" should now have italic
    // After splitting and re-normalization, find the italic runs
    const vaRun = para.runs.find((r) => r.text === "va");
    const lidRun = para.runs.find((r) => r.text === "lid");
    expect(vaRun!.formatting.italic).toBe(true);
    expect(lidRun!.formatting.italic).toBe(true);
  });
});

// =============================================================================
// apply_formatting — hyperlink runs
// =============================================================================

describe("apply_formatting — hyperlink runs", () => {
  it("should format text inside a hyperlink", async () => {
    const buffer = await createHyperlinkDocx();
    // "Visit our website for details." — "our website" is inside w:hyperlink
    const { doc, results } = await loadAndPatch(buffer, [
      {
        type: "apply_formatting",
        target: "p0",
        find: "our website",
        formatting: { bold: true },
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    const hlRun = para.runs.find((r) => r.text === "our website");
    expect(hlRun!.formatting.bold).toBe(true);
    // Should still have original hyperlink formatting
    expect(hlRun!.formatting.color).toBe("0000FF");
    expect(hlRun!.formatting.underline).toBe("single");
  });
});

// =============================================================================
// apply_formatting — table cell paragraphs
// =============================================================================

describe("apply_formatting — table cells", () => {
  it("should format text in a table cell paragraph", async () => {
    const buffer = await createTableDocx();
    const { doc, results } = await loadAndPatch(buffer, [
      {
        type: "apply_formatting",
        target: "tbl0.r1.c0.p0",
        find: "Cell 1",
        formatting: { bold: true, color: "0000FF" },
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const table = doc.body[1] as NormalizedTable;
    const cellPara = table.rows[1]!.cells[0]!.paragraphs[0]!;
    expect(cellPara.runs[0]!.formatting.bold).toBe(true);
    expect(cellPara.runs[0]!.formatting.color).toBe("0000FF");
  });
});

// =============================================================================
// apply_formatting — preserves unmentioned properties
// =============================================================================

describe("apply_formatting — preserves unmentioned properties", () => {
  it("should not affect properties not in the formatting update", async () => {
    const buffer = await createFormattedDocx();
    // p2 has "Red large text" with color=FF0000 and sz=28
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p2", find: "Red large text", formatting: { bold: true } },
    ]);

    const para = doc.body[2] as NormalizedParagraph;
    // Bold should be added
    expect(para.runs[0]!.formatting.bold).toBe(true);
    // Existing color and fontSize should be preserved
    expect(para.runs[0]!.formatting.color).toBe("FF0000");
    expect(para.runs[0]!.formatting.fontSize).toBe(28);
  });
});

// =============================================================================
// apply_formatting — no-op when removing from plain run
// =============================================================================

describe("apply_formatting — edge cases", () => {
  it("should be a no-op when removing formatting from a plain run", async () => {
    const buffer = await createSimpleDocx("Plain text");
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Plain", formatting: { bold: false } },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    // Text should be unchanged
    expect(para.text).toBe("Plain text");
    // Should still have no formatting
    expect(para.runs[0]!.formatting.bold).toBeUndefined();
  });

  it("should error when find text does not exist", async () => {
    const buffer = await createSimpleDocx("Hello");
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "nonexistent", formatting: { bold: true } },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("not found");
  });
});

// =============================================================================
// apply_style — paragraph style
// =============================================================================

describe("apply_style — paragraph style", () => {
  it("should apply a paragraph style to a plain paragraph", async () => {
    const buffer = await createStyledDocx();
    // p2 is "Plain paragraph for testing." — no style
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p2", style: "Heading2" },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[2] as NormalizedParagraph;
    expect(para.style).toBe("Heading2");
  });

  it("should change an existing paragraph style", async () => {
    const buffer = await createStyledDocx();
    // p0 has Heading1 style
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p0", style: "Title" },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.style).toBe("Title");
  });

  it("should reject a character style used as paragraph style", async () => {
    const buffer = await createStyledDocx();
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p2", style: "Strong" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("paragraph");
  });

  it("should reject a nonexistent style", async () => {
    const buffer = await createStyledDocx();
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p2", style: "NoSuchStyle" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("not found");
  });
});

// =============================================================================
// apply_style — character style
// =============================================================================

describe("apply_style — character style", () => {
  it("should apply a character style to a substring", async () => {
    const buffer = await createStyledDocx();
    // p1 has "This is bold and italic text here."
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p1", find: "bold", style: "Strong" },
    ]);

    expect(results[0]!.success).toBe(true);
    // Verify the w:rStyle was set on the run — check XML directly
    const para = doc.body[1] as NormalizedParagraph;
    const boldRun = para.runs.find((r) => r.text === "bold");
    expect(boldRun).toBeDefined();
    const rPr = findChild(boldRun!.xmlRef, "w:rPr");
    expect(rPr).toBeDefined();
    const rStyle = findChild(rPr!, "w:rStyle");
    expect(rStyle).toBeDefined();
    expect(getAttr(rStyle!, "w:val")).toBe("Strong");
  });

  it("should reject a paragraph style used as character style", async () => {
    const buffer = await createStyledDocx();
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p1", find: "bold", style: "Heading1" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("character");
  });

  it("should apply character style to text that requires run splitting", async () => {
    const buffer = await createStyledDocx();
    // p1 has "This is bold and italic text here." — "and" spans the plain run
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p1", find: "and", style: "Emphasis" },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[1] as NormalizedParagraph;
    const andRun = para.runs.find((r) => r.text === "and");
    expect(andRun).toBeDefined();
    const rPr = findChild(andRun!.xmlRef, "w:rPr");
    expect(rPr).toBeDefined();
    const rStyle = findChild(rPr!, "w:rStyle");
    expect(rStyle).toBeDefined();
    expect(getAttr(rStyle!, "w:val")).toBe("Emphasis");
  });
});

// =============================================================================
// Validator — apply_formatting
// =============================================================================

describe("validator — apply_formatting", () => {
  it("should reject empty formatting object", async () => {
    const buffer = await createSimpleDocx("Hello");
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", formatting: {} } as Patch,
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("Empty formatting");
  });

  it("should reject empty find string", async () => {
    const buffer = await createSimpleDocx("Hello");
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "", formatting: { bold: true } },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("Empty find");
  });

  it("should reject invalid occurrence", async () => {
    const buffer = await createSimpleDocx("Hello");
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Hello", occurrence: 0, formatting: { bold: true } },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("Invalid occurrence");
  });

  it("should reject nonexistent target", async () => {
    const buffer = await createSimpleDocx("Hello");
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p999", formatting: { bold: true } },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("not found");
  });
});

// =============================================================================
// Validator — apply_style
// =============================================================================

describe("validator — apply_style", () => {
  it("should reject empty style ID", async () => {
    const buffer = await createStyledDocx();
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p0", style: "" } as Patch,
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("Empty style");
  });

  it("should reject nonexistent style", async () => {
    const buffer = await createStyledDocx();
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p0", style: "FakeStyle" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("not found");
  });

  it("should reject table style used as paragraph style", async () => {
    const buffer = await createStyledDocx();
    const { results } = await loadAndPatch(buffer, [
      { type: "apply_style", target: "p0", style: "TableGrid" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("table");
  });
});

// =============================================================================
// Full round-trip with formatting patches
// =============================================================================

describe("round-trip — formatting patches", () => {
  it("should preserve formatting patch through round-trip", async () => {
    const buffer = await createSimpleDocx("Test formatting round trip");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "apply_formatting", target: "p0", find: "formatting", formatting: { bold: true, color: "FF0000" } },
    ]);

    // Serialize and write back
    const newDocXml = serializeXml(doc.xmlParts.document);
    const output = await writeDocx({ zip: entries.zip, documentXml: newDocXml });

    // Reload and verify
    const reloaded = await readDocx(output);
    const reDoc = normalizeDocument(reloaded);
    const para = reDoc.body[0] as NormalizedParagraph;

    expect(para.text).toBe("Test formatting round trip");
    const fmtRun = para.runs.find((r) => r.text === "formatting");
    expect(fmtRun!.formatting.bold).toBe(true);
    expect(fmtRun!.formatting.color).toBe("FF0000");
  });

  it("should preserve style patch through round-trip", async () => {
    const buffer = await createStyledDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "apply_style", target: "p2", style: "Heading2" },
    ]);

    const newDocXml = serializeXml(doc.xmlParts.document);
    const output = await writeDocx({ zip: entries.zip, documentXml: newDocXml });

    const reloaded = await readDocx(output);
    const reDoc = normalizeDocument(reloaded);
    const para = reDoc.body[2] as NormalizedParagraph;
    expect(para.style).toBe("Heading2");
  });
});

// =============================================================================
// Styles parsing / projection
// =============================================================================

describe("styles parsing", () => {
  it("should parse styles from styles.xml", async () => {
    const buffer = await createStyledDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    expect(doc.styles.length).toBeGreaterThan(0);
    const heading1 = doc.styles.find((s) => s.id === "Heading1");
    expect(heading1).toBeDefined();
    expect(heading1!.name).toBe("heading 1");
    expect(heading1!.type).toBe("paragraph");

    const strong = doc.styles.find((s) => s.id === "Strong");
    expect(strong).toBeDefined();
    expect(strong!.type).toBe("character");
  });

  it("should return empty styles when no styles.xml", async () => {
    const buffer = await createSimpleDocx("No styles");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    expect(doc.styles).toEqual([]);
  });
});

// =============================================================================
// Sequential formatting + text patches
// =============================================================================

describe("mixed formatting and text patches", () => {
  it("should apply formatting then text replacement on same paragraph", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Hello", formatting: { bold: true } },
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
    ]);

    expect(results.every((r) => r.success)).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Earth");
    // Bold should still be on the "Hello" run
    const helloRun = para.runs.find((r) => r.text === "Hello");
    expect(helloRun!.formatting.bold).toBe(true);
  });
});

// =============================================================================
// Bug fix tests — occurrence validation in whole-paragraph mode
// =============================================================================

describe("bug fix — occurrence ignored in whole-paragraph mode", () => {
  it("should accept apply_formatting with invalid occurrence when find is omitted", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // occurrence: 0 is invalid, but should be ignored when find is omitted
    const result = validatePatch(
      { type: "apply_formatting", target: "p0", occurrence: 0, formatting: { bold: true } } as Patch,
      doc,
    );
    expect(result.valid).toBe(true);
  });

  it("should still reject apply_formatting with invalid occurrence when find is provided", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const result = validatePatch(
      { type: "apply_formatting", target: "p0", find: "Hello", occurrence: 0, formatting: { bold: true } } as Patch,
      doc,
    );
    expect(result.valid).toBe(false);
    expect(result.error).toContain("Invalid occurrence");
  });

  it("should accept apply_style with invalid occurrence when find is omitted", async () => {
    const buffer = await createStyledDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const result = validatePatch(
      { type: "apply_style", target: "p2", occurrence: -1, style: "Heading2" } as Patch,
      doc,
    );
    expect(result.valid).toBe(true);
  });

  it("should still reject apply_style with invalid occurrence when find is provided", async () => {
    const buffer = await createStyledDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const result = validatePatch(
      { type: "apply_style", target: "p1", find: "bold", occurrence: 0, style: "Strong" } as Patch,
      doc,
    );
    expect(result.valid).toBe(false);
    expect(result.error).toContain("Invalid occurrence");
  });
});

// =============================================================================
// Bug fix tests — all-undefined formatting rejected
// =============================================================================

describe("bug fix — all-undefined formatting rejected", () => {
  it("should reject { bold: undefined } as empty formatting", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const result = validatePatch(
      { type: "apply_formatting", target: "p0", formatting: { bold: undefined } } as Patch,
      doc,
    );
    expect(result.valid).toBe(false);
    expect(result.error).toContain("Empty formatting");
  });

  it("should reject { bold: undefined, italic: undefined } as empty formatting", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const result = validatePatch(
      { type: "apply_formatting", target: "p0", formatting: { bold: undefined, italic: undefined } } as Patch,
      doc,
    );
    expect(result.valid).toBe(false);
    expect(result.error).toContain("Empty formatting");
  });

  it("should accept { bold: null } (removal is a meaningful action)", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const result = validatePatch(
      { type: "apply_formatting", target: "p0", formatting: { bold: null } } as Patch,
      doc,
    );
    expect(result.valid).toBe(true);
  });

  it("should accept { bold: false } (removal is a meaningful action)", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const result = validatePatch(
      { type: "apply_formatting", target: "p0", formatting: { bold: false } } as Patch,
      doc,
    );
    expect(result.valid).toBe(true);
  });
});

// =============================================================================
// Bug fix tests — fontFamily removal preserves other w:rFonts attributes
// =============================================================================

describe("bug fix — fontFamily removal preserves other rFonts attributes", () => {
  it("should preserve w:eastAsia when removing fontFamily", async () => {
    // Create a DOCX where a run has w:rFonts with w:ascii, w:hAnsi, AND w:eastAsia
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho"/>
        </w:rPr>
        <w:t>Mixed fonts text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    // Build DOCX manually using the same JSZip approach
    const JSZip = (await import("jszip")).default;
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
    zip.file("word/document.xml", documentXml);
    const buffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }) as Buffer;

    const { doc, results } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Mixed fonts text", formatting: { fontFamily: null } },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;

    // fontFamily (w:ascii/w:hAnsi) should be removed
    expect(para.runs[0]!.formatting.fontFamily).toBeUndefined();

    // But w:rFonts element should still exist with w:eastAsia
    const rPr = findChild(para.runs[0]!.xmlRef, "w:rPr");
    expect(rPr).toBeDefined();
    const rFonts = findChild(rPr!, "w:rFonts");
    expect(rFonts).toBeDefined();
    const attrs = getAttrs(rFonts!);
    expect(attrs["w:eastAsia"]).toBe("MS Mincho");
    expect(attrs["w:ascii"]).toBeUndefined();
    expect(attrs["w:hAnsi"]).toBeUndefined();
  });

  it("should remove w:rFonts entirely when no other attributes remain", async () => {
    const buffer = await createSimpleDocx("Font test");
    // Add fontFamily then remove it — should remove the whole element since it only has w:ascii and w:hAnsi
    const { doc } = await loadAndPatch(buffer, [
      { type: "apply_formatting", target: "p0", find: "Font", formatting: { fontFamily: "Arial" } },
      { type: "apply_formatting", target: "p0", find: "Font", formatting: { fontFamily: null } },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.runs[0]!.formatting.fontFamily).toBeUndefined();

    // The w:rFonts element should be completely gone since it had no other attributes
    const rPr = findChild(para.runs[0]!.xmlRef, "w:rPr");
    if (rPr) {
      const rFonts = findChild(rPr, "w:rFonts");
      expect(rFonts).toBeUndefined();
    }
  });
});
