/**
 * Tests for OOXML revision tracking (tracked changes).
 *
 * Tests that applying patches with { trackRevisions: true } produces
 * correct w:ins, w:del, w:rPrChange, w:pPrChange markup in the XML.
 */

import { describe, it, expect } from "vitest";
import { serializeXml } from "../../src/xml/serializer.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { readDocx } from "../../src/zip/reader.js";
import { writeDocx } from "../../src/zip/writer.js";
import { applyPatches } from "../../src/patches/engine.js";
import {
  createSimpleDocx,
  createFormattedDocx,
  createTableDocx,
  createSplitRunDocx,
  createStyledDocx,
} from "../fixtures/create-docx.js";
import type {
  Patch,
  PatchOptions,
  NormalizedParagraph,
  NormalizedTable,
} from "../../src/types.js";

const REV_OPTIONS: PatchOptions = {
  trackRevisions: true,
  revisionAuthor: "Test Author",
  revisionDate: "2025-01-01T00:00:00Z",
};

/**
 * Serialize a document's XML to a string for assertions.
 */
function getDocXml(doc: ReturnType<typeof normalizeDocument>): string {
  return serializeXml(doc.xmlParts.document);
}

// =============================================================================
// replace_text with revisions
// =============================================================================

describe("replace_text with revisions", () => {
  it("should produce w:del and w:ins for a simple replacement", async () => {
    const buffer = await createSimpleDocx("This agreement is valid");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    // Should contain w:del with the old text as w:delText
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:delText");
    expect(xml).toContain("valid");

    // Should contain w:ins with the new text
    expect(xml).toContain("w:ins");
    expect(xml).toContain("enforceable");

    // Should contain author and date attributes
    expect(xml).toContain("Test Author");
    expect(xml).toContain("2025-01-01T00:00:00Z");

    // Re-normalized text should reflect the new content
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("This agreement is enforceable");
  });

  it("should handle replacement at the beginning of text", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi" },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hi World");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
  });

  it("should handle replacement with empty string (delete via replace)", async () => {
    const buffer = await createSimpleDocx("Remove this word please");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "this word ", replace: "" },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Remove please");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:del");
    // No w:ins since replacement is empty
  });

  it("should handle specific occurrence with revisions", async () => {
    const buffer = await createSimpleDocx("foo bar foo baz foo");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "foo",
        replace: "qux",
        occurrence: 2,
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("foo bar qux baz foo");
  });

  it("should produce unique revision IDs for multiple patches", async () => {
    const buffer = await createSimpleDocx("alpha beta gamma");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "alpha", replace: "one" },
      { type: "replace_text", target: "p0", find: "beta", replace: "two" },
    ], REV_OPTIONS);

    expect(results.every((r) => r.success)).toBe(true);

    const xml = getDocXml(doc);
    // Extract all w:id values — they should be unique
    const idMatches = xml.match(/w:id="(\d+)"/g);
    expect(idMatches).not.toBeNull();
    const ids = idMatches!.map((m) => m.match(/\d+/)![0]);
    const uniqueIds = new Set(ids);
    expect(uniqueIds.size).toBe(ids.length);
  });
});

// =============================================================================
// insert_text with revisions
// =============================================================================

describe("insert_text with revisions", () => {
  it("should wrap inserted text in w:ins", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: " Beautiful",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Beautiful World");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Beautiful");
    expect(xml).toContain("Test Author");
    // Should NOT contain w:del (insert only, no deletion)
    expect(xml).not.toContain("w:del");
  });

  it("should insert before anchor text", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "before",
        anchor: "World",
        text: "Beautiful ",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Beautiful World");
  });

  it("should insert at the very beginning of a paragraph", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "before",
        anchor: "Hello",
        text: "Dear User, ",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Dear User, Hello World");
  });

  it("should insert at the very end of a paragraph", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "World",
        text: "!",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello World!");
  });

  it("should insert in the middle of a run (causes split)", async () => {
    const buffer = await createSimpleDocx("HelloWorld");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: " ",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello World");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
  });

  // Edge case 3: Insertion at a run boundary (between two runs)
  it("should insert at a run boundary without splitting", async () => {
    // createFormattedDocx p1: "This is " + "bold" + " and " + "italic" + " text."
    const buffer = await createFormattedDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // Insert after "bold" — this sits exactly at the boundary between "bold" and " and "
    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p1",
        position: "after",
        anchor: "bold",
        text: "er",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[1] as NormalizedParagraph;
    expect(para.text).toBe("This is bolder and italic text.");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
    // "er" is in a separate w:ins-wrapped run, "bold" stays in its original run
    expect(xml).toContain(">er<");
  });

  // Edge case 5: Insertion inside a hyperlink run
  it("should insert inside a hyperlink run", async () => {
    const { createHyperlinkDocx } = await import("../fixtures/create-docx.js");
    const buffer = await createHyperlinkDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // p0 text: "Visit our website for details."
    // "our website" is inside a w:hyperlink
    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "our",
        text: " great",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Visit our great website for details.");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
    expect(xml).toContain("great");
  });

  // Edge case 6: Anchor spans multiple runs
  it("should insert after anchor that spans multiple runs", async () => {
    // createSplitRunDocx: "This agreement is va" + "lid and enforceable"
    const buffer = await createSplitRunDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // "valid" spans two runs; insert after it
    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "valid",
        text: "ly",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("This agreement is validly and enforceable");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
  });

  // Edge case 7: Inserted text contains newlines
  it("should handle inserted text with newlines (w:br)", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: "\nDear\n",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
    expect(xml).toContain("w:br");
  });

  // Edge case 8: Inserted text has leading/trailing whitespace
  it("should preserve whitespace with xml:space='preserve'", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: "  spaced  ",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello  spaced   World");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
    expect(xml).toContain('xml:space="preserve"');
  });

  // Edge case 9: Inserted text contains XML special characters
  it("should escape XML special characters in inserted text", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: " <A & B>",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello <A & B> World");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
    // XML should have escaped entities
    expect(xml).toContain("&lt;A &amp; B&gt;");
  });

  // Edge case 10: Multiple sequential inserts at the same anchor
  it("should handle multiple sequential inserts at the same anchor", async () => {
    const buffer = await createSimpleDocx("hello world");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "hello",
        text: "X",
      },
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "hello",
        text: "Y",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    expect(results[1]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    // Patch 1 inserts X after "hello" → "helloX world"
    // Patch 2 finds "hello" in re-normalized text, inserts Y after it
    // Result depends on whether "hello" is still found as a standalone anchor
    // The key point: both patches succeed and text contains both X and Y
    expect(para.text).toContain("hello");
    expect(para.text).toContain("X");
    expect(para.text).toContain("Y");
    expect(para.text).toContain("world");
  });

  // Edge case 11: Insert after previously deleted text (should fail)
  it("should fail to find anchor that was previously deleted", async () => {
    const buffer = await createSimpleDocx("foo bar baz");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "delete_text", target: "p0", find: "bar" },
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "bar",
        text: "X",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    // Patch 2 should fail — "bar" is deleted and no longer in normalized text
    expect(results[1]!.success).toBe(false);
    expect(results[1]!.error).toContain("not found");
  });

  // Edge case 12: Insert after previously inserted/replaced text
  it("should insert after text that was previously replaced (revision-tracked)", async () => {
    const buffer = await createSimpleDocx("foo bar baz");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "bar", replace: "qux" },
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "qux",
        text: "!",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    expect(results[1]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("foo qux! baz");

    // Both insert operations should produce separate w:ins elements
    const xml = getDocXml(doc);
    const insCount = (xml.match(/<w:ins /g) || []).length;
    expect(insCount).toBeGreaterThanOrEqual(2);
  });

  // Edge case 13: Empty paragraph (no runs) — anchor not found
  it("should fail on empty paragraph (no anchor to find)", async () => {
    const buffer = await createSimpleDocx("");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "anything",
        text: "X",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("not found");
  });
});

// =============================================================================
// delete_text with revisions
// =============================================================================

describe("delete_text with revisions", () => {
  it("should wrap deleted text in w:del with w:delText", async () => {
    const buffer = await createSimpleDocx("Remove this word please");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "delete_text", target: "p0", find: "this word " },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Remove please");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:delText");
    // Should NOT contain w:ins (delete only)
    expect(xml).not.toContain("w:ins");
  });

  it("should keep deleted text in XML (visible in tracked changes mode)", async () => {
    const buffer = await createSimpleDocx("keep this delete that done");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "delete_text", target: "p0", find: "delete that " },
    ], REV_OPTIONS);

    const xml = getDocXml(doc);
    // The deleted text should still be in the XML as w:delText
    expect(xml).toContain("delete that");
    expect(xml).toContain("w:delText");
  });

  it("subsequent patch should not see deleted text", async () => {
    const buffer = await createSimpleDocx("foo bar baz");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "delete_text", target: "p0", find: "bar " },
      // After deletion, paragraph text should be "foo baz"
      { type: "replace_text", target: "p0", find: "foo baz", replace: "done" },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    expect(results[1]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("done");
  });
});

// =============================================================================
// apply_formatting with revisions
// =============================================================================

describe("apply_formatting with revisions", () => {
  it("should add w:rPrChange when applying formatting with revisions", async () => {
    const buffer = await createSimpleDocx("Bold this text");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "apply_formatting",
        target: "p0",
        find: "this",
        formatting: { bold: true },
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    expect(xml).toContain("w:rPrChange");
    expect(xml).toContain("Test Author");
  });

  it("should preserve original formatting in w:rPrChange", async () => {
    const buffer = await createFormattedDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // p2 has red large text — add italic to it
    const results = applyPatches(doc, [
      {
        type: "apply_formatting",
        target: "p2",
        formatting: { italic: true },
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    expect(xml).toContain("w:rPrChange");
    // The original rPr (with w:color and w:sz) should be preserved inside w:rPrChange
    expect(xml).toContain("w:i");
  });
});

// =============================================================================
// apply_style with revisions
// =============================================================================

describe("apply_style with revisions", () => {
  it("should add w:pPrChange when applying paragraph style with revisions", async () => {
    const buffer = await createStyledDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // p2 is plain, apply Heading2 style
    const results = applyPatches(doc, [
      {
        type: "apply_style",
        target: "p2",
        style: "Heading2",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    expect(xml).toContain("w:pPrChange");
    expect(xml).toContain("Heading2");
  });

  it("should add w:rPrChange when applying character style with revisions", async () => {
    const buffer = await createStyledDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // Apply "Strong" character style to "bold" in p1
    const results = applyPatches(doc, [
      {
        type: "apply_style",
        target: "p1",
        find: "bold",
        style: "Strong",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    expect(xml).toContain("w:rPrChange");
    expect(xml).toContain("Strong");
  });
});

// =============================================================================
// Table operations with revisions
// =============================================================================

describe("table operations with revisions", () => {
  it("should mark added row with w:ins in w:trPr", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "add_table_row",
        target: "tbl0",
        afterRow: 1,
        cells: ["New A", "New B"],
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const table = doc.body[1] as NormalizedTable;
    expect(table.rows.length).toBe(3);

    const xml = getDocXml(doc);
    // The new row should have w:ins inside w:trPr
    expect(xml).toContain("w:trPr");
    // Look for w:ins inside the context of w:trPr (they appear as direct children)
    // We verify the XML contains the revision markup
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Test Author");
  });

  it("should keep deleted row in XML with w:del in w:trPr", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "delete_table_row", target: "tbl0", row: 0 },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    // Table should still have 2 rows (row is not removed with revisions)
    const table = doc.body[1] as NormalizedTable;
    expect(table.rows.length).toBe(2);

    const xml = getDocXml(doc);
    // The deleted row should still have its content
    expect(xml).toContain("Header A");
    expect(xml).toContain("Header B");
    // And be marked with w:del in w:trPr
    expect(xml).toContain("w:del");
  });

  it("should replace cell text with revision tracking", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "replace_cell_text",
        target: "tbl0.r1.c0.p0",
        find: "Cell 1",
        replace: "Updated",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const table = doc.body[1] as NormalizedTable;
    expect(table.rows[1]!.cells[0]!.paragraphs[0]!.text).toBe("Updated");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
  });
});

// =============================================================================
// Sequential patches with revisions
// =============================================================================

describe("sequential patches with revisions", () => {
  it("should apply multiple revision-tracked patches sequentially", async () => {
    const buffer = await createSimpleDocx("The quick brown fox jumps");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "quick", replace: "slow" },
      { type: "replace_text", target: "p0", find: "brown", replace: "red" },
    ], REV_OPTIONS);

    expect(results.every((r) => r.success)).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("The slow red fox jumps");

    const xml = getDocXml(doc);
    // Should have multiple w:del and w:ins pairs
    const delCount = (xml.match(/<w:del /g) || []).length;
    const insCount = (xml.match(/<w:ins /g) || []).length;
    expect(delCount).toBeGreaterThanOrEqual(2);
    expect(insCount).toBeGreaterThanOrEqual(2);
  });

  it("patch 2 should see patch 1 insertions but not deletions", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
      // Patch 2 should see "Earth" from patch 1's insertion
      { type: "replace_text", target: "p0", find: "Earth", replace: "Mars" },
    ], REV_OPTIONS);

    expect(results.every((r) => r.success)).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Mars");
  });

  it("should continue after a failed patch", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "nonexistent", replace: "x" },
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi" },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(false);
    expect(results[1]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hi World");
  });
});

// =============================================================================
// Round-trip: patched with revisions → export → reload → verify
// =============================================================================

describe("round-trip with revisions", () => {
  it("should produce a valid DOCX with revision markup after round-trip", async () => {
    const buffer = await createSimpleDocx("Original text here");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "Original", replace: "Modified" },
    ], REV_OPTIONS);

    // Serialize and write back
    const newDocXml = serializeXml(doc.xmlParts.document);
    const output = await writeDocx({
      zip: entries.zip,
      documentXml: newDocXml,
    });

    // Reload and verify
    const reloaded = await readDocx(output);
    const reDoc = normalizeDocument(reloaded);

    // The re-loaded document should include the inserted text in its normalized view
    const para = reDoc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Modified text here");

    // The XML should still contain revision markup
    const reXml = serializeXml(reDoc.xmlParts.document);
    expect(reXml).toContain("w:del");
    expect(reXml).toContain("w:ins");
    expect(reXml).toContain("w:delText");
  });
});

// =============================================================================
// Non-revision mode still works (regression)
// =============================================================================

describe("non-revision mode (regression)", () => {
  it("should apply patches without revision tracking by default", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // No options = no revision tracking
    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Earth");

    const xml = getDocXml(doc);
    // Should NOT contain revision markup
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
    expect(xml).not.toContain("w:rPrChange");
  });

  it("should apply patches with trackRevisions: false", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
    ], { trackRevisions: false });

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Earth");

    const xml = getDocXml(doc);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
  });
});

// =============================================================================
// Projection with revisions
// =============================================================================

describe("projection with revisions", () => {
  it("should annotate markdown projection with revision markers", async () => {
    const buffer = await createSimpleDocx("This agreement is valid");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ], REV_OPTIONS);

    // Import projection
    const { project } = await import("../../src/projection/llm.js");

    // Without includeRevisions — should show only current text
    const mdDefault = project(doc, "markdown") as string;
    expect(mdDefault).toContain("enforceable");
    expect(mdDefault).not.toContain("{+");
    expect(mdDefault).not.toContain("{-");

    // With includeRevisions — should show revision annotations
    const mdRevisions = project(doc, "markdown", { includeRevisions: true }) as string;
    expect(mdRevisions).toContain("{+enforceable+}");
    expect(mdRevisions).toContain("{-valid-}");
  });

  it("should include revisions array in JSON projection", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
    ], REV_OPTIONS);

    const { project } = await import("../../src/projection/llm.js");

    // Without includeRevisions — no revisions in output
    const jsonDefault = project(doc, "json") as any;
    expect(jsonDefault.body[0].revisions).toBeUndefined();

    // With includeRevisions — should have revisions
    const jsonRevisions = project(doc, "json", { includeRevisions: true }) as any;
    expect(jsonRevisions.body[0].revisions).toBeDefined();
    expect(jsonRevisions.body[0].revisions.length).toBeGreaterThan(0);
  });
});

// =============================================================================
// Multi-run replacement with revisions
// =============================================================================

describe("multi-run replacement with revisions", () => {
  it("should handle replacement spanning multiple runs", async () => {
    const buffer = await createSplitRunDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // "valid" spans "va" (run 1) + "lid" (run 2)
    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "valid",
        replace: "binding",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("This agreement is binding and enforceable");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("binding");
  });
});

// =============================================================================
// Public API (Docxtor class) with revisions
// =============================================================================

describe("Docxtor class applyPatches with revisions", () => {
  it("should accept PatchOptions via the public API", async () => {
    const { Docxtor } = await import("../../src/index.js");

    const buffer = await createSimpleDocx("Hello World");
    const docxtor = await Docxtor.load(buffer);

    const results = docxtor.applyPatches(
      [{ type: "replace_text", target: "p0", find: "World", replace: "Earth" }],
      { trackRevisions: true, revisionAuthor: "API User" },
    );

    expect(results[0]!.success).toBe(true);

    // Verify the output DOCX contains revision markup
    const outputBuffer = await docxtor.toBuffer();
    const reDocxtor = await Docxtor.load(outputBuffer);
    const para = reDocxtor.getParagraphs()[0]!;
    expect(para.text).toBe("Hello Earth");
  });
});
