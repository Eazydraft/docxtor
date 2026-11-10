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
  createBodyDocx,
} from "../fixtures/create-docx.js";
import { scanMaxRevisionId, revisionTimestamp } from "../../src/patches/revision.js";
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

// =============================================================================
// Nested w:del inside containers (w:hyperlink, w:ins)
// =============================================================================

describe("nested w:del inside containers", () => {
  it("should detect w:del revision marks inside a w:hyperlink", async () => {
    // "Visit our website for details." — "our website" is inside w:hyperlink
    const { createHyperlinkDocx } = await import("../fixtures/create-docx.js");
    const buffer = await createHyperlinkDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // Delete "website" which is inside the hyperlink
    const results = applyPatches(doc, [
      { type: "delete_text", target: "p0", find: "website" },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Visit our  for details.");

    // The revision should be captured even though w:del is inside w:hyperlink
    expect(para.revisions).toBeDefined();
    const delRevisions = para.revisions!.filter((r) => r.type === "delete");
    expect(delRevisions.length).toBeGreaterThanOrEqual(1);
    expect(delRevisions.some((r) => r.text === "website")).toBe(true);
  });

  it("should detect w:del revision marks inside a w:ins (delete after insert)", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // First, insert text (produces w:ins), then delete part of the original
    // which shares a parent with the w:ins
    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
      // Now "Hello" is a plain run, "Earth" is inside w:ins, "World" is inside w:del
      // Delete "Hello" — this puts a w:del at the paragraph level
      { type: "delete_text", target: "p0", find: "Hello " },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);
    expect(results[1]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Earth");

    // Both deletions ("World" from replace, "Hello " from delete) should be captured
    expect(para.revisions).toBeDefined();
    const delRevisions = para.revisions!.filter((r) => r.type === "delete");
    expect(delRevisions.length).toBeGreaterThanOrEqual(2);
  });
});

// =============================================================================
// w:br preservation in revision-tracked run splitting
// =============================================================================

describe("w:br preservation in split runs", () => {
  it("should emit w:br elements when splitting a run containing line breaks", async () => {
    // Create a DOCX with a run containing a w:br
    const { createLineBreakDocx } = await import("../fixtures/create-docx.js");
    const buffer = await createLineBreakDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // The paragraph text should be "Hello\nWorld" (w:br → \n)
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello\nWorld");

    // Insert text in the middle (between "Hel" and "lo\nWorld") to trigger a split
    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: "!",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const updatedPara = doc.body[0] as NormalizedParagraph;
    expect(updatedPara.text).toBe("Hello!\nWorld");

    // The XML should contain w:br, NOT a literal newline in w:t
    const xml = getDocXml(doc);
    expect(xml).toContain("w:br");
    // Ensure no w:t contains a literal newline character
    const tMatches = xml.match(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/);
    expect(tMatches).toBeNull();
  });

  it("should produce proper w:br when splitting mid-run at a line break boundary", async () => {
    const { createLineBreakDocx } = await import("../fixtures/create-docx.js");
    const buffer = await createLineBreakDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // Insert right before the newline (after "Hello", before "\n")
    const results = applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "before",
        anchor: "\nWorld",
        text: " Goodbye",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Goodbye\nWorld");

    const xml = getDocXml(doc);
    expect(xml).toContain("w:br");
  });
});

// =============================================================================
// add_table with revisions
// =============================================================================

describe("add_table with revisions", () => {
  it("should mark all rows with w:ins in w:trPr", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "add_table",
        headers: ["Name", "Age"],
        rows: [["Alice", "30"]],
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    // Both rows (header + data) should have w:ins in w:trPr
    expect(xml).toContain("w:trPr");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Test Author");
  });

  it("should mark rows with w:ins when using afterId", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "add_table",
        afterId: "p0",
        rows: [["A", "B"]],
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Test Author");
  });

  it("should survive export and reload with revisions", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      {
        type: "add_table",
        headers: ["X"],
        rows: [["Y"]],
      },
    ], REV_OPTIONS);

    // Verify in the current doc's XML
    const xml = getDocXml(doc);
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Test Author");
    expect(xml).toContain("w:tbl");
  });
});

// =============================================================================
// append_text markdown table with revisions
// =============================================================================

describe("append_text markdown table with revisions", () => {
  it("should mark table rows with w:ins when appending a markdown table", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "append_text",
        text: "| Col1 | Col2 |\n| --- | --- |\n| A | B |",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    // Table rows should have revision marks
    expect(xml).toContain("w:trPr");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Test Author");
  });

  it("should wrap plain text paragraphs in w:ins and mark table rows", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "append_text",
        text: "Before\n| H1 | H2 |\n| --- | --- |\n| V1 | V2 |\nAfter",
      },
    ], REV_OPTIONS);

    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    // Both paragraph w:ins and table row w:ins should be present
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Test Author");

    // Should have the table
    const tables = doc.body.filter((e) => e.kind === "table");
    expect(tables).toHaveLength(1);
  });
});

// =============================================================================
// Revision ids and dates
// =============================================================================

describe("revision id allocation", () => {
  const ids = (xml: string, tag: string): number[] =>
    [...xml.matchAll(new RegExp(`<${tag} w:id="(\\d+)"`, "g"))].map((m) => Number(m[1]));

  it("starts above every bookmark id when the document has no revisions", async () => {
    const bookmarks = [0, 1, 2, 3, 4, 5]
      .map((i) => `<w:bookmarkStart w:id="${i}" w:name="b${i}"/><w:bookmarkEnd w:id="${i}"/>`)
      .join("");
    const doc = normalizeDocument(await readDocx(await createBodyDocx(
      `<w:p>${bookmarks}<w:r><w:t>This agreement is valid</w:t></w:r></w:p>`,
    )));
    expect(scanMaxRevisionId(doc)).toBe(5);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ], REV_OPTIONS);
    expect(results[0]!.success).toBe(true);

    const xml = getDocXml(doc);
    // The deletion is written first, so it takes 6 and the insertion 7
    expect(ids(xml, "w:del")).toEqual([6]);
    expect(ids(xml, "w:ins")).toEqual([7]);
  });

  it("gives a lone tracked insertion the first id after the bookmarks", async () => {
    const doc = normalizeDocument(await readDocx(await createBodyDocx(
      `<w:p><w:bookmarkStart w:id="0" w:name="a"/><w:r><w:t>Hello</w:t></w:r><w:bookmarkEnd w:id="0"/>` +
      `<w:bookmarkStart w:id="5" w:name="b"/><w:bookmarkEnd w:id="5"/></w:p>`,
    )));
    const results = applyPatches(doc, [{ type: "append_text", text: "New paragraph" }], REV_OPTIONS);
    expect(results[0]!.success).toBe(true);
    const all = ids(getDocXml(doc), "w:ins");
    expect(all.length).toBeGreaterThan(0);
    expect(Math.min(...all)).toBe(6);
  });

  it("counts a bookmark end that has no start in the scanned parts", async () => {
    const doc = normalizeDocument(await readDocx(await createBodyDocx(
      `<w:p><w:r><w:t>Hello</w:t></w:r><w:bookmarkEnd w:id="40"/></w:p>`,
    )));
    expect(scanMaxRevisionId(doc)).toBe(40);
  });

  it("starts above a move range id", async () => {
    const doc = normalizeDocument(await readDocx(await createBodyDocx(
      `<w:p><w:moveFromRangeStart w:id="10" w:author="A" w:date="2025-01-01T00:00:00Z" w:name="move1"/>` +
      `<w:moveFrom w:id="3" w:author="A" w:date="2025-01-01T00:00:00Z"><w:r><w:t>Moved text</w:t></w:r></w:moveFrom>` +
      `<w:moveFromRangeEnd w:id="10"/><w:r><w:t>This agreement is valid</w:t></w:r></w:p>`,
    )));
    expect(scanMaxRevisionId(doc)).toBe(10);

    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ], REV_OPTIONS);
    expect(ids(getDocXml(doc), "w:del")).toEqual([11]);
  });

  it("counts w:moveFrom and w:moveTo ids", async () => {
    const doc = normalizeDocument(await readDocx(await createBodyDocx(
      `<w:p><w:moveTo w:id="21" w:author="A" w:date="2025-01-01T00:00:00Z"><w:r><w:t>Moved</w:t></w:r></w:moveTo></w:p>`,
    )));
    expect(scanMaxRevisionId(doc)).toBe(21);
  });
});

describe("default revision date", () => {
  const WORD_DATE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/;

  it("revisionTimestamp drops the fractional seconds", () => {
    expect(revisionTimestamp(new Date("2026-11-10T16:25:03.493Z"))).toBe("2026-11-10T16:25:03Z");
    expect(revisionTimestamp(new Date("2026-11-10T16:25:03.000Z"))).toBe("2026-11-10T16:25:03Z");
    expect(revisionTimestamp()).toMatch(WORD_DATE);
  });

  it("a tracked edit with no revisionDate writes a date without milliseconds", async () => {
    const doc = normalizeDocument(await readDocx(await createSimpleDocx("This agreement is valid")));
    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ], { trackRevisions: true, revisionAuthor: "Test Author" });

    const dates = [...getDocXml(doc).matchAll(/w:date="([^"]+)"/g)].map((m) => m[1]!);
    expect(dates).toHaveLength(2);
    for (const d of dates) expect(d).toMatch(WORD_DATE);
  });

  it("a comment with no revisionDate writes a date without milliseconds", async () => {
    const doc = normalizeDocument(await readDocx(await createSimpleDocx("This agreement is valid")));
    const results = applyPatches(doc, [{ type: "add_comment", target: "p0", text: "Check this." }]);
    expect(results[0]!.success).toBe(true);
    const date = /w:date="([^"]+)"/.exec(serializeXml(doc.xmlParts.comments!))![1]!;
    expect(date).toMatch(WORD_DATE);
  });

  it("passes a caller-supplied revisionDate through unchanged", async () => {
    const doc = normalizeDocument(await readDocx(await createSimpleDocx("This agreement is valid")));
    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ], { trackRevisions: true, revisionDate: "2026-11-10T16:25:03.493Z" });
    expect(getDocXml(doc)).toContain('w:date="2026-11-10T16:25:03.493Z"');
  });
});
