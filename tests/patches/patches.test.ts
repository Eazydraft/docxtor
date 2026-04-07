import { describe, it, expect } from "vitest";
import { parseXml } from "../../src/xml/parser.js";
import { serializeXml } from "../../src/xml/serializer.js";
import { normalizeParagraph } from "../../src/normalizer/paragraph.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { readDocx } from "../../src/zip/reader.js";
import { writeDocx } from "../../src/zip/writer.js";
import { applyPatches } from "../../src/patches/engine.js";
import {
  createSimpleDocx,
  createFormattedDocx,
  createTableDocx,
  createHeaderFooterDocx,
  createSplitRunDocx,
} from "../fixtures/create-docx.js";
import type { Patch, NormalizedParagraph, NormalizedTable } from "../../src/types.js";

// =============================================================================
// Text replacement — single run
// =============================================================================

describe("replace_text — single run", () => {
  it("should replace text within a single run", async () => {
    const buffer = await createSimpleDocx("This agreement is valid");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("This agreement is enforceable");
  });

  it("should replace text at the beginning of a run", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi" },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hi World");
  });

  it("should replace text at the end of a run", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Earth");
  });

  it("should replace entire run text", async () => {
    const buffer = await createSimpleDocx("Replace me");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "Replace me", replace: "Done" },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Done");
  });

  it("should handle specific occurrence", async () => {
    const buffer = await createSimpleDocx("foo bar foo baz foo");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      {
        type: "replace_text",
        target: "p0",
        find: "foo",
        replace: "qux",
        occurrence: 2,
      },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("foo bar qux baz foo");
  });

  it("should fail when text not found", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "nonexistent", replace: "x" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("not found");
  });

  it("should fail when target paragraph not found", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p999", find: "Hello", replace: "x" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("not found");
  });
});

// =============================================================================
// Text replacement — multi-run (with run splitting)
// =============================================================================

describe("replace_text — multi-run with splitting", () => {
  it("should replace text spanning two runs", async () => {
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
    ]);

    expect(results[0]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("This agreement is binding and enforceable");
  });

  it("should replace text in multi-run formatted paragraph", async () => {
    const buffer = await createFormattedDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // "bold and italic" spans multiple runs in paragraph 1
    // p1 text: "This is bold and italic text."
    const para = doc.body[1] as NormalizedParagraph;
    expect(para.text).toBe("This is bold and italic text.");

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p1", find: "bold", replace: "strong" },
    ]);

    expect(results[0]!.success).toBe(true);
    const updated = doc.body[1] as NormalizedParagraph;
    expect(updated.text).toContain("strong");
    expect(updated.text).toContain("italic");
  });
});

// =============================================================================
// Text deletion
// =============================================================================

describe("delete_text", () => {
  it("should delete text from a paragraph", async () => {
    const buffer = await createSimpleDocx("Remove this word please");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "delete_text", target: "p0", find: "this word " },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Remove please");
  });
});

// =============================================================================
// Text insertion
// =============================================================================

describe("insert_text", () => {
  it("should insert text before an anchor", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "before",
        anchor: "World",
        text: "Beautiful ",
      },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Beautiful World");
  });

  it("should insert text after an anchor", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      {
        type: "insert_text",
        target: "p0",
        position: "after",
        anchor: "Hello",
        text: " Beautiful",
      },
    ]);

    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hello Beautiful World");
  });
});

// =============================================================================
// Table operations
// =============================================================================

describe("table operations", () => {
  it("should add a row to a table", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const table = doc.body[1] as NormalizedTable;
    expect(table.rows.length).toBe(2);

    applyPatches(doc, [
      {
        type: "add_table_row",
        target: "tbl0",
        afterRow: 1,
        cells: ["New A", "New B"],
      },
    ]);

    const updated = doc.body[1] as NormalizedTable;
    expect(updated.rows.length).toBe(3);
    expect(updated.rows[2]!.cells[0]!.paragraphs[0]!.text).toBe("New A");
    expect(updated.rows[2]!.cells[1]!.paragraphs[0]!.text).toBe("New B");
  });

  it("should add a row at the beginning", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      {
        type: "add_table_row",
        target: "tbl0",
        afterRow: -1,
        cells: ["First A", "First B"],
      },
    ]);

    const updated = doc.body[1] as NormalizedTable;
    expect(updated.rows.length).toBe(3);
    expect(updated.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe("First A");
  });

  it("should delete a row from a table", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "delete_table_row", target: "tbl0", row: 0 },
    ]);

    const updated = doc.body[1] as NormalizedTable;
    expect(updated.rows.length).toBe(1);
    // Should be the former row 1 (Cell 1, Cell 2)
    expect(updated.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe("Cell 1");
  });

  it("should replace text in a table cell", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      {
        type: "replace_cell_text",
        target: "tbl0.r1.c0.p0",
        find: "Cell 1",
        replace: "Updated",
      },
    ]);

    const table = doc.body[1] as NormalizedTable;
    expect(table.rows[1]!.cells[0]!.paragraphs[0]!.text).toBe("Updated");
  });

  it("should fail on cell count mismatch", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "add_table_row",
        target: "tbl0",
        afterRow: 0,
        cells: ["Only one"], // Table has 2 columns
      },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("Cell count mismatch");
  });
});

// =============================================================================
// Sequential patches
// =============================================================================

describe("sequential patches", () => {
  it("should apply multiple patches in sequence", async () => {
    const buffer = await createSimpleDocx("The quick brown fox jumps");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "quick", replace: "slow" },
      { type: "replace_text", target: "p0", find: "brown", replace: "red" },
      { type: "replace_text", target: "p0", find: "fox", replace: "cat" },
    ]);

    expect(results.every((r) => r.success)).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("The slow red cat jumps");
  });

  it("should handle patches that depend on prior patches", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
      // This second patch operates on updated text
      { type: "replace_text", target: "p0", find: "Earth", replace: "Mars" },
    ]);

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
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[1]!.success).toBe(true);
    const para = doc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Hi World");
  });
});

// =============================================================================
// Header/footer patches
// =============================================================================

describe("header/footer patches", () => {
  it("should patch text in a header", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "hdr.default.p0",
        find: "Header text",
        replace: "New Header",
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const header = [...doc.headers.values()][0]!;
    expect(header.paragraphs[0]!.text).toBe("New Header");
  });

  it("should patch text in a footer", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "replace_text",
        target: "ftr.default.p0",
        find: "Footer text",
        replace: "Page 1 of 10",
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const footer = [...doc.footers.values()][0]!;
    expect(footer.paragraphs[0]!.text).toBe("Page 1 of 10");
  });
});

// =============================================================================
// Full round-trip (load → patch → serialize → reload → verify)
// =============================================================================

describe("full round-trip with patches", () => {
  it("should produce a valid DOCX after patching", async () => {
    const buffer = await createSimpleDocx("Original text here");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "Original", replace: "Modified" },
    ]);

    // Serialize the mutated FXP back to XML
    const { serializeXml } = await import("../../src/xml/serializer.js");
    const newDocXml = serializeXml(doc.xmlParts.document);

    // Write back to DOCX
    const output = await writeDocx({
      zip: entries.zip,
      documentXml: newDocXml,
    });

    // Reload and verify
    const reloaded = await readDocx(output);
    const reDoc = normalizeDocument(reloaded);
    const para = reDoc.body[0] as NormalizedParagraph;
    expect(para.text).toBe("Modified text here");
  });

  it("should preserve formatting through round-trip", async () => {
    const buffer = await createFormattedDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // Patch the heading
    applyPatches(doc, [
      { type: "replace_text", target: "p0", find: "Document Title", replace: "New Title" },
    ]);

    const { serializeXml } = await import("../../src/xml/serializer.js");
    const newDocXml = serializeXml(doc.xmlParts.document);
    const output = await writeDocx({ zip: entries.zip, documentXml: newDocXml });

    const reloaded = await readDocx(output);
    const reDoc = normalizeDocument(reloaded);

    const heading = reDoc.body[0] as NormalizedParagraph;
    expect(heading.text).toBe("New Title");
    expect(heading.style).toBe("Heading1");
    expect(heading.runs[0]!.formatting.bold).toBe(true);
  });

  it("should preserve table structure through round-trip after adding a row", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      {
        type: "add_table_row",
        target: "tbl0",
        afterRow: 1,
        cells: ["Added A", "Added B"],
      },
    ]);

    const { serializeXml } = await import("../../src/xml/serializer.js");
    const newDocXml = serializeXml(doc.xmlParts.document);
    const output = await writeDocx({ zip: entries.zip, documentXml: newDocXml });

    const reloaded = await readDocx(output);
    const reDoc = normalizeDocument(reloaded);

    const table = reDoc.body[1] as NormalizedTable;
    expect(table.rows.length).toBe(3);
    expect(table.rows[2]!.cells[0]!.paragraphs[0]!.text).toBe("Added A");
  });
});
