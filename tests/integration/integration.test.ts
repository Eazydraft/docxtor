import { describe, it, expect } from "vitest";
import { Docxtor } from "../../src/index.js";
import {
  createSimpleDocx,
  createFormattedDocx,
  createTableDocx,
  createHeaderFooterDocx,
  createSplitRunDocx,
} from "../fixtures/create-docx.js";

// =============================================================================
// Loading
// =============================================================================

describe("Docxtor.load", () => {
  it("should load a DOCX from a buffer", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const doc = await Docxtor.load(buffer);

    expect(doc).toBeInstanceOf(Docxtor);
    expect(doc.getParagraphs()).toHaveLength(1);
    expect(doc.getParagraphs()[0]!.text).toBe("Hello World");
  });

  it("should expose body elements", async () => {
    const buffer = await createTableDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(3); // paragraph, table, paragraph

    expect(doc.getParagraphs()).toHaveLength(2);
    expect(doc.getTables()).toHaveLength(1);
  });
});

// =============================================================================
// Projection — JSON
// =============================================================================

describe("Docxtor.project — JSON", () => {
  it("should project a simple document to JSON", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    expect(projection.body).toHaveLength(1);
    expect(projection.body[0]).toEqual({
      id: "p0",
      text: "Hello World",
    });
  });

  it("should project a formatted document with styles", async () => {
    const buffer = await createFormattedDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    expect(projection.body).toHaveLength(3);
    expect(projection.body[0]).toMatchObject({
      id: "p0",
      text: "Document Title",
      style: "Heading1",
    });
  });

  it("should project a document with a table", async () => {
    const buffer = await createTableDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    // Find the table in the projection
    const table = projection.body.find((e: any) => e.id === "tbl0") as any;
    expect(table).toBeDefined();
    expect(table.rows).toHaveLength(2);
    expect(table.rows[0][0].text).toBe("Header A");
    expect(table.rows[1][1].text).toBe("Cell 2");
  });

  it("should include headers and footers", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    expect(projection.headers).toHaveLength(1);
    expect(projection.headers![0]!.paragraphs[0]!.text).toBe("Header text");
    expect(projection.footers).toHaveLength(1);
    expect(projection.footers![0]!.paragraphs[0]!.text).toBe("Footer text");
  });

  it("should include formatting when requested", async () => {
    const buffer = await createFormattedDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includeFormatting: true });

    const heading = projection.body[0] as any;
    expect(heading.formatting).toBeDefined();
    expect(heading.formatting[0]).toEqual({ bold: true });
  });

  it("should exclude headers/footers when requested", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includeHeadersFooters: false });

    expect(projection.headers).toBeUndefined();
    expect(projection.footers).toBeUndefined();
  });
});

// =============================================================================
// Projection — Markdown
// =============================================================================

describe("Docxtor.project — Markdown", () => {
  it("should project a simple document to markdown", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toBe("[p0] Hello World");
  });

  it("should project styles as hints", async () => {
    const buffer = await createFormattedDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toContain("[p0] (Heading1) Document Title");
    expect(markdown).toContain("[p1] This is bold and italic text.");
  });

  it("should project tables as pipe-delimited rows", async () => {
    const buffer = await createTableDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toContain("| [tbl0.r0.c0.p0] Header A");
    expect(markdown).toContain("[tbl0.r0.c1.p0] Header B |");
    expect(markdown).toContain("[tbl0.r1.c0.p0] Cell 1");
  });

  it("should include header/footer sections", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toContain("--- Header (default) ---");
    expect(markdown).toContain("[hdr.default.p0] Header text");
    expect(markdown).toContain("--- Footer (default) ---");
    expect(markdown).toContain("[ftr.default.p0] Footer text");
  });
});

// =============================================================================
// Patching through public API
// =============================================================================

describe("Docxtor.applyPatches", () => {
  it("should apply text replacement", async () => {
    const buffer = await createSimpleDocx("This agreement is valid");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("This agreement is enforceable");
  });

  it("should report failures without throwing", async () => {
    const buffer = await createSimpleDocx("Hello");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "nonexistent", replace: "x" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toBeDefined();
  });
});

// =============================================================================
// Full pipeline: load → project → patch → export → reload
// =============================================================================

describe("Full pipeline", () => {
  it("should survive a complete edit cycle", async () => {
    // 1. Create a DOCX
    const buffer = await createFormattedDocx();

    // 2. Load it
    const doc = await Docxtor.load(buffer);

    // 3. Get projection (what the LLM would see)
    const projection = doc.project("json");
    expect(projection.body).toHaveLength(3);

    // 4. Apply patches (what the LLM would produce)
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Document Title", replace: "Contract v2" },
      { type: "replace_text", target: "p2", find: "Red large text", replace: "Blue small text" },
    ]);
    expect(results.every((r) => r.success)).toBe(true);

    // 5. Export
    const output = await doc.toBuffer();
    expect(Buffer.isBuffer(output)).toBe(true);

    // 6. Reload and verify
    const reloaded = await Docxtor.load(output);
    expect(reloaded.getParagraphs()[0]!.text).toBe("Contract v2");
    expect(reloaded.getParagraphs()[0]!.style).toBe("Heading1");

    // Formatting should be preserved
    expect(reloaded.getParagraphs()[0]!.runs[0]!.formatting.bold).toBe(true);
  });

  it("should handle table edits end-to-end", async () => {
    const buffer = await createTableDocx();
    const doc = await Docxtor.load(buffer);

    // Replace cell text + add a row
    doc.applyPatches([
      {
        type: "replace_cell_text",
        target: "tbl0.r1.c0.p0",
        find: "Cell 1",
        replace: "Updated 1",
      },
      {
        type: "add_table_row",
        target: "tbl0",
        afterRow: 1,
        cells: ["New A", "New B"],
      },
    ]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);

    const table = reloaded.getTables()[0]!;
    expect(table.rows).toHaveLength(3);
    expect(table.rows[1]!.cells[0]!.paragraphs[0]!.text).toBe("Updated 1");
    expect(table.rows[2]!.cells[0]!.paragraphs[0]!.text).toBe("New A");
  });

  it("should handle header/footer edits end-to-end", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      {
        type: "replace_text",
        target: "hdr.default.p0",
        find: "Header text",
        replace: "CONFIDENTIAL",
      },
    ]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);

    const headers = [...reloaded.getHeaders().values()];
    expect(headers[0]!.paragraphs[0]!.text).toBe("CONFIDENTIAL");
  });

  it("should handle cross-run replacement end-to-end", async () => {
    const buffer = await createSplitRunDocx();
    const doc = await Docxtor.load(buffer);

    // "valid" spans two runs
    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "valid", replace: "binding" },
    ]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);
    expect(reloaded.getParagraphs()[0]!.text).toBe(
      "This agreement is binding and enforceable",
    );
  });

  it("projection after patching should reflect changes", async () => {
    const buffer = await createSimpleDocx("Original text");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Original", replace: "Updated" },
    ]);

    const projection = doc.project("json");
    expect(projection.body[0]).toMatchObject({
      id: "p0",
      text: "Updated text",
    });
  });
});
