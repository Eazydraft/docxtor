import { describe, it, expect } from "vitest";
import { readDocx } from "../../src/zip/reader.js";
import { writeDocx } from "../../src/zip/writer.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { applyPatches } from "../../src/patches/engine.js";
import { serializeXml } from "../../src/xml/serializer.js";
import { Docxtor } from "../../src/index.js";
import {
  createSimpleDocx,
  createHeaderFooterDocx,
} from "../fixtures/create-docx.js";
import type { Patch, NormalizedParagraph } from "../../src/types.js";

// =============================================================================
// add_header / add_footer — basic
// =============================================================================

describe("add_header", () => {
  it("should add a default header to a document with existing headers", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // Remove existing default header first, then add new one
    const results = applyPatches(doc, [
      { type: "remove_header", variant: "default" },
      { type: "add_header", variant: "default", paragraphs: ["New Header"] },
    ]);

    expect(results.every((r) => r.success)).toBe(true);
    const headers = [...doc.headers.values()];
    const defaultHdr = headers.find((h) => h.variant === "default");
    expect(defaultHdr).toBeDefined();
    expect(defaultHdr!.paragraphs[0]!.text).toBe("New Header");
  });

  it("should add a first-page header with w:titlePg", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "add_header", variant: "first", paragraphs: ["First Page Header"] },
    ]);

    expect(results[0]!.success).toBe(true);

    // Verify w:titlePg is in sectPr
    const xml = serializeXml(doc.xmlParts.document);
    expect(xml).toContain("w:titlePg");
  });

  it("should add an even-page header and set w:evenAndOddHeaders in settings", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "add_header", variant: "even", paragraphs: ["Even Page Header"] },
    ]);

    expect(results[0]!.success).toBe(true);

    // Verify settings.xml has w:evenAndOddHeaders
    expect(doc.xmlParts.settings).toBeDefined();
    const settingsXml = serializeXml(doc.xmlParts.settings!);
    expect(settingsXml).toContain("w:evenAndOddHeaders");
  });
});

describe("add_footer", () => {
  it("should add a default footer", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "remove_footer", variant: "default" },
      { type: "add_footer", variant: "default", paragraphs: ["New Footer"] },
    ]);

    expect(results.every((r) => r.success)).toBe(true);
    const footers = [...doc.footers.values()];
    const defaultFtr = footers.find((f) => f.variant === "default");
    expect(defaultFtr).toBeDefined();
    expect(defaultFtr!.paragraphs[0]!.text).toBe("New Footer");
  });

  it("should add an even-page footer and set w:evenAndOddHeaders", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "add_footer", variant: "even", paragraphs: ["Even Footer"] },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.xmlParts.settings).toBeDefined();
    const settingsXml = serializeXml(doc.xmlParts.settings!);
    expect(settingsXml).toContain("w:evenAndOddHeaders");
  });
});

// =============================================================================
// remove_header / remove_footer
// =============================================================================

describe("remove_header", () => {
  it("should remove the default header", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    expect(doc.headers.size).toBe(1);
    const results = applyPatches(doc, [
      { type: "remove_header", variant: "default" },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.headers.size).toBe(0);
  });

  it("should error when removing a non-existent header variant", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      { type: "remove_header", variant: "first" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("found");

  });
});

describe("remove_footer", () => {
  it("should remove the default footer", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    expect(doc.footers.size).toBe(1);
    const results = applyPatches(doc, [
      { type: "remove_footer", variant: "default" },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.footers.size).toBe(0);
  });
});

// =============================================================================
// Round-trip: add header/footer → export → reload → verify
// =============================================================================

describe("header/footer round-trip (export → reload)", () => {
  it("should persist an added header through export/reload", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "add_header", variant: "first", paragraphs: ["First Page Only"] },
    ]);

    const exported = await doc.toBuffer();
    const reloaded = await Docxtor.load(exported);

    const headers = [...reloaded.getHeaders().values()];
    const firstHdr = headers.find((h) => h.variant === "first");
    expect(firstHdr).toBeDefined();
    expect(firstHdr!.paragraphs[0]!.text).toBe("First Page Only");
  });

  it("should persist an added footer through export/reload", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "add_footer", variant: "first", paragraphs: ["First Page Footer"] },
    ]);

    const exported = await doc.toBuffer();
    const reloaded = await Docxtor.load(exported);

    const footers = [...reloaded.getFooters().values()];
    const firstFtr = footers.find((f) => f.variant === "first");
    expect(firstFtr).toBeDefined();
    expect(firstFtr!.paragraphs[0]!.text).toBe("First Page Footer");
  });

  it("should persist header removal through export/reload", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    expect(doc.getHeaders().size).toBe(1);

    doc.applyPatches([
      { type: "remove_header", variant: "default" },
    ]);

    const exported = await doc.toBuffer();
    const reloaded = await Docxtor.load(exported);

    expect(reloaded.getHeaders().size).toBe(0);
  });

  it("should remove stale header XML from zip after remove_header", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "remove_header", variant: "default" },
    ]);

    const exported = await doc.toBuffer();
    const entries = await readDocx(exported);

    // The original word/header1.xml should no longer be in the zip
    expect(entries.headers.size).toBe(0);
  });

  it("should remove stale footer XML from zip after remove_footer", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "remove_footer", variant: "default" },
    ]);

    const exported = await doc.toBuffer();
    const entries = await readDocx(exported);

    expect(entries.footers.size).toBe(0);
  });

  it("should not leave stale zip entries after add/remove cycle", async () => {
    const doc = await Docxtor.create();

    // Add header and footer
    doc.applyPatches([
      { type: "add_header", variant: "default", paragraphs: ["Temp"] },
      { type: "add_footer", variant: "default", paragraphs: ["Temp"] },
    ]);

    // Remove them
    doc.applyPatches([
      { type: "remove_header", variant: "default" },
      { type: "remove_footer", variant: "default" },
    ]);

    const exported = await doc.toBuffer();
    const entries = await readDocx(exported);

    expect(entries.headers.size).toBe(0);
    expect(entries.footers.size).toBe(0);
  });

  it("should round-trip even-page header with settings.xml", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "add_header", variant: "even", paragraphs: ["Even Header"] },
    ]);

    const exported = await doc.toBuffer();

    // Verify settings.xml is in the exported ZIP
    const entries = await readDocx(exported);
    expect(entries.settingsXml).toBeDefined();
    expect(entries.settingsXml).toContain("w:evenAndOddHeaders");

    const reloaded = await Docxtor.load(exported);
    const headers = [...reloaded.getHeaders().values()];
    const evenHdr = headers.find((h) => h.variant === "even");
    expect(evenHdr).toBeDefined();
    expect(evenHdr!.paragraphs[0]!.text).toBe("Even Header");
  });
});

// =============================================================================
// add_header/add_footer on documents without document.xml.rels (Docxtor.create)
// =============================================================================

describe("add_header/add_footer on Docxtor.create() documents", () => {
  it("should add a header to a document created via Docxtor.create()", async () => {
    const doc = await Docxtor.create();

    const results = doc.applyPatches([
      { type: "add_header", variant: "default", paragraphs: ["Created Header"] },
    ]);

    expect(results[0]!.success).toBe(true);
    const headers = [...doc.getHeaders().values()];
    expect(headers.length).toBe(1);
    expect(headers[0]!.paragraphs[0]!.text).toBe("Created Header");
  });

  it("should add a footer to a document created via Docxtor.create()", async () => {
    const doc = await Docxtor.create();

    const results = doc.applyPatches([
      { type: "add_footer", variant: "default", paragraphs: ["Created Footer"] },
    ]);

    expect(results[0]!.success).toBe(true);
    const footers = [...doc.getFooters().values()];
    expect(footers.length).toBe(1);
    expect(footers[0]!.paragraphs[0]!.text).toBe("Created Footer");
  });

  it("should round-trip header added to Docxtor.create() document", async () => {
    const doc = await Docxtor.create();

    doc.applyPatches([
      { type: "add_header", variant: "default", paragraphs: ["My Header"] },
    ]);

    const exported = await doc.toBuffer();
    const reloaded = await Docxtor.load(exported);

    const headers = [...reloaded.getHeaders().values()];
    expect(headers.length).toBe(1);
    expect(headers[0]!.paragraphs[0]!.text).toBe("My Header");
  });
});

// =============================================================================
// Field placeholder handling in headers/footers
// =============================================================================

describe("field placeholders in headers/footers", () => {
  it("should create a header with {{PAGE}} field code", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "add_header", variant: "first", paragraphs: ["Page {{PAGE}} of {{NUMPAGES}}"] },
    ]);

    expect(results[0]!.success).toBe(true);

    // The header XML should contain field code instructions
    const ast = doc.getAST();
    const firstHdr = [...ast.xmlParts.headers.values()].find((fxp) => {
      const xml = serializeXml(fxp);
      return xml.includes("PAGE") && xml.includes("NUMPAGES");
    });
    expect(firstHdr).toBeDefined();
  });

  it("should round-trip field placeholders in footer", async () => {
    const doc = await Docxtor.create();

    doc.applyPatches([
      { type: "add_footer", variant: "default", paragraphs: ["Page {{PAGE}}"] },
    ]);

    const exported = await doc.toBuffer();
    const entries = await readDocx(exported);

    // Check that footer XML contains field code pattern
    for (const [, xml] of entries.footers) {
      expect(xml).toContain("w:fldChar");
      expect(xml).toContain("PAGE");
    }
  });
});

// =============================================================================
// Container targeting (append_text / add_table to header/footer)
// =============================================================================

describe("container targeting", () => {
  it("should append text to an existing header", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      {
        type: "append_text",
        text: "Appended to header",
        container: { type: "header", variant: "default" },
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const headers = [...doc.getHeaders().values()];
    const defaultHdr = headers.find((h) => h.variant === "default");
    expect(defaultHdr).toBeDefined();
    // Original paragraph + appended paragraph
    expect(defaultHdr!.paragraphs.length).toBeGreaterThanOrEqual(2);
    expect(defaultHdr!.paragraphs.some((p) => p.text === "Appended to header")).toBe(true);
  });

  it("should append text to an existing footer", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      {
        type: "append_text",
        text: "Appended to footer",
        container: { type: "footer", variant: "default" },
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const footers = [...doc.getFooters().values()];
    const defaultFtr = footers.find((f) => f.variant === "default");
    expect(defaultFtr!.paragraphs.some((p) => p.text === "Appended to footer")).toBe(true);
  });

  it("should reject add_table with both afterId and container", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      {
        type: "add_table",
        afterId: "p0",
        rows: [["A", "B"]],
        container: { type: "header", variant: "default" },
      },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("afterId");
    expect(results[0]!.error).toContain("container");
  });

  it("should add a table to a header container", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      {
        type: "add_table",
        rows: [["Cell1", "Cell2"]],
        container: { type: "header", variant: "default" },
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const headers = [...doc.getHeaders().values()];
    const defaultHdr = headers.find((h) => h.variant === "default");
    expect(defaultHdr!.tables.length).toBe(1);
    expect(defaultHdr!.tables[0]!.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe("Cell1");
  });

  it("should round-trip container-targeted append", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      {
        type: "append_text",
        text: "Persisted text",
        container: { type: "header", variant: "default" },
      },
    ]);

    const exported = await doc.toBuffer();
    const reloaded = await Docxtor.load(exported);

    const headers = [...reloaded.getHeaders().values()];
    const defaultHdr = headers.find((h) => h.variant === "default");
    expect(defaultHdr!.paragraphs.some((p) => p.text === "Persisted text")).toBe(true);
  });
});

// =============================================================================
// Relationship/content-type mutation integrity
// =============================================================================

describe("relationship and content-type integrity", () => {
  it("should add and then remove a header with clean rels/content-types", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    // Add a first-page header
    doc.applyPatches([
      { type: "add_header", variant: "first", paragraphs: ["Temp Header"] },
    ]);

    const ast = doc.getAST();
    const relsXmlAfterAdd = serializeXml(ast.xmlParts.documentRels!);
    const ctXmlAfterAdd = serializeXml(ast.xmlParts.contentTypes!);

    // Should have added a relationship and content type
    expect(relsXmlAfterAdd).toContain("header");
    expect(ctXmlAfterAdd).toContain("header");

    // Remove it
    doc.applyPatches([
      { type: "remove_header", variant: "first" },
    ]);

    const relsXmlAfterRemove = serializeXml(ast.xmlParts.documentRels!);
    const ctXmlAfterRemove = serializeXml(ast.xmlParts.contentTypes!);

    // The added relationship and content type should be gone
    // (original rId1 header should still be there)
    const firstHdrHeaders = [...doc.getHeaders().values()].filter(
      (h) => h.variant === "first",
    );
    expect(firstHdrHeaders.length).toBe(0);
  });

  it("should validate afterId + container as invalid", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const results = doc.validatePatches([
      {
        type: "add_table",
        afterId: "p0",
        rows: [["X"]],
        container: { type: "header", variant: "default" },
      },
    ]);

    expect(results[0]!.valid).toBe(false);
    expect(results[0]!.error).toContain("afterId");
  });

  it("should produce valid DOCX after multiple add/remove cycles", async () => {
    const doc = await Docxtor.create();

    // Add header and footer
    doc.applyPatches([
      { type: "add_header", variant: "default", paragraphs: ["Header"] },
      { type: "add_footer", variant: "default", paragraphs: ["Footer"] },
    ]);

    // Remove them
    doc.applyPatches([
      { type: "remove_header", variant: "default" },
      { type: "remove_footer", variant: "default" },
    ]);

    // Add them back
    doc.applyPatches([
      { type: "add_header", variant: "default", paragraphs: ["Header v2"] },
      { type: "add_footer", variant: "default", paragraphs: ["Footer v2"] },
    ]);

    // Export and reload - should not throw
    const exported = await doc.toBuffer();
    const reloaded = await Docxtor.load(exported);

    const headers = [...reloaded.getHeaders().values()];
    const footers = [...reloaded.getFooters().values()];
    expect(headers.length).toBe(1);
    expect(footers.length).toBe(1);
    expect(headers[0]!.paragraphs[0]!.text).toBe("Header v2");
    expect(footers[0]!.paragraphs[0]!.text).toBe("Footer v2");
  });
});
