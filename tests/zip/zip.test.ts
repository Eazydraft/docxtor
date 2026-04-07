import { describe, it, expect } from "vitest";
import { readDocx } from "../../src/zip/reader.js";
import { writeDocx } from "../../src/zip/writer.js";
import {
  createSimpleDocx,
  createHeaderFooterDocx,
  createTableDocx,
} from "../fixtures/create-docx.js";

describe("zip/reader", () => {
  it("should read document.xml from a simple DOCX", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);

    expect(entries.documentXml).toBeDefined();
    expect(entries.documentXml).toContain("Hello World");
    expect(entries.documentXml).toContain("w:document");
    expect(entries.documentXml).toContain("w:body");
    expect(entries.documentXml).toContain("w:p");
    expect(entries.documentXml).toContain("w:r");
    expect(entries.documentXml).toContain("w:t");
  });

  it("should read headers and footers", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);

    expect(entries.headers.size).toBe(1);
    expect(entries.footers.size).toBe(1);

    const header = entries.headers.get("word/header1.xml");
    expect(header).toContain("Header text");

    const footer = entries.footers.get("word/footer1.xml");
    expect(footer).toContain("Footer text");
  });

  it("should read document relationships", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);

    expect(entries.documentRels).toBeDefined();
    expect(entries.documentRels).toContain("header1.xml");
    expect(entries.documentRels).toContain("footer1.xml");
  });

  it("should throw on invalid DOCX (missing document.xml)", async () => {
    const JSZip = (await import("jszip")).default;
    const zip = new JSZip();
    zip.file("dummy.txt", "not a docx");
    const buffer = await zip.generateAsync({ type: "nodebuffer" });

    await expect(readDocx(buffer)).rejects.toThrow(
      "Invalid DOCX: missing word/document.xml",
    );
  });

  it("should preserve the JSZip instance for writing", async () => {
    const buffer = await createSimpleDocx("Test");
    const entries = await readDocx(buffer);

    expect(entries.zip).toBeDefined();
    expect(entries.zip.file("word/document.xml")).toBeDefined();
  });
});

describe("zip/writer", () => {
  it("should write modified document.xml back to a valid DOCX", async () => {
    const buffer = await createSimpleDocx("Original text");
    const entries = await readDocx(buffer);

    // Modify the document XML
    const modified = entries.documentXml.replace(
      "Original text",
      "Modified text",
    );

    const output = await writeDocx({
      zip: entries.zip,
      documentXml: modified,
    });

    // Read back the output and verify
    const result = await readDocx(output);
    expect(result.documentXml).toContain("Modified text");
    expect(result.documentXml).not.toContain("Original text");
  });

  it("should preserve unmodified entries", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);

    // Only modify document.xml
    const modified = entries.documentXml.replace(
      "Document body text",
      "Changed body",
    );

    const output = await writeDocx({
      zip: entries.zip,
      documentXml: modified,
    });

    // Header and footer should be preserved
    const result = await readDocx(output);
    expect(result.headers.get("word/header1.xml")).toContain("Header text");
    expect(result.footers.get("word/footer1.xml")).toContain("Footer text");
  });

  it("should write modified headers and footers", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);

    const modifiedHeader = entries.headers
      .get("word/header1.xml")!
      .replace("Header text", "New header");

    const output = await writeDocx({
      zip: entries.zip,
      documentXml: entries.documentXml,
      headers: new Map([["word/header1.xml", modifiedHeader]]),
    });

    const result = await readDocx(output);
    expect(result.headers.get("word/header1.xml")).toContain("New header");
  });

  it("should produce a buffer", async () => {
    const buffer = await createSimpleDocx("Test");
    const entries = await readDocx(buffer);

    const output = await writeDocx({
      zip: entries.zip,
      documentXml: entries.documentXml,
    });

    expect(Buffer.isBuffer(output)).toBe(true);
    expect(output.length).toBeGreaterThan(0);
  });
});

describe("zip round-trip", () => {
  it("should round-trip a simple DOCX without corruption", async () => {
    const buffer = await createSimpleDocx("Round trip test");
    const entries = await readDocx(buffer);

    // Write back without changes
    const output = await writeDocx({
      zip: entries.zip,
      documentXml: entries.documentXml,
    });

    // Read back and verify
    const result = await readDocx(output);
    expect(result.documentXml).toContain("Round trip test");
  });

  it("should round-trip a DOCX with table", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);

    const output = await writeDocx({
      zip: entries.zip,
      documentXml: entries.documentXml,
    });

    const result = await readDocx(output);
    expect(result.documentXml).toContain("Header A");
    expect(result.documentXml).toContain("Cell 1");
    expect(result.documentXml).toContain("Cell 2");
  });
});
