import { describe, it, expect } from "vitest";
import { parseXml } from "../../src/xml/parser.js";
import { normalizeParagraph } from "../../src/normalizer/paragraph.js";
import { normalizeTable } from "../../src/normalizer/table.js";
import { normalizeHeaderFooter } from "../../src/normalizer/header-footer.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { readDocx } from "../../src/zip/reader.js";
import {
  createSimpleDocx,
  createFormattedDocx,
  createTableDocx,
  createHeaderFooterDocx,
  createSplitRunDocx,
} from "../fixtures/create-docx.js";
import { findChild, findChildren } from "../../src/xml/helpers.js";

// =============================================================================
// Paragraph normalizer
// =============================================================================

describe("normalizeParagraph", () => {
  it("should normalize a simple paragraph with one run", () => {
    const xml = `<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>`;
    const parsed = parseXml(xml);
    const pNode = parsed[0]!;

    const result = normalizeParagraph(pNode, "p0");

    expect(result.kind).toBe("paragraph");
    expect(result.id).toBe("p0");
    expect(result.text).toBe("Hello World");
    expect(result.runs.length).toBe(1);
    expect(result.runs[0]!.text).toBe("Hello World");
    expect(result.runs[0]!.startOffset).toBe(0);
    expect(result.runs[0]!.endOffset).toBe(11);
  });

  it("should normalize a paragraph with multiple runs", () => {
    const xml = `<w:p>
      <w:r><w:t xml:space="preserve">This is </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:t xml:space="preserve"> text.</w:t></w:r>
    </w:p>`;
    const parsed = parseXml(xml);
    const pNode = parsed[0]!;

    const result = normalizeParagraph(pNode, "p0");

    expect(result.text).toBe("This is bold text.");
    expect(result.runs.length).toBe(3);

    // First run
    expect(result.runs[0]!.text).toBe("This is ");
    expect(result.runs[0]!.startOffset).toBe(0);
    expect(result.runs[0]!.endOffset).toBe(8);

    // Second run (bold)
    expect(result.runs[1]!.text).toBe("bold");
    expect(result.runs[1]!.startOffset).toBe(8);
    expect(result.runs[1]!.endOffset).toBe(12);
    expect(result.runs[1]!.formatting.bold).toBe(true);

    // Third run
    expect(result.runs[2]!.text).toBe(" text.");
    expect(result.runs[2]!.startOffset).toBe(12);
    expect(result.runs[2]!.endOffset).toBe(18);
  });

  it("should extract formatting from runs", () => {
    const xml = `<w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
          <w:u w:val="single"/>
          <w:color w:val="FF0000"/>
          <w:sz w:val="28"/>
          <w:rFonts w:ascii="Arial"/>
          <w:highlight w:val="yellow"/>
          <w:strike/>
        </w:rPr>
        <w:t>Formatted</w:t>
      </w:r>
    </w:p>`;
    const parsed = parseXml(xml);
    const pNode = parsed[0]!;

    const result = normalizeParagraph(pNode, "p0");
    const fmt = result.runs[0]!.formatting;

    expect(fmt.bold).toBe(true);
    expect(fmt.italic).toBe(true);
    expect(fmt.underline).toBe("single");
    expect(fmt.color).toBe("FF0000");
    expect(fmt.fontSize).toBe(28);
    expect(fmt.fontFamily).toBe("Arial");
    expect(fmt.highlight).toBe("yellow");
    expect(fmt.strikethrough).toBe(true);
  });

  it("should handle bold explicitly off", () => {
    const xml = `<w:p><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>Not bold</w:t></w:r></w:p>`;
    const parsed = parseXml(xml);
    const result = normalizeParagraph(parsed[0]!, "p0");
    expect(result.runs[0]!.formatting.bold).toBe(false);
  });

  it("should extract paragraph style", () => {
    const xml = `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>`;
    const parsed = parseXml(xml);
    const result = normalizeParagraph(parsed[0]!, "p0");
    expect(result.style).toBe("Heading1");
  });

  it("should extract list info", () => {
    const xml = `<w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="1"/>
          <w:numId w:val="3"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>List item</w:t></w:r>
    </w:p>`;
    const parsed = parseXml(xml);
    const result = normalizeParagraph(parsed[0]!, "p0");

    expect(result.listInfo).toBeDefined();
    expect(result.listInfo!.numId).toBe("3");
    expect(result.listInfo!.level).toBe(1);
  });

  it("should skip runs with no text (e.g., break-only runs)", () => {
    const xml = `<w:p>
      <w:r><w:t>Before</w:t></w:r>
      <w:r><w:br/></w:r>
      <w:r><w:t>After</w:t></w:r>
    </w:p>`;
    const parsed = parseXml(xml);
    const result = normalizeParagraph(parsed[0]!, "p0");

    expect(result.text).toBe("BeforeAfter");
    expect(result.runs.length).toBe(2);
  });

  it("should hold xmlRef pointers to original FXP nodes", () => {
    const xml = `<w:p><w:r><w:t>Test</w:t></w:r></w:p>`;
    const parsed = parseXml(xml);
    const pNode = parsed[0]!;

    const result = normalizeParagraph(pNode, "p0");

    // xmlRef should point to the original node (same reference)
    expect(result.xmlRef).toBe(pNode);
    // Run xmlRef should point to the w:r node
    const runs = findChildren(pNode, "w:r");
    expect(result.runs[0]!.xmlRef).toBe(runs[0]);
  });

  it("should handle empty paragraphs", () => {
    const xml = `<w:p></w:p>`;
    const parsed = parseXml(xml);
    const result = normalizeParagraph(parsed[0]!, "p0");

    expect(result.text).toBe("");
    expect(result.runs.length).toBe(0);
  });

  it("should handle paragraphs with only pPr (no runs)", () => {
    const xml = `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>`;
    const parsed = parseXml(xml);
    const result = normalizeParagraph(parsed[0]!, "p0");

    expect(result.text).toBe("");
    expect(result.runs.length).toBe(0);
    expect(result.style).toBe("Normal");
  });
});

// =============================================================================
// Table normalizer
// =============================================================================

describe("normalizeTable", () => {
  it("should normalize a simple 2x2 table", () => {
    const xml = `<w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4788"/><w:gridCol w:w="4788"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const parsed = parseXml(xml);
    const result = normalizeTable(parsed[0]!, "tbl0");

    expect(result.kind).toBe("table");
    expect(result.id).toBe("tbl0");
    expect(result.style).toBe("TableGrid");
    expect(result.rows.length).toBe(2);

    // First row
    expect(result.rows[0]!.cells.length).toBe(2);
    expect(result.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe("A1");
    expect(result.rows[0]!.cells[1]!.paragraphs[0]!.text).toBe("B1");

    // Second row
    expect(result.rows[1]!.cells[0]!.paragraphs[0]!.text).toBe("A2");
    expect(result.rows[1]!.cells[1]!.paragraphs[0]!.text).toBe("B2");
  });

  it("should assign correct IDs to inner paragraphs", () => {
    const xml = `<w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const parsed = parseXml(xml);
    const result = normalizeTable(parsed[0]!, "tbl2");

    expect(result.rows[0]!.cells[0]!.paragraphs[0]!.id).toBe("tbl2.r0.c0.p0");
  });

  it("should handle cells with multiple paragraphs", () => {
    const xml = `<w:tbl>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>First para</w:t></w:r></w:p>
          <w:p><w:r><w:t>Second para</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;
    const parsed = parseXml(xml);
    const result = normalizeTable(parsed[0]!, "tbl0");

    const cell = result.rows[0]!.cells[0]!;
    expect(cell.paragraphs.length).toBe(2);
    expect(cell.paragraphs[0]!.text).toBe("First para");
    expect(cell.paragraphs[1]!.text).toBe("Second para");
    expect(cell.paragraphs[0]!.id).toBe("tbl0.r0.c0.p0");
    expect(cell.paragraphs[1]!.id).toBe("tbl0.r0.c0.p1");
  });

  it("should hold xmlRef pointers", () => {
    const xml = `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const parsed = parseXml(xml);
    const tblNode = parsed[0]!;
    const result = normalizeTable(tblNode, "tbl0");

    expect(result.xmlRef).toBe(tblNode);
  });
});

// =============================================================================
// Header/footer normalizer
// =============================================================================

describe("normalizeHeaderFooter", () => {
  it("should normalize a header", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Header text</w:t></w:r></w:p>
</w:hdr>`;
    const parsed = parseXml(xml);
    const result = normalizeHeaderFooter(parsed, "header", "default", "rId1");

    expect(result.type).toBe("header");
    expect(result.variant).toBe("default");
    expect(result.rId).toBe("rId1");
    expect(result.paragraphs.length).toBe(1);
    expect(result.paragraphs[0]!.text).toBe("Header text");
    expect(result.paragraphs[0]!.id).toBe("hdr.default.p0");
  });

  it("should normalize a footer", () => {
    const xml = `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Page 1</w:t></w:r></w:p>
</w:ftr>`;
    const parsed = parseXml(xml);
    const result = normalizeHeaderFooter(parsed, "footer", "default", "rId2");

    expect(result.type).toBe("footer");
    expect(result.paragraphs[0]!.text).toBe("Page 1");
    expect(result.paragraphs[0]!.id).toBe("ftr.default.p0");
  });
});

// =============================================================================
// Document normalizer (integration)
// =============================================================================

describe("normalizeDocument", () => {
  it("should normalize a simple document", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    expect(doc.body.length).toBe(1);
    expect(doc.body[0]!.kind).toBe("paragraph");
    const para = doc.body[0] as any;
    expect(para.text).toBe("Hello World");
    expect(para.id).toBe("p0");
  });

  it("should normalize a formatted document with multiple paragraphs", async () => {
    const buffer = await createFormattedDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    expect(doc.body.length).toBe(3);

    // Heading paragraph
    const heading = doc.body[0] as any;
    expect(heading.text).toBe("Document Title");
    expect(heading.style).toBe("Heading1");
    expect(heading.runs[0].formatting.bold).toBe(true);

    // Multi-run paragraph
    const multiRun = doc.body[1] as any;
    expect(multiRun.text).toBe("This is bold and italic text.");

    // Red large text
    const redText = doc.body[2] as any;
    expect(redText.text).toBe("Red large text");
    expect(redText.runs[0].formatting.color).toBe("FF0000");
    expect(redText.runs[0].formatting.fontSize).toBe(28);
  });

  it("should normalize a document with a table", async () => {
    const buffer = await createTableDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // Should have: paragraph, table, paragraph
    expect(doc.body.length).toBe(3);
    expect(doc.body[0]!.kind).toBe("paragraph");
    expect(doc.body[1]!.kind).toBe("table");
    expect(doc.body[2]!.kind).toBe("paragraph");

    const table = doc.body[1] as any;
    expect(table.id).toBe("tbl0");
    expect(table.rows.length).toBe(2);
    expect(table.rows[0].cells[0].paragraphs[0].text).toBe("Header A");
    expect(table.rows[1].cells[1].paragraphs[0].text).toBe("Cell 2");
  });

  it("should normalize headers and footers", async () => {
    const buffer = await createHeaderFooterDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    expect(doc.headers.size).toBe(1);
    expect(doc.footers.size).toBe(1);

    const header = [...doc.headers.values()][0]!;
    expect(header.paragraphs[0]!.text).toBe("Header text");
    expect(header.paragraphs[0]!.id).toBe("hdr.default.p0");

    const footer = [...doc.footers.values()][0]!;
    expect(footer.paragraphs[0]!.text).toBe("Footer text");
    expect(footer.paragraphs[0]!.id).toBe("ftr.default.p0");
  });

  it("should normalize a document with split runs", async () => {
    const buffer = await createSplitRunDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const para = doc.body[0] as any;
    expect(para.text).toBe("This agreement is valid and enforceable");
    expect(para.runs.length).toBe(2);
    expect(para.runs[0].text).toBe("This agreement is va");
    expect(para.runs[0].startOffset).toBe(0);
    expect(para.runs[0].endOffset).toBe(20);
    expect(para.runs[1].text).toBe("lid and enforceable");
    expect(para.runs[1].startOffset).toBe(20);
    expect(para.runs[1].endOffset).toBe(39);
    expect(para.runs[1].formatting.bold).toBe(true);
  });

  it("should hold xmlParts with mutable FXP objects", async () => {
    const buffer = await createSimpleDocx("Test");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    expect(doc.xmlParts.document).toBeDefined();
    expect(Array.isArray(doc.xmlParts.document)).toBe(true);
  });
});
