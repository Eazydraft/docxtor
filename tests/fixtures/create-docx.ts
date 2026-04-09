/**
 * Test helper: creates minimal DOCX buffers for testing.
 *
 * These are real, valid DOCX files built using JSZip + raw OOXML.
 * We don't use a DOCX library here because we need to control
 * the exact XML structure for testing our parser/normalizer.
 */

import JSZip from "jszip";

/** Build content types XML, conditionally including styles.xml override. */
function buildContentTypes(options: { stylesXml?: string }): string {
  let overrides = `
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`;

  if (options.stylesXml) {
    overrides += `
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>${overrides}
</Types>`;
}

/** Top-level relationships. */
const TOP_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

/** Document relationships (references headers/footers). */
const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;

/**
 * Create a minimal DOCX buffer with a single paragraph.
 */
export async function createSimpleDocx(text: string): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>${escapeXml(text)}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with multiple paragraphs, some with formatting.
 */
export async function createFormattedDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Document Title</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">This is </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>bold</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> and </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>italic</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:color w:val="FF0000"/>
          <w:sz w:val="28"/>
        </w:rPr>
        <w:t>Red large text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with a table.
 */
export async function createTableDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Before table</w:t>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="5000" w:type="pct"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="4788"/>
        <w:gridCol w:w="4788"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="4788" w:type="dxa"/>
          </w:tcPr>
          <w:p>
            <w:r>
              <w:rPr><w:b/></w:rPr>
              <w:t>Header A</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="4788" w:type="dxa"/>
          </w:tcPr>
          <w:p>
            <w:r>
              <w:rPr><w:b/></w:rPr>
              <w:t>Header B</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="4788" w:type="dxa"/>
          </w:tcPr>
          <w:p>
            <w:r>
              <w:t>Cell 1</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="4788" w:type="dxa"/>
          </w:tcPr>
          <w:p>
            <w:r>
              <w:t>Cell 2</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:t>After table</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with headers and footers.
 */
export async function createHeaderFooterDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Document body text</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId1"/>
      <w:footerReference w:type="default" r:id="rId2"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:t>Header text</w:t>
    </w:r>
  </w:p>
</w:hdr>`;

  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:t>Footer text</w:t>
    </w:r>
  </w:p>
</w:ftr>`;

  return buildDocx({
    documentXml,
    headerXml,
    footerXml,
    includeDocRels: true,
  });
}

/**
 * Create a DOCX with a paragraph containing multiple runs
 * that split a word across runs (the hard case for patching).
 */
export async function createSplitRunDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">This agreement is va</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>lid and enforceable</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with hyperlinks in a paragraph.
 *
 * Structure: "Visit our website for details."
 * Where "our website" is a hyperlink with r:id="rId10".
 */
export async function createHyperlinkDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">Visit </w:t>
      </w:r>
      <w:hyperlink r:id="rId10">
        <w:r>
          <w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr>
          <w:t>our website</w:t>
        </w:r>
      </w:hyperlink>
      <w:r>
        <w:t xml:space="preserve"> for details.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with a nested table (table inside a table cell).
 */
export async function createNestedTableDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
      </w:tblPr>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Outer cell text</w:t>
            </w:r>
          </w:p>
          <w:tbl>
            <w:tblPr>
              <w:tblStyle w:val="TableGrid"/>
            </w:tblPr>
            <w:tr>
              <w:tc>
                <w:p>
                  <w:r>
                    <w:t>Inner cell A</w:t>
                  </w:r>
                </w:p>
              </w:tc>
              <w:tc>
                <w:p>
                  <w:r>
                    <w:t>Inner cell B</w:t>
                  </w:r>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Regular cell</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

// =============================================================================
// Internal helpers
// =============================================================================

interface BuildDocxOptions {
  documentXml: string;
  headerXml?: string;
  footerXml?: string;
  stylesXml?: string;
  includeDocRels?: boolean;
}

async function buildDocx(options: BuildDocxOptions): Promise<Buffer> {
  const zip = new JSZip();

  zip.file("[Content_Types].xml", buildContentTypes(options));
  zip.file("_rels/.rels", TOP_RELS);
  zip.file("word/document.xml", options.documentXml);

  if (options.includeDocRels || options.headerXml || options.footerXml) {
    zip.file("word/_rels/document.xml.rels", DOC_RELS);
  }

  if (options.headerXml) {
    zip.file("word/header1.xml", options.headerXml);
  }

  if (options.footerXml) {
    zip.file("word/footer1.xml", options.footerXml);
  }

  if (options.stylesXml) {
    zip.file("word/styles.xml", options.stylesXml);
  }

  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  }) as Promise<Buffer>;
}

/**
 * A minimal styles.xml with common styles for testing.
 */
export const TEST_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="240" w:after="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
      <w:color w:val="2F5496"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="26"/>
      <w:color w:val="2F5496"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:sz w:val="56"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Strong">
    <w:name w:val="Strong"/>
    <w:rPr>
      <w:b/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Emphasis">
    <w:name w:val="Emphasis"/>
    <w:rPr>
      <w:i/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr>
      <w:color w:val="0563C1"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
  </w:style>
</w:styles>`;

/**
 * Create a DOCX with styles.xml included for testing style patches.
 *
 * Contains two paragraphs: one with Heading1 style, one plain.
 * Includes multi-run formatting on the second paragraph (bold/italic).
 */
export async function createStyledDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="2F5496"/></w:rPr>
        <w:t>Document Title</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">This is </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>bold</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> and </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>italic</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> text here.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Plain paragraph for testing.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml, stylesXml: TEST_STYLES_XML });
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Create a DOCX with a single paragraph containing a w:br (line break).
 *
 * Structure: a single w:r with "Hello", w:br, "World"
 * Normalized text: "Hello\nWorld"
 */
export async function createLineBreakDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello</w:t>
        <w:br/>
        <w:t>World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with a page break between two paragraphs.
 *
 * Structure:
 * - Paragraph: "Before page break"
 * - Paragraph with only a page break (w:br w:type="page")
 * - Paragraph: "After page break"
 *
 * The middle paragraph should be normalized as a NormalizedPageBreak.
 */
export async function createPageBreakDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Before page break</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:br w:type="page"/>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>After page break</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX where a paragraph contains both text and a page break.
 *
 * Structure: a single paragraph with "Some text" followed by w:br w:type="page"
 * This should be normalized as a NormalizedParagraph with hasPageBreak: true.
 */
export async function createInlinePageBreakDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Text with page break</w:t>
        <w:br w:type="page"/>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Next page content</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with a w:lastRenderedPageBreak inside a run that also has text.
 *
 * This is the most common page break form in real Word documents —
 * Word inserts it as a rendering hint to mark where the page broke
 * when the document was last paginated.
 *
 * Structure:
 * - Paragraph: "First page content"
 * - Paragraph: w:lastRenderedPageBreak + "Second page content" (inline)
 */
export async function createLastRenderedPageBreakDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First page content</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:lastRenderedPageBreak/>
        <w:t>Second page content</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with a standalone w:lastRenderedPageBreak paragraph
 * (a paragraph whose only run contains just the break hint, no text).
 *
 * This should be normalized as a NormalizedPageBreak (page-break-only paragraph).
 */
export async function createLastRenderedPageBreakOnlyDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Before break</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:lastRenderedPageBreak/>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>After break</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with w:pageBreakBefore in a paragraph's properties.
 *
 * This is a paragraph-level setting that forces the paragraph to start
 * on a new page. It should trigger hasPageBreak on the paragraph.
 *
 * Structure:
 * - Paragraph: "First page content"
 * - Paragraph with w:pageBreakBefore: "Starts on new page"
 */
export async function createPageBreakBeforeDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>First page content</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pageBreakBefore/>
      </w:pPr>
      <w:r>
        <w:t>Starts on new page</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with w:pageBreakBefore w:val="0" (disabled).
 *
 * This should NOT trigger hasPageBreak since the property is explicitly off.
 */
export async function createPageBreakBeforeDisabledDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pageBreakBefore w:val="0"/>
      </w:pPr>
      <w:r>
        <w:t>No page break here</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with a page break between two text runs in the same paragraph.
 *
 * Structure: "Before break" + w:br w:type="page" + "After break" in a single paragraph.
 * The page break offset should be 12 (length of "Before break").
 */
export async function createMidTextPageBreakDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Before break</w:t>
      </w:r>
      <w:r>
        <w:br w:type="page"/>
      </w:r>
      <w:r>
        <w:t>After break</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}

/**
 * Create a DOCX with both a tracked insertion and a mid-text page break
 * in the same paragraph.
 *
 * Structure:
 *   "Hello " + w:ins["world"] + w:br w:type="page" + "After break"
 *
 * The normalized text is "Hello worldAfter break" with:
 *   - An insert revision for "world" at [6, 11)
 *   - pageBreakOffset = 11 (after "Hello world", before "After break")
 *
 * When both includeRevisions and includePageBreaks are enabled, the
 * markdown should read:
 *   "Hello {+world+} [PAGE BREAK] After break"
 * (the [PAGE BREAK] marker must land after the revision annotation).
 */
export async function createRevisionWithPageBreakDocx(): Promise<Buffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">Hello </w:t>
      </w:r>
      <w:ins w:id="1" w:author="Test" w:date="2025-01-01T00:00:00Z">
        <w:r>
          <w:t>world</w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:br w:type="page"/>
      </w:r>
      <w:r>
        <w:t>After break</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  return buildDocx({ documentXml });
}
