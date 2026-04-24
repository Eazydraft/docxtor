import { describe, it, expect } from "vitest";
import { Docxtor } from "../../src/index.js";
import { parseFlatOpc } from "../../src/flat-opc/parser.js";
import { serializeFlatOpc } from "../../src/flat-opc/serializer.js";

// =============================================================================
// Minimal Flat OPC fixtures
// =============================================================================

const MINIMAL_FLAT_OPC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:r>
              <w:t>Hello World</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r>
              <w:t>Second paragraph</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

const FLAT_OPC_WITH_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
            <w:r>
              <w:t>Title</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r>
              <w:t>Body text here</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
    <pkg:xmlData>
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Heading1">
          <w:name w:val="heading 1"/>
        </w:style>
      </w:styles>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

const FLAT_OPC_WITH_HEADERS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:r>
              <w:t>Document body</w:t>
            </w:r>
          </w:p>
          <w:sectPr>
            <w:headerReference w:type="default" r:id="rId1"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/header1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml">
    <pkg:xmlData>
      <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
          <w:r>
            <w:t>Header text</w:t>
          </w:r>
        </w:p>
      </w:hdr>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/[Content_Types].xml" pkg:contentType="application/vnd.openxmlformats-package.content-types+xml">
    <pkg:xmlData>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
      </Types>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

// =============================================================================
// Parser tests
// =============================================================================

describe("Flat OPC parser", () => {
  it("extracts document.xml from minimal Flat OPC", () => {
    const { entries, parts } = parseFlatOpc(MINIMAL_FLAT_OPC);

    expect(entries.documentXml).toContain("Hello World");
    expect(entries.documentXml).toContain("Second paragraph");
    expect(parts.length).toBe(1);
  });

  it("extracts styles.xml when present", () => {
    const { entries, parts } = parseFlatOpc(FLAT_OPC_WITH_STYLES);

    expect(entries.documentXml).toContain("Title");
    expect(entries.stylesXml).toContain("Heading1");
    expect(parts.length).toBe(2);
  });

  it("extracts header files and document rels", () => {
    const { entries } = parseFlatOpc(FLAT_OPC_WITH_HEADERS);

    expect(entries.documentXml).toContain("Document body");
    expect(entries.headers.size).toBe(1);
    expect(entries.headers.get("word/header1.xml")).toContain("Header text");
    expect(entries.documentRels).toContain("rId1");
    expect(entries.contentTypes).toContain("PartName");
  });

  it("throws on missing document.xml", () => {
    const invalid = `<?xml version="1.0"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/styles.xml" pkg:contentType="text/xml">
    <pkg:xmlData><w:styles/></pkg:xmlData>
  </pkg:part>
</pkg:package>`;

    expect(() => parseFlatOpc(invalid)).toThrow("missing word/document.xml");
  });

  it("handles binary parts without crashing", () => {
    const withBinary = `<?xml version="1.0"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">
    <pkg:binaryData>iVBORw0KGgoAAAANSUhEUg==</pkg:binaryData>
  </pkg:part>
</pkg:package>`;

    const { entries, parts } = parseFlatOpc(withBinary);
    expect(entries.documentXml).toContain("Text");
    expect(parts.length).toBe(2);
    expect(parts[1]!.binaryData).toBe("iVBORw0KGgoAAAANSUhEUg==");
  });
});

// =============================================================================
// Serializer tests
// =============================================================================

describe("Flat OPC serializer", () => {
  it("round-trips a parsed Flat OPC", () => {
    const { parts } = parseFlatOpc(MINIMAL_FLAT_OPC);
    const output = serializeFlatOpc(parts);

    expect(output).toContain("pkg:package");
    expect(output).toContain("Hello World");
    expect(output).toContain("Second paragraph");
    expect(output).toContain("pkg:name");
    expect(output).toContain("pkg:contentType");
  });

  it("preserves binary parts", () => {
    const withBinary = `<?xml version="1.0"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">
    <pkg:binaryData>abc123==</pkg:binaryData>
  </pkg:part>
</pkg:package>`;

    const { parts } = parseFlatOpc(withBinary);
    const output = serializeFlatOpc(parts);

    expect(output).toContain("abc123==");
    expect(output).toContain("pkg:binaryData");
  });
});

// =============================================================================
// Docxtor.loadFlatOpc integration tests
// =============================================================================

describe("Docxtor.loadFlatOpc", () => {
  it("loads and normalizes a minimal Flat OPC document", () => {
    const doc = Docxtor.loadFlatOpc(MINIMAL_FLAT_OPC);

    const paras = doc.getParagraphs();
    expect(paras.length).toBe(2);
    expect(paras[0]!.text).toBe("Hello World");
    expect(paras[1]!.text).toBe("Second paragraph");
    expect(paras[0]!.id).toBe("p0");
    expect(paras[1]!.id).toBe("p1");
  });

  it("projects to markdown", () => {
    const doc = Docxtor.loadFlatOpc(MINIMAL_FLAT_OPC);
    const md = doc.project("markdown");

    expect(md).toContain("[p0]");
    expect(md).toContain("Hello World");
    expect(md).toContain("[p1]");
    expect(md).toContain("Second paragraph");
  });

  it("projects to JSON", () => {
    const doc = Docxtor.loadFlatOpc(MINIMAL_FLAT_OPC);
    const json = doc.project("json");

    expect(json.body.length).toBe(2);
    expect(json.body[0]!.id).toBe("p0");
  });

  it("applies patches and exports to Flat OPC", () => {
    const doc = Docxtor.loadFlatOpc(MINIMAL_FLAT_OPC);

    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "World", replace: "Docxtor" },
    ]);

    expect(results[0]!.success).toBe(true);

    const output = doc.toFlatOpc();
    expect(output).toContain("pkg:package");
    expect(output).toContain("Docxtor");
    expect(output).not.toContain(">Hello World<");
  });

  it("applies patches with revision tracking", () => {
    const doc = Docxtor.loadFlatOpc(MINIMAL_FLAT_OPC);

    const results = doc.applyPatches(
      [
        {
          type: "replace_text",
          target: "p0",
          find: "World",
          replace: "Docxtor",
        },
      ],
      { trackRevisions: true, revisionAuthor: "AI Assistant" },
    );

    expect(results[0]!.success).toBe(true);

    const output = doc.toFlatOpc();
    expect(output).toContain("w:del");
    expect(output).toContain("w:ins");
    expect(output).toContain("Docxtor");
  });

  it("handles document with styles", () => {
    const doc = Docxtor.loadFlatOpc(FLAT_OPC_WITH_STYLES);

    const paras = doc.getParagraphs();
    expect(paras[0]!.style).toBe("Heading1");
    expect(paras[0]!.text).toBe("Title");
  });

  it("isFromFlatOpc returns true", () => {
    const doc = Docxtor.loadFlatOpc(MINIMAL_FLAT_OPC);
    expect(doc.isFromFlatOpc()).toBe(true);
  });

  it("toFlatOpc round-trips through parse again", () => {
    const doc = Docxtor.loadFlatOpc(MINIMAL_FLAT_OPC);

    doc.applyPatches([
      {
        type: "replace_text",
        target: "p1",
        find: "Second",
        replace: "Modified",
      },
    ]);

    const flatOpc = doc.toFlatOpc();

    // Re-parse the output
    const doc2 = Docxtor.loadFlatOpc(flatOpc);
    const paras = doc2.getParagraphs();
    expect(paras[0]!.text).toBe("Hello World");
    expect(paras[1]!.text).toBe("Modified paragraph");
  });
});

// =============================================================================
// Error cases
// =============================================================================

describe("Flat OPC error handling", () => {
  it("toFlatOpc throws for ZIP-loaded documents", async () => {
    const doc = await Docxtor.create();
    expect(() => doc.toFlatOpc()).toThrow("loadFlatOpc");
  });

  it("isFromFlatOpc returns false for ZIP-loaded documents", async () => {
    const doc = await Docxtor.create();
    expect(doc.isFromFlatOpc()).toBe(false);
  });
});
