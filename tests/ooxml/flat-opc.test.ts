import { describe, it, expect } from "vitest";
import {
  isFlatOpc,
  parseFlatOpc,
  buildFlatOpc,
  buildDocumentFragment,
  opcContentType,
  opcPathToZipPath,
  zipPathToOpcPath,
  type FlatOpcPart,
} from "../../src/ooxml/flat-opc.js";
import { Docxtor } from "../../src/index.js";
import { parseXml } from "../../src/xml/parser.js";
import { serializeXml } from "../../src/xml/serializer.js";

// =============================================================================
// Shared fixtures
// =============================================================================

const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

const HEADER_XML = `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Page header</w:t></w:r></w:p></w:hdr>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`;

const DOC_XML_WITH_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:body></w:document>`;

/** Build a minimal Flat OPC XML string with a document.xml and optional extra parts. */
function makeFlatOpc(
  documentXml: string = DOCUMENT_XML,
  extraParts: Array<{
    name: string;
    contentType: string;
    content: string;
    isBinary?: boolean;
    padding?: string;
  }> = [],
): string {
  const docPart = `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>${documentXml}</pkg:xmlData></pkg:part>`;
  const rest = extraParts
    .map((p) => {
      const tag = p.isBinary ? "pkg:binaryData" : "pkg:xmlData";
      const paddingAttr = p.padding ? ` pkg:padding="${p.padding}"` : "";
      return `<pkg:part pkg:name="${p.name}" pkg:contentType="${p.contentType}"${paddingAttr}><${tag}>${p.content}</${tag}></pkg:part>`;
    })
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><?mso-application progid="Word.Document"?><pkg:package xmlns:pkg="${PKG_NS}">${docPart}${rest}</pkg:package>`;
}

// =============================================================================
// isFlatOpc
// =============================================================================

describe("isFlatOpc", () => {
  it("returns true for a pkg:package root element", () => {
    expect(isFlatOpc(makeFlatOpc())).toBe(true);
  });

  it("returns true when xmlPackage namespace string appears in the head", () => {
    expect(isFlatOpc(`<?xml version="1.0"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>`)).toBe(true);
  });

  it("returns false for raw w:document XML", () => {
    expect(isFlatOpc(DOCUMENT_XML)).toBe(false);
  });

  it("returns false for plain text", () => {
    expect(isFlatOpc("Hello World")).toBe(false);
  });

  it("returns false for empty string", () => {
    expect(isFlatOpc("")).toBe(false);
  });
});

// =============================================================================
// parseFlatOpc
// =============================================================================

describe("parseFlatOpc", () => {
  it("parses a single XML part", () => {
    const parts = parseFlatOpc(makeFlatOpc());

    expect(parts).toHaveLength(1);
    expect(parts[0]!.name).toBe("/word/document.xml");
    expect(parts[0]!.contentType).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    );
    expect(parts[0]!.isBinary).toBe(false);
    expect(parts[0]!.content).toContain("Hello World");
    expect(parts[0]!.content).toContain("w:document");
  });

  it("parses multiple XML parts", () => {
    const xml = makeFlatOpc(DOCUMENT_XML, [
      {
        name: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        content: STYLES_XML,
      },
      {
        name: "/word/_rels/document.xml.rels",
        contentType: "application/vnd.openxmlformats-package.relationships+xml",
        content: RELS_XML,
      },
    ]);

    const parts = parseFlatOpc(xml);

    expect(parts).toHaveLength(3);
    const names = parts.map((p) => p.name);
    expect(names).toContain("/word/document.xml");
    expect(names).toContain("/word/styles.xml");
    expect(names).toContain("/word/_rels/document.xml.rels");
  });

  it("preserves inner XML content verbatim", () => {
    const parts = parseFlatOpc(makeFlatOpc());
    // The content should be parseable as XML
    const fxp = parseXml(parts[0]!.content);
    expect(serializeXml(fxp)).toContain("Hello World");
  });

  it("detects binary parts (pkg:binaryData)", () => {
    // Valid base64 — JSZip validates during storage
    const b64 = Buffer.from("fake png data").toString("base64");
    const xml = makeFlatOpc(DOCUMENT_XML, [
      {
        name: "/word/media/image1.png",
        contentType: "image/png",
        content: b64,
        isBinary: true,
      },
    ]);

    const parts = parseFlatOpc(xml);
    const imgPart = parts.find((p) => p.name === "/word/media/image1.png");

    expect(imgPart).toBeDefined();
    expect(imgPart!.isBinary).toBe(true);
    expect(imgPart!.content).toBe(b64);
  });

  it("preserves the padding attribute", () => {
    const xml = `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" pkg:padding="512"><pkg:xmlData>${DOCUMENT_XML}</pkg:xmlData></pkg:part></pkg:package>`;
    const parts = parseFlatOpc(xml);

    expect(parts[0]!.padding).toBe("512");
  });

  it("returns empty array for non-flat-OPC XML", () => {
    expect(parseFlatOpc(DOCUMENT_XML)).toEqual([]);
    expect(parseFlatOpc("")).toEqual([]);
    expect(parseFlatOpc("not xml at all")).toEqual([]);
  });

  it("parses header and footer parts by path", () => {
    const xml = makeFlatOpc(DOCUMENT_XML, [
      {
        name: "/word/header1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        content: HEADER_XML,
      },
    ]);

    const parts = parseFlatOpc(xml);
    const hdrPart = parts.find((p) => p.name === "/word/header1.xml");

    expect(hdrPart).toBeDefined();
    expect(hdrPart!.content).toContain("Page header");
  });
});

// =============================================================================
// buildFlatOpc
// =============================================================================

describe("buildFlatOpc", () => {
  const sampleParts: FlatOpcPart[] = [
    {
      name: "/word/document.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      content: DOCUMENT_XML,
      isBinary: false,
    },
  ];

  it("produces a valid pkg:package wrapper", () => {
    const result = buildFlatOpc(sampleParts);

    expect(result).toContain("pkg:package");
    expect(result).toContain(PKG_NS);
  });

  it("wraps XML parts in pkg:xmlData", () => {
    const result = buildFlatOpc(sampleParts);

    expect(result).toContain("pkg:xmlData");
    expect(result).toContain("Hello World");
  });

  it("wraps binary parts in pkg:binaryData", () => {
    const b64 = Buffer.from("binary content").toString("base64");
    const parts: FlatOpcPart[] = [
      {
        name: "/word/media/img.png",
        contentType: "image/png",
        content: b64,
        isBinary: true,
      },
    ];

    const result = buildFlatOpc(parts);

    expect(result).toContain("pkg:binaryData");
    expect(result).toContain(b64);
    expect(result).not.toContain("pkg:xmlData");
  });

  it("includes pkg:name and pkg:contentType on each part", () => {
    const result = buildFlatOpc(sampleParts);

    expect(result).toContain('pkg:name="/word/document.xml"');
    expect(result).toContain("pkg:contentType=");
  });

  it("includes pkg:padding when set", () => {
    const parts: FlatOpcPart[] = [
      { ...sampleParts[0]!, padding: "512" },
    ];

    const result = buildFlatOpc(parts);

    expect(result).toContain('pkg:padding="512"');
  });

  it("omits pkg:padding when not set", () => {
    const result = buildFlatOpc(sampleParts);

    expect(result).not.toContain("pkg:padding");
  });

  it("applies overrides to matching parts", () => {
    const overrides = new Map([["/word/document.xml", "<w:document>REPLACED</w:document>"]]);
    const result = buildFlatOpc(sampleParts, overrides);

    expect(result).toContain("REPLACED");
    expect(result).not.toContain("Hello World");
  });

  it("preserves unaffected parts when overrides are applied", () => {
    const parts: FlatOpcPart[] = [
      ...sampleParts,
      {
        name: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        content: STYLES_XML,
        isBinary: false,
      },
    ];
    const overrides = new Map([["/word/document.xml", "<w:document>CHANGED</w:document>"]]);
    const result = buildFlatOpc(parts, overrides);

    expect(result).toContain("CHANGED");
    expect(result).toContain("w:styles"); // styles part preserved
  });

  it("round-trips through parseFlatOpc", () => {
    const original = makeFlatOpc(DOCUMENT_XML, [
      {
        name: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        content: STYLES_XML,
      },
    ]);

    const parts = parseFlatOpc(original);
    const rebuilt = buildFlatOpc(parts);

    // Both should have the same parts
    const rebuiltParts = parseFlatOpc(rebuilt);
    expect(rebuiltParts).toHaveLength(parts.length);
    expect(rebuiltParts.map((p) => p.name).sort()).toEqual(
      parts.map((p) => p.name).sort(),
    );
  });

  it("escapes special characters in part names and content types", () => {
    const parts: FlatOpcPart[] = [
      {
        name: '/word/document&special".xml',
        contentType: "application/xml",
        content: "",
        isBinary: false,
      },
    ];

    const result = buildFlatOpc(parts);
    // The attribute value should be escaped
    expect(result).toContain("&amp;");
    expect(result).toContain("&quot;");
    // The raw unescaped characters should not appear in attribute context
    expect(result).not.toContain('="' + '/word/document&special".xml' + '"');
  });
});

// =============================================================================
// buildDocumentFragment
// =============================================================================

describe("buildDocumentFragment", () => {
  it("wraps FXP nodes in a w:document/w:body shell", () => {
    const fxp = parseXml(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Test</w:t></w:r></w:p>`);
    const docFxp = parseXml(DOCUMENT_XML);

    const result = buildDocumentFragment(fxp, docFxp, serializeXml);

    expect(result).toContain("w:document");
    expect(result).toContain("w:body");
    expect(result).toContain("w:p");
    expect(result).toContain("Test");
  });

  it("copies namespace declarations from the document root", () => {
    const docWithNs = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body></w:body></w:document>`;
    const docFxp = parseXml(docWithNs);
    const bodyFxp = parseXml(`<w:p><w:r><w:t>Hi</w:t></w:r></w:p>`);

    const result = buildDocumentFragment(bodyFxp, docFxp, serializeXml);

    expect(result).toContain('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
    expect(result).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  });

  it("falls back to minimal w namespace when doc root has no attributes", () => {
    const result = buildDocumentFragment([], [], serializeXml);

    expect(result).toContain("xmlns:w=");
    expect(result).toContain("w:document");
    expect(result).toContain("w:body");
  });

  it("serializes multiple nodes into the body", () => {
    const docFxp = parseXml(DOCUMENT_XML);
    const p1 = parseXml(`<w:p><w:r><w:t>First</w:t></w:r></w:p>`);
    const p2 = parseXml(`<w:p><w:r><w:t>Second</w:t></w:r></w:p>`);

    const result = buildDocumentFragment([...p1, ...p2], docFxp, serializeXml);

    expect(result).toContain("First");
    expect(result).toContain("Second");
    // Both should appear inside w:body
    const bodyStart = result.indexOf("<w:body>");
    const firstIdx = result.indexOf("First", bodyStart);
    const secondIdx = result.indexOf("Second", bodyStart);
    expect(firstIdx).toBeGreaterThan(bodyStart);
    expect(secondIdx).toBeGreaterThan(firstIdx);
  });
});

// =============================================================================
// Path helpers and opcContentType
// =============================================================================

describe("path helpers", () => {
  it("opcPathToZipPath strips leading slash", () => {
    expect(opcPathToZipPath("/word/document.xml")).toBe("word/document.xml");
    expect(opcPathToZipPath("/[Content_Types].xml")).toBe("[Content_Types].xml");
  });

  it("opcPathToZipPath is a no-op when there is no leading slash", () => {
    expect(opcPathToZipPath("word/document.xml")).toBe("word/document.xml");
  });

  it("zipPathToOpcPath prepends a slash", () => {
    expect(zipPathToOpcPath("word/document.xml")).toBe("/word/document.xml");
    expect(zipPathToOpcPath("word/header1.xml")).toBe("/word/header1.xml");
  });

  it("zipPathToOpcPath is a no-op when slash already present", () => {
    expect(zipPathToOpcPath("/word/document.xml")).toBe("/word/document.xml");
  });
});

describe("opcContentType", () => {
  it("returns correct type for document.xml", () => {
    expect(opcContentType("/word/document.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    );
  });

  it("returns correct type for styles.xml", () => {
    expect(opcContentType("/word/styles.xml")).toContain("styles+xml");
  });

  it("returns header content type for header paths", () => {
    expect(opcContentType("/word/header1.xml")).toContain("header+xml");
    expect(opcContentType("/word/header99.xml")).toContain("header+xml");
  });

  it("returns footer content type for footer paths", () => {
    expect(opcContentType("/word/footer1.xml")).toContain("footer+xml");
  });

  it("falls back to application/xml for unknown paths", () => {
    expect(opcContentType("/word/unknown.xml")).toBe("application/xml");
  });
});

// =============================================================================
// Docxtor.fromOoxml — raw document XML
// =============================================================================

describe("Docxtor.fromOoxml — raw document XML", () => {
  it("loads a raw w:document XML string synchronously", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);

    expect(doc).toBeInstanceOf(Docxtor);
    expect(doc.getParagraphs()).toHaveLength(1);
    expect(doc.getParagraphs()[0]!.text).toBe("Hello World");
  });

  it("normalizes multiple paragraphs", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`;
    const doc = Docxtor.fromOoxml(xml);

    expect(doc.getParagraphs()).toHaveLength(2);
    expect(doc.getParagraphs()[0]!.text).toBe("First");
    expect(doc.getParagraphs()[1]!.text).toBe("Second");
  });

  it("supports applyPatches after loading from raw XML", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "Patched" },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("Patched");
  });

  it("returns a document whose toString() contains the XML body", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);

    expect(doc.toString()).toContain("w:document");
    expect(doc.toString()).toContain("Hello World");
  });

  it("respects LoadOptions.includeHeadersFooters = false", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML, { includeHeadersFooters: false });

    expect(doc.getHeaders().size).toBe(0);
    expect(doc.getFooters().size).toBe(0);
  });
});

// =============================================================================
// Docxtor.fromOoxml — Flat OPC
// =============================================================================

describe("Docxtor.fromOoxml — Flat OPC", () => {
  it("loads a Flat OPC package and normalizes the document body", () => {
    const doc = Docxtor.fromOoxml(makeFlatOpc());

    expect(doc).toBeInstanceOf(Docxtor);
    expect(doc.getParagraphs()).toHaveLength(1);
    expect(doc.getParagraphs()[0]!.text).toBe("Hello World");
  });

  it("extracts styles from /word/styles.xml part", () => {
    const xml = makeFlatOpc(DOCUMENT_XML, [
      {
        name: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        content: STYLES_XML,
      },
    ]);

    const doc = Docxtor.fromOoxml(xml);

    expect(doc.getStyles()).toHaveLength(1);
    expect(doc.getStyles()[0]!.id).toBe("Normal");
  });

  it("loads header from /word/header1.xml part", () => {
    const xml = makeFlatOpc(DOC_XML_WITH_HEADER, [
      {
        name: "/word/_rels/document.xml.rels",
        contentType: "application/vnd.openxmlformats-package.relationships+xml",
        content: RELS_XML,
      },
      {
        name: "/word/header1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        content: HEADER_XML,
      },
    ]);

    const doc = Docxtor.fromOoxml(xml);

    expect(doc.getHeaders().size).toBe(1);
    const header = [...doc.getHeaders().values()][0]!;
    expect(header.paragraphs[0]!.text).toBe("Page header");
  });

  it("supports applyPatches on content loaded from Flat OPC", () => {
    const doc = Docxtor.fromOoxml(makeFlatOpc());
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "OPC Patched" },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("OPC Patched");
  });

  it("throws when /word/document.xml part is missing", () => {
    // A package with only a styles part, no document.xml
    const xml = `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/xml"><pkg:xmlData>${STYLES_XML}</pkg:xmlData></pkg:part></pkg:package>`;

    expect(() => Docxtor.fromOoxml(xml)).toThrow(
      "Invalid OOXML: missing /word/document.xml part",
    );
  });

  it("preserves binary parts internally so toFlatOpc can emit them", () => {
    const b64 = Buffer.from("fake image bytes").toString("base64");
    const xml = makeFlatOpc(DOCUMENT_XML, [
      { name: "/word/media/img.png", contentType: "image/png", content: b64, isBinary: true },
    ]);

    const doc = Docxtor.fromOoxml(xml);
    const flatOpc = doc.toFlatOpc();

    expect(flatOpc).toContain("pkg:binaryData");
    expect(flatOpc).toContain(b64);
  });
});

// =============================================================================
// doc.toFlatOpc
// =============================================================================

describe("doc.toFlatOpc", () => {
  it("returns a Flat OPC XML string with a pkg:package root", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const result = doc.toFlatOpc();

    expect(result).toContain("pkg:package");
    expect(result).toContain(PKG_NS);
  });

  it("includes the document content", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const result = doc.toFlatOpc();

    expect(result).toContain("Hello World");
    expect(result).toContain("w:document");
  });

  it("contains the /word/document.xml part", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const result = doc.toFlatOpc();

    expect(result).toContain('pkg:name="/word/document.xml"');
  });

  it("reflects patches applied before the call", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "Changed" },
    ]);

    const result = doc.toFlatOpc();

    expect(result).toContain("Changed");
    expect(result).not.toContain("Hello World");
  });

  it("preserves original non-document parts when loaded from Flat OPC", () => {
    const xml = makeFlatOpc(DOCUMENT_XML, [
      {
        name: "/word/styles.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        content: STYLES_XML,
      },
    ]);

    const doc = Docxtor.fromOoxml(xml);
    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "Patched" },
    ]);

    const result = doc.toFlatOpc();
    const parts = parseFlatOpc(result);

    // Styles part must still be present and unchanged
    const stylesPart = parts.find((p) => p.name === "/word/styles.xml");
    expect(stylesPart).toBeDefined();
    expect(stylesPart!.content).toContain("Normal");

    // Document part must contain the patch
    const docPart = parts.find((p) => p.name === "/word/document.xml");
    expect(docPart!.content).toContain("Patched");
  });

  it("produces output that can be loaded back via fromOoxml", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "Round trip" },
    ]);

    const flatOpc = doc.toFlatOpc();
    const reloaded = Docxtor.fromOoxml(flatOpc);

    expect(reloaded.getParagraphs()[0]!.text).toBe("Round trip");
  });

  it("includes header parts when loaded from Flat OPC with headers", () => {
    const xml = makeFlatOpc(DOC_XML_WITH_HEADER, [
      {
        name: "/word/_rels/document.xml.rels",
        contentType: "application/vnd.openxmlformats-package.relationships+xml",
        content: RELS_XML,
      },
      {
        name: "/word/header1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        content: HEADER_XML,
      },
    ]);

    const doc = Docxtor.fromOoxml(xml);
    const result = doc.toFlatOpc();
    const parts = parseFlatOpc(result);

    const names = parts.map((p) => p.name);
    expect(names).toContain("/word/header1.xml");
  });

  it("produces a minimal single-part package when loaded from raw XML", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const result = doc.toFlatOpc();
    const parts = parseFlatOpc(result);

    // Should have at least the document part
    expect(parts.length).toBeGreaterThanOrEqual(1);
    expect(parts.some((p) => p.name === "/word/document.xml")).toBe(true);
  });

  it("round-trips through load() → toFlatOpc() → fromOoxml() correctly", async () => {
    // Start from a real DOCX file loaded via the async path
    const JSZip = (await import("jszip")).default;
    const zip = new JSZip();
    zip.file(
      "word/document.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>From load</w:t></w:r></w:p></w:body></w:document>`,
    );
    const buffer = await zip.generateAsync({ type: "nodebuffer" });

    const docFromLoad = await Docxtor.load(buffer);
    // load() does not produce a flat OPC — toFlatOpc() builds a minimal one
    const flatOpc = docFromLoad.toFlatOpc();
    const docFromOpc = Docxtor.fromOoxml(flatOpc);

    expect(docFromOpc.getParagraphs()[0]!.text).toBe("From load");
  });
});

// =============================================================================
// doc.patchResultsToOoxml
// =============================================================================

describe("doc.patchResultsToOoxml", () => {
  it("returns a Flat OPC XML string", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "Changed" },
    ]);

    const result = doc.patchResultsToOoxml(results);

    expect(result).toContain("pkg:package");
    expect(result).toContain(PKG_NS);
  });

  it("includes the modified paragraph content", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "Modified" },
    ]);

    const result = doc.patchResultsToOoxml(results);

    expect(result).toContain("Modified");
  });

  it("only includes elements targeted by successful patches", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Para zero</w:t></w:r></w:p><w:p><w:r><w:t>Para one</w:t></w:r></w:p></w:body></w:document>`;
    const doc = Docxtor.fromOoxml(xml);

    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Para zero", replace: "Changed zero" },
    ]);

    const fragment = doc.patchResultsToOoxml(results);

    expect(fragment).toContain("Changed zero");
    expect(fragment).not.toContain("Para one");
  });

  it("skips failed patches", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const results = doc.applyPatches([
      // this will fail — target does not exist
      { type: "replace_text", target: "p99", find: "x", replace: "y" },
    ]);

    expect(results[0]!.success).toBe(false);

    const fragment = doc.patchResultsToOoxml(results);
    const parts = parseFlatOpc(fragment);
    const docPart = parts.find((p) => p.name === "/word/document.xml");

    // Fragment document body should be empty (no successful targets)
    expect(docPart!.content).toContain("w:body");
    // No paragraph content from the failed patch
    expect(docPart!.content).not.toContain("p99");
  });

  it("does not duplicate elements when the same target appears in multiple patches", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello", replace: "Hi" },
      { type: "replace_text", target: "p0", find: "World", replace: "Earth" },
    ]);

    const fragment = doc.patchResultsToOoxml(results);

    // Count occurrences of w:p in the body — should be exactly one
    const bodyContent = fragment.slice(fragment.indexOf("<w:body>"));
    const pCount = (bodyContent.match(/<w:p[> ]/g) ?? []).length;
    expect(pCount).toBe(1);
  });

  it("works with tracked-change patches", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const results = doc.applyPatches(
      [{ type: "replace_text", target: "p0", find: "Hello World", replace: "Tracked" }],
      { trackRevisions: true, revisionAuthor: "Test" },
    );

    expect(results[0]!.success).toBe(true);

    const fragment = doc.patchResultsToOoxml(results);

    // Should contain revision markup
    expect(fragment).toContain("w:ins");
    expect(fragment).toContain("Tracked");
  });

  it("returns a parseable Flat OPC that fromOoxml can load", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "Fragment" },
    ]);

    const fragment = doc.patchResultsToOoxml(results);
    const reloaded = Docxtor.fromOoxml(fragment);

    expect(reloaded.getParagraphs()[0]!.text).toBe("Fragment");
  });

  it("handles patches without a target (append_text) gracefully", () => {
    const doc = Docxtor.fromOoxml(DOCUMENT_XML);
    const results = doc.applyPatches([
      { type: "append_text", text: "Appended" },
    ]);

    expect(results[0]!.success).toBe(true);

    // append_text has no target — the fragment should be empty but valid
    const fragment = doc.patchResultsToOoxml(results);

    expect(fragment).toContain("pkg:package");
    expect(fragment).toContain("w:body");
    // Should not contain the appended text (use toFlatOpc for that)
    expect(fragment).not.toContain("Appended");
  });

  it("includes table nodes for table-targeted patches", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`;
    const doc = Docxtor.fromOoxml(xml);

    const results = doc.applyPatches([
      { type: "add_table_row", target: "tbl0", afterRow: 0, cells: ["New"] },
    ]);

    expect(results[0]!.success).toBe(true);

    const fragment = doc.patchResultsToOoxml(results);

    expect(fragment).toContain("w:tbl");
    expect(fragment).toContain("New");
  });

  it("preserves namespace declarations from the source document", () => {
    const xmlWithNs = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>`;
    const doc = Docxtor.fromOoxml(xmlWithNs);
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hi", replace: "Bye" },
    ]);

    const fragment = doc.patchResultsToOoxml(results);

    expect(fragment).toContain('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
    expect(fragment).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  });

  it("uses the original document content type when loaded from Flat OPC", () => {
    const customContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    const doc = Docxtor.fromOoxml(makeFlatOpc());
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Hello World", replace: "X" },
    ]);

    const fragment = doc.patchResultsToOoxml(results);

    expect(fragment).toContain(customContentType);
  });
});
