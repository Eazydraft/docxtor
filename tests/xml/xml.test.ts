import { describe, it, expect } from "vitest";
import { parseXml } from "../../src/xml/parser.js";
import { serializeXml } from "../../src/xml/serializer.js";
import {
  getTagName,
  isTag,
  isTextNode,
  getChildren,
  findChild,
  findChildren,
  findByPath,
  getAttr,
  setAttr,
  removeAttr,
  getTextContent,
  setTextContent,
  getRunText,
  findTextNode,
  cloneNode,
  insertAfter,
  removeChild,
  createElement,
  createTextNode,
  createWt,
  createRun,
} from "../../src/xml/helpers.js";

// =============================================================================
// XML Parse/Serialize round-trip
// =============================================================================

describe("xml parser + serializer round-trip", () => {
  const SIMPLE_PARA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  it("should parse and re-serialize XML without losing content", () => {
    const parsed = parseXml(SIMPLE_PARA);
    const serialized = serializeXml(parsed);

    expect(serialized).toContain("w:document");
    expect(serialized).toContain("w:body");
    expect(serialized).toContain("w:p");
    expect(serialized).toContain("w:r");
    expect(serialized).toContain("w:t");
    expect(serialized).toContain("Hello World");
  });

  it("should preserve namespace declarations", () => {
    const parsed = parseXml(SIMPLE_PARA);
    const serialized = serializeXml(parsed);

    expect(serialized).toContain(
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    );
  });

  it("should preserve attributes on elements", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:p w:rsidR="00A77427" w:rsidRDefault="00A77427">
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
  </w:pPr>
  <w:r>
    <w:rPr><w:b/></w:rPr>
    <w:t>Text</w:t>
  </w:r>
</w:p>`;

    const parsed = parseXml(xml);
    const serialized = serializeXml(parsed);

    expect(serialized).toContain('w:rsidR="00A77427"');
    expect(serialized).toContain('w:val="Heading1"');
    expect(serialized).toContain("w:b");
  });

  it("should preserve xml:space attribute", () => {
    const xml = `<w:r><w:t xml:space="preserve"> text </w:t></w:r>`;
    const parsed = parseXml(xml);
    const serialized = serializeXml(parsed);

    expect(serialized).toContain('xml:space="preserve"');
    expect(serialized).toContain(" text ");
  });

  it("should handle self-closing tags", () => {
    const xml = `<w:rPr><w:b/><w:i/></w:rPr>`;
    const parsed = parseXml(xml);
    const serialized = serializeXml(parsed);

    // FXP might serialize as <w:b></w:b> or <w:b/> — both are valid
    expect(serialized).toContain("w:b");
    expect(serialized).toContain("w:i");
  });

  it("should preserve element ordering", () => {
    const xml = `<w:p>
  <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
  <w:r><w:t>First</w:t></w:r>
  <w:r><w:t>Second</w:t></w:r>
  <w:r><w:t>Third</w:t></w:r>
</w:p>`;

    const parsed = parseXml(xml);
    const serialized = serializeXml(parsed);

    const firstIdx = serialized.indexOf("First");
    const secondIdx = serialized.indexOf("Second");
    const thirdIdx = serialized.indexOf("Third");

    expect(firstIdx).toBeLessThan(secondIdx);
    expect(secondIdx).toBeLessThan(thirdIdx);
  });

  it("should not mangle numeric-looking attribute values", () => {
    const xml = `<w:p w:rsidR="00A77427"><w:r><w:t>42</w:t></w:r></w:p>`;
    const parsed = parseXml(xml);
    const serialized = serializeXml(parsed);

    // "00A77427" must not be parsed as a hex number
    expect(serialized).toContain('w:rsidR="00A77427"');
    // "42" text content must not become a number
    expect(serialized).toContain("42");
  });
});

// =============================================================================
// XML Helpers
// =============================================================================

describe("xml helpers", () => {
  // Parse a test document for helper tests
  const TEST_XML = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p w:rsidR="00A7">
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Bold text</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> normal text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  function getParsed() {
    return parseXml(TEST_XML);
  }

  describe("getTagName", () => {
    it("should return the tag name of an element", () => {
      const nodes = getParsed();
      expect(getTagName(nodes[0]!)).toBe("w:document");
    });

    it("should return undefined for text nodes", () => {
      const textNode = { "#text": "hello" };
      expect(getTagName(textNode)).toBeUndefined();
    });
  });

  describe("isTag", () => {
    it("should return true for matching tag", () => {
      const nodes = getParsed();
      expect(isTag(nodes[0]!, "w:document")).toBe(true);
    });

    it("should return false for non-matching tag", () => {
      const nodes = getParsed();
      expect(isTag(nodes[0]!, "w:p")).toBe(false);
    });
  });

  describe("isTextNode", () => {
    it("should identify text nodes", () => {
      expect(isTextNode({ "#text": "hello" })).toBe(true);
    });

    it("should return false for element nodes", () => {
      const nodes = getParsed();
      expect(isTextNode(nodes[0]!)).toBe(false);
    });
  });

  describe("getChildren / findChild / findChildren", () => {
    it("should get children of an element", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const children = getChildren(doc);
      expect(children.length).toBeGreaterThan(0);
    });

    it("should find a specific child by tag name", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body");
      expect(body).toBeDefined();
      expect(isTag(body!, "w:body")).toBe(true);
    });

    it("should find all children with a tag name", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const paragraphs = findChildren(body, "w:p");
      expect(paragraphs.length).toBe(1);
    });

    it("should return empty for non-existent children", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const missing = findChildren(doc, "w:nonexistent");
      expect(missing.length).toBe(0);
    });
  });

  describe("findByPath", () => {
    it("should find nested elements by path", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const para = findByPath(doc, ["w:body", "w:p"]);
      expect(para).toBeDefined();
      expect(isTag(para!, "w:p")).toBe(true);
    });

    it("should return undefined for invalid path", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const result = findByPath(doc, ["w:body", "w:nonexistent"]);
      expect(result).toBeUndefined();
    });
  });

  describe("attribute access", () => {
    it("should get attribute values", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const para = findChild(body, "w:p")!;
      expect(getAttr(para, "w:rsidR")).toBe("00A7");
    });

    it("should set attribute values", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const para = findChild(body, "w:p")!;
      setAttr(para, "w:newAttr", "test123");
      expect(getAttr(para, "w:newAttr")).toBe("test123");
    });

    it("should remove attributes", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const para = findChild(body, "w:p")!;
      removeAttr(para, "w:rsidR");
      expect(getAttr(para, "w:rsidR")).toBeUndefined();
    });

    it("should return undefined for missing attributes", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      expect(getAttr(doc, "nonexistent")).toBeUndefined();
    });
  });

  describe("text access", () => {
    it("should get text content from text nodes", () => {
      const node = { "#text": "hello" };
      expect(getTextContent(node)).toBe("hello");
    });

    it("should set text content on text nodes", () => {
      const node = { "#text": "hello" };
      setTextContent(node, "world");
      expect(getTextContent(node)).toBe("world");
    });

    it("should get run text from w:r elements", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const para = findChild(body, "w:p")!;
      const runs = findChildren(para, "w:r");

      expect(getRunText(runs[0]!)).toBe("Bold text");
      expect(getRunText(runs[1]!)).toBe(" normal text");
    });

    it("should find the w:t node inside a run", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const para = findChild(body, "w:p")!;
      const run = findChildren(para, "w:r")[0]!;
      const wt = findTextNode(run);
      expect(wt).toBeDefined();
      expect(isTag(wt!, "w:t")).toBe(true);
    });
  });

  describe("node manipulation", () => {
    it("should deep clone nodes", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const clone = cloneNode(doc);

      // Should be equal but not the same reference
      expect(JSON.stringify(clone)).toBe(JSON.stringify(doc));
      expect(clone).not.toBe(doc);

      // Modifying clone should not affect original
      const cloneBody = findChild(clone, "w:body")!;
      const clonePara = findChild(cloneBody, "w:p")!;
      setAttr(clonePara, "w:rsidR", "CHANGED");
      const origBody = findChild(doc, "w:body")!;
      const origPara = findChild(origBody, "w:p")!;
      expect(getAttr(origPara, "w:rsidR")).toBe("00A7");
    });

    it("should insert nodes after a reference", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const para = findChild(body, "w:p")!;
      const newPara = createElement("w:p", [
        createElement("w:r", [createWt("New paragraph")]),
      ]);

      insertAfter(body, para, newPara);

      const allParas = findChildren(body, "w:p");
      expect(allParas.length).toBe(2);
    });

    it("should remove child nodes", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const para = findChild(body, "w:p")!;
      const runs = findChildren(para, "w:r");

      expect(runs.length).toBe(2);
      removeChild(para, runs[1]!);
      expect(findChildren(para, "w:r").length).toBe(1);
    });

    it("should throw when reference child not found", () => {
      const nodes = getParsed();
      const doc = nodes[0]!;
      const body = findChild(doc, "w:body")!;
      const fakeChild = createElement("w:fake");

      expect(() => insertAfter(body, fakeChild, createElement("w:new"))).toThrow(
        "Reference child not found",
      );
    });
  });

  describe("node creation", () => {
    it("should create element nodes", () => {
      const node = createElement("w:p", [], { "w:rsidR": "001A" });
      expect(isTag(node, "w:p")).toBe(true);
      expect(getAttr(node, "w:rsidR")).toBe("001A");
    });

    it("should create text nodes", () => {
      const node = createTextNode("hello");
      expect(isTextNode(node)).toBe(true);
      expect(getTextContent(node)).toBe("hello");
    });

    it("should create w:t elements with xml:space for whitespace", () => {
      const node = createWt(" spaced ");
      expect(isTag(node, "w:t")).toBe(true);
      expect(getAttr(node, "xml:space")).toBe("preserve");
    });

    it("should create w:t elements without xml:space for clean text", () => {
      const node = createWt("clean");
      expect(getAttr(node, "xml:space")).toBeUndefined();
    });

    it("should create run elements with formatting", () => {
      const rPr = createElement("w:rPr", [createElement("w:b")]);
      const run = createRun("Bold text", rPr);

      expect(isTag(run, "w:r")).toBe(true);
      const children = getChildren(run);
      // Should have rPr + w:t
      expect(children.length).toBe(2);
      expect(findChild(run, "w:rPr")).toBeDefined();
      expect(findTextNode(run)).toBeDefined();
      expect(getRunText(run)).toBe("Bold text");
    });
  });
});
