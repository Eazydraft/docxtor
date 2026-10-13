/**
 * §3.5 Metadata scrub and core properties.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor } from "../../src/index.js";
import { createLeaseDocx, createSimpleDocx } from "../fixtures/create-docx.js";

async function entry(buffer: Buffer, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(buffer);
  const f = zip.file(path);
  return f ? f.async("string") : undefined;
}

describe("getCoreProperties", () => {
  it("reads docProps/core.xml", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    expect(doc.getCoreProperties()).toEqual({
      creator: "Jane Doe (Opposing Counsel LLP)",
      lastModifiedBy: "Jane Doe",
      title: "Office Lease",
      subject: "Lease",
      description: "Landlord form",
      keywords: "lease;landlord",
      category: "Real Estate",
      contentStatus: "Draft",
      revision: "17",
      created: "2025-12-01T10:00:00Z",
      modified: "2026-01-01T12:00:00Z",
    });
  });

  it("returns an empty object when there is no core.xml", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    expect(doc.getCoreProperties()).toEqual({});
  });
});

describe("setCoreProperties", () => {
  it("updates existing fields and persists through save/reload", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.setCoreProperties({ creator: "USI Legal", title: "Office Lease (Tenant Draft)", modified: "2026-10-13T00:00:00Z" });
    const reloaded = await Docxtor.load(await doc.toBuffer());
    const props = reloaded.getCoreProperties();
    expect(props.creator).toBe("USI Legal");
    expect(props.title).toBe("Office Lease (Tenant Draft)");
    expect(props.modified).toBe("2026-10-13T00:00:00Z");
    expect(props.subject).toBe("Lease"); // untouched
    const core = (await entry(await reloaded.toBuffer(), "docProps/core.xml"))!;
    expect(core).toContain('<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-13T00:00:00Z</dcterms:modified>');
  });

  it("creates core.xml, its relationship and content type when missing", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    doc.setCoreProperties({ creator: "USI Legal", title: "T & C <draft>", created: "2026-10-13T00:00:00Z" });
    const out = await doc.toBuffer();

    const core = (await entry(out, "docProps/core.xml"))!;
    expect(core).toContain("<dc:creator>USI Legal</dc:creator>");
    expect(core).toContain("<dc:title>T &amp; C &lt;draft&gt;</dc:title>");
    expect(core).toContain('xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"');
    expect((await entry(out, "_rels/.rels"))!).toContain('Target="docProps/core.xml"');
    expect((await entry(out, "[Content_Types].xml"))!).toContain('PartName="/docProps/core.xml"');

    const reloaded = await Docxtor.load(out);
    expect(reloaded.getCoreProperties()).toEqual({ creator: "USI Legal", title: "T & C <draft>", created: "2026-10-13T00:00:00Z" });
  });

  it("an empty string empties the element", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.setCoreProperties({ keywords: "" });
    expect(doc.getCoreProperties().keywords).toBe("");
  });
});

describe("scrubMetadata", () => {
  it("clears author/company fields, removes custom.xml, keeps title/created/modified", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.scrubMetadata();
    const out = await doc.toBuffer();
    const reloaded = await Docxtor.load(out);

    const props = reloaded.getCoreProperties();
    expect(props.creator).toBeUndefined();
    expect(props.lastModifiedBy).toBeUndefined();
    expect(props.description).toBeUndefined();
    expect(props.keywords).toBeUndefined();
    expect(props.category).toBeUndefined();
    expect(props.contentStatus).toBeUndefined();
    expect(props.revision).toBe("1");
    expect(props.title).toBe("Office Lease");
    expect(props.created).toBe("2025-12-01T10:00:00Z");
    expect(props.modified).toBe("2026-01-01T12:00:00Z");

    const core = (await entry(out, "docProps/core.xml"))!;
    expect(core).not.toContain("Jane Doe");
    expect(core).not.toContain("Opposing Counsel");

    const app = (await entry(out, "docProps/app.xml"))!;
    expect(app).not.toContain("Company");
    expect(app).not.toContain("Manager");
    expect(app).not.toContain("HyperlinkBase");
    expect(app).toContain("<Application>Microsoft Office Word</Application>");

    expect(await entry(out, "docProps/custom.xml")).toBeUndefined();
    expect((await entry(out, "_rels/.rels"))!).not.toContain("custom-properties");
    expect((await entry(out, "_rels/.rels"))!).toContain("core-properties");
    expect((await entry(out, "[Content_Types].xml"))!).not.toContain("custom.xml");
    expect((await entry(out, "[Content_Types].xml"))!).toContain("/docProps/core.xml");
  });

  it("is safe on a document with no docProps at all", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    expect(() => doc.scrubMetadata()).not.toThrow();
    const out = await doc.toBuffer();
    expect(await entry(out, "docProps/core.xml")).toBeUndefined();
    expect((await Docxtor.load(out)).getParagraphs()[0]!.text).toBe("Hello");
  });

  it("scrubbed output survives a second edit/save cycle", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.scrubMetadata();
    const once = await Docxtor.load(await doc.toBuffer());
    once.applyPatches([{ type: "append_text", text: "Tail" }]);
    const out = await once.toBuffer();
    expect(await entry(out, "docProps/custom.xml")).toBeUndefined();
    expect((await Docxtor.load(out)).getCoreProperties().creator).toBeUndefined();
  });
});

describe("round trip of extra parts", () => {
  it("preserves footnotes, core, app and rels content when nothing changed", async () => {
    const buffer = await createLeaseDocx();
    const doc = await Docxtor.load(buffer);
    const out = await doc.toBuffer();
    for (const path of ["docProps/core.xml", "docProps/app.xml", "docProps/custom.xml", "_rels/.rels", "word/footnotes.xml"]) {
      const before = (await entry(buffer, path))!;
      const after = (await entry(out, path))!;
      // Same content modulo serializer formatting (self-closing tags, whitespace between tags)
      const norm = (s: string): string =>
        s.replace(/>\s+</g, "><").replace(/<([\w:]+)([^>]*?)\s*\/>/g, "<$1$2></$1>").trim();
      expect(norm(after)).toBe(norm(before));
    }
  });
});
