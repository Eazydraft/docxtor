/**
 * Native paragraph identity (w14:paraId).
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor } from "../../src/index.js";
import { createLeaseDocx, createSimpleDocx, createTableDocx, createHeaderFooterDocx } from "../fixtures/create-docx.js";

const TRACKED = {
  trackRevisions: true,
  revisionAuthor: "USI Legal AI",
  revisionDate: "2026-10-13T00:00:00Z",
};

async function readZipEntry(buffer: Buffer, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(path);
  return file ? file.async("string") : undefined;
}

describe("paraId: reading native ids", () => {
  it("exposes w14:paraId on normalized paragraphs", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    expect(doc.findParagraphById("p0")!.paraId).toBe("0A000001");
    expect(doc.findParagraphById("p4")!.paraId).toBe("0A000005");
    expect(doc.findParagraphById("tbl0.r1.c1.p0")!.paraId).toBe("0A000013");
    expect(doc.findParagraphById("hdr.default.p0")!.paraId).toBe("0A000020");
    expect(doc.findParagraphById("ftr.default.p0")!.paraId).toBe("0A000021");
  });

  it("leaves paraId undefined on documents without w14 ids", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    expect(doc.findParagraphById("p0")!.paraId).toBeUndefined();
  });

  it("findParagraphByParaId resolves body, table, header and footer paragraphs (case-insensitive)", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    expect(doc.findParagraphByParaId("0a000005")!.text).toContain("Tenant shall indemnify");
    expect(doc.findParagraphByParaId("0A000012")!.text).toBe("Operating Expenses");
    expect(doc.findParagraphByParaId("0A000020")!.id).toBe("hdr.default.p0");
    expect(doc.findParagraphByParaId("0A000021")!.id).toBe("ftr.default.p0");
    expect(doc.findParagraphByParaId("DEADBEEF")).toBeUndefined();
  });

  it("accepts para: targets in patches", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const results = doc.applyPatches([
      { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "may not recapture" },
      { type: "insert_text", target: "para:0a000003", position: "after", anchor: "Exhibit A", text: " attached hereto" },
      { type: "add_list_item", target: "para:0A000005", position: "after", text: "Mirror clause." },
    ]);
    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.findParagraphByParaId("0A000007")!.text).toContain("may not recapture");
    expect(doc.findParagraphByParaId("0A000003")!.text).toContain("Exhibit A attached hereto");
    // add_list_item inserted after p4 → new paragraph sits at body index 5
    expect(doc.getParagraphs()[5]!.text).toBe("Mirror clause.");
  });

  it("validation reports an unknown para: id", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const [v] = doc.validatePatches([
      { type: "replace_text", target: "para:FFFFFFFF", find: "x", replace: "y" },
    ]);
    expect(v!.valid).toBe(false);
    expect(v!.error).toContain("para:FFFFFFFF");
  });
});

describe("paraId: projections", () => {
  it("markdown shows [p0 ¶XXXXXXXX] tokens", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const md = doc.project("markdown");
    expect(md).toContain("[p0 ¶0A000001]");
    expect(md).toContain("[tbl0.r0.c0.p0 ¶0A000010]");
    expect(md).toContain("[hdr.default.p0 ¶0A000020]");
  });

  it("markdown is unchanged for documents without paraIds", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    expect(doc.project("markdown")).toContain("[p0] Hello");
    expect(doc.project("markdown")).not.toContain("¶");
  });

  it("json carries paraId per paragraph", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const json = doc.project("json");
    const p0 = json.body[0] as { id: string; paraId?: string };
    expect(p0.paraId).toBe("0A000001");
  });
});

describe("paraId: assignParaIds load option", () => {
  it("stamps every paragraph in body, tables, headers and footers", async () => {
    const doc = await Docxtor.load(await createLeaseDocx({ withParaIds: false }), { assignParaIds: true });

    const all = [
      ...doc.getParagraphs(),
      ...doc.getTables().flatMap((t) => t.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs))),
      ...[...doc.getHeaders().values()].flatMap((h) => h.paragraphs),
      ...[...doc.getFooters().values()].flatMap((f) => f.paragraphs),
    ];
    expect(all.length).toBeGreaterThan(10);

    const ids = new Set<string>();
    for (const p of all) {
      expect(p.paraId).toMatch(/^[0-9A-F]{8}$/);
      expect(parseInt(p.paraId!, 16)).toBeLessThan(0x80000000);
      expect(parseInt(p.paraId!, 16)).toBeGreaterThan(0);
      ids.add(p.paraId!);
    }
    expect(ids.size).toBe(all.length);
  });

  it("declares the namespace and mc:Ignorable, sets w14:textId, and survives a round trip", async () => {
    const doc = await Docxtor.load(await createLeaseDocx({ withParaIds: false }), { assignParaIds: true });
    const before = new Map(doc.getParagraphs().map((p) => [p.paraId!, p.text]));

    const out = await doc.toBuffer();
    const docXml = (await readZipEntry(out, "word/document.xml"))!;
    expect(docXml).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
    expect(docXml).toMatch(/mc:Ignorable="[^"]*w14/);
    expect(docXml).toMatch(/w14:paraId="[0-9A-F]{8}" w14:textId="[0-9A-F]{8}"/);
    const hdrXml = (await readZipEntry(out, "word/header1.xml"))!;
    expect(hdrXml).toContain("xmlns:w14=");
    expect(hdrXml).toContain("w14:paraId=");

    const reloaded = await Docxtor.load(out);
    for (const [paraId, text] of before) {
      expect(reloaded.findParagraphByParaId(paraId)!.text).toBe(text);
    }
  });

  it("appends w14 to an existing mc:Ignorable list and keeps existing ids", async () => {
    // The lease already has ids and mc:Ignorable="w14": loading with assignParaIds is a no-op
    const doc = await Docxtor.load(await createLeaseDocx(), { assignParaIds: true });
    expect(doc.findParagraphById("p0")!.paraId).toBe("0A000001");
    const xml = doc.toString();
    expect(xml.match(/mc:Ignorable="w14"/g)!.length).toBe(1);
  });

  it("is a no-op when the option is off", async () => {
    const doc = await Docxtor.load(await createLeaseDocx({ withParaIds: false }));
    expect(doc.findParagraphById("p0")!.paraId).toBeUndefined();
    expect(doc.toString()).not.toContain("w14:");
  });
});

describe("paraId: new paragraphs created by patches", () => {
  it("insert_paragraph, append_text, add_list_item and add_table get fresh unique ids", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const results = doc.applyPatches([
      { type: "insert_paragraph", target: "p4", position: "after", text: "Mirror indemnity." },
      { type: "append_text", text: "Appended one.\nAppended two." },
      { type: "add_list_item", target: "p2", position: "after", text: "New item." },
      { type: "add_table", afterId: "p10", rows: [["a", "b"]] },
      { type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["Taxes", "Pro rata"] },
    ]);
    expect(results.every((r) => r.success)).toBe(true);

    const all = [
      ...doc.getParagraphs(),
      ...doc.getTables().flatMap((t) => t.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs))),
    ];
    const ids = all.map((p) => p.paraId);
    expect(ids.every((id) => id !== undefined && /^[0-9A-F]{8}$/.test(id))).toBe(true);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("add_table_row does not duplicate the cloned row's paraIds", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([{ type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["Taxes", "Pro rata"] }]);
    const table = doc.getTables()[0]!;
    const row1 = table.rows[1]!.cells.map((c) => c.paragraphs[0]!.paraId);
    const row2 = table.rows[2]!.cells.map((c) => c.paragraphs[0]!.paraId);
    expect(row1).toEqual(["0A000012", "0A000013"]);
    expect(row2[0]).not.toBe("0A000012");
    expect(row2[1]).not.toBe("0A000013");
    expect(row2[0]).toMatch(/^[0-9A-F]{8}$/);
  });

  it("does not stamp ids when the part does not declare w14", async () => {
    const doc = await Docxtor.load(await createSimpleDocx("Hello"));
    doc.applyPatches([
      { type: "insert_paragraph", target: "p0", position: "after", text: "World" },
      { type: "append_text", text: "Tail" },
    ]);
    expect(doc.toString()).not.toContain("w14:paraId");
    expect(doc.getParagraphs()[1]!.paraId).toBeUndefined();
  });

  it("stamps header/footer paragraphs when that part declares w14", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "append_text", text: "Second header line", container: { type: "header", variant: "default" } },
      { type: "insert_paragraph", target: "hdr.default.p0", position: "before", text: "Top line" },
    ]);
    const hdr = [...doc.getHeaders().values()][0]!;
    expect(hdr.paragraphs.length).toBe(3);
    expect(hdr.paragraphs.every((p) => p.paraId !== undefined)).toBe(true);
    expect(new Set(hdr.paragraphs.map((p) => p.paraId)).size).toBe(3);
  });

  it("a part created without w14 (plain header/footer fixture) is left alone", async () => {
    const doc = await Docxtor.load(await createHeaderFooterDocx());
    doc.applyPatches([{ type: "append_text", text: "x", container: { type: "header", variant: "default" } }]);
    const hdr = [...doc.getHeaders().values()][0]!;
    expect(hdr.paragraphs[1]!.paraId).toBeUndefined();
  });

  it("a table fixture without ids stays without ids after row insert", async () => {
    const doc = await Docxtor.load(await createTableDocx());
    doc.applyPatches([{ type: "add_table_row", target: "tbl0", afterRow: 0, cells: ["x", "y"] }]);
    expect(doc.toString()).not.toContain("w14:");
  });
});

describe("paraId: identity stability (acceptance)", () => {
  it("ten tracked patches incl. two paragraph inserts: paraIds stable, positional ids shift", async () => {
    const original = await Docxtor.load(await createLeaseDocx());
    const before = new Map(original.getParagraphs().map((p) => [p.paraId!, { id: p.id, text: p.text }]));

    const results = original.applyPatches([
      { type: "insert_paragraph", target: "para:0A000005", position: "after", text: "Landlord shall indemnify, defend and hold harmless Tenant from all claims arising from Landlord's negligence." },
      { type: "insert_paragraph", target: "para:0A000002", position: "before", text: "RECITALS" },
      { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "shall have no right to recapture" },
      { type: "insert_text", target: "para:0A000003", position: "after", anchor: "Exhibit A", text: " [CONFIRM RSF]" },
      { type: "delete_text", target: "para:0A000009", find: " per rentable square foot per year" },
      { type: "insert_text", target: "para:0A000009", position: "after", anchor: "$42.00", text: " per RSF per annum" },
      { type: "apply_formatting", target: "para:0A000001", formatting: { bold: true } },
      { type: "replace_text", target: "para:0A000014", find: "follows", replace: "is attached" },
      { type: "insert_text", target: "para:0A00000A", position: "before", anchor: "Security", text: "[ISSUE-7] " },
      { type: "replace_cell_text", target: "para:0A000012", find: "Operating Expenses", replace: "Operating Expenses (see Exhibit C)" },
    ], TRACKED);
    expect(results.every((r) => r.success)).toBe(true);

    const reloaded = await Docxtor.load(await original.toBuffer());

    // Every original paraId still resolves; untouched paragraphs read the same
    const untouched = ["0A000001", "0A000002", "0A000004", "0A000006", "0A000008", "0A000010", "0A000011", "0A000013"];
    for (const [paraId, prev] of before) {
      const p = reloaded.findParagraphByParaId(paraId);
      expect(p, `paraId ${paraId} should still resolve`).toBeDefined();
      if (untouched.includes(paraId)) expect(p!.text).toBe(prev.text);
    }

    // Positional ids shifted by the two inserts
    expect(reloaded.findParagraphByParaId("0A000002")!.id).toBe("p2");    // was p1
    expect(reloaded.findParagraphByParaId("0A000006")!.id).toBe("p7");    // was p5
    expect(before.get("0A000006")!.id).toBe("p5");

    // The inserted paragraphs have their own ids and carry the tracked mark
    const mirror = reloaded.getParagraphs().find((p) => p.text.startsWith("Landlord shall indemnify"))!;
    expect(mirror.paraId).toMatch(/^[0-9A-F]{8}$/);
    expect(mirror.listInfo).toEqual({ numId: "1", level: 1, listType: "multilevel" });
  });
});

describe("paraId: table cells in the JSON projection", () => {
  it("single-paragraph cells carry paraId without extra options", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const json = doc.project("json");
    const table = json.body.find((el) => "rows" in el) as { rows: Array<Array<{ id: string; paraId?: string; paragraphs?: unknown }>> };
    expect(table.rows[0]![0]).toEqual({ id: "tbl0.r0.c0.p0", paraId: "0A000010", text: "Item" });
    expect(table.rows[1]![1]!.paraId).toBe("0A000013");
    expect(table.rows[1]![1]!.paragraphs).toBeUndefined();
  });

  it("cells without ids have no paraId key", async () => {
    const doc = await Docxtor.load(await createTableDocx());
    const table = doc.project("json").body.find((el) => "rows" in el) as { rows: Array<Array<{ paraId?: string }>> };
    expect("paraId" in table.rows[0]![0]!).toBe(false);
  });
});
