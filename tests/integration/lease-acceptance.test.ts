/**
 * §5 Acceptance tests for the lease review additions, run against the
 * synthetic lease fixture. No language model involved.
 *
 * Each `it` corresponds to one checkbox in DOCXTOR_ADDITIONS.md §5; the
 * detailed per-feature behaviour is covered by the unit test files.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor } from "../../src/index.js";
import type { Patch } from "../../src/index.js";
import { createLeaseDocx } from "../fixtures/create-docx.js";

const AUTHOR = "USI Legal AI";
const TRACKED = { trackRevisions: true, revisionAuthor: AUTHOR, revisionDate: "2026-10-13T00:00:00Z" };

async function entry(buffer: Buffer, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(buffer);
  const f = zip.file(path);
  return f ? f.async("string") : undefined;
}

function allParagraphs(doc: Docxtor) {
  return [
    ...doc.getParagraphs(),
    ...doc.getTables().flatMap((t) => t.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs))),
    ...[...doc.getHeaders().values()].flatMap((h) => h.paragraphs),
    ...[...doc.getFooters().values()].flatMap((f) => f.paragraphs),
  ];
}

describe("lease acceptance (§5)", () => {
  it("Round trip: load, save with no patches — paragraph count, text and paraIds unchanged", async () => {
    const buffer = await createLeaseDocx();
    const doc = await Docxtor.load(buffer);
    const before = allParagraphs(doc).map((p) => ({ id: p.id, paraId: p.paraId, text: p.text }));

    const out = await doc.toBuffer();
    const reloaded = await Docxtor.load(out);
    const after = allParagraphs(reloaded).map((p) => ({ id: p.id, paraId: p.paraId, text: p.text }));
    expect(after).toEqual(before);

    // Package parts all still there
    const zip = await JSZip.loadAsync(out);
    for (const path of [
      "[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/_rels/document.xml.rels",
      "word/header1.xml", "word/footer1.xml", "word/footnotes.xml", "word/numbering.xml", "word/styles.xml",
      "docProps/core.xml", "docProps/app.xml", "docProps/custom.xml",
    ]) {
      expect(zip.file(path), path).not.toBeNull();
    }
    const docXml = (await entry(out, "word/document.xml"))!;
    expect(docXml).toContain('w14:paraId="0A000001"');
    expect(docXml).toContain('mc:Ignorable="w14"');
  });

  it("Identity stability: paraIds resolve to the same text after ten tracked patches, save and reload", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const before = new Map(allParagraphs(doc).map((p) => [p.paraId!, p.text]));
    const touched = new Set(["0A000007", "0A000003", "0A000009", "0A000012", "0A000014"]);

    const results = doc.applyPatches([
      { type: "insert_paragraph", target: "para:0A000005", position: "after", text: "Landlord shall indemnify, defend and hold harmless Tenant." },
      { type: "insert_paragraph", target: "para:0A000002", position: "before", text: "RECITALS" },
      { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "shall have no right to recapture" },
      { type: "insert_text", target: "para:0A000003", position: "after", anchor: "Exhibit A", text: " attached hereto" },
      { type: "delete_text", target: "para:0A000009", find: " per year" },
      { type: "insert_text", target: "para:0A000009", position: "after", anchor: "$42.00", text: " (subject to escalation)" },
      { type: "apply_formatting", target: "para:0A000001", formatting: { bold: true } },
      { type: "replace_text", target: "para:0A000014", find: "follows", replace: "is attached" },
      { type: "replace_cell_text", target: "para:0A000012", find: "Operating Expenses", replace: "Operating Expenses (Exhibit C)" },
      { type: "insert_text", target: "para:0A000020", position: "after", anchor: "CONFIDENTIAL", text: " - TENANT DRAFT" },
    ], TRACKED);
    expect(results.every((r) => r.success)).toBe(true);

    const reloaded = await Docxtor.load(await doc.toBuffer());
    for (const [paraId, text] of before) {
      const p = reloaded.findParagraphByParaId(paraId);
      expect(p, paraId).toBeDefined();
      if (!touched.has(paraId) && paraId !== "0A000020") expect(p!.text).toBe(text);
    }
    // Positional ids shifted, paraIds did not
    expect(reloaded.findParagraphByParaId("0A000005")!.id).toBe("p5");
    expect(reloaded.findParagraphByParaId("0A000014")!.id).toBe("p12");
  });

  it("Engine fixed point: tracked + acceptAll equals untracked (byte-level with pinned date, modulo fresh paraIds)", async () => {
    const batch: Patch[] = [
      { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "shall have no right to recapture" },
      { type: "insert_paragraph", target: "para:0A000005", position: "after", text: "Landlord shall indemnify Tenant." },
      { type: "delete_paragraph", target: "para:0A000006" },
      { type: "insert_paragraph", target: "para:0A000014", position: "before", text: "IN WITNESS WHEREOF" },
      { type: "replace_cell_text", target: "para:0A000012", find: "Operating", replace: "Building Operating" },
    ];
    const tracked = await Docxtor.load(await createLeaseDocx());
    const untracked = await Docxtor.load(await createLeaseDocx());
    tracked.applyPatches(batch, TRACKED);
    untracked.applyPatches(batch);
    tracked.acceptAllRevisions({ authors: [AUTHOR] });

    const strip = (xml: string): string => xml.replace(/ w14:paraId="[0-9A-F]{8}" w14:textId="[0-9A-F]{8}"/g, "");
    expect(strip(tracked.toString())).toBe(strip(untracked.toString()));
    // Projections match once id tokens are removed (positional ids are
    // contiguous after accept but next-free in the untracked path).
    const stripIds = (md: string): string => md.replace(/^\[[^\]]+\] ?/gm, "").replace(/ ¶[0-9A-F]{8}/g, "");
    expect(stripIds(tracked.project("markdown"))).toBe(stripIds(untracked.project("markdown")));
  });

  it("Correspondence: N bracket tokens at N distinct anchors are found exactly once each, at their paragraph", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const issues = [
      { paraId: "0A000003", anchor: "Exhibit A", token: "[ISSUE-1: confirm RSF]" },
      { paraId: "0A000005", anchor: "Landlord", token: "[ISSUE-2: mutual indemnity?]" },
      { paraId: "0A000009", anchor: "$42.00", token: "[ISSUE-3: confirm rent]" },
      { paraId: "0A00000A", anchor: "two months", token: "[ISSUE-4: deposit]" },
      { paraId: "0A000013", anchor: "5%", token: "[ISSUE-5: cap]" },
    ];
    const results = doc.applyPatches(
      issues.map((i) => ({ type: "insert_text", target: `para:${i.paraId}`, position: "after", anchor: i.anchor, text: ` ${i.token}` } as Patch)),
      TRACKED,
    );
    expect(results.every((r) => r.success)).toBe(true);

    const reloaded = await Docxtor.load(await doc.toBuffer());
    const md = reloaded.project("markdown");
    const found = md.match(/\[ISSUE-\d+: [^\]]+\]/g) ?? [];
    expect(found.length).toBe(issues.length);
    expect(new Set(found).size).toBe(issues.length);
    for (const issue of issues) {
      const p = reloaded.findParagraphByParaId(issue.paraId)!;
      expect(p.text).toContain(issue.token);
      expect(p.text.split(issue.token).length - 1).toBe(1);
    }
  });

  it("Paragraph insert in a list: same numId and ilvl; after accept exactly one new paragraph, no blank", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const count = doc.getParagraphs().length;
    doc.applyPatches([
      { type: "insert_paragraph", target: "para:0A000005", position: "after", text: "Landlord shall indemnify, defend and hold harmless Tenant." },
    ], TRACKED);
    const inserted = doc.getParagraphs()[5]!;
    expect(inserted.listInfo).toEqual({ numId: "1", level: 1, listType: "multilevel" });
    expect(doc.project("markdown")).toContain("2.2 Landlord shall indemnify");

    doc.acceptAllRevisions();
    expect(doc.getParagraphs().length).toBe(count + 1);
    expect(doc.getParagraphs().some((p) => p.text.trim() === "")).toBe(false);
    expect(doc.toString()).not.toContain("w:ins");
  });

  it("Paragraph delete: delete + accept drops the count by one with no empty paragraph", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const count = doc.getParagraphs().length;
    doc.applyPatches([{ type: "delete_paragraph", target: "para:0A000007" }], TRACKED);
    expect(doc.getParagraphs().length).toBe(count);
    doc.acceptAllRevisions();
    expect(doc.getParagraphs().length).toBe(count - 1);
    expect(doc.getParagraphs().some((p) => p.text.trim() === "")).toBe(false);
    expect(doc.findParagraphByParaId("0A000007")).toBeUndefined();
  });

  it("Author filter: accepting only the tool's author leaves the other author's marks untouched", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.applyPatches([
      { type: "replace_text", target: "para:0A000007", find: "may recapture", replace: "may not recapture" },
      { type: "delete_paragraph", target: "para:0A000006" },
    ], TRACKED);
    doc.acceptAllRevisions({ authors: [AUTHOR] });

    const out = await doc.toBuffer();
    const docXml = (await entry(out, "word/document.xml"))!;
    expect(docXml).not.toContain(AUTHOR);
    expect((docXml.match(/w:author="Opposing Counsel"/g) ?? []).length).toBe(2);
    expect(docXml).toContain("<w:delText>$40.00</w:delText>");
    expect((await entry(out, "word/footnotes.xml"))!).toContain('w:author="Opposing Counsel"');
  });

  it("Red-text read: colour filter returns only the red runs with correct ids and offsets", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const json = doc.project("json", { runFilter: { color: "FF0000" } });
    const hits: Array<{ id: string; paraId?: string; text: string; start: number; end: number }> = [];
    const visit = (p: { id: string; paraId?: string; text: string; formatting?: Array<{ text: string; startOffset: number; endOffset: number }> }) => {
      for (const r of p.formatting ?? []) hits.push({ id: p.id, paraId: p.paraId, text: r.text, start: r.startOffset, end: r.endOffset });
    };
    for (const el of json.body as any[]) {
      if ("rows" in el) for (const row of el.rows) for (const cell of row) for (const p of cell.paragraphs ?? []) visit(p);
      else visit(el);
    }
    expect(hits).toEqual([
      { id: "p9", paraId: "0A00000A", text: "two months", start: 25, end: 35 },
      { id: "p9", paraId: "0A00000A", text: "Parent Co", start: 48, end: 57 },
      { id: "tbl0.r1.c1.p0", paraId: "0A000013", text: "5%", start: 10, end: 12 },
    ]);
  });

  it("Atomic failure: a failing batch leaves the exported file identical to a no-patch export", async () => {
    const buffer = await createLeaseDocx();
    const pristine = await (await Docxtor.load(buffer)).toBuffer();

    const doc = await Docxtor.load(buffer);
    const patches: Patch[] = [];
    for (let i = 0; i < 49; i++) {
      patches.push({ type: "insert_text", target: "para:0A000005", position: "after", anchor: "Tenant", text: `[${i}]`, occurrence: 1 });
    }
    patches.push({ type: "replace_text", target: "para:0A000005", find: "NOT THERE", replace: "x" });
    const results = doc.applyPatches(patches, { ...TRACKED, atomic: true });
    expect(results.filter((r) => r.success).length).toBe(49);
    expect(results.filter((r) => !r.success).length).toBe(1);

    const out = await doc.toBuffer();
    const a = await JSZip.loadAsync(pristine);
    const b = await JSZip.loadAsync(out);
    expect(Object.keys(b.files).sort()).toEqual(Object.keys(a.files).sort());
    for (const path of Object.keys(a.files)) {
      if (a.files[path]!.dir) continue;
      expect(await b.file(path)!.async("string"), path).toBe(await a.file(path)!.async("string"));
    }
  });

  it("Metadata scrub: no author/company fields remain and custom.xml is gone", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    doc.scrubMetadata();
    const out = await doc.toBuffer();
    const reloaded = await Docxtor.load(out);
    const props = reloaded.getCoreProperties();
    expect(props.creator).toBeUndefined();
    expect(props.lastModifiedBy).toBeUndefined();
    expect(props.title).toBe("Office Lease");
    expect(props.revision).toBe("1");
    expect((await entry(out, "docProps/app.xml"))!).not.toContain("Company");
    expect(await entry(out, "docProps/custom.xml")).toBeUndefined();
    expect((await entry(out, "[Content_Types].xml"))!).not.toContain("custom.xml");
    expect((await entry(out, "_rels/.rels"))!).not.toContain("custom-properties");
  });
});
