/**
 * The skill's CLI (cli/docxtor.ts), run as a process the way an agent runs it.
 */

import { describe, it, expect, beforeAll } from "vitest";
import { execFile } from "node:child_process";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { Docxtor } from "../../src/index.js";
import { createLeaseDocx, createMultiParagraphDocx } from "../fixtures/create-docx.js";

const run = promisify(execFile);
const CLI = join(import.meta.dirname, "../../cli/docxtor.ts");
const TSX = join(import.meta.dirname, "../../node_modules/.bin/tsx");

async function cli(...args: string[]): Promise<{ code: number; out: string; err: string }> {
  try {
    const { stdout, stderr } = await run(TSX, [CLI, ...args]);
    return { code: 0, out: stdout, err: stderr };
  } catch (e) {
    const x = e as { code: number; stdout: string; stderr: string };
    return { code: x.code, out: x.stdout, err: x.stderr };
  }
}

const texts = async (file: string): Promise<string[]> => (await Docxtor.load(file)).getParagraphs().map((p) => p.text);

let dir: string;
let v1: string, lease: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "docxtor-cli-"));
  v1 = join(dir, "v1.docx");
  lease = join(dir, "lease.docx");
  await writeFile(v1, await createMultiParagraphDocx([
    "AGREEMENT",
    "1. Payment. Buyer shall pay Seller within 30 days of invoice.",
    "2. Law. This Agreement is governed by the laws of Texas.",
    "3. Notices. Notices go to the Seller's address.",
    "4. Term. One year.",
  ]));
  await writeFile(lease, await createLeaseDocx());
}, 30_000);

describe("docxtor CLI", () => {
  it("view prints ids and inline revisions", async () => {
    const r = await cli("view", lease);
    expect(r.code).toBe(0);
    expect(r.out).toContain("[p8 ¶0A000009]");
    expect(r.out).toContain("{-$40.00-}{+$42.00+}");
    expect(r.out).not.toContain("Available Styles");
    const g = await cli("view", lease, "--grep", "base rent");
    expect(g.out.trim().split("\n")).toHaveLength(1);
  }, 30_000);

  it("replace refuses an ambiguous find and writes nothing", async () => {
    const out = join(dir, "amb.docx");
    const r = await cli("replace", v1, "-o", out, "--find", "Seller", "--replace", "Vendor");
    expect(r.code).toBe(1);
    expect(r.err).toContain("matches 2 times");
    expect(r.err).toContain("[p1]");
  }, 30_000);

  it("replace --all tracks every occurrence, and quote variants are tried", async () => {
    const out = join(dir, "all.docx");
    const r = await cli("replace", v1, "-o", out, "--find", "Seller", "--replace", "Vendor", "--all", "--author", "Jane");
    expect(r.code).toBe(0);
    expect(await texts(out)).toContain("3. Notices. Notices go to the Vendor's address.");
    const rev = await cli("revisions", out, "--json");
    const changes = JSON.parse(rev.out) as { kind: string; authors: string[] }[];
    expect(changes).toHaveLength(2);
    expect(changes.every((c) => c.kind === "replacement" && c.authors[0] === "Jane")).toBe(true);

    const curly = await cli("replace", v1, "-o", join(dir, "q.docx"), "--find", "Seller’s", "--replace", "Vendor's");
    expect(curly.code).toBe(0);
  }, 30_000);

  it("replace marks only the words that differ", async () => {
    const out = join(dir, "min.docx");
    const r = await cli("replace", v1, "-o", out, "--find", "within 30 days of invoice", "--replace", "within 60 days of invoice");
    expect(r.code).toBe(0);
    const changes = JSON.parse((await cli("revisions", out, "--json")).out) as { deleted: string; inserted: string }[];
    expect(changes.map((c) => [c.deleted, c.inserted])).toEqual([["30", "60"]]);
    expect((await texts(out))[1]).toBe("1. Payment. Buyer shall pay Seller within 60 days of invoice.");

    const del = join(dir, "del.docx");
    expect((await cli("replace", v1, "-o", del, "--find", "the laws of Texas", "--replace", "Texas")).code).toBe(0);
    expect((await texts(del))[2]).toBe("2. Law. This Agreement is governed by Texas.");
  }, 30_000);

  it("several --find/--replace pairs apply in sequence", async () => {
    const out = join(dir, "pairs.docx");
    const r = await cli("replace", v1, "-o", out, "--find", "30 days", "--replace", "60 days", "--find", "Texas", "--replace", "New York");
    expect(r.code).toBe(0);
    const t = await texts(out);
    expect(t[1]).toContain("60 days");
    expect(t[2]).toContain("New York");
  }, 30_000);

  it("reject --ids --others accept resolves a mixed set", async () => {
    const red = join(dir, "red.docx");
    await cli("replace", v1, "-o", red, "--find", "30 days", "--replace", "60 days", "--find", "Texas", "--replace", "New York");
    const changes = JSON.parse((await cli("revisions", red, "--json")).out) as { ids: string[] }[];
    const out = join(dir, "resolved.docx");
    const r = await cli("reject", red, "-o", out, "--ids", changes[0]!.ids.join(","), "--others", "accept");
    expect(r.code).toBe(0);
    expect(r.out).toContain("0 changes still tracked");
    const t = await texts(out);
    expect(t[1]).toContain("30 days");
    expect(t[2]).toContain("New York");

    const bad = await cli("reject", red, "-o", out, "--ids", "999");
    expect(bad.code).toBe(1);
  }, 30_000);

  it("revisions treats a deletion and insertion split by whitespace as one replacement", async () => {
    const { createBodyDocx } = await import("../fixtures/create-docx.js");
    const a = 'w:author="Counterparty" w:date="2024-01-01T00:00:00Z"';
    const file = join(dir, "ws.docx");
    await writeFile(file, await createBodyDocx(
      `<w:p><w:r><w:t xml:space="preserve">(b) </w:t></w:r><w:del w:id="7" ${a}><w:r><w:delText>Baa3</w:delText></w:r></w:del>` +
      `<w:r><w:t xml:space="preserve"> </w:t></w:r><w:ins w:id="8" ${a}><w:r><w:t>Ba2</w:t></w:r></w:ins>` +
      `<w:r><w:t xml:space="preserve"> by Moody's</w:t></w:r></w:p>`,
    ));
    const changes = JSON.parse((await cli("revisions", file, "--json")).out) as { kind: string; ids: string[]; deleted: string; inserted: string }[];
    expect(changes).toHaveLength(1);
    expect(changes[0]).toMatchObject({ kind: "replacement", ids: ["7", "8"], deleted: "Baa3", inserted: "Ba2" });
  }, 30_000);

  it("replace marks a pure deletion or insertion exactly", async () => {
    const del = join(dir, "pure-del.docx");
    await cli("replace", v1, "-o", del, "--find", "Buyer shall pay Seller within", "--replace", "Buyer shall pay within");
    const d = JSON.parse((await cli("revisions", del, "--json")).out) as { kind: string; deleted: string }[];
    expect(d).toMatchObject([{ kind: "deletion", deleted: "Seller " }]);

    const ins = join(dir, "pure-ins.docx");
    await cli("replace", v1, "-o", ins, "--find", "One year.", "--replace", "One calendar year.");
    const i = JSON.parse((await cli("revisions", ins, "--json")).out) as { kind: string; inserted: string }[];
    expect(i).toMatchObject([{ kind: "insertion", inserted: "calendar " }]);
    expect((await texts(ins))[4]).toBe("4. Term. One calendar year.");
  }, 30_000);

  it("view --grep --context shows neighbours, separated by --", async () => {
    const r = await cli("view", v1, "--grep", "Payment|Term\\.", "--context", "0");
    expect(r.out.trim().split("\n")).toHaveLength(2);
    const c = await cli("view", v1, "--grep", "Law", "--context", "1");
    expect(c.out.trim().split("\n").map((l) => l.slice(0, 4))).toEqual(["[p1]", "[p2]", "[p3]"]);
  }, 30_000);

  it("edit --dry-run reports each patch and explains failures against the batch's state", async () => {
    const patches = join(dir, "dry.json");
    await writeFile(patches, JSON.stringify([
      { type: "insert_paragraph", target: "p0", position: "after", text: "Preamble." },
      { type: "replace_text", target: "p1", find: "Invoices", replace: "Fees" },
    ]));
    const r = await cli("edit", v1, "--dry-run", patches);
    expect(r.code).toBe(1);
    expect(r.out).toContain("1. ok insert_paragraph p0");
    expect(r.out).toContain("2. FAILS replace_text p1");
    expect(r.out).toContain('p1 currently reads: "1. Payment. Buyer shall pay Seller within 30 days of invoice."');
  }, 30_000);

  it("metadata reads, sets and scrubs core properties", async () => {
    const out = join(dir, "meta.docx");
    const w = await cli("metadata", v1, "-o", out, "--set", "creator=Jane Doe", "--set", "title=Agreement");
    expect(w.code).toBe(0);
    expect((await cli("metadata", out)).out).toContain("creator: Jane Doe");
    const scrubbed = join(dir, "meta2.docx");
    await cli("metadata", out, "-o", scrubbed, "--scrub");
    const after = (await cli("metadata", scrubbed)).out;
    expect(after).not.toContain("Jane Doe");
    expect(after).toContain("title: Agreement");
    expect((await cli("metadata", v1, "-o", out, "--set", "colour=red")).code).toBe(1);
  }, 30_000);

  it("comments lists a comment with its author and anchored passage", async () => {
    const out = join(dir, "cm.docx");
    await cli("comment", v1, "-o", out, "--find", "laws of Texas", "--text", "Prefer New York.\nSee the term sheet.", "--author", "Jane");
    const list = JSON.parse((await cli("comments", out, "--json")).out) as { author: string; text: string; anchor: string; paragraph: string }[];
    expect(list).toMatchObject([{ author: "Jane", text: "Prefer New York.\nSee the term sheet.", anchor: "laws of Texas", paragraph: "p2" }]);
    expect((await cli("comments", v1)).out.trim()).toBe("No comments.");
  }, 30_000);

  it("view labels tracked table rows", async () => {
    const { createTableDocx } = await import("../fixtures/create-docx.js");
    const file = join(dir, "tbl.docx");
    await writeFile(file, await createTableDocx());
    const patches = join(dir, "row.json");
    await writeFile(patches, JSON.stringify([{ type: "add_table_row", target: "tbl0", afterRow: 0, cells: ["X", "Y"] }]));
    const out = join(dir, "tbl-row.docx");
    expect((await cli("edit", file, "-o", out, patches, "--author", "Jane")).code).toBe(0);
    const v = await cli("view", out, "--grep", "\\[tbl0\\.r1\\.");
    expect(v.out).toContain("{+row inserted by Jane+}");
    expect((await cli("view", out, "--no-revisions")).out).not.toContain("row inserted");
  }, 30_000);

  it("comments reports the paragraph of a point comment that has only a reference", async () => {
    const JSZip = (await import("jszip")).default;
    const ranged = join(dir, "ranged.docx");
    await cli("comment", v1, "-o", ranged, "--in", "p3", "--text", "Check the address.");
    const zip = await JSZip.loadAsync(await import("node:fs/promises").then((fs) => fs.readFile(ranged)));
    const xml = await zip.file("word/document.xml")!.async("string");
    const stripped = xml.replace(/<w:commentRange(Start|End)\b[^>]*?(\/>|><\/w:commentRange(Start|End)>)/g, "");
    expect(stripped).toContain("w:commentReference");
    expect(stripped).not.toContain("w:commentRangeStart");
    zip.file("word/document.xml", stripped);
    const point = join(dir, "point.docx");
    await writeFile(point, await zip.generateAsync({ type: "nodebuffer" }));
    const list = JSON.parse((await cli("comments", point, "--json")).out) as { paragraph: string; anchor: string }[];
    expect(list).toMatchObject([{ paragraph: "p3", anchor: "" }]);
  }, 30_000);

  it("revisions lists tracked changes outside paragraphs too, and their ids resolve", async () => {
    const changes = JSON.parse((await cli("revisions", lease, "--json")).out) as { kind: string; ids: string[]; context: string }[];
    const note = changes.find((c) => c.context.startsWith("footnotes:"));
    expect(note).toBeDefined();
    const out = join(dir, "lease-fn.docx");
    const r = await cli("accept", lease, "-o", out, "--ids", note!.ids.join(","));
    expect(r.code).toBe(0);
    expect(r.out).toMatch(/^accepted 1 revision/);
  }, 30_000);

  it("comment --in accepts a para: id", async () => {
    const out = join(dir, "cp.docx");
    const r = await cli("comment", lease, "-o", out, "--find", "recapture", "--in", "para:0A000007", "--text", "No recapture.");
    expect(r.code).toBe(0);
    expect(r.out).toContain("comment on [p6]");
  }, 30_000);

  it("comment anchors to unique text", async () => {
    const out = join(dir, "c.docx");
    const r = await cli("comment", v1, "-o", out, "--find", "laws of Texas", "--text", "Prefer New York.", "--author", "Jane");
    expect(r.code).toBe(0);
    const doc = await Docxtor.load(out);
    expect(doc.getAST().xmlParts.comments).toBeDefined();
    expect(await texts(out)).toEqual(await texts(v1));
  }, 30_000);

  it("edit applies a JSON patch array atomically", async () => {
    const patches = join(dir, "p.json");
    await writeFile(patches, JSON.stringify([
      { type: "insert_paragraph", target: "p2", position: "after", text: "2A. New clause." },
      { type: "replace_text", target: "p99", find: "x", replace: "y" },
    ]));
    const out = join(dir, "e.docx");
    const bad = await cli("edit", v1, "-o", out, patches);
    expect(bad.code).toBe(1);
    expect(bad.err).toContain("nothing was written");

    await writeFile(patches, JSON.stringify([{ type: "insert_paragraph", target: "p2", position: "after", text: "2A. New clause." }]));
    const ok = await cli("edit", v1, "-o", out, patches);
    expect(ok.code).toBe(0);
    expect((await texts(out))[3]).toBe("2A. New clause.");
  }, 30_000);
});
