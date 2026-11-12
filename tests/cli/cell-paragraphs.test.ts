/**
 * The CLI on table cells with several paragraphs and on empty cells: ids in
 * `view`, "not found" errors that say where the text is, and filling a
 * blank cell through `edit`.
 */

import { describe, it, expect, beforeAll } from "vitest";
import { execFile } from "node:child_process";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { Docxtor } from "../../src/index.js";
import { createBodyDocx } from "../fixtures/create-docx.js";
import { allParagraphs, explainNoMatch, findAcrossBreak } from "../../cli/paragraphs.js";

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

const cell = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const para = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/**
 * - p0..p2: body paragraphs; "Seller" is in p1 and p2 only
 * - tbl0 r0: a cell with "Heading" and "Body text, Version 2.1", and an empty cell
 * - tbl0 r1: "[Drafting note: Select one option.]" and "( x ) Date of last signature" in one cell; "Date:"
 * - tbl0 r2: "Signed:" and an empty cell
 */
const BODY =
  para("AGREEMENT") + para("The Seller delivers.") + para("The Seller invoices.") +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
  `<w:tr>${cell(para("Heading") + para("Body text, Version 2.1"))}${cell("<w:p/>")}</w:tr>` +
  `<w:tr>${cell(para("[Drafting note: Select one option.]") + para("( x ) Date of last signature"))}${cell(para("Date:"))}</w:tr>` +
  `<w:tr>${cell(para("Signed:"))}${cell("<w:p/>")}</w:tr>` +
  `</w:tbl><w:p/>`;

let dir: string;
let file: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "docxtor-cells-"));
  file = join(dir, "cells.docx");
  await writeFile(file, await createBodyDocx(BODY));
}, 30_000);

describe("view", () => {
  it("prints an id before every paragraph of a cell", async () => {
    const r = await cli("view", file);
    expect(r.out).toContain("| [tbl0.r0.c0.p0] Heading [tbl0.r0.c0.p1] Body text, Version 2.1 | [tbl0.r0.c1.p0]  |");
    expect(r.out).toContain("| [tbl0.r2.c0.p0] Signed: | [tbl0.r2.c1.p0]  |");
  }, 30_000);

  it("--grep on a word in a later paragraph returns the whole row, ids included", async () => {
    const r = await cli("view", file, "--grep", "Version 2\\.1");
    expect(r.out.trim()).toBe("| [tbl0.r0.c0.p0] Heading [tbl0.r0.c0.p1] Body text, Version 2.1 | [tbl0.r0.c1.p0]  |");
  }, 30_000);
});

describe("comment", () => {
  it("anchors on text in the second paragraph of a cell with --in", async () => {
    const out = join(dir, "c1.docx");
    const r = await cli("comment", file, "-o", out, "--in", "tbl0.r0.c0.p1", "--find", "Body", "--text", "Check this");
    expect(r.code).toBe(0);
    expect(r.out).toContain('comment on [tbl0.r0.c0.p1] anchored to "Body"');
  }, 30_000);

  it("says which paragraph holds text that is not in --in", async () => {
    const r = await cli("comment", file, "-o", join(dir, "c2.docx"), "--in", "tbl0.r0.c0.p0", "--find", "Version 2.1", "--text", "Check this");
    expect(r.code).toBe(1);
    expect(r.err).toContain('"Version 2.1" is not in tbl0.r0.c0.p0; it is in tbl0.r0.c0.p1. Use --in tbl0.r0.c0.p1.');
  }, 30_000);

  it("names --in when the text is nowhere", async () => {
    const r = await cli("comment", file, "-o", join(dir, "c3.docx"), "--in", "p0", "--find", "Nowhere", "--text", "Check this");
    expect(r.code).toBe(1);
    expect(r.err).toContain('"Nowhere" not found in p0; nothing was written.');
    expect(r.err).toContain("view FILE --grep");
  }, 30_000);
});

describe("replace", () => {
  it("lists every paragraph that holds text which is not in --in", async () => {
    const r = await cli("replace", file, "-o", join(dir, "r1.docx"), "--in", "p0", "--find", "Seller", "--replace", "Vendor");
    expect(r.code).toBe(1);
    expect(r.err).toContain('"Seller" is not in p0; it matches 2 times elsewhere. Use --in with one of these ids:');
    expect(r.err).toContain("  1. [p1] The Seller delivers.");
    expect(r.err).toContain("  2. [p2] The Seller invoices.");
    expect(r.err).toContain("Nothing was written.");
  }, 30_000);

  it("says a find copied across two paragraphs of a cell must be split", async () => {
    const r = await cli("replace", file, "-o", join(dir, "r2.docx"),
      "--find", "[Drafting note: Select one option.] ( x ) Date of last signature", "--replace", "( x ) Effective Date");
    expect(r.code).toBe(1);
    expect(r.err).toContain("runs across two paragraphs (tbl0.r1.c0.p0 and tbl0.r1.c0.p1); a find must stay inside one paragraph. Split it into two --find/--replace pairs.");
  }, 30_000);

  it("keeps the error for text that is nowhere", async () => {
    const r = await cli("replace", file, "-o", join(dir, "r3.docx"), "--in", "p0", "--find", "Nowhere", "--replace", "x");
    expect(r.code).toBe(1);
    expect(r.err).toContain(`"Nowhere" not found in p0; nothing was written. Text in a docx may differ from what you expect`);
    const all = await cli("replace", file, "-o", join(dir, "r3.docx"), "--find", "Nowhere", "--replace", "x");
    expect(all.err).toContain(`"Nowhere" not found; nothing was written.`);
  }, 30_000);
});

describe("findAcrossBreak and explainNoMatch", () => {
  it("matches only neighbours in one container, whatever spaces are typed at the break", async () => {
    const every = allParagraphs(await Docxtor.load(file));
    const ids = (find: string): string[] | undefined => findAcrossBreak(every, find)?.map((p) => p.id);
    expect(ids("option.]  ( x ) Date")).toEqual(["tbl0.r1.c0.p0", "tbl0.r1.c0.p1"]);
    expect(ids("option.]( x )")).toEqual(["tbl0.r1.c0.p0", "tbl0.r1.c0.p1"]);
    expect(ids("AGREEMENT The Seller")).toEqual(["p0", "p1"]);
    // Two cells of one row, and the body paragraph before a table, are not one container
    expect(ids("signature Date:")).toBeUndefined();
    expect(ids("invoices. Heading")).toBeUndefined();
    // Text inside one paragraph, or absent, is not a break
    expect(ids("Seller delivers")).toBeUndefined();
    expect(ids("Nowhere at all")).toBeUndefined();
  });

  it("with whole-word, the find must start and end at a word boundary", async () => {
    const doc = await Docxtor.load(await createBodyDocx(para("XAgreement") + para("SellerY") + para("the Agreement") + para("Seller pays")));
    const every = allParagraphs(doc);
    expect(findAcrossBreak(every, "Agreement Seller")?.map((p) => p.id)).toEqual(["p0", "p1"]);
    expect(findAcrossBreak(every, "Agreement Seller", true)?.map((p) => p.id)).toEqual(["p2", "p3"]);
    expect(explainNoMatch(every.slice(0, 2), "Agreement Seller", true, undefined, "Split it.")).toBeUndefined();
  });

  it("does not pair two paragraphs of a cell that have a nested table between them", async () => {
    const nested = `<w:tbl><w:tr>${cell(para("inner"))}</w:tr></w:tbl>`;
    const doc = await Docxtor.load(await createBodyDocx(
      `<w:tbl><w:tr>${cell(para("first part") + nested + para("second part"))}${cell(para("left side") + para("right side"))}</w:tr></w:tbl><w:p/>`,
    ));
    const every = allParagraphs(doc);
    expect(findAcrossBreak(every, "first part second part")).toBeUndefined();
    expect(findAcrossBreak(every, "left side right side")?.map((p) => p.id)).toEqual(["tbl0.r0.c1.p0", "tbl0.r0.c1.p1"]);
  });

  it("returns nothing when the text is nowhere", async () => {
    const every = allParagraphs(await Docxtor.load(file));
    expect(explainNoMatch(every, "Nowhere", false, "p0", "Split it.")).toBeUndefined();
    expect(explainNoMatch(every, "Nowhere", false, undefined, "Split it.")).toBeUndefined();
  });
});

describe("edit", () => {
  it("fills blank cells with insert_text and no anchor, tracked", async () => {
    const out = join(dir, "e1.docx");
    const patches = join(dir, "e1.json");
    await writeFile(patches, JSON.stringify([
      { type: "insert_text", target: "tbl0.r2.c1.p0", position: "after", text: "June 22, 2026" },
      { type: "insert_text", target: "tbl0.r0.c1.p0", anchor: "", position: "end", text: "Acme Inc." },
    ]));
    const r = await cli("edit", file, "-o", out, patches, "--author", "Jane");
    expect(r.code).toBe(0);
    const v = await cli("view", out);
    expect(v.out).toContain("| [tbl0.r2.c0.p0] Signed: | [tbl0.r2.c1.p0] {+June 22, 2026+} |");
    expect(v.out).toContain("[tbl0.r0.c1.p0] {+Acme Inc.+} |");
  }, 30_000);

  it("--untracked writes the fill plain", async () => {
    const out = join(dir, "e2.docx");
    const patches = join(dir, "e2.json");
    await writeFile(patches, JSON.stringify([{ type: "insert_text", target: "tbl0.r2.c1.p0", position: "after", text: "June 22, 2026" }]));
    const r = await cli("edit", file, "-o", out, patches, "--untracked");
    expect(r.code).toBe(0);
    expect((await cli("view", out)).out).toContain("[tbl0.r2.c1.p0] June 22, 2026 |");
    expect((await cli("revisions", out)).out).not.toContain("June 22");
  }, 30_000);

  it("a patch whose find is in a neighbouring paragraph says where it is", async () => {
    const patches = join(dir, "e3.json");
    await writeFile(patches, JSON.stringify([{ type: "replace_text", target: "tbl0.r0.c0.p0", find: "Version 2.1", replace: "Version 3" }]));
    const r = await cli("edit", file, "--dry-run", patches);
    expect(r.code).toBe(1);
    expect(r.out).toContain("it is in tbl0.r0.c0.p1 (use that as the target)");
    expect(r.out).toContain('tbl0.r0.c0.p0 currently reads: "Heading"');
  }, 30_000);
});
