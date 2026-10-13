/**
 * §3.7 Colour filter on projection (reading red-text answers).
 */

import { describe, it, expect } from "vitest";
import { Docxtor } from "../../src/index.js";
import { createLeaseDocx, createFormattedDocx } from "../fixtures/create-docx.js";

type ParaProjection = {
  id: string;
  paraId?: string;
  text: string;
  formatting?: Array<{ text: string; startOffset: number; endOffset: number; color?: string; themeColor?: string; bold?: boolean }>;
};

describe("runFilter", () => {
  it("returns only runs of the given colour with ids and offsets; black text excluded", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const json = doc.project("json", { runFilter: { color: "FF0000" } });

    const withRuns = (json.body as ParaProjection[]).filter((p) => p.formatting && p.formatting.length > 0);
    expect(withRuns.map((p) => p.id)).toEqual(["p9"]);

    const p9 = withRuns[0]!;
    expect(p9.paraId).toBe("0A00000A");
    expect(p9.text).toBe("Security deposit amount: two months. Guarantor: Parent Co.");
    expect(p9.formatting).toEqual([
      { text: "two months", startOffset: 25, endOffset: 35, color: "FF0000" },
      { text: "Parent Co", startOffset: 48, endOffset: 57, color: "FF0000", themeColor: "accent2" },
    ]);
    for (const run of p9.formatting!) {
      expect(p9.text.slice(run.startOffset, run.endOffset)).toBe(run.text);
    }

    // Paragraphs with no matching runs carry no formatting array at all
    const p4 = (json.body as ParaProjection[]).find((p) => p.id === "p4")!;
    expect(p4.formatting).toBeUndefined();
  });

  it("is case-insensitive on the colour value", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const json = doc.project("json", { runFilter: { color: "ff0000" } });
    const p9 = (json.body as ParaProjection[]).find((p) => p.id === "p9")!;
    expect(p9.formatting!.length).toBe(2);
  });

  it("reports matching runs inside table cells", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const json = doc.project("json", { runFilter: { color: "FF0000" } });
    const table = json.body.find((el) => "rows" in el) as { rows: Array<Array<{ id: string; paragraphs?: ParaProjection[] }>> };
    const cell = table.rows[1]![1]!;
    expect(cell.id).toBe("tbl0.r1.c1.p0");
    expect(cell.paragraphs![0]!.formatting).toEqual([
      { text: "5%", startOffset: 10, endOffset: 12, color: "FF0000" },
    ]);
    const plainCell = table.rows[1]![0]!;
    expect(plainCell.paragraphs![0]!.formatting).toBeUndefined();
  });

  it("filters on other properties too, requiring all to match", async () => {
    // createFormattedDocx p1: "This is " + bold + " and " + italic + " text."; p2: red size 28
    const doc = await Docxtor.load(await createFormattedDocx());
    const bold = doc.project("json", { runFilter: { bold: true } });
    expect((bold.body as ParaProjection[]).map((p) => p.formatting?.map((r) => r.text) ?? [])).toEqual([
      ["Document Title"],
      ["bold"],
      [],
    ]);
    const boldRed = doc.project("json", { runFilter: { bold: true, color: "FF0000" } });
    expect((boldRed.body as ParaProjection[]).every((p) => p.formatting === undefined)).toBe(true);
  });

  it("does not change the default projection", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const json = doc.project("json");
    const p9 = (json.body as ParaProjection[]).find((p) => p.id === "p9")!;
    expect(p9.formatting).toBeUndefined();
    const withFormatting = doc.project("json", { includeFormatting: true });
    const p9f = (withFormatting.body as ParaProjection[]).find((p) => p.id === "p9")!;
    expect(p9f.formatting!.length).toBe(2); // only formatted runs, as before
  });
});
