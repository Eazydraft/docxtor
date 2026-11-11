/**
 * Content controls (`w:sdt`): reading through the wrapper, filling an empty
 * control, and the typed controls (drop-down list, date, checkbox), locks
 * and data bindings.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor, type NormalizedTable, type Patch, type PatchOptions } from "../../src/index.js";
import { getChildren, getTagName } from "../../src/xml/helpers.js";
import { parseDate } from "../../src/patches/control-ops.js";
import { createContentControlDocx, createTableDocx, createBodyDocx } from "../fixtures/create-docx.js";

const load = async (): Promise<Docxtor> => Docxtor.load(await createContentControlDocx());

const TRACKED: PatchOptions = { trackRevisions: true, revisionAuthor: "Reviewer", revisionDate: "2026-11-11T00:00:00Z" };

/** Apply one patch and return its result. */
function apply(doc: Docxtor, patch: Patch, options?: PatchOptions) {
  return doc.applyPatches([patch], options)[0]!;
}

/** The serialized `<w:sdt>…</w:sdt>` whose `w:id` is `id`. */
function sdt(xml: string, id: string): string {
  const at = xml.indexOf(`<w:id w:val="${id}"/>`);
  expect(at, `control ${id}`).toBeGreaterThan(-1);
  const start = xml.lastIndexOf("<w:sdt>", at);
  // The matching close: step over nested controls
  let depth = 0;
  const re = /<w:sdt>|<\/w:sdt>/g;
  re.lastIndex = start;
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    depth += m[0] === "<w:sdt>" ? 1 : -1;
    if (depth === 0) return xml.slice(start, m.index + m[0].length);
  }
  throw new Error(`control ${id} is not closed`);
}

const sdtPr = (xml: string, id: string): string => /<w:sdtPr>.*?<\/w:sdtPr>/.exec(sdt(xml, id))![0];
const sdtContent = (xml: string, id: string): string => {
  const s = sdt(xml, id);
  return s.slice(s.indexOf("<w:sdtContent>"), s.lastIndexOf("</w:sdtContent>") + "</w:sdtContent>".length);
};

// =============================================================================
// Stage 1: reading through the wrapper
// =============================================================================

describe("content controls in the view", () => {
  it("shows the text of every control, prompts marked as fields", async () => {
    const md = (await load()).project("markdown");
    expect(md).toContain("[p0] Title: {field: Insert Title for Contract} ref ACME-42");
    // A filled rich-text or plain-text control is content, nested or not
    expect(md).toContain("[p1] Contact: Jane Roe Dept Legal");
    expect(md).toContain("[p2] Scope of work.\n[p3] Deliverables follow.");
    expect(md).toContain("[p4] ☐ Agency has implemented the plan");
    expect(md).toContain("[p7] {field: Click here to enter text.}");
    expect(md).toContain('[p8] Mode: {field dropdown "One-off purchase" OR "Ongoing supply"= Ongoing supply}');
    expect(md).toContain('[p9] Colour: {field combo "Red" OR "Blue": Pick or type}');
    expect(md).toContain("[hdr.default.p0] Doc ref: {field: Insert ref}");
  });

  it("shows a wrapped cell as a cell, with its control", async () => {
    const md = (await load()).project("markdown");
    expect(md).toContain("| [tbl0.r0.c0.p0] Start date | [tbl0.r0.c1.p0] {field date: Select date} |");
    expect(md).toContain(
      '| [tbl0.r1.c0.p0] Type of contract | [tbl0.r1.c1.p0] {field dropdown "One-off purchase" OR "Ongoing supply": Choose an item.} |',
    );
  });

  it("counts as many cells per row as the XML has w:tc, wrapped or not", async () => {
    const doc = await load();
    const tcCount = (node: Parameters<typeof getChildren>[0]): number =>
      getChildren(node).reduce(
        (n, c) => n + (getTagName(c) === "w:tc" ? 1 : getTagName(c) === "w:sdt" || getTagName(c) === "w:sdtContent" ? tcCount(c) : 0),
        0,
      );
    const table = doc.findTableById("tbl0") as NormalizedTable;
    expect(table.rows.map((r) => r.cells.length)).toEqual([2, 2, 2]);
    for (const row of table.rows) expect(row.cells.length).toBe(tcCount(row.xmlRef));
    expect(table.rows[0]!.cells[1]!.control?.kind).toBe("date");
    expect(table.rows[2]!.cells[1]!.control).toBeUndefined();
  });

  it("keeps marker nesting with tracked changes and formatting shown", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p8", find: "Ongoing supply", replace: "One-off purchase" }, TRACKED).success).toBe(true);
    const md = doc.project("markdown", { includeRevisions: true, includeFormatting: true });
    expect(md).toContain(
      '[p8] Mode: {field dropdown "One-off purchase" OR "Ongoing supply"= {-Ongoing supply-}{+One-off purchase+}}',
    );
  });

  it("puts every control and its range in the JSON projection", async () => {
    const json = (await load()).project("json");
    const p0 = json.body[0] as { text: string; controls: { id: string; kind: string; placeholder: boolean; startOffset: number; endOffset: number; alias?: string }[] };
    expect(p0.text).toBe("Title: Insert Title for Contract ref ACME-42");
    expect(p0.controls).toEqual([
      { id: "101", kind: "richText", placeholder: true, startOffset: 7, endOffset: 32, alias: "Title", tag: "title" },
      { id: "102", kind: "richText", placeholder: false, startOffset: 37, endOffset: 44, alias: "Reference" },
    ]);

    // A nested control follows the one that holds it
    const p1 = json.body[1] as { controls: { id: string; startOffset: number; endOffset: number }[] };
    expect(p1.controls.map((c) => [c.id, c.startOffset, c.endOffset])).toEqual([
      ["103", 9, 17],
      ["104", 18, 28],
      ["105", 23, 28],
    ]);

    // A control around a cell is reported on the cell's paragraph, over all of its text
    const table = json.body[2] as { rows: { paragraphs?: { controls?: unknown[] }[] }[][] };
    expect(table.rows[1]![1]!.paragraphs![0]!.controls).toEqual([
      {
        id: "107", kind: "dropDownList", placeholder: true, startOffset: 0, endOffset: 15,
        alias: "Type of contract", options: ["One-off purchase", "Ongoing supply"],
      },
    ]);

    const p4 = json.body.find((e) => e.id === "p4") as { controls: { kind: string; checked: boolean }[] };
    expect(p4.controls[0]).toMatchObject({ kind: "checkbox", checked: false });
    const p5 = json.body.find((e) => e.id === "p5") as { controls: { locked: string }[] };
    expect(p5.controls[0]!.locked).toBe("content");
    const p6 = json.body.find((e) => e.id === "p6") as { controls: { bound: boolean }[] };
    expect(p6.controls[0]!.bound).toBe(true);
  });

  it("numbers paragraphs inside a block-level control like any other", async () => {
    const doc = await load();
    expect(doc.getParagraphs().map((p) => p.id)).toEqual(["p0", "p1", "p2", "p3", "p4", "p5", "p6", "p7", "p8", "p9"]);
    expect(doc.findParagraphById("p2")!.control?.id).toBe("109");
    expect(doc.findParagraphById("p3")!.control).toBe(doc.findParagraphById("p2")!.control);
    expect(doc.findParagraphById("p4")!.control).toBeUndefined();
  });
});

describe("text edits inside a filled control", () => {
  it("replace_text changes the run inside w:sdtContent and nothing else", async () => {
    const doc = await load();
    const before = doc.toString();
    expect(apply(doc, { type: "replace_text", target: "p0", find: "ACME-42", replace: "ACME-43" }).success).toBe(true);
    const after = doc.toString();

    expect(sdtPr(after, "102")).toBe(sdtPr(before, "102"));
    expect(sdtContent(after, "102")).toBe(`<w:sdtContent><w:r><w:t xml:space="preserve">ACME-43</w:t></w:r></w:sdtContent>`);
    // Only that text changed: no run landed outside the control
    expect(after).toBe(before.replace("ACME-42", "ACME-43"));
  });

  it("a partial replacement splits the run inside the control", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p1", find: "Jane", replace: "John" }).success).toBe(true);
    expect(doc.findParagraphById("p1")!.text).toBe("Contact: John Roe Dept Legal");
    const content = sdtContent(doc.toString(), "103");
    expect(content.match(/<w:r>/g)).toHaveLength(2);
    expect(content).toContain("John");
    expect(content).toContain(" Roe");
  });

  it("edits a nested control in place", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p1", find: "Legal", replace: "Finance" }).success).toBe(true);
    const xml = doc.toString();
    expect(sdtContent(xml, "105")).toContain("Finance");
    expect(sdt(xml, "104")).toContain(sdt(xml, "105"));
  });

  it("a tracked edit then accept-all equals the untracked edit; reject-all restores the input", async () => {
    const patch: Patch = { type: "replace_text", target: "p1", find: "Jane Roe", replace: "John Doe" };
    const input = (await load()).toString();

    const direct = await load();
    apply(direct, patch);

    const accepted = await load();
    expect(apply(accepted, patch, TRACKED).success).toBe(true);
    const tracked = accepted.toString();
    expect(sdtContent(tracked, "103")).toMatch(/^<w:sdtContent><w:del [^>]*>.*<\/w:del><w:ins [^>]*>.*<\/w:ins><\/w:sdtContent>$/);
    accepted.acceptAllRevisions();
    // The same XML, but for xml:space on a w:t that no longer needs it
    const plain = (xml: string): string => xml.replaceAll(' xml:space="preserve"', "");
    expect(plain(accepted.toString())).toBe(plain(direct.toString()));

    const rejected = await load();
    apply(rejected, patch, TRACKED);
    rejected.rejectAllRevisions();
    expect(rejected.toString()).toBe(input);
  });

  it("refuses to clear a control", async () => {
    const doc = await load();
    const r = apply(doc, { type: "delete_text", target: "p0", find: "ACME-42" });
    expect(r.success).toBe(false);
    expect(r.error).toBe(
      'p0: clearing a content control is not supported (the edit would leave the content control "Reference" empty)',
    );
  });

  it("refuses a match that crosses the edge of a control", async () => {
    const doc = await load();
    const r = apply(doc, { type: "replace_text", target: "p0", find: "ref ACME", replace: "no. ACME" });
    expect(r.success).toBe(false);
    expect(r.error).toContain('"ref ACME" crosses the edge of a content control "Reference"');
    // Text that only touches the control from outside is ordinary text
    expect(apply(doc, { type: "replace_text", target: "p0", find: " ref ", replace: " reference " }).success).toBe(true);
    expect(sdtContent(doc.toString(), "102")).toContain("ACME-42");
  });

  it("comments and formatting reach text in a control", async () => {
    const doc = await load();
    const results = doc.applyPatches([
      { type: "add_comment", target: "p1", find: "Jane Roe", text: "Check the name" },
      { type: "apply_formatting", target: "p0", find: "ACME-42", formatting: { bold: true } },
    ]);
    expect(results.map((r) => r.success)).toEqual([true, true]);
    const xml = doc.toString();
    expect(sdtContent(xml, "103")).toMatch(/<w:commentRangeStart [^>]*\/><w:r>.*<w:commentRangeEnd /);
    expect(sdtContent(xml, "102")).toContain("<w:b/>");
  });
});

describe("edits a control refuses", () => {
  it.each([
    [
      "a locked control",
      { type: "replace_text", target: "p5", find: "Locked", replace: "Open" },
      'p5: the text is in a locked content control "Fixed"; its content cannot be edited',
    ],
    [
      "a data-bound control",
      { type: "replace_text", target: "p6", find: "Bound value", replace: "Acme Ltd" },
      'p6: the text is in a content control "Company" bound to a custom XML part; Word overwrites its content from that part when the file is opened, so an edit here would be lost',
    ],
    [
      "a drop-down value that is not in the list",
      { type: "replace_text", target: "tbl0.r1.c1.p0", find: "Choose an item.", replace: "Lease" },
      'tbl0.r1.c1.p0: "Lease" is not an option of the dropdown content control "Type of contract"; choose one of: "One-off purchase", "Ongoing supply"',
    ],
    [
      "deleting a prompt",
      { type: "delete_text", target: "p0", find: "Insert Title" },
      'p0: "Insert Title for Contract" is the placeholder of an empty content control "Title"; there is nothing to delete',
    ],
    [
      "formatting in a locked control",
      { type: "apply_formatting", target: "p5", find: "Locked", formatting: { bold: true } },
      'p5: "Locked text" is in a locked content control "Fixed"; its content cannot be edited',
    ],
    [
      "a paragraph added to an empty block control",
      { type: "insert_paragraph", target: "p7", position: "after", text: "More" },
      'p7: the paragraph is the placeholder of an empty content control "Background"; fill the control with replace_text before adding content to it',
    ],
  ] as [string, Patch, string][])("%s fails validation with a reason", async (_name, patch, message) => {
    const doc = await load();
    const before = doc.toString();
    expect(doc.validatePatches([patch])).toEqual([{ valid: false, error: message }]);
    const r = apply(doc, patch, TRACKED);
    expect(r.success).toBe(false);
    expect(r.error).toBe(message);
    expect(doc.toString()).toBe(before);
  });

  it("text outside a locked control in the same paragraph is still editable", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p5", find: "Fixed", replace: "Set" }).success).toBe(true);
    expect(doc.findParagraphById("p5")!.text).toBe("Set: Locked text");
  });
});

describe("table rows with a wrapped cell", () => {
  it("add_table_row takes the full cell count and clones no control", async () => {
    const doc = await load();
    const short = apply(doc, { type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["Only one"] });
    expect(short.error).toBe("Cell count mismatch: patch has 1 cells, but row has 2 cells");

    const r = apply(doc, { type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["Renewal", "Two years"] });
    expect(r.success).toBe(true);

    const xml = doc.toString();
    const ids = [...xml.matchAll(/<w:sdtPr>.*?<w:id w:val="(\d+)"\/>/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(xml.match(/<w:id w:val="107"\/>/g)).toHaveLength(1);

    const table = doc.findTableById("tbl0")!;
    expect(table.rows).toHaveLength(4);
    expect(table.rows[2]!.cells.map((c) => c.paragraphs[0]!.text)).toEqual(["Renewal", "Two years"]);
    expect(table.rows[2]!.cells[1]!.control).toBeUndefined();
    // The source row keeps its control and its prompt
    expect(table.rows[1]!.cells[1]!.paragraphs[0]!.text).toBe("Choose an item.");
    expect(xml).not.toContain("Two yearsChoose");
  });

  it("delete_table_row removes a row whose cell is wrapped", async () => {
    const doc = await load();
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 0 }).success).toBe(true);
    expect(doc.findTableById("tbl0")!.rows.map((r) => r.cells[0]!.paragraphs[0]!.text)).toEqual(["Type of contract", "Notes"]);
    expect(doc.toString()).not.toContain('<w:id w:val="106"/>');
  });

  it("tables without controls are unchanged", async () => {
    const doc = await Docxtor.load(await createTableDocx());
    const table = doc.getTables()[0]!;
    expect(apply(doc, { type: "add_table_row", target: table.id, afterRow: 0, cells: table.rows[0]!.cells.map((_, i) => `c${i}`) }).success).toBe(true);
  });
});

describe("paragraphs inside a block-level control", () => {
  it("insert_paragraph lands inside the same w:sdtContent", async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_paragraph", target: "p2", position: "after", text: "Milestones are monthly." }).success).toBe(true);
    const content = sdtContent(doc.toString(), "109");
    expect(content.match(/<w:p[ >]/g)).toHaveLength(3);
    expect(content.indexOf("Scope of work.")).toBeLessThan(content.indexOf("Milestones are monthly."));
    expect(content.indexOf("Milestones are monthly.")).toBeLessThan(content.indexOf("Deliverables follow."));
    // The new paragraph belongs to the control like its neighbours
    const added = doc.getParagraphs().find((p) => p.text === "Milestones are monthly.")!;
    expect(added.control).toBe(doc.findParagraphById("p2")!.control);
  });

  it("add_table after a paragraph in a control goes next to it", async () => {
    const doc = await load();
    expect(apply(doc, { type: "add_table", afterId: "p3", rows: [["a", "b"]] }).success).toBe(true);
    const content = sdtContent(doc.toString(), "109");
    expect(content).toMatch(/Deliverables follow\.<\/w:t><\/w:r><\/w:p><w:tbl>/);
  });

  it("delete_paragraph of the only paragraph in a control leaves an empty paragraph", async () => {
    const doc = await load();
    const results = doc.applyPatches([
      { type: "delete_paragraph", target: "p3" },
      { type: "delete_paragraph", target: "p2" },
    ]);
    expect(results.map((r) => r.success)).toEqual([true, true]);
    expect(sdtContent(doc.toString(), "109")).toBe("<w:sdtContent><w:p/></w:sdtContent>");
    expect(doc.findParagraphById("p2")!.text).toBe("");
  });
});

// =============================================================================
// Stage 2: filling an empty control
// =============================================================================

describe("filling an empty control", () => {
  it("replaces the whole prompt, clears the flag and keeps w:placeholder", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p0", find: "Insert Title for Contract", replace: "Supply of Widgets" }).success).toBe(true);
    const xml = doc.toString();

    expect(sdtPr(xml, "101")).toBe(
      `<w:sdtPr><w:rPr><w:sz w:val="22"/></w:rPr><w:alias w:val="Title"/><w:tag w:val="title"/><w:id w:val="101"/>` +
        `<w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder></w:sdtPr>`,
    );
    // One run, with the control's own run properties, never the prompt style
    expect(sdtContent(xml, "101")).toBe(
      `<w:sdtContent><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>Supply of Widgets</w:t></w:r></w:sdtContent>`,
    );

    const p0 = doc.findParagraphById("p0")!;
    expect(p0.text).toBe("Title: Supply of Widgets ref ACME-42");
    expect(p0.controls![0]!.control.placeholder).toBe(false);
    expect(doc.project("markdown")).toContain("[p0] Title: Supply of Widgets ref ACME-42");
  });

  it("a partial match on the prompt still clears the whole prompt", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p0", find: "Insert", replace: "Supply" }).success).toBe(true);
    expect(doc.findParagraphById("p0")!.text).toBe("Title: Supply ref ACME-42");
    expect(sdtContent(doc.toString(), "101")).not.toContain("PlaceholderText");
  });

  it("insert_text anchored on the prompt fills the control too", async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_text", target: "hdr.default.p0", anchor: "Insert ref", position: "after", text: "HR-7" }).success).toBe(true);
    expect(doc.findParagraphById("hdr.default.p0")!.text).toBe("Doc ref: HR-7");
    const header = await (await JSZip.loadAsync(await doc.toBuffer())).file("word/header1.xml")!.async("string");
    expect(header).not.toContain("showingPlcHdr");
    expect(sdtContent(header, "116")).toBe("<w:sdtContent><w:r><w:t>HR-7</w:t></w:r></w:sdtContent>");
  });

  it("a filled control has no run properties of its own unless the control gives some", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p7", find: "Click here to enter text.", replace: "Background of the engagement." }).success).toBe(true);
    const xml = doc.toString();
    expect(sdtContent(xml, "113")).toBe("<w:sdtContent><w:p><w:r><w:t>Background of the engagement.</w:t></w:r></w:p></w:sdtContent>");
    expect(sdtPr(xml, "113")).not.toContain("showingPlcHdr");
    // The control now holds content, so it takes another paragraph
    expect(apply(doc, { type: "insert_paragraph", target: "p7", position: "after", text: "More." }).success).toBe(true);
  });

  it("tracked: one insertion and no deletion; reject-all leaves the control empty", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p0", find: "Insert Title for Contract", replace: "Supply of Widgets" }, TRACKED).success).toBe(true);
    const xml = doc.toString();
    expect(sdtContent(xml, "101")).toBe(
      `<w:sdtContent><w:ins w:id="1" w:author="Reviewer" w:date="2026-11-11T00:00:00Z">` +
        `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>Supply of Widgets</w:t></w:r></w:ins></w:sdtContent>`,
    );
    expect(xml).not.toContain("<w:del ");
    expect(doc.project("markdown", { includeRevisions: true })).toContain("[p0] Title: {+Supply of Widgets+} ref ACME-42");

    doc.rejectAllRevisions();
    expect(sdt(doc.toString(), "101")).toMatch(/<\/w:sdtPr><w:sdtContent\/><\/w:sdt>$/);
    expect(doc.findParagraphById("p0")!.text).toBe("Title:  ref ACME-42");
    // An empty control is still shown as a field
    expect(doc.project("markdown")).toContain("[p0] Title: {field:} ref ACME-42");
  });

  it("tracked then accept-all equals the untracked fill", async () => {
    const patch: Patch = { type: "replace_text", target: "p0", find: "Insert Title for Contract", replace: "Supply of Widgets" };
    const direct = await load();
    apply(direct, patch);
    const tracked = await load();
    apply(tracked, patch, TRACKED);
    tracked.acceptAllRevisions();
    expect(tracked.toString()).toBe(direct.toString());
  });
});

// =============================================================================
// Stage 3: typed controls
// =============================================================================

describe("drop-down lists and combo boxes", () => {
  it("a listed value fills the control and sets w:lastValue", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "tbl0.r1.c1.p0", find: "Choose an item.", replace: "Ongoing supply" }).success).toBe(true);
    const xml = doc.toString();
    expect(sdtPr(xml, "107")).toContain('<w:dropDownList w:lastValue="ongoing">');
    expect(sdtPr(xml, "107")).not.toContain("showingPlcHdr");
    expect(doc.project("markdown")).toContain(
      '[tbl0.r1.c1.p0] {field dropdown "One-off purchase" OR "Ongoing supply"= Ongoing supply} |',
    );
    expect(doc.findTableById("tbl0")!.rows[1]!.cells[1]!.control!.placeholder).toBe(false);
  });

  it("changing a filled drop-down checks the new value and updates w:lastValue", async () => {
    const doc = await load();
    const bad = apply(doc, { type: "replace_text", target: "p8", find: "Ongoing", replace: "Occasional" });
    expect(bad.error).toBe(
      'p8: "Occasional supply" is not an option of the dropdown content control "Mode"; choose one of: "One-off purchase", "Ongoing supply"',
    );
    expect(apply(doc, { type: "replace_text", target: "p8", find: "Ongoing supply", replace: "One-off purchase" }).success).toBe(true);
    expect(sdtPr(doc.toString(), "114")).toContain('<w:dropDownList w:lastValue="one-off">');
  });

  it("a combo box takes a listed value or free text", async () => {
    const listed = await load();
    expect(apply(listed, { type: "replace_text", target: "p9", find: "Pick or type", replace: "Blue" }).success).toBe(true);
    expect(sdtPr(listed.toString(), "115")).toContain('<w:comboBox w:lastValue="b">');

    const free = await load();
    expect(apply(free, { type: "replace_text", target: "p9", find: "Pick or type", replace: "Teal" }).success).toBe(true);
    expect(sdtPr(free.toString(), "115")).toContain('<w:comboBox w:lastValue="Teal">');
    expect(free.findParagraphById("p9")!.text).toBe("Colour: Teal");
  });
});

describe("date pickers", () => {
  it.each([
    ["2026-11-11", "2026-11-11T00:00:00Z"],
    ["11 November 2026", "2026-11-11T00:00:00Z"],
    ["Nov 11, 2026", "2026-11-11T00:00:00Z"],
  ])("%s sets w:fullDate", async (text, full) => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "tbl0.r0.c1.p0", find: "Select date", replace: text }).success).toBe(true);
    const xml = doc.toString();
    expect(sdtPr(xml, "106")).toContain(`<w:date w:fullDate="${full}">`);
    expect(sdtPr(xml, "106")).not.toContain("showingPlcHdr");
    expect(doc.findParagraphById("tbl0.r0.c1.p0")!.text).toBe(text);
  });

  it("free text is written as given, without w:fullDate", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "tbl0.r0.c1.p0", find: "Select date", replace: "On signing" }).success).toBe(true);
    expect(sdtPr(doc.toString(), "106")).toContain("<w:date>");
    expect(doc.project("markdown")).toContain("[tbl0.r0.c1.p0] {field date= On signing} |");
  });

  it("reads unambiguous dates only", () => {
    expect(parseDate("1 May 2026")).toBe("2026-05-01T00:00:00Z");
    expect(parseDate("February 29, 2028")).toBe("2028-02-29T00:00:00Z");
    expect(parseDate("29 February 2026")).toBeUndefined();
    expect(parseDate("03/04/2026")).toBeUndefined();
    expect(parseDate("next Tuesday")).toBeUndefined();
  });
});

describe("checkboxes", () => {
  const boxPr = '<w14:checkbox><w14:checked w14:val="1"/>';

  it("replacing the glyph flips w14:checked with it", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p4", find: "☐", replace: "☒" }).success).toBe(true);
    const xml = doc.toString();
    expect(sdtPr(xml, "110")).toContain(boxPr);
    expect(sdtContent(xml, "110")).toBe(
      `<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:cs="MS Gothic"/></w:rPr><w:t>☒</w:t></w:r></w:sdtContent>`,
    );
    expect(doc.findParagraphById("p4")!.controls![0]!.control.checked).toBe(true);
  });

  it("set_checkbox does the same without the glyph, and back again", async () => {
    const doc = await load();
    expect(apply(doc, { type: "set_checkbox", target: "p4", checked: true }).success).toBe(true);
    expect(doc.findParagraphById("p4")!.text).toBe("☒ Agency has implemented the plan");
    expect(sdtPr(doc.toString(), "110")).toContain(boxPr);

    expect(apply(doc, { type: "set_checkbox", target: "p4", checked: false }).success).toBe(true);
    expect(doc.findParagraphById("p4")!.text).toBe("☐ Agency has implemented the plan");
    expect(sdtPr(doc.toString(), "110")).toContain('<w14:checked w14:val="0"/>');
  });

  it("setting the state a checkbox already has changes nothing", async () => {
    const doc = await load();
    const before = doc.toString();
    expect(apply(doc, { type: "set_checkbox", target: "p4", checked: false }, TRACKED).success).toBe(true);
    expect(doc.toString()).toBe(before);
  });

  it("tracked: the old glyph is a deletion and the new one an insertion", async () => {
    const doc = await load();
    expect(apply(doc, { type: "set_checkbox", target: "p4", checked: true }, TRACKED).success).toBe(true);
    expect(doc.project("markdown", { includeRevisions: true })).toContain("[p4] {-☐-}{+☒+} Agency has implemented the plan");
    expect(sdtContent(doc.toString(), "110")).toMatch(/^<w:sdtContent><w:del [^>]*><w:r>.*<w:delText>☐<\/w:delText><\/w:r><\/w:del><w:ins [^>]*><w:r>.*<w:t>☒<\/w:t><\/w:r><\/w:ins><\/w:sdtContent>$/);
  });

  it("refuses anything that is not one of the two glyphs, and a paragraph without a checkbox", async () => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target: "p4", find: "☐", replace: "yes" }).error).toBe(
      'p4: a checkbox control holds "☐" (unchecked) or "☒" (checked); replace one with the other, or use set_checkbox',
    );
    expect(apply(doc, { type: "set_checkbox", target: "p1", checked: true }).error).toBe("p1: the paragraph has no checkbox control");
    expect(apply(doc, { type: "set_checkbox", target: "p4", checked: true, occurrence: 2 }).error).toBe(
      "p4: occurrence 2 asked for, but the paragraph has 1 checkbox control",
    );
  });
});

// =============================================================================
// Locks on controls that a structural edit would remove or change
// =============================================================================

describe("locks seen through nesting, tables and rows", () => {
  const run = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;
  const control = (id: number, pr: string, content: string): string =>
    `<w:sdt><w:sdtPr>${pr}<w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const cell = (inner: string): string => `<w:tc><w:p>${inner}</w:p></w:tc>`;
  const table = (rows: string): string => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>${rows}</w:tbl>`;
  const from = async (body: string): Promise<Docxtor> => Docxtor.load(await createBodyDocx(body));
  const AGAINST = 'tbl0 row 1: the row holds a content control "Amount" that is locked against deletion';
  const CONTENT = 'tbl0 row 1: the row holds a locked content control "Amount"; it cannot be deleted';

  it("delete_paragraph is refused when a control around an inline control is locked against deletion", async () => {
    const doc = await from(
      `<w:p>${run("Party: ")}${control(1, `<w:alias w:val="Outer"/><w:lock w:val="sdtLocked"/>`, run("Dept ") + control(2, "<w:text/>", run("Legal")))}</w:p>` +
        `<w:p>${run("Next")}</w:p>`,
    );
    for (const options of [undefined, TRACKED]) {
      const r = apply(doc, { type: "delete_paragraph", target: "p0" }, options);
      expect(r.error).toBe('p0: the paragraph holds a content control "Outer" that is locked against deletion');
    }
    expect(doc.getParagraphs()).toHaveLength(2);
  });

  it("delete_paragraph inside a block control locked against deletion is allowed: the control stays", async () => {
    const doc = await from(
      control(1, `<w:lock w:val="sdtLocked"/>`, `<w:p>${run("One ")}${control(2, "", run("inner"))}</w:p><w:p>${run("Two")}</w:p>`),
    );
    expect(apply(doc, { type: "delete_paragraph", target: "p0" }).success).toBe(true);
    expect(doc.toString()).toContain('<w:id w:val="1"/>');
    expect(doc.getParagraphs().map((p) => p.text)).toEqual(["Two"]);
  });

  it("add_table and add_list after a table check the control around that table", async () => {
    const doc = await from(
      control(1, `<w:alias w:val="Schedule"/><w:lock w:val="sdtContentLocked"/>`, table(`<w:tr>${cell(run("a"))}${cell(run("b"))}</w:tr>`)) +
        `<w:p>${run("After")}</w:p>`,
    );
    const message = 'tbl0: the table is in a locked content control "Schedule"; its content cannot be edited';
    expect(apply(doc, { type: "add_table", afterId: "tbl0", rows: [["x", "y"]] }).error).toBe(message);
    expect(apply(doc, { type: "add_list", afterId: "tbl0", listType: "bullet", items: ["x"] }).error).toBe(message);
    // Rows of a table inside a content-locked control cannot change either
    const locked = 'tbl0: the table is in a locked content control "Schedule"; its content cannot be edited';
    expect(apply(doc, { type: "add_table_row", target: "tbl0", afterRow: 0, cells: ["x", "y"] }).error).toBe(locked);
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 0 }).error).toBe(locked);
    expect(doc.getTables()).toHaveLength(1);
    expect(doc.getTables()[0]!.rows).toHaveLength(1);
  });

  it("add_table after a table in an open rich-text control lands inside it", async () => {
    const doc = await from(control(1, "", table(`<w:tr>${cell(run("a"))}${cell(run("b"))}</w:tr>`)) + `<w:p>${run("After")}</w:p>`);
    expect(apply(doc, { type: "add_table", afterId: "tbl0", rows: [["x", "y"]] }).success).toBe(true);
    expect(doc.toString()).toMatch(/<\/w:tbl><w:tbl>.*<\/w:tbl><\/w:sdtContent>/);
  });

  it.each([
    ["a locked control around a cell", AGAINST, `<w:tr>${cell(run("k"))}${control(1, `<w:alias w:val="Amount"/><w:lock w:val="sdtLocked"/>`, cell(run("v")))}</w:tr>`],
    ["a locked control inside a cell", AGAINST, `<w:tr>${cell(run("k"))}${cell(control(1, `<w:alias w:val="Amount"/><w:lock w:val="sdtLocked"/>`, run("v")))}</w:tr>`],
    ["a content-locked control inside a cell", CONTENT, `<w:tr>${cell(run("k"))}${cell(control(1, `<w:alias w:val="Amount"/><w:lock w:val="contentLocked"/>`, run("v")))}</w:tr>`],
    ["a locked control around a nested control", AGAINST, `<w:tr>${cell(run("k"))}${cell(control(1, `<w:alias w:val="Amount"/><w:lock w:val="sdtLocked"/>`, control(2, "", run("v"))))}</w:tr>`],
    ["a locked control nested in an open one", AGAINST, `<w:tr>${cell(run("k"))}${cell(control(2, "", control(1, `<w:alias w:val="Amount"/><w:lock w:val="sdtLocked"/>`, run("v"))))}</w:tr>`],
    [
      "a locked control around a row of a nested table",
      AGAINST,
      `<w:tr>${cell(run("k"))}<w:tc>${table(control(1, `<w:alias w:val="Amount"/><w:lock w:val="sdtLocked"/>`, `<w:tr>${cell(run("n1"))}${cell(run("n2"))}</w:tr>`))}<w:p/></w:tc></w:tr>`,
    ],
    [
      "a locked control around a nested table",
      AGAINST,
      `<w:tr>${cell(run("k"))}<w:tc>${control(1, `<w:alias w:val="Amount"/><w:lock w:val="sdtLocked"/>`, table(`<w:tr>${cell(run("n1"))}${cell(run("n2"))}</w:tr>`))}<w:p/></w:tc></w:tr>`,
    ],
  ])("delete_table_row is refused for a row holding %s", async (_name, expected, lockedRow) => {
    const doc = await from(table(`<w:tr>${cell(run("h1"))}${cell(run("h2"))}</w:tr>${lockedRow}`));
    const before = doc.toString();
    for (const options of [undefined, TRACKED]) {
      const r = apply(doc, { type: "delete_table_row", target: "tbl0", row: 1 }, options);
      expect(r.error).toBe(expected);
    }
    expect(doc.toString()).toBe(before);
    // A row without the control is still deletable
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 0 }).success).toBe(true);
  });

  it.each([
    ["locked against deletion", `<w:lock w:val="sdtLocked"/>`, 'tbl0 row 1: the row holds a content control "Item" that is locked against deletion', undefined],
    [
      "content-locked",
      `<w:lock w:val="contentLocked"/>`,
      'tbl0 row 1: the row is in a locked content control "Item"; its content cannot be edited',
      'tbl0 row 1: the row is in a locked content control "Item"; its content cannot be edited',
    ],
    [
      "data-bound",
      `<w:dataBinding w:xpath="/a" w:storeItemID="{0}"/>`,
      'tbl0 row 1: the row is in a content control "Item" bound to a custom XML part; Word overwrites its content from that part when the file is opened, so an edit here would be lost',
      'tbl0 row 1: the row is in a content control "Item" bound to a custom XML part; Word overwrites its content from that part when the file is opened, so an edit here would be lost',
    ],
  ])("a control wrapped around a row is kept on the row and checked: %s", async (_name, pr, deleteError, insertError) => {
    const doc = await from(
      table(
        `<w:tr>${cell(run("h1"))}${cell(run("h2"))}</w:tr>` +
          control(1, `<w:alias w:val="Item"/>${pr}`, `<w:tr>${cell(run("a"))}${cell(run("b"))}</w:tr>`),
      ),
    );
    const t = doc.findTableById("tbl0")!;
    expect(t.rows.map((r) => r.control?.alias)).toEqual([undefined, "Item"]);
    expect(t.rows[1]!.cells.map((c) => c.control)).toEqual([undefined, undefined]);

    const before = doc.toString();
    for (const options of [undefined, TRACKED]) {
      expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 1 }, options).error).toBe(deleteError);
    }
    const added = apply(doc, { type: "add_table_row", target: "tbl0", afterRow: 1, cells: ["x", "y"] });
    if (insertError) {
      expect(added.error).toBe(insertError);
      expect(doc.toString()).toBe(before);
    } else {
      // Locked against deletion only: its content may grow, and the new row lands inside it
      expect(added.success).toBe(true);
      expect(doc.toString()).toMatch(/<w:sdtContent><w:tr>.*<\/w:tr><w:tr>.*x.*<\/w:tr><\/w:sdtContent>/);
    }
    // The row outside the control is unaffected
    expect(apply(doc, { type: "add_table_row", target: "tbl0", afterRow: 0, cells: ["p", "q"] }).success).toBe(true);
  });

  it("add_table_row copies no prompt, whichever paragraph or cell of the row holds it", async () => {
    const prompt = (text: string): string => `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${text}</w:t></w:r>`;
    const shown = `<w:showingPlcHdr/>`;
    const doc = await from(
      table(
        `<w:tr>` +
          // a prompt in the second paragraph of a cell, next to a filled control
          `<w:tc><w:p>${run("Label")}</w:p><w:p>${run("see ")}${control(1, shown, prompt("Insert note"))}${control(2, "", run("kept"))}</w:p></w:tc>` +
          // a control around a cell whose prompt spans two paragraphs
          control(3, shown, `<w:tc><w:p>${prompt("Enter")}</w:p><w:p>${prompt("details")}</w:p></w:tc>`) +
          `</w:tr>`,
      ),
    );
    expect(apply(doc, { type: "add_table_row", target: "tbl0", afterRow: 0, cells: ["A", "B"] }).success).toBe(true);
    const rows = doc.findTableById("tbl0")!.rows;
    expect(rows[1]!.cells.map((c) => c.paragraphs.map((p) => p.text))).toEqual([
      ["A", "see kept"],
      ["B", ""],
    ]);
    // The source row is untouched, and no control is repeated
    expect(rows[0]!.cells.map((c) => c.paragraphs.map((p) => p.text))).toEqual([
      ["Label", "see Insert notekept"],
      ["Enter", "details"],
    ]);
    const xml = doc.toString();
    expect(xml.match(/<w:sdt>/g)).toHaveLength(3);
    expect(xml.match(/PlaceholderText/g)).toHaveLength(3);
  });

  it("rows of a table inside a data-bound control cannot change", async () => {
    const doc = await from(
      control(1, `<w:alias w:val="Feed"/><w:dataBinding w:xpath="/a" w:storeItemID="{0}"/>`, table(`<w:tr>${cell(run("a"))}${cell(run("b"))}</w:tr>`)),
    );
    const message =
      'tbl0: the table is in a content control "Feed" bound to a custom XML part; Word overwrites its content from that part when the file is opened, so an edit here would be lost';
    expect(apply(doc, { type: "add_table_row", target: "tbl0", afterRow: 0, cells: ["x", "y"] }).error).toBe(message);
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 0 }).error).toBe(message);
  });

  it("an empty control flagged as showing its prompt prints without a separator space", async () => {
    const doc = await from(
      `<w:p>${run("Name: ")}<w:sdt><w:sdtPr><w:id w:val="1"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent/></w:sdt></w:p>` +
        `<w:p>${run("When: ")}<w:sdt><w:sdtPr><w:id w:val="2"/><w:showingPlcHdr/><w:date/></w:sdtPr><w:sdtContent/></w:sdt></w:p>`,
    );
    const md = doc.project("markdown");
    expect(md).toContain("[p0] Name: {field:}");
    expect(md).toContain("[p1] When: {field date:}");
  });
});

describe("insert_text at the edge of a control", () => {
  it.each([[undefined], [TRACKED]])("before a word that follows an empty control stays outside it and leaves the prompt (%#)", async (options) => {
    const doc = await load();
    const before = doc.toString();
    // p0: "Title: " [prompt] " ref " [ACME-42]; " ref " starts where the prompt ends
    expect(apply(doc, { type: "insert_text", target: "p0", anchor: " ref", position: "before", text: " (draft)" }, options).success).toBe(true);
    const xml = doc.toString();
    expect(sdt(xml, "101")).toBe(sdt(before, "101"));
    const p0 = doc.findParagraphById("p0")!;
    expect(p0.text).toBe("Title: Insert Title for Contract (draft) ref ACME-42");
    expect(p0.controls![0]!.control.placeholder).toBe(true);
    expect(p0.controls![0]).toMatchObject({ startOffset: 7, endOffset: 32 });
  });

  it("before a word that follows a filled control stays outside it", async () => {
    const doc = await load();
    const before = doc.toString();
    // p1: "Contact: " [Jane Roe] " " [Dept [Legal]]
    expect(apply(doc, { type: "insert_text", target: "p1", anchor: " Dept", position: "before", text: "," }).success).toBe(true);
    expect(sdt(doc.toString(), "103")).toBe(sdt(before, "103"));
    expect(doc.findParagraphById("p1")!.text).toBe("Contact: Jane Roe, Dept Legal");
  });

  it("before the first word of a filled control goes inside it; after its last word too", async () => {
    const doc = await load();
    expect(apply(doc, { type: "insert_text", target: "p1", anchor: "Jane", position: "before", text: "Dr " }).success).toBe(true);
    expect(apply(doc, { type: "insert_text", target: "p1", anchor: "Roe", position: "after", text: " QC" }).success).toBe(true);
    expect(sdtContent(doc.toString(), "103")).toContain("Dr Jane Roe QC");
    expect(doc.findParagraphById("p1")!.text).toBe("Contact: Dr Jane Roe QC Dept Legal");
  });

  it("after a word that precedes a control stays outside it", async () => {
    const doc = await load();
    const before = doc.toString();
    expect(apply(doc, { type: "insert_text", target: "p1", anchor: "Contact: ", position: "after", text: "Ms " }).success).toBe(true);
    expect(sdt(doc.toString(), "103")).toBe(sdt(before, "103"));
    expect(doc.findParagraphById("p1")!.text).toBe("Contact: Ms Jane Roe Dept Legal");
  });
});

describe("controls left empty by rejecting a tracked fill", () => {
  const fillThenReject = async (target: string, find: string, replace: string): Promise<string> => {
    const doc = await load();
    expect(apply(doc, { type: "replace_text", target, find, replace }, TRACKED).success).toBe(true);
    doc.rejectAllRevisions();
    return doc.project("markdown");
  };

  it("a typed control shows as an empty field, not as a field with an empty value", async () => {
    expect(await fillThenReject("tbl0.r0.c1.p0", "Select date", "1 May 2026")).toContain("[tbl0.r0.c1.p0] {field date:} |");
    expect(await fillThenReject("tbl0.r1.c1.p0", "Choose an item.", "Ongoing supply")).toContain(
      '[tbl0.r1.c1.p0] {field dropdown "One-off purchase" OR "Ongoing supply":} |',
    );
    expect(await fillThenReject("p9", "Pick or type", "Teal")).toContain('[p9] Colour: {field combo "Red" OR "Blue":}');
  });

  it("a block-level control with nothing in it still shows as a field", async () => {
    expect(await fillThenReject("p7", "Click here to enter text.", "Background.")).toContain("[p7] {field:}");
  });

  it("an empty paragraph in a block control that holds text elsewhere is just an empty paragraph", async () => {
    const doc = await load();
    expect(apply(doc, { type: "delete_text", target: "p3", find: "Deliverables follow." }).success).toBe(true);
    const md = doc.project("markdown");
    expect(md).toContain("[p2] Scope of work.\n[p3] \n[p4]");
  });
});

describe("every patch type respects a closed control", () => {
  const run = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;
  const control = (id: number, pr: string, content: string): string =>
    `<w:sdt><w:sdtPr>${pr}<w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const LOCK = `<w:alias w:val="Terms"/><w:lock w:val="contentLocked"/>`;
  const BOUND = `<w:alias w:val="Terms"/><w:dataBinding w:xpath="/a" w:storeItemID="{0}"/>`;
  const numPr = `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>`;
  // p0 and p1 (a list item) sit in the control; p2 is outside it
  const body = (pr: string): string =>
    control(1, pr, `<w:p>${run("First clause")}</w:p><w:p>${numPr}${run("Listed clause")}</w:p>`) + `<w:p>${run("Outside")}</w:p>`;

  const patches: Patch[] = [
    { type: "replace_text", target: "p0", find: "First", replace: "Second" },
    { type: "insert_text", target: "p0", anchor: "First", position: "after", text: " and only" },
    { type: "delete_text", target: "p0", find: "First " },
    { type: "apply_formatting", target: "p0", formatting: { bold: true } },
    { type: "apply_formatting", target: "p0", find: "First", formatting: { bold: true } },
    { type: "insert_paragraph", target: "p0", position: "after", text: "New" },
    { type: "delete_paragraph", target: "p0" },
    { type: "add_table", afterId: "p0", rows: [["a"]] },
    { type: "add_list", afterId: "p0", listType: "bullet", items: ["a"] },
    { type: "convert_to_list", targets: ["p2", "p0"], listType: "bullet" },
    { type: "add_list_item", target: "p1", position: "after", text: "Another" },
    { type: "remove_list_item", target: "p1" },
    { type: "remove_list", targets: ["p1"] },
    { type: "set_list_properties", target: "p1", level: 1 },
  ];

  it.each([
    ["content-locked", LOCK, 'is in a locked content control "Terms"; its content cannot be edited'],
    ["data-bound", BOUND, 'is in a content control "Terms" bound to a custom XML part'],
  ])("%s: each of them is refused and the document is unchanged", async (_name, pr, reason) => {
    for (const options of [undefined, TRACKED]) {
      for (const patch of patches) {
        const doc = await Docxtor.load(await createBodyDocx(body(pr)));
        const before = doc.toString();
        const r = apply(doc, patch, options);
        expect(r.success, patch.type).toBe(false);
        expect(r.error, patch.type).toContain(reason);
        expect(doc.toString(), patch.type).toBe(before);
      }
    }
  });

  it("the same patches apply when the control is open", async () => {
    for (const patch of patches) {
      const doc = await Docxtor.load(await createBodyDocx(body(`<w:alias w:val="Terms"/>`)));
      expect(apply(doc, patch).error, patch.type).toBeUndefined();
    }
  });

  it("a comment may still be put on text in a locked control: it changes no content", async () => {
    const doc = await Docxtor.load(await createBodyDocx(body(LOCK)));
    expect(apply(doc, { type: "add_comment", target: "p0", find: "First", text: "Why?" }).success).toBe(true);
    expect(doc.findParagraphById("p0")!.text).toBe("First clause");
  });

  it("delete_paragraph is refused for a content-locked inline control, so a tracked delete cannot mark its text deleted", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p>${run("Rate: ")}${control(1, `<w:alias w:val="Rate"/><w:lock w:val="contentLocked"/>`, run("5%"))}</w:p><w:p>${run("Next")}</w:p>`),
    );
    for (const options of [undefined, TRACKED]) {
      expect(apply(doc, { type: "delete_paragraph", target: "p0" }, options).error).toBe(
        'p0: the paragraph holds a locked content control "Rate"; it cannot be deleted',
      );
    }
  });
});

describe("prompts and inline controls in structural edits", () => {
  const run = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;
  const prompt = (text: string): string => `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${text}</w:t></w:r>`;
  const control = (id: number, pr: string, content: string): string =>
    `<w:sdt><w:sdtPr>${pr}<w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

  it("insert_paragraph does not give the new text the grey style of a prompt that ends the target", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Name: </w:t></w:r>${control(1, "<w:showingPlcHdr/>", prompt("Insert name"))}</w:p>`),
    );
    expect(apply(doc, { type: "insert_paragraph", target: "p0", position: "after", text: "Address: 1 High St" }).success).toBe(true);
    const xml = doc.toString();
    expect(xml.match(/PlaceholderText/g)).toHaveLength(1);
    // It takes the formatting of the last typed run instead
    expect(xml).toContain("<w:r><w:rPr><w:b/></w:rPr><w:t>Address: 1 High St</w:t></w:r>");
  });

  it("a tracked remove_list_item marks text in an inline control deleted too", async () => {
    const numPr = `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>`;
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p>${numPr}${run("Supplier: ")}${control(1, "", run("Acme Ltd"))}</w:p><w:p>${numPr}${run("Second")}</w:p>`),
    );
    expect(apply(doc, { type: "remove_list_item", target: "p0" }, TRACKED).success).toBe(true);
    expect(doc.findParagraphById("p0")!.text).toBe("");
    expect(doc.toString()).toMatch(/<w:sdtContent><w:del [^>]*><w:r><w:delText>Acme Ltd<\/w:delText><\/w:r><\/w:del><\/w:sdtContent>/);
    doc.rejectAllRevisions();
    expect(doc.findParagraphById("p0")!.text).toBe("Supplier: Acme Ltd");
  });

  it("a tracked remove_list_item reaches a hyperlink and a control inside it", async () => {
    const numPr = `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>`;
    const doc = await Docxtor.load(
      await createBodyDocx(
        `<w:p>${numPr}${run("See ")}<w:hyperlink w:anchor="x">${run("clause ")}${control(1, "", run("12"))}</w:hyperlink>${run(".")}</w:p><w:p>${numPr}${run("Second")}</w:p>`,
      ),
    );
    expect(doc.findParagraphById("p0")!.text).toBe("See clause 12.");
    expect(apply(doc, { type: "remove_list_item", target: "p0" }, TRACKED).success).toBe(true);
    expect(doc.findParagraphById("p0")!.text).toBe("");
    expect(doc.toString()).toContain('<w:id w:val="1"/>');
    doc.rejectAllRevisions();
    expect(doc.findParagraphById("p0")!.text).toBe("See clause 12.");
  });

  it("runs of a paragraph wrapped by a block control carry no inline control", async () => {
    const doc = await load();
    // p7 is the prompt paragraph of an empty block-level control; p2 sits in a filled one
    for (const id of ["p7", "p2", "tbl0.r0.c1.p0"]) {
      const para = doc.findParagraphById(id)!;
      expect(para.control, id).toBeDefined();
      expect(para.controls, id).toBeUndefined();
      expect(para.runs.every((r) => r.control === undefined), id).toBe(true);
    }
  });

  it("a page break inside a block control carries the control", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(control(1, `<w:alias w:val="Part"/>`, `<w:p>${run("Before")}</w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p>`) + `<w:p>${run("After")}</w:p>`),
    );
    const pb = doc.getBody().find((e) => e.kind === "page_break")!;
    expect(pb.control?.alias).toBe("Part");
    expect(pb.control).toBe(doc.findParagraphById("p0")!.control);
  });
});

describe("comments, validation options and emptied row controls", () => {
  const run = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;
  const control = (id: number, pr: string, content: string): string =>
    `<w:sdt><w:sdtPr>${pr}<w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const cell = (inner: string): string => `<w:tc><w:p>${inner}</w:p></w:tc>`;

  it.each([[undefined], [TRACKED]])("filling a commented prompt keeps the comment's reference (%#)", async (options) => {
    const doc = await load();
    expect(apply(doc, { type: "add_comment", target: "p0", find: "Insert Title for Contract", text: "Confirm the title" }).success).toBe(true);
    expect(apply(doc, { type: "replace_text", target: "p0", find: "Insert Title", replace: "Supply of Widgets" }, options).success).toBe(true);
    const content = sdtContent(doc.toString(), "101");
    expect(content).toMatch(/^<w:sdtContent><w:commentRangeStart w:id="0"\/>(<w:ins [^>]*>)?<w:r>.*Supply of Widgets<\/w:t><\/w:r>(<\/w:ins>)?<w:commentRangeEnd w:id="0"\/><w:r>.*<w:commentReference w:id="0"\/><\/w:r><\/w:sdtContent>$/);
  });

  it("ticking a commented checkbox keeps the comment's reference", async () => {
    const doc = await load();
    expect(apply(doc, { type: "add_comment", target: "p4", find: "☐", text: "Tick this" }).success).toBe(true);
    expect(apply(doc, { type: "set_checkbox", target: "p4", checked: true }).success).toBe(true);
    const content = sdtContent(doc.toString(), "110");
    expect(content).toContain("☒");
    expect(content).toMatch(/<w:commentRangeEnd w:id="0"\/><w:r>.*<w:commentReference w:id="0"\/><\/w:r><\/w:sdtContent>$/);
  });

  it("validatePatches matches text the way the patches will be applied", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(`<w:p>${run("Code: ")}${control(1, `<w:alias w:val="Code"/><w:lock w:val="contentLocked"/>`, run("**A**"))}</w:p>`),
    );
    const patch: Patch = { type: "replace_text", target: "p0", find: "**A**", replace: "B" };
    // With markers the find reads as "A" in bold, which is in the locked control as well
    expect(doc.validatePatches([patch])[0]!.valid).toBe(false);
    expect(doc.validatePatches([patch], { inlineMarkers: false })[0]!.error).toContain('locked content control "Code"');
    const outside: Patch = { type: "replace_text", target: "p0", find: "Code", replace: "**Ref**" };
    expect(doc.validatePatches([outside], { inlineMarkers: false })).toEqual([{ valid: true }]);
  });

  it.each([["untracked"], ["tracked, then accepted"]])("a control around a deleted row goes with it, leaving no empty control in the table (%s)", async (mode) => {
    const doc = await Docxtor.load(
      await createBodyDocx(
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>${cell(run("keep"))}</w:tr>` +
          control(1, "", `<w:tr>${cell(run("go"))}</w:tr>`) +
          `</w:tbl><w:p/>`,
      ),
    );
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 1 }, mode === "untracked" ? undefined : TRACKED).success).toBe(true);
    if (mode !== "untracked") doc.acceptAllRevisions();
    const xml = doc.toString();
    expect(xml).not.toContain("<w:sdt>");
    expect(doc.findTableById("tbl0")!.rows.map((r) => r.cells[0]!.paragraphs[0]!.text)).toEqual(["keep"]);
  });

  it.each([["untracked"], ["tracked, then accepted"]])("nested controls around a deleted row all go with it (%s)", async (mode) => {
    const doc = await Docxtor.load(
      await createBodyDocx(
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>${cell(run("keep"))}</w:tr>` +
          control(1, "", control(2, "", `<w:tr>${cell(run("go"))}</w:tr>`)) +
          `</w:tbl><w:p/>`,
      ),
    );
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 1 }, mode === "untracked" ? undefined : TRACKED).success).toBe(true);
    if (mode !== "untracked") doc.acceptAllRevisions();
    expect(doc.toString()).not.toContain("<w:sdt>");
    expect(doc.findTableById("tbl0")!.rows).toHaveLength(1);
  });

  it("an outer control that still holds a row stays when a nested one is emptied", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>` +
          control(1, "", `<w:tr>${cell(run("stay"))}</w:tr>` + control(2, "", `<w:tr>${cell(run("go"))}</w:tr>`)) +
          `</w:tbl><w:p/>`,
      ),
    );
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 1 }, TRACKED).success).toBe(true);
    doc.acceptAllRevisions();
    const xml = doc.toString();
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
    expect(xml).toContain('<w:id w:val="1"/>');
    expect(xml).toMatch(/<w:sdtContent><w:tr>.*stay.*<\/w:tr><\/w:sdtContent>/);
  });

  it("rejecting an inserted row inside a nested control leaves no empty control either", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>${cell(run("keep"))}</w:tr>` +
          control(1, "", control(2, "", `<w:tr><w:trPr><w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:trPr>${cell(run("new"))}</w:tr>`)) +
          `</w:tbl><w:p/>`,
      ),
    );
    doc.rejectAllRevisions();
    expect(doc.toString()).not.toContain("<w:sdt>");
  });

  it("a control around two rows stays when one of them is deleted", async () => {
    const doc = await Docxtor.load(
      await createBodyDocx(
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>` +
          control(1, "", `<w:tr>${cell(run("a"))}</w:tr><w:tr>${cell(run("b"))}</w:tr>`) +
          `</w:tbl><w:p/>`,
      ),
    );
    expect(apply(doc, { type: "delete_table_row", target: "tbl0", row: 0 }).success).toBe(true);
    expect(doc.toString()).toMatch(/<w:sdtContent><w:tr>.*b.*<\/w:tr><\/w:sdtContent>/);
  });
});
