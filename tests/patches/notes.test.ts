/**
 * Footnotes and endnotes: the mark in the text, the notes section of the
 * view, and edits to a note's own paragraphs.
 */

import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { Docxtor } from "../../src/index.js";
import { createNotesDocx, createLeaseDocx } from "../fixtures/create-docx.js";

const fnRef = (id: number, attrs = ""): string =>
  `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"${attrs}/></w:r>`;
const enRef = (id: number): string =>
  `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
const t = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const note = (tag: "footnote" | "endnote", id: number, text: string): string =>
  `<w:${tag} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
  `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${tag}Ref/></w:r>${t(text)}</w:p></w:${tag}>`;

/**
 * Two footnotes referenced out of id order (the first reference points at
 * w:id 5, the second at w:id 3), one endnote, and a reference in a table cell.
 */
function twoNotes(): Promise<Buffer> {
  return createNotesDocx({
    body:
      `<w:p>${t("The Bank")}${fnRef(5)}${t(" lends to the Borrower")}${enRef(2)}${t(".")}</w:p>` +
      `<w:tbl><w:tr><w:tc><w:p>${t("Fees")}${fnRef(3)}</w:p></w:tc></w:tr></w:tbl>` +
      `<w:p>${t("No note here.")}</w:p>`,
    footnotes: note("footnote", 3, " Fees are set out in Schedule 2.") + note("footnote", 5, " IBRD and IDA are the World Bank."),
    endnotes: note("endnote", 2, " As defined in the Agreement."),
  });
}

async function entry(buffer: Buffer, path: string): Promise<string> {
  return (await JSZip.loadAsync(buffer)).file(path)!.async("string");
}

async function changedParts(before: Buffer, after: Buffer): Promise<string[]> {
  const [a, b] = [await JSZip.loadAsync(before), await JSZip.loadAsync(after)];
  const changed: string[] = [];
  for (const path of Object.keys(a.files)) {
    if (a.files[path]!.dir) continue;
    if ((await a.file(path)!.async("string")) !== (await b.file(path)?.async("string"))) changed.push(path);
  }
  for (const path of Object.keys(b.files)) if (!(path in a.files)) changed.push(path);
  return changed.sort();
}

const TRACK = { trackRevisions: true, revisionAuthor: "Reviewer", revisionDate: "2026-11-11T00:00:00Z" } as const;

describe("note marks in the view", () => {
  it("numbers the marks by position in the text, not by w:id", async () => {
    const doc = await Docxtor.load(await twoNotes());
    const md = doc.project("markdown");
    expect(md).toContain("[p0] The Bank[^1] lends to the Borrower[^e1].");
    expect(md).toContain("| [tbl0.r0.c0.p0] Fees[^2] |");
    expect(md).toContain("[p1] No note here.");
  });

  it("fn1 is the note the first reference points at", async () => {
    const doc = await Docxtor.load(await twoNotes());
    expect(doc.findParagraphById("fn1.p0")!.text).toBe(" IBRD and IDA are the World Bank.");
    expect(doc.findParagraphById("fn2.p0")!.text).toBe(" Fees are set out in Schedule 2.");
    expect([...doc.getFootnotes().values()].map((n) => [n.noteId, n.number, n.idPrefix, n.referenced])).toEqual([
      ["5", 1, "fn1", true],
      ["3", 2, "fn2", true],
    ]);
    expect([...doc.getEndnotes().values()].map((n) => [n.noteId, n.idPrefix])).toEqual([["2", "en1"]]);
  });

  it("the mark is not part of the paragraph's text", async () => {
    const doc = await Docxtor.load(await twoNotes());
    const para = doc.getParagraphs()[0]!;
    expect(para.text).toBe("The Bank lends to the Borrower.");
    expect(para.noteRefs).toEqual([
      { kind: "footnote", noteId: "5", offset: 8 },
      { kind: "endnote", noteId: "2", offset: 30 },
    ]);
    const json = doc.project("json");
    expect(json.body[0]).toMatchObject({
      text: "The Bank lends to the Borrower.",
      noteRefs: [{ note: "fn1", offset: 8 }, { note: "en1", offset: 30 }],
    });
  });

  it("places the mark after formatting markers and inside revision marks", async () => {
    const doc = await Docxtor.load(await createNotesDocx({
      body:
        `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bank</w:t></w:r>${fnRef(1)}${t(" lends ")}` +
        `<w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">${t("money")}${fnRef(2)}</w:ins>${t(".")}</w:p>`,
      footnotes: note("footnote", 1, " One.") + note("footnote", 2, " Two."),
    }));
    expect(doc.project("markdown", { includeFormatting: true, includeRevisions: true, includeNotes: false }))
      .toBe("[p0] **Bank**[^1] lends {+money+}[^2].");
  });

  it("a reference with its own mark prints the mark text and no number, and is left out of the count", async () => {
    const doc = await Docxtor.load(await createNotesDocx({
      body:
        `<w:p>${t("Alpha")}<w:r><w:footnoteReference w:id="1" w:customMarkFollows="1"/><w:t>*</w:t></w:r>` +
        `${t(" beta")}${fnRef(2)}</w:p>`,
      footnotes: note("footnote", 1, " Starred.") + note("footnote", 2, " Numbered."),
    }));
    expect(doc.project("markdown", { includeNotes: false })).toBe("[p0] Alpha* beta[^1]");
    expect(doc.findParagraphById("fn1.p0")!.text).toBe(" Numbered.");
    expect(doc.findParagraphById("fn2.p0")!.text).toBe(" Starred.");
  });

  it("a replace_text beside the mark keeps the reference run where it is", async () => {
    for (const options of [undefined, TRACK]) {
      const doc = await Docxtor.load(await twoNotes());
      const r = doc.applyPatches([{ type: "replace_text", target: "p0", find: "Bank", replace: "Lender" }], options);
      expect(r[0]).toMatchObject({ success: true });
      if (options) doc.acceptAllRevisions();
      expect(doc.toString()).toContain(`<w:footnoteReference w:id="5"/>`);
      expect(doc.project("markdown", { includeNotes: false }).split("\n")[0]).toBe("[p0] The Lender[^1] lends to the Borrower[^e1].");
    }
  });

  it("a replacement that spans the mark keeps the reference, after the new text", async () => {
    for (const options of [undefined, TRACK]) {
      const doc = await Docxtor.load(await twoNotes());
      const r = doc.applyPatches([{ type: "replace_text", target: "p0", find: "Bank lends", replace: "Lender gives" }], options);
      expect(r[0]).toMatchObject({ success: true });
      if (options) doc.acceptAllRevisions();
      expect(doc.project("markdown", { includeNotes: false }).split("\n")[0]).toBe("[p0] The Lender gives[^1] to the Borrower[^e1].");
    }
  });
});

describe("the notes section", () => {
  it("lists each note once, without the separator entries or the note's own mark", async () => {
    const doc = await Docxtor.load(await twoNotes());
    const md = doc.project("markdown");
    expect(md.slice(md.indexOf("--- Footnotes ---"))).toBe(
      [
        "--- Footnotes ---",
        "[fn1.p0] (FootnoteText)  IBRD and IDA are the World Bank.",
        "[fn2.p0] (FootnoteText)  Fees are set out in Schedule 2.",
        "",
        "--- Endnotes ---",
        "[en1.p0] (FootnoteText)  As defined in the Agreement.",
      ].join("\n"),
    );
  });

  it("is in the JSON projection as footnotes and endnotes", async () => {
    const json = (await Docxtor.load(await twoNotes())).project("json");
    expect(json.footnotes).toMatchObject([
      { id: "fn1", number: 1, paragraphs: [{ id: "fn1.p0", text: " IBRD and IDA are the World Bank." }] },
      { id: "fn2", number: 2, paragraphs: [{ id: "fn2.p0" }] },
    ]);
    expect(json.endnotes).toMatchObject([{ id: "en1", number: 1 }]);
  });

  it("includeNotes: false and includeHeadersFooters: false each drop their own section only", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    const all = doc.project("markdown");
    expect(all).toContain("--- Header (default) ---");
    expect(all).toContain("--- Footnotes ---");

    const noNotes = doc.project("markdown", { includeNotes: false });
    expect(noNotes).toContain("--- Header (default) ---");
    expect(noNotes).not.toContain("--- Footnotes ---");
    expect(doc.project("json", { includeNotes: false }).footnotes).toBeUndefined();

    const noHeaders = doc.project("markdown", { includeHeadersFooters: false });
    expect(noHeaders).not.toContain("--- Header (default) ---");
    expect(noHeaders).toContain("--- Footnotes ---");
    expect(doc.project("json", { includeHeadersFooters: false }).footnotes).toHaveLength(1);
  });

  it("shows a tracked insertion inside a footnote as {+…+}", async () => {
    const doc = await Docxtor.load(await createLeaseDocx());
    expect(doc.project("markdown", { includeRevisions: true })).toContain("[fn1.p0] See {+revised +}Exhibit B.");
  });

  it("a note nothing refers to is numbered after the ones that are referenced", async () => {
    const doc = await Docxtor.load(await createNotesDocx({
      body: `<w:p>${t("Text")}${fnRef(7)}</w:p>`,
      footnotes: note("footnote", 4, " Orphan.") + note("footnote", 7, " Referenced."),
    }));
    expect([...doc.getFootnotes().values()].map((n) => [n.idPrefix, n.noteId, n.referenced])).toEqual([
      ["fn1", "7", true],
      ["fn2", "4", false],
    ]);
  });

  it("a document without note parts has no notes and no section", async () => {
    const doc = await Docxtor.load(await createNotesDocx({ body: `<w:p>${t("Plain")}</w:p>` }));
    expect(doc.getFootnotes().size).toBe(0);
    expect(doc.project("markdown")).toBe("[p0] Plain");
    expect(doc.project("json").footnotes).toBeUndefined();
  });
});

describe("editing a note", () => {
  it.each([
    ["untracked", undefined],
    ["tracked", TRACK],
  ])("replace_text on fn1.p0 changes word/footnotes.xml only (%s)", async (_name, options) => {
    const input = await twoNotes();
    const doc = await Docxtor.load(input);
    const r = doc.applyPatches([{ type: "replace_text", target: "fn1.p0", find: "World Bank", replace: "Bank Group" }], options);
    expect(r[0]).toMatchObject({ success: true });
    const output = await doc.toBuffer();
    expect(await changedParts(input, output)).toEqual(["word/footnotes.xml"]);

    const xml = await entry(output, "word/footnotes.xml");
    // The note's own mark and the other note are untouched
    expect(xml).toContain(`<w:footnote w:id="5"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>`);
    expect(xml).toContain(note("footnote", 3, " Fees are set out in Schedule 2."));
    if (options) {
      expect(xml).toContain(`<w:del w:id="1" w:author="Reviewer" w:date="2026-11-11T00:00:00Z"><w:r><w:delText>World Bank</w:delText></w:r></w:del>`);
      expect(doc.project("markdown", { includeRevisions: true })).toContain("[fn1.p0] (FootnoteText)  IBRD and IDA are the {-World Bank-}{+Bank Group+}.");
    } else {
      expect(doc.findParagraphById("fn1.p0")!.text).toBe(" IBRD and IDA are the Bank Group.");
    }
  });

  it("accept and reject behave in a note as in the body", async () => {
    const input = await twoNotes();
    const patch = { type: "replace_text", target: "fn1.p0", find: "World Bank", replace: "Bank Group" } as const;

    const untracked = await Docxtor.load(input);
    untracked.applyPatches([patch]);

    const accepted = await Docxtor.load(input);
    accepted.applyPatches([patch], TRACK);
    expect(accepted.acceptAllRevisions()).toEqual({ accepted: 2 });
    expect(await entry(await accepted.toBuffer(), "word/footnotes.xml"))
      .toBe(await entry(await untracked.toBuffer(), "word/footnotes.xml"));

    const rejected = await Docxtor.load(input);
    rejected.applyPatches([patch], TRACK);
    expect(rejected.rejectAllRevisions()).toEqual({ rejected: 2 });
    expect(rejected.findParagraphById("fn1.p0")!.text).toBe(" IBRD and IDA are the World Bank.");
    // The ids are still the same notes after the document is re-normalized
    expect(rejected.findParagraphById("fn2.p0")!.text).toBe(" Fees are set out in Schedule 2.");
  });

  it("insert_text, delete_text and apply_formatting work on an endnote", async () => {
    const input = await twoNotes();
    const doc = await Docxtor.load(input);
    const results = doc.applyPatches([
      { type: "insert_text", target: "en1.p0", anchor: "Agreement", position: "after", text: " and its Schedules" },
      { type: "delete_text", target: "en1.p0", find: "As defined " },
      { type: "apply_formatting", target: "en1.p0", find: "Agreement", formatting: { italic: true } },
    ]);
    expect(results.map((r) => r.success)).toEqual([true, true, true]);
    expect(doc.findParagraphById("en1.p0")!.text).toBe(" in the Agreement and its Schedules.");
    expect(await changedParts(input, await doc.toBuffer())).toEqual(["word/endnotes.xml"]);
    expect(await entry(await doc.toBuffer(), "word/endnotes.xml")).toContain(`<w:rPr><w:i/></w:rPr><w:t>Agreement</w:t>`);
  });

  it("apply_style works on a note paragraph", async () => {
    const input = await createLeaseDocx();
    const doc = await Docxtor.load(input);
    const r = doc.applyPatches([{ type: "apply_style", target: "fn1.p0", style: "Heading2" }], TRACK);
    expect(r[0]).toMatchObject({ success: true });
    expect(doc.findParagraphById("fn1.p0")!.style).toBe("Heading2");
    expect(await changedParts(input, await doc.toBuffer())).toEqual(["word/footnotes.xml"]);
  });

  it("add_comment on a note anchors inside the note", async () => {
    const doc = await Docxtor.load(await twoNotes());
    const r = doc.applyPatches([{ type: "add_comment", target: "fn2.p0", find: "Schedule 2", text: "Check the schedule number.", author: "Reviewer" }]);
    expect(r[0]).toMatchObject({ success: true });
    const output = await doc.toBuffer();
    const notes = await entry(output, "word/footnotes.xml");
    expect(notes).toMatch(/<w:commentRangeStart w:id="0"\/><w:r><w:t>Schedule 2<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/>/);
    expect(await entry(output, "word/document.xml")).not.toContain("commentRange");
    expect(await entry(output, "word/comments.xml")).toContain("Check the schedule number.");
  });

  it("insert_paragraph adds a second paragraph to a note", async () => {
    const doc = await Docxtor.load(await twoNotes());
    const r = doc.applyPatches([{ type: "insert_paragraph", target: "fn1.p0", position: "after", text: "See also the Articles." }]);
    expect(r[0]).toMatchObject({ success: true });
    expect(doc.findParagraphById("fn1.p1")!.text).toBe("See also the Articles.");
    const xml = await entry(await doc.toBuffer(), "word/footnotes.xml");
    expect(xml.match(/<w:footnoteRef\/>/g)).toHaveLength(2);
    expect(xml).toContain(`<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:t>See also the Articles.</w:t></w:r></w:p></w:footnote>`);
  });

  it("refuses to delete the paragraph that holds the note's number", async () => {
    const doc = await Docxtor.load(await twoNotes());
    const r = doc.applyPatches([{ type: "delete_paragraph", target: "fn1.p0" }]);
    expect(r[0]!.success).toBe(false);
    expect(r[0]!.error).toContain("deleting a footnote is not supported");

    // A paragraph added to the note can be deleted again
    doc.applyPatches([{ type: "insert_paragraph", target: "fn1.p0", position: "after", text: "Extra." }]);
    expect(doc.applyPatches([{ type: "delete_paragraph", target: "fn1.p1" }])[0]).toMatchObject({ success: true });
    expect(doc.findParagraphById("fn1.p1")).toBeUndefined();
  });

  it("an atomic batch that fails leaves the note untouched, and one that succeeds keeps the note ids", async () => {
    const input = await twoNotes();
    const doc = await Docxtor.load(input);
    const failed = doc.applyPatches(
      [
        { type: "replace_text", target: "fn1.p0", find: "World Bank", replace: "Bank Group" },
        { type: "replace_text", target: "fn2.p0", find: "no such text", replace: "x" },
      ],
      { atomic: true },
    );
    expect(failed.map((r) => r.success)).toEqual([true, false]);
    expect(await changedParts(input, await doc.toBuffer())).toEqual([]);

    const ok = doc.applyPatches(
      [
        { type: "replace_text", target: "fn1.p0", find: "World Bank", replace: "Bank Group" },
        { type: "replace_text", target: "p0", find: "Borrower", replace: "Recipient" },
      ],
      { atomic: true },
    );
    expect(ok.every((r) => r.success)).toBe(true);
    expect(doc.findParagraphById("fn1.p0")!.text).toBe(" IBRD and IDA are the Bank Group.");
    expect(doc.project("markdown")).toContain("[fn1.p0] (FootnoteText)  IBRD and IDA are the Bank Group.");
  });

  it("a tab inside a note reads and writes as a tab", async () => {
    const doc = await Docxtor.load(await createNotesDocx({
      body: `<w:p>${t("Text")}${fnRef(1)}</w:p>`,
      footnotes: `<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:tab/></w:r>${t("Indented note.")}</w:p></w:footnote>`,
    }));
    expect(doc.findParagraphById("fn1.p0")!.text).toBe("\tIndented note.");
    doc.applyPatches([{ type: "replace_text", target: "fn1.p0", find: "\tIndented", replace: "\tTabbed" }]);
    expect(doc.findParagraphById("fn1.p0")!.text).toBe("\tTabbed note.");
  });
});
