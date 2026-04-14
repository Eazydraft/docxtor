import { describe, it, expect } from "vitest";
import { Docxtor } from "../../src/index.js";
import {
  createSimpleDocx,
  createFormattedDocx,
  createTableDocx,
  createHeaderFooterDocx,
  createSplitRunDocx,
  createPageBreakDocx,
  createInlinePageBreakDocx,
  createLastRenderedPageBreakDocx,
  createLastRenderedPageBreakOnlyDocx,
  createPageBreakBeforeDocx,
  createPageBreakBeforeDisabledDocx,
  createMidTextPageBreakDocx,
  createRevisionWithPageBreakDocx,
} from "../fixtures/create-docx.js";

// =============================================================================
// Creating
// =============================================================================

describe("Docxtor.create", () => {
  it("should create a valid empty DOCX when called with no arguments", async () => {
    const doc = await Docxtor.create();

    expect(doc).toBeInstanceOf(Docxtor);
    expect(doc.getParagraphs()).toHaveLength(1);
    expect(doc.getParagraphs()[0]!.text).toBe("");
  });

  it("should produce a DOCX buffer that Docxtor.load can read back", async () => {
    const doc = await Docxtor.create();
    const buffer = await doc.toBuffer();

    const reloaded = await Docxtor.load(buffer);
    expect(reloaded).toBeInstanceOf(Docxtor);
    expect(reloaded.getParagraphs()).toHaveLength(1);
    expect(reloaded.getParagraphs()[0]!.text).toBe("");
  });

  it("should delegate to load() when input is provided", async () => {
    const buffer = await createSimpleDocx("Hello from create");
    const doc = await Docxtor.create(buffer);

    expect(doc).toBeInstanceOf(Docxtor);
    expect(doc.getParagraphs()).toHaveLength(1);
    expect(doc.getParagraphs()[0]!.text).toBe("Hello from create");
  });

  it("should support patching an empty document with append_text", async () => {
    const doc = await Docxtor.create();

    const results = doc.applyPatches([
      { type: "append_text", text: "New content" },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.getParagraphs()).toHaveLength(2);
    expect(doc.getParagraphs()[1]!.text).toBe("New content");
  });

  it("should support the full create → patch → export → reload cycle", async () => {
    const doc = await Docxtor.create();

    doc.applyPatches([
      { type: "append_text", text: "Created from scratch" },
    ]);

    const buffer = await doc.toBuffer();
    const reloaded = await Docxtor.load(buffer);

    expect(reloaded.getParagraphs()[1]!.text).toBe("Created from scratch");
  });

  it("should produce valid projections for an empty document", async () => {
    const doc = await Docxtor.create();

    const json = doc.project("json");
    expect(json.body).toHaveLength(1);
    expect(json.body[0]).toMatchObject({ id: "p0", text: "" });

    const markdown = doc.project("markdown");
    expect(markdown).toContain("[p0]");
  });
});

// =============================================================================
// Loading
// =============================================================================

describe("Docxtor.load", () => {
  it("should load a DOCX from a buffer", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const doc = await Docxtor.load(buffer);

    expect(doc).toBeInstanceOf(Docxtor);
    expect(doc.getParagraphs()).toHaveLength(1);
    expect(doc.getParagraphs()[0]!.text).toBe("Hello World");
  });

  it("should expose body elements", async () => {
    const buffer = await createTableDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(3); // paragraph, table, paragraph

    expect(doc.getParagraphs()).toHaveLength(2);
    expect(doc.getTables()).toHaveLength(1);
  });
});

// =============================================================================
// Projection — JSON
// =============================================================================

describe("Docxtor.project — JSON", () => {
  it("should project a simple document to JSON", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    expect(projection.body).toHaveLength(1);
    expect(projection.body[0]).toEqual({
      id: "p0",
      text: "Hello World",
    });
  });

  it("should project a formatted document with styles", async () => {
    const buffer = await createFormattedDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    expect(projection.body).toHaveLength(3);
    expect(projection.body[0]).toMatchObject({
      id: "p0",
      text: "Document Title",
      style: "Heading1",
    });
  });

  it("should project a document with a table", async () => {
    const buffer = await createTableDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    // Find the table in the projection
    const table = projection.body.find((e: any) => e.id === "tbl0") as any;
    expect(table).toBeDefined();
    expect(table.rows).toHaveLength(2);
    expect(table.rows[0][0].text).toBe("Header A");
    expect(table.rows[1][1].text).toBe("Cell 2");
  });

  it("should include headers and footers", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    expect(projection.headers).toHaveLength(1);
    expect(projection.headers![0]!.paragraphs[0]!.text).toBe("Header text");
    expect(projection.footers).toHaveLength(1);
    expect(projection.footers![0]!.paragraphs[0]!.text).toBe("Footer text");
  });

  it("should include formatting when requested", async () => {
    const buffer = await createFormattedDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includeFormatting: true });

    const heading = projection.body[0] as any;
    expect(heading.formatting).toBeDefined();
    expect(heading.formatting[0]).toEqual({ bold: true });
  });

  it("should exclude headers/footers when requested", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includeHeadersFooters: false });

    expect(projection.headers).toBeUndefined();
    expect(projection.footers).toBeUndefined();
  });
});

// =============================================================================
// Projection — Markdown
// =============================================================================

describe("Docxtor.project — Markdown", () => {
  it("should project a simple document to markdown", async () => {
    const buffer = await createSimpleDocx("Hello World");
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toBe("[p0] Hello World");
  });

  it("should project styles as hints", async () => {
    const buffer = await createFormattedDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toContain("[p0] (Heading1) Document Title");
    expect(markdown).toContain("[p1] This is bold and italic text.");
  });

  it("should project tables as pipe-delimited rows", async () => {
    const buffer = await createTableDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toContain("| [tbl0.r0.c0.p0] Header A");
    expect(markdown).toContain("[tbl0.r0.c1.p0] Header B |");
    expect(markdown).toContain("[tbl0.r1.c0.p0] Cell 1");
  });

  it("should include header/footer sections", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toContain("--- Header (default) ---");
    expect(markdown).toContain("[hdr.default.p0] Header text");
    expect(markdown).toContain("--- Footer (default) ---");
    expect(markdown).toContain("[ftr.default.p0] Footer text");
  });
});

// =============================================================================
// Projection — Page Breaks (JSON)
// =============================================================================

describe("Docxtor.project — Page Breaks (JSON)", () => {
  it("should hide page breaks from JSON projection by default", async () => {
    const buffer = await createPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    // Only the two text paragraphs should appear — page break is hidden
    expect(projection.body).toHaveLength(2);
    expect(projection.body[0]).toMatchObject({ id: "p0", text: "Before page break" });
    expect(projection.body[1]).toMatchObject({ id: "p1", text: "After page break" });
  });

  it("should hide hasPageBreak from JSON projection by default", async () => {
    const buffer = await createInlinePageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");

    expect(projection.body).toHaveLength(2);
    expect((projection.body[0] as any).hasPageBreak).toBeUndefined();
  });

  it("should include page_break elements when includePageBreaks is true", async () => {
    const buffer = await createPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includePageBreaks: true });

    expect(projection.body).toHaveLength(3);
    expect(projection.body[0]).toMatchObject({ id: "p0", text: "Before page break" });
    expect(projection.body[1]).toEqual({ id: "pb0", type: "page_break" });
    expect(projection.body[2]).toMatchObject({ id: "p1", text: "After page break" });
  });

  it("should include hasPageBreak and pageBreakOffset on paragraphs when includePageBreaks is true", async () => {
    const buffer = await createInlinePageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includePageBreaks: true });

    expect(projection.body).toHaveLength(2);
    expect(projection.body[0]).toMatchObject({
      id: "p0",
      text: "Text with page break",
      hasPageBreak: true,
      pageBreakOffset: 20, // break is after "Text with page break"
    });
    expect(projection.body[1]).toMatchObject({
      id: "p1",
      text: "Next page content",
    });
    expect((projection.body[1] as any).hasPageBreak).toBeUndefined();
    expect((projection.body[1] as any).pageBreakOffset).toBeUndefined();
  });
});

// =============================================================================
// Projection — Page Breaks (Markdown)
// =============================================================================

describe("Docxtor.project — Page Breaks (Markdown)", () => {
  it("should hide page breaks from Markdown projection by default", async () => {
    const buffer = await createPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toContain("[p0] Before page break");
    expect(markdown).toContain("[p1] After page break");
    expect(markdown).not.toContain("PAGE BREAK");
  });

  it("should hide [PAGE BREAK] hint from Markdown projection by default", async () => {
    const buffer = await createInlinePageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown");

    expect(markdown).toContain("[p0] Text with page break");
    expect(markdown).not.toContain("[PAGE BREAK]");
  });

  it("should show --- PAGE BREAK --- when includePageBreaks is true", async () => {
    const buffer = await createPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", { includePageBreaks: true });

    expect(markdown).toContain("[p0] Before page break");
    expect(markdown).toContain("[pb0] --- PAGE BREAK ---");
    expect(markdown).toContain("[p1] After page break");
  });

  it("should show [PAGE BREAK] hint at end when break is after text", async () => {
    const buffer = await createInlinePageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", { includePageBreaks: true });

    // Break is after "Text with page break", so marker goes at end
    expect(markdown).toContain("[p0] Text with page break [PAGE BREAK]");
    expect(markdown).toContain("[p1] Next page content");
    expect(markdown).not.toContain("[p1] [PAGE BREAK]");
  });
});

// =============================================================================
// Page Break — AST Inspection
// =============================================================================

describe("Page Break AST", () => {
  it("should normalize page-break-only paragraphs as NormalizedPageBreak", async () => {
    const buffer = await createPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(3);
    expect(body[0]!.kind).toBe("paragraph");
    expect(body[1]!.kind).toBe("page_break");
    expect(body[1]!.id).toBe("pb0");
    expect(body[2]!.kind).toBe("paragraph");
  });

  it("should set hasPageBreak and pageBreakOffset on paragraphs with text and page break", async () => {
    const buffer = await createInlinePageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(2);
    expect(body[0]!.kind).toBe("paragraph");
    expect((body[0] as any).hasPageBreak).toBe(true);
    expect((body[0] as any).pageBreakOffset).toBe(20); // after "Text with page break"
    expect(body[1]!.kind).toBe("paragraph");
    expect((body[1] as any).hasPageBreak).toBeUndefined();
    expect((body[1] as any).pageBreakOffset).toBeUndefined();
  });

  it("page break paragraphs should not count toward paragraph IDs", async () => {
    const buffer = await createPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    // Paragraphs should be p0 and p1 (page break gets pb0, not p1)
    const paragraphs = doc.getParagraphs();
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]!.id).toBe("p0");
    expect(paragraphs[1]!.id).toBe("p1");
  });
});

// =============================================================================
// Page Break — lastRenderedPageBreak (real-world Word documents)
// =============================================================================

describe("Page Break — lastRenderedPageBreak", () => {
  it("should detect lastRenderedPageBreak as inline page break (hasPageBreak)", async () => {
    const buffer = await createLastRenderedPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(2);
    expect(body[0]!.kind).toBe("paragraph");
    expect((body[0] as any).hasPageBreak).toBeUndefined();
    expect(body[1]!.kind).toBe("paragraph");
    expect((body[1] as any).hasPageBreak).toBe(true);
    expect((body[1] as any).text).toBe("Second page content");
  });

  it("should show lastRenderedPageBreak in JSON with offset when includePageBreaks is true", async () => {
    const buffer = await createLastRenderedPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includePageBreaks: true });
    expect(projection.body).toHaveLength(2);
    expect(projection.body[1]).toMatchObject({
      id: "p1",
      text: "Second page content",
      hasPageBreak: true,
      pageBreakOffset: 0, // lastRenderedPageBreak appears before text in the run
    });
  });

  it("should hide lastRenderedPageBreak from JSON by default", async () => {
    const buffer = await createLastRenderedPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");
    expect(projection.body).toHaveLength(2);
    expect((projection.body[1] as any).hasPageBreak).toBeUndefined();
  });

  it("should show [PAGE BREAK] in markdown when includePageBreaks is true", async () => {
    const buffer = await createLastRenderedPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", { includePageBreaks: true });
    expect(markdown).toContain("[p1] [PAGE BREAK] Second page content");
  });

  it("should normalize standalone lastRenderedPageBreak as NormalizedPageBreak", async () => {
    const buffer = await createLastRenderedPageBreakOnlyDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(3);
    expect(body[0]!.kind).toBe("paragraph");
    expect(body[0]!.id).toBe("p0");
    expect(body[1]!.kind).toBe("page_break");
    expect(body[1]!.id).toBe("pb0");
    expect(body[2]!.kind).toBe("paragraph");
    expect(body[2]!.id).toBe("p1");
  });

  it("should show standalone lastRenderedPageBreak in markdown projection", async () => {
    const buffer = await createLastRenderedPageBreakOnlyDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", { includePageBreaks: true });
    expect(markdown).toContain("[pb0] --- PAGE BREAK ---");
    expect(markdown).toContain("[p0] Before break");
    expect(markdown).toContain("[p1] After break");
  });
});

// =============================================================================
// Page Break — pageBreakBefore (paragraph property)
// =============================================================================

describe("Page Break — pageBreakBefore", () => {
  it("should detect w:pageBreakBefore as hasPageBreak", async () => {
    const buffer = await createPageBreakBeforeDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(2);
    expect(body[0]!.kind).toBe("paragraph");
    expect((body[0] as any).hasPageBreak).toBeUndefined();
    expect(body[1]!.kind).toBe("paragraph");
    expect((body[1] as any).hasPageBreak).toBe(true);
    expect((body[1] as any).text).toBe("Starts on new page");
  });

  it("should show pageBreakBefore in JSON with offset 0 when includePageBreaks is true", async () => {
    const buffer = await createPageBreakBeforeDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includePageBreaks: true });
    expect(projection.body).toHaveLength(2);
    expect(projection.body[1]).toMatchObject({
      id: "p1",
      text: "Starts on new page",
      hasPageBreak: true,
      pageBreakOffset: 0, // pageBreakBefore is always offset 0
    });
  });

  it("should hide pageBreakBefore from JSON by default", async () => {
    const buffer = await createPageBreakBeforeDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");
    expect((projection.body[1] as any).hasPageBreak).toBeUndefined();
  });

  it("should show [PAGE BREAK] hint in markdown when includePageBreaks is true", async () => {
    const buffer = await createPageBreakBeforeDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", { includePageBreaks: true });
    expect(markdown).toContain("[p1] [PAGE BREAK] Starts on new page");
  });

  it("should NOT detect w:pageBreakBefore w:val=0 as a page break", async () => {
    const buffer = await createPageBreakBeforeDisabledDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(1);
    expect(body[0]!.kind).toBe("paragraph");
    expect((body[0] as any).hasPageBreak).toBeUndefined();
    expect((body[0] as any).text).toBe("No page break here");
  });
});

// =============================================================================
// Page Break — positional offset (mid-text break)
// =============================================================================

describe("Page Break — positional offset", () => {
  it("should place [PAGE BREAK] between text in markdown", async () => {
    const buffer = await createMidTextPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", { includePageBreaks: true });
    // "Before break" (12 chars) + [PAGE BREAK] + "After break"
    expect(markdown).toContain("[p0] Before break [PAGE BREAK] After break");
  });

  it("should set pageBreakOffset to 12 in AST for mid-text break", async () => {
    const buffer = await createMidTextPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const body = doc.getBody();
    expect(body).toHaveLength(1);
    expect(body[0]!.kind).toBe("paragraph");
    expect((body[0] as any).hasPageBreak).toBe(true);
    expect((body[0] as any).pageBreakOffset).toBe(12); // "Before break" = 12 chars
    expect((body[0] as any).text).toBe("Before breakAfter break");
  });

  it("should expose pageBreakOffset in JSON projection for mid-text break", async () => {
    const buffer = await createMidTextPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json", { includePageBreaks: true });
    expect(projection.body).toHaveLength(1);
    expect(projection.body[0]).toMatchObject({
      id: "p0",
      text: "Before breakAfter break",
      hasPageBreak: true,
      pageBreakOffset: 12,
    });
  });

  it("should hide pageBreakOffset from JSON when includePageBreaks is false", async () => {
    const buffer = await createMidTextPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const projection = doc.project("json");
    expect((projection.body[0] as any).hasPageBreak).toBeUndefined();
    expect((projection.body[0] as any).pageBreakOffset).toBeUndefined();
  });
});

// =============================================================================
// Page Break + Revisions interaction
// =============================================================================

describe("Page Break — revision annotation interaction", () => {
  it("should place [PAGE BREAK] marker at the correct position when revisions are also annotated", async () => {
    const buffer = await createRevisionWithPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", {
      includePageBreaks: true,
      includeRevisions: true,
    });

    // Original text: "Hello worldAfter break" (pageBreakOffset = 11)
    // With revisions: "Hello {+world+}After break"
    // With page break: "Hello {+world+} [PAGE BREAK] After break"
    expect(markdown).toContain("[p0] Hello {+world+} [PAGE BREAK] After break");
  });

  it("should place [PAGE BREAK] correctly without revisions even when revisions exist", async () => {
    const buffer = await createRevisionWithPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", {
      includePageBreaks: true,
      includeRevisions: false,
    });

    // Without revision annotation, text is "Hello worldAfter break"
    // pageBreakOffset = 11, so marker goes after "Hello world"
    expect(markdown).toContain("[p0] Hello world [PAGE BREAK] After break");
  });

  it("should not show [PAGE BREAK] when only revisions are enabled", async () => {
    const buffer = await createRevisionWithPageBreakDocx();
    const doc = await Docxtor.load(buffer);

    const markdown = doc.project("markdown", {
      includePageBreaks: false,
      includeRevisions: true,
    });

    expect(markdown).toContain("[p0] Hello {+world+}After break");
    expect(markdown).not.toContain("[PAGE BREAK]");
  });
});

// =============================================================================
// Patching through public API
// =============================================================================

describe("Docxtor.applyPatches", () => {
  it("should apply text replacement", async () => {
    const buffer = await createSimpleDocx("This agreement is valid");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "valid", replace: "enforceable" },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe("This agreement is enforceable");
  });

  it("should report failures without throwing", async () => {
    const buffer = await createSimpleDocx("Hello");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "nonexistent", replace: "x" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toBeDefined();
  });
});

// =============================================================================
// append_text patch
// =============================================================================

describe("append_text", () => {
  it("should append a new paragraph to the document body", async () => {
    const buffer = await createSimpleDocx("First paragraph");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "append_text", text: "Second paragraph" },
    ]);

    expect(results[0]!.success).toBe(true);
    expect(doc.getParagraphs()).toHaveLength(2);
    expect(doc.getParagraphs()[1]!.text).toBe("Second paragraph");
  });

  it("should assign the correct paragraph ID", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([{ type: "append_text", text: "Second" }]);

    expect(doc.getParagraphs()[1]!.id).toBe("p1");
  });

  it("should support multiple sequential appends", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "append_text", text: "Second" },
      { type: "append_text", text: "Third" },
    ]);

    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.getParagraphs()).toHaveLength(3);
    expect(doc.getParagraphs()[0]!.text).toBe("First");
    expect(doc.getParagraphs()[1]!.text).toBe("Second");
    expect(doc.getParagraphs()[2]!.text).toBe("Third");
    expect(doc.getParagraphs()[1]!.id).toBe("p1");
    expect(doc.getParagraphs()[2]!.id).toBe("p2");
  });

  it("should survive export and reload", async () => {
    const buffer = await createSimpleDocx("Original");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([{ type: "append_text", text: "Appended" }]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);

    expect(reloaded.getParagraphs()).toHaveLength(2);
    expect(reloaded.getParagraphs()[0]!.text).toBe("Original");
    expect(reloaded.getParagraphs()[1]!.text).toBe("Appended");
  });

  it("should appear in projections", async () => {
    const buffer = await createSimpleDocx("Existing");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([{ type: "append_text", text: "New" }]);

    const json = doc.project("json");
    expect(json.body).toHaveLength(2);
    expect(json.body[1]).toMatchObject({ id: "p1", text: "New" });

    const markdown = doc.project("markdown");
    expect(markdown).toContain("[p0] Existing");
    expect(markdown).toContain("[p1] New");
  });

  it("should reject empty text", async () => {
    const buffer = await createSimpleDocx("Hello");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "append_text", text: "" },
    ]);

    expect(results[0]!.success).toBe(false);
    expect(results[0]!.error).toContain("Empty");
  });

  it("should work on an empty document created with Docxtor.create()", async () => {
    const doc = await Docxtor.create();

    doc.applyPatches([{ type: "append_text", text: "Hello World" }]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);

    expect(reloaded.getParagraphs()).toHaveLength(2);
    expect(reloaded.getParagraphs()[1]!.text).toBe("Hello World");
  });

  it("should be targetable by subsequent patches", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "append_text", text: "Appended text here" },
      { type: "replace_text", target: "p1", find: "text here", replace: "content" },
    ]);

    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.getParagraphs()[1]!.text).toBe("Appended content");
  });

  it("should handle text with newlines by creating separate paragraphs", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "append_text", text: "Line one\nLine two\nLine three" },
    ]);

    expect(results[0]!.success).toBe(true);
    // Each line becomes its own paragraph
    expect(doc.getParagraphs()).toHaveLength(4); // "First" + 3 lines
    expect(doc.getParagraphs()[1]!.text).toBe("Line one");
    expect(doc.getParagraphs()[2]!.text).toBe("Line two");
    expect(doc.getParagraphs()[3]!.text).toBe("Line three");
  });

  it("should assign correct paragraph IDs for each line", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "append_text", text: "Alpha\nBravo\nCharlie" },
    ]);

    expect(doc.getParagraphs()[1]!.id).toBe("p1");
    expect(doc.getParagraphs()[2]!.id).toBe("p2");
    expect(doc.getParagraphs()[3]!.id).toBe("p3");
  });

  it("should make newline-split paragraphs targetable by subsequent patches", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "append_text", text: "Hello world\nGoodbye world" },
      { type: "replace_text", target: "p2", find: "Goodbye", replace: "Farewell" },
    ]);

    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.getParagraphs()[1]!.text).toBe("Hello world");
    expect(doc.getParagraphs()[2]!.text).toBe("Farewell world");
  });

  it("should show newline-split paragraphs in projections", async () => {
    const buffer = await createSimpleDocx("Existing");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "append_text", text: "Line A\nLine B" },
    ]);

    const json = doc.project("json");
    expect(json.body).toHaveLength(3);
    expect(json.body[1]).toMatchObject({ id: "p1", text: "Line A" });
    expect(json.body[2]).toMatchObject({ id: "p2", text: "Line B" });

    const markdown = doc.project("markdown");
    expect(markdown).toContain("[p0] Existing");
    expect(markdown).toContain("[p1] Line A");
    expect(markdown).toContain("[p2] Line B");
  });

  it("should handle trailing newline by creating an empty paragraph", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "append_text", text: "Hello\n" },
    ]);

    expect(doc.getParagraphs()).toHaveLength(3);
    expect(doc.getParagraphs()[1]!.text).toBe("Hello");
    expect(doc.getParagraphs()[2]!.text).toBe("");
  });

  it("should handle leading newline by creating an empty paragraph first", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "append_text", text: "\nHello" },
    ]);

    expect(doc.getParagraphs()).toHaveLength(3);
    expect(doc.getParagraphs()[1]!.text).toBe("");
    expect(doc.getParagraphs()[2]!.text).toBe("Hello");
  });

  it("should survive export/reload with newlines", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "append_text", text: "Hello\nWorld" },
    ]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);

    // Each line becomes its own paragraph
    expect(reloaded.getParagraphs()).toHaveLength(3); // "First" + "Hello" + "World"
    expect(reloaded.getParagraphs()[1]!.text).toBe("Hello");
    expect(reloaded.getParagraphs()[2]!.text).toBe("World");
  });

  it("should handle newlines with revision tracking", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches(
      [{ type: "append_text", text: "Line A\nLine B" }],
      { trackRevisions: true, revisionAuthor: "test" },
    );

    expect(results[0]!.success).toBe(true);
    // Each line becomes its own paragraph
    expect(doc.getParagraphs()).toHaveLength(3); // "First" + "Line A" + "Line B"
    expect(doc.getParagraphs()[1]!.text).toBe("Line A");
    expect(doc.getParagraphs()[2]!.text).toBe("Line B");

    // Verify it survives export/reload
    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);
    expect(reloaded.getParagraphs()).toHaveLength(3);
    expect(reloaded.getParagraphs()[1]!.text).toBe("Line A");
    expect(reloaded.getParagraphs()[2]!.text).toBe("Line B");
  });

  it("should wrap each paragraph run in w:ins when revision tracking", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches(
      [{ type: "append_text", text: "Alpha\nBravo" }],
      { trackRevisions: true, revisionAuthor: "tester" },
    );

    // Inspect the raw XML to confirm w:ins wrapping
    const xml = doc.toString();
    // Each appended paragraph should contain a w:ins element
    const insMatches = xml.match(/<w:ins /g);
    expect(insMatches).not.toBeNull();
    expect(insMatches!.length).toBeGreaterThanOrEqual(2);
  });

  it("should handle multiple appends each with newlines", async () => {
    const buffer = await createSimpleDocx("First");
    const doc = await Docxtor.load(buffer);

    const results = doc.applyPatches([
      { type: "append_text", text: "A\nB" },
      { type: "append_text", text: "C\nD" },
    ]);

    expect(results.every((r) => r.success)).toBe(true);
    expect(doc.getParagraphs()).toHaveLength(5); // "First" + A + B + C + D
    expect(doc.getParagraphs()[1]!.text).toBe("A");
    expect(doc.getParagraphs()[2]!.text).toBe("B");
    expect(doc.getParagraphs()[3]!.text).toBe("C");
    expect(doc.getParagraphs()[4]!.text).toBe("D");
    // IDs should be sequential
    expect(doc.getParagraphs()[1]!.id).toBe("p1");
    expect(doc.getParagraphs()[2]!.id).toBe("p2");
    expect(doc.getParagraphs()[3]!.id).toBe("p3");
    expect(doc.getParagraphs()[4]!.id).toBe("p4");
  });
});

// =============================================================================
// Full pipeline: load → project → patch → export → reload
// =============================================================================

describe("Full pipeline", () => {
  it("should survive a complete edit cycle", async () => {
    // 1. Create a DOCX
    const buffer = await createFormattedDocx();

    // 2. Load it
    const doc = await Docxtor.load(buffer);

    // 3. Get projection (what the LLM would see)
    const projection = doc.project("json");
    expect(projection.body).toHaveLength(3);

    // 4. Apply patches (what the LLM would produce)
    const results = doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Document Title", replace: "Contract v2" },
      { type: "replace_text", target: "p2", find: "Red large text", replace: "Blue small text" },
    ]);
    expect(results.every((r) => r.success)).toBe(true);

    // 5. Export
    const output = await doc.toBuffer();
    expect(Buffer.isBuffer(output)).toBe(true);

    // 6. Reload and verify
    const reloaded = await Docxtor.load(output);
    expect(reloaded.getParagraphs()[0]!.text).toBe("Contract v2");
    expect(reloaded.getParagraphs()[0]!.style).toBe("Heading1");

    // Formatting should be preserved
    expect(reloaded.getParagraphs()[0]!.runs[0]!.formatting.bold).toBe(true);
  });

  it("should handle table edits end-to-end", async () => {
    const buffer = await createTableDocx();
    const doc = await Docxtor.load(buffer);

    // Replace cell text + add a row
    doc.applyPatches([
      {
        type: "replace_cell_text",
        target: "tbl0.r1.c0.p0",
        find: "Cell 1",
        replace: "Updated 1",
      },
      {
        type: "add_table_row",
        target: "tbl0",
        afterRow: 1,
        cells: ["New A", "New B"],
      },
    ]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);

    const table = reloaded.getTables()[0]!;
    expect(table.rows).toHaveLength(3);
    expect(table.rows[1]!.cells[0]!.paragraphs[0]!.text).toBe("Updated 1");
    expect(table.rows[2]!.cells[0]!.paragraphs[0]!.text).toBe("New A");
  });

  it("should handle header/footer edits end-to-end", async () => {
    const buffer = await createHeaderFooterDocx();
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      {
        type: "replace_text",
        target: "hdr.default.p0",
        find: "Header text",
        replace: "CONFIDENTIAL",
      },
    ]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);

    const headers = [...reloaded.getHeaders().values()];
    expect(headers[0]!.paragraphs[0]!.text).toBe("CONFIDENTIAL");
  });

  it("should handle cross-run replacement end-to-end", async () => {
    const buffer = await createSplitRunDocx();
    const doc = await Docxtor.load(buffer);

    // "valid" spans two runs
    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "valid", replace: "binding" },
    ]);

    const output = await doc.toBuffer();
    const reloaded = await Docxtor.load(output);
    expect(reloaded.getParagraphs()[0]!.text).toBe(
      "This agreement is binding and enforceable",
    );
  });

  it("projection after patching should reflect changes", async () => {
    const buffer = await createSimpleDocx("Original text");
    const doc = await Docxtor.load(buffer);

    doc.applyPatches([
      { type: "replace_text", target: "p0", find: "Original", replace: "Updated" },
    ]);

    const projection = doc.project("json");
    expect(projection.body[0]).toMatchObject({
      id: "p0",
      text: "Updated text",
    });
  });
});
