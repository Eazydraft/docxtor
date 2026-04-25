import { describe, it, expect } from "vitest";
import { readDocx } from "../../src/zip/reader.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { applyPatches } from "../../src/patches/engine.js";
import { project, type JsonProjection } from "../../src/projection/llm.js";
import {
  createListDocx,
  createMultiParagraphDocx,
  createSimpleDocx,
} from "../fixtures/create-docx.js";
import type { NormalizedParagraph } from "../../src/types.js";

// =============================================================================
// Markdown projection of lists
// =============================================================================

describe("list markdown projection", () => {
  it("should render bullet list items with dash prefix", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const md = project(doc, "markdown") as string;

    // Should contain bullet markers
    expect(md).toContain("- Item one");
    expect(md).toContain("- Item two");
    expect(md).toContain("- Item three");
  });

  it("should render ordered list items with numbers", async () => {
    const buffer = await createMultiParagraphDocx(["Alpha", "Beta", "Gamma"]);
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // Convert to ordered list
    const ids = doc.body
      .filter((e): e is NormalizedParagraph => e.kind === "paragraph")
      .map((p) => p.id);

    applyPatches(doc, [
      { type: "convert_to_list", targets: ids, listType: "ordered" },
    ]);

    const md = project(doc, "markdown") as string;

    expect(md).toContain("1. Alpha");
    expect(md).toContain("2. Beta");
    expect(md).toContain("3. Gamma");
  });

  it("should render nested bullet items with indentation", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    applyPatches(doc, [
      {
        type: "add_list",
        listType: "bullet",
        items: [
          { text: "Top", level: 0 },
          { text: "Nested", level: 1 },
          { text: "Deep", level: 2 },
        ],
      },
    ]);

    const md = project(doc, "markdown") as string;

    expect(md).toContain("- Top");
    expect(md).toContain("  - Nested");
    expect(md).toContain("    - Deep");
  });
});

// =============================================================================
// JSON projection of lists
// =============================================================================

describe("list JSON projection", () => {
  it("should include listType and numId in JSON projection", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const json = project(doc, "json") as JsonProjection;

    // Find list paragraph entries
    const listEntries = json.body.filter(
      (e: any) => e.listType,
    ) as any[];
    expect(listEntries.length).toBe(3);
    expect(listEntries[0].listType).toBe("bullet");
    expect(listEntries[0].numId).toBeDefined();
  });

  it("should include listType for ordered lists", async () => {
    const buffer = await createMultiParagraphDocx(["A", "B"]);
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const ids = doc.body
      .filter((e): e is NormalizedParagraph => e.kind === "paragraph")
      .map((p) => p.id);

    applyPatches(doc, [
      { type: "convert_to_list", targets: ids, listType: "ordered" },
    ]);

    const json = project(doc, "json") as JsonProjection;
    const listEntries = json.body.filter(
      (e: any) => e.listType,
    ) as any[];
    expect(listEntries.length).toBe(2);
    expect(listEntries[0].listType).toBe("ordered");
  });
});
