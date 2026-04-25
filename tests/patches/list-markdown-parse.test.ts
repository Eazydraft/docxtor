import { describe, it, expect } from "vitest";
import { readDocx } from "../../src/zip/reader.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { applyPatches } from "../../src/patches/engine.js";
import { createSimpleDocx } from "../fixtures/create-docx.js";
import type { NormalizedParagraph } from "../../src/types.js";

// =============================================================================
// append_text with markdown list syntax
// =============================================================================

describe("append_text with markdown lists", () => {
  it("should parse bullet list from markdown", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "append_text",
        text: "- Apple\n- Banana\n- Cherry",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const listParas = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    expect(listParas.length).toBe(3);
    expect(listParas[0]!.text).toBe("Apple");
    expect(listParas[1]!.text).toBe("Banana");
    expect(listParas[2]!.text).toBe("Cherry");
    expect(listParas[0]!.listInfo!.listType).toBe("bullet");
  });

  it("should parse ordered list from markdown", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "append_text",
        text: "1. First\n2. Second\n3. Third",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const listParas = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    expect(listParas.length).toBe(3);
    expect(listParas[0]!.text).toBe("First");
    expect(listParas[0]!.listInfo!.listType).toBe("ordered");
  });

  it("should parse nested bullet list from markdown", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "append_text",
        text: "- Top\n  - Nested\n    - Deep",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const listParas = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    expect(listParas.length).toBe(3);
    expect(listParas[0]!.listInfo!.level).toBe(0);
    expect(listParas[1]!.listInfo!.level).toBe(1);
    expect(listParas[2]!.listInfo!.level).toBe(2);
  });

  it("should handle mixed text and list in append_text", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "append_text",
        text: "Some intro text\n- Item A\n- Item B\nSome outro text",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );

    // Should have: original "Hello", "Some intro text", 2 list items, "Some outro text"
    const listParas = paras.filter((p) => p.listInfo);
    expect(listParas.length).toBe(2);
    expect(listParas[0]!.text).toBe("Item A");
    expect(listParas[1]!.text).toBe("Item B");
  });

  it("should use * and + as bullet markers too", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "append_text",
        text: "* Star\n+ Plus",
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const listParas = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    expect(listParas.length).toBe(2);
    expect(listParas[0]!.text).toBe("Star");
    expect(listParas[1]!.text).toBe("Plus");
  });
});
