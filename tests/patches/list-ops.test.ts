import { describe, it, expect } from "vitest";
import { readDocx } from "../../src/zip/reader.js";
import { normalizeDocument } from "../../src/normalizer/document.js";
import { applyPatches } from "../../src/patches/engine.js";
import { serializeXml } from "../../src/xml/serializer.js";
import {
  createSimpleDocx,
  createListDocx,
  createMultiParagraphDocx,
} from "../fixtures/create-docx.js";
import type { Patch, NormalizedParagraph } from "../../src/types.js";

// =============================================================================
// add_list
// =============================================================================

describe("add_list", () => {
  it("should add a bullet list to a simple document", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "add_list",
        listType: "bullet",
        items: ["First", "Second", "Third"],
      },
    ]);

    expect(results[0]!.success).toBe(true);

    // Should have original paragraph + 3 list items
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras.length).toBe(4);

    // The new paragraphs should have listInfo
    const listParas = paras.filter((p) => p.listInfo);
    expect(listParas.length).toBe(3);
    expect(listParas[0]!.text).toBe("First");
    expect(listParas[1]!.text).toBe("Second");
    expect(listParas[2]!.text).toBe("Third");
    expect(listParas[0]!.listInfo!.listType).toBe("bullet");

    // Numbering XML should exist
    expect(doc.xmlParts.numbering).toBeDefined();
  });

  it("should add an ordered list", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "add_list",
        listType: "ordered",
        items: ["Alpha", "Beta"],
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const listParas = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    expect(listParas.length).toBe(2);
    expect(listParas[0]!.listInfo!.listType).toBe("ordered");
  });

  it("should add list items with levels", async () => {
    const buffer = await createSimpleDocx("Hello");
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "add_list",
        listType: "bullet",
        items: [
          { text: "Top", level: 0 },
          { text: "Nested", level: 1 },
          { text: "Back", level: 0 },
        ],
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
    expect(listParas[2]!.listInfo!.level).toBe(0);
  });

  it("should add list after a specific paragraph", async () => {
    const buffer = await createMultiParagraphDocx(["First", "Second", "Third"]);
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const p1Id = (doc.body[0] as NormalizedParagraph).id;

    const results = applyPatches(doc, [
      {
        type: "add_list",
        listType: "bullet",
        items: ["ListItem"],
        afterId: p1Id,
      },
    ]);

    expect(results[0]!.success).toBe(true);
    // List item should be after p0 and before p1(original)
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras[1]!.text).toBe("ListItem");
    expect(paras[1]!.listInfo).toBeDefined();
  });
});

// =============================================================================
// add_list_item
// =============================================================================

describe("add_list_item", () => {
  it("should add an item after an existing list item", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const firstItem = doc.body[0] as NormalizedParagraph;
    expect(firstItem.listInfo).toBeDefined();

    const results = applyPatches(doc, [
      {
        type: "add_list_item",
        target: firstItem.id,
        position: "after",
        text: "New item",
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras.length).toBe(4);
    expect(paras[1]!.text).toBe("New item");
    expect(paras[1]!.listInfo).toBeDefined();
  });

  it("should add an item before an existing list item", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const secondItem = doc.body[1] as NormalizedParagraph;

    const results = applyPatches(doc, [
      {
        type: "add_list_item",
        target: secondItem.id,
        position: "before",
        text: "Inserted before",
      },
    ]);

    expect(results[0]!.success).toBe(true);
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras.length).toBe(4);
    expect(paras[1]!.text).toBe("Inserted before");
  });
});

// =============================================================================
// remove_list_item
// =============================================================================

describe("remove_list_item", () => {
  it("should remove a list item", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const secondItem = doc.body[1] as NormalizedParagraph;

    const results = applyPatches(doc, [
      { type: "remove_list_item", target: secondItem.id },
    ]);

    expect(results[0]!.success).toBe(true);
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras.length).toBe(2);
    expect(paras[0]!.text).toBe("Item one");
    expect(paras[1]!.text).toBe("Item three");
  });

  it("should remove a list item with revision tracking", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const secondItem = doc.body[1] as NormalizedParagraph;
    expect(secondItem.text).toBe("Item two");

    const results = applyPatches(doc, [
      { type: "remove_list_item", target: secondItem.id },
    ], { trackRevisions: true, revisionAuthor: "test" });

    expect(results[0]!.success).toBe(true);

    // With revision tracking, the paragraph should still exist in doc.body
    // but its text should be empty (runs are wrapped in w:del)
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras.length).toBe(3);
    // The deleted paragraph's text should be empty
    expect(paras[1]!.text).toBe("");
    // It should have a deletion revision
    expect(paras[1]!.revisions).toBeDefined();
    expect(paras[1]!.revisions!.length).toBeGreaterThan(0);
    expect(paras[1]!.revisions![0]!.type).toBe("delete");
  });

  it("should remove a revision-inserted list item", async () => {
    // Simulate: add item with revisions, then remove it with revisions
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const opts = { trackRevisions: true, revisionAuthor: "test" };

    // Add item (will be wrapped in w:ins)
    const firstItem = doc.body[0] as NormalizedParagraph;
    applyPatches(doc, [
      { type: "add_list_item", target: firstItem.id, position: "after", text: "New item" },
    ], opts);

    const addedItem = doc.body[1] as NormalizedParagraph;
    expect(addedItem.text).toBe("New item");

    // Now remove the added item (its runs are inside w:ins)
    const results = applyPatches(doc, [
      { type: "remove_list_item", target: addedItem.id },
    ], opts);

    expect(results[0]!.success).toBe(true);

    // The paragraph text should be empty after removal
    const updated = doc.body.find((e) => e.id === addedItem.id) as NormalizedParagraph;
    expect(updated.text).toBe("");
  });
});

// =============================================================================
// convert_to_list
// =============================================================================

describe("convert_to_list", () => {
  it("should convert plain paragraphs to a bullet list", async () => {
    const buffer = await createMultiParagraphDocx(["Apple", "Banana", "Cherry"]);
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const ids = doc.body
      .filter((e): e is NormalizedParagraph => e.kind === "paragraph")
      .map((p) => p.id);

    const results = applyPatches(doc, [
      { type: "convert_to_list", targets: ids, listType: "bullet" },
    ]);

    expect(results[0]!.success).toBe(true);
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras.every((p) => p.listInfo != null)).toBe(true);
    expect(paras[0]!.listInfo!.listType).toBe("bullet");

    // XML should contain w:numPr
    const xml = serializeXml(doc.xmlParts.document);
    expect(xml).toContain("w:numPr");
  });

  it("should convert paragraphs to an ordered list", async () => {
    const buffer = await createMultiParagraphDocx(["One", "Two"]);
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const ids = doc.body
      .filter((e): e is NormalizedParagraph => e.kind === "paragraph")
      .map((p) => p.id);

    const results = applyPatches(doc, [
      { type: "convert_to_list", targets: ids, listType: "ordered" },
    ]);

    expect(results[0]!.success).toBe(true);
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras[0]!.listInfo!.listType).toBe("ordered");
  });
});

// =============================================================================
// remove_list
// =============================================================================

describe("remove_list", () => {
  it("should strip list formatting from paragraphs", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const ids = doc.body
      .filter((e): e is NormalizedParagraph => e.kind === "paragraph")
      .map((p) => p.id);

    // All should be list items initially
    const before = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    expect(before.length).toBe(3);

    const results = applyPatches(doc, [
      { type: "remove_list", targets: ids },
    ]);

    expect(results[0]!.success).toBe(true);
    const after = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    expect(after.length).toBe(0);

    // Text should still be there
    const paras = doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
    expect(paras.length).toBe(3);
    expect(paras[0]!.text).toBe("Item one");
  });
});

// =============================================================================
// set_list_properties
// =============================================================================

describe("set_list_properties", () => {
  it("should change the level of a list item", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const secondItem = doc.body[1] as NormalizedParagraph;
    expect(secondItem.listInfo!.level).toBe(0);

    const results = applyPatches(doc, [
      { type: "set_list_properties", target: secondItem.id, level: 2 },
    ]);

    expect(results[0]!.success).toBe(true);
    const updated = doc.body[1] as NormalizedParagraph;
    expect(updated.listInfo!.level).toBe(2);
  });

  it("should change the list type of an item", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const firstItem = doc.body[0] as NormalizedParagraph;
    expect(firstItem.listInfo!.listType).toBe("bullet");

    const results = applyPatches(doc, [
      { type: "set_list_properties", target: firstItem.id, listType: "ordered" },
    ]);

    expect(results[0]!.success).toBe(true);
    const updated = doc.body[0] as NormalizedParagraph;
    expect(updated.listInfo!.listType).toBe("ordered");
  });
});

// =============================================================================
// Round-trip: add list, export, reload
// =============================================================================

describe("list round-trip", () => {
  it("should survive a write-then-read cycle", async () => {
    const buffer = await createSimpleDocx("Hello");
    const doc = await (await import("../../src/index.js")).Docxtor.load(buffer);

    doc.applyPatches([
      {
        type: "add_list",
        listType: "ordered",
        items: ["Step 1", "Step 2"],
      },
    ]);

    const output = await doc.toBuffer();

    // Reload
    const doc2 = await (await import("../../src/index.js")).Docxtor.load(output);
    const paras = doc2.getParagraphs();
    const listParas = paras.filter((p) => p.listInfo);
    expect(listParas.length).toBe(2);
    expect(listParas[0]!.text).toBe("Step 1");
    expect(listParas[0]!.listInfo!.listType).toBe("ordered");
  });
});

// =============================================================================
// add_list with numId reuse
// =============================================================================

describe("add_list with numId reuse", () => {
  it("should continue numbering when reusing numId from existing list", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    // The existing list uses numId "1"
    const lastItem = doc.body[2] as NormalizedParagraph;
    expect(lastItem.listInfo!.numId).toBe("1");

    const results = applyPatches(doc, [
      {
        type: "add_list",
        listType: "bullet",
        numId: "1",
        items: ["Item four", "Item five"],
        afterId: lastItem.id,
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const listParas = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    expect(listParas.length).toBe(5);

    // All should share the same numId
    expect(listParas.every((p) => p.listInfo!.numId === "1")).toBe(true);
    expect(listParas[3]!.text).toBe("Item four");
    expect(listParas[4]!.text).toBe("Item five");
  });

  it("should create a new numId when numId is not provided", async () => {
    const buffer = await createListDocx();
    const entries = await readDocx(buffer);
    const doc = normalizeDocument(entries);

    const results = applyPatches(doc, [
      {
        type: "add_list",
        listType: "ordered",
        items: ["New list item"],
      },
    ]);

    expect(results[0]!.success).toBe(true);

    const listParas = doc.body.filter(
      (e): e is NormalizedParagraph =>
        e.kind === "paragraph" && !!e.listInfo,
    );
    // The new item should have a different numId than the existing list
    const newItem = listParas[listParas.length - 1]!;
    expect(newItem.listInfo!.numId).not.toBe("1");
  });
});
