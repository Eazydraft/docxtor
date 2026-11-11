/**
 * Shared helpers for locating a paragraph's containers — both in the FXP
 * tree (the parent children array holding its `w:p`) and in the normalized
 * view (the `doc.body` / `hf.paragraphs` / `cell.paragraphs` array holding
 * its NormalizedParagraph).
 *
 * Used by list ops, paragraph ops, and the engine for insertions and
 * removals that must keep both views in sync.
 */

import type {
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  NormalizedPageBreak,
  FXPNode,
} from "../types.js";
import { findChild, getChildren, getTagName, contentChildrenOf } from "../xml/helpers.js";
import { normalizeParagraph } from "../normalizer/paragraph.js";
import { enrichListType } from "../normalizer/document.js";

/** Any array that holds normalized paragraphs (body, hf.paragraphs, cell.paragraphs). */
export type NormalizedContainerArray = Array<
  NormalizedParagraph | NormalizedTable | NormalizedPageBreak
>;

/**
 * Get the w:body children array from the document FXP tree.
 */
export function getBodyChildren(doc: NormalizedDocument): FXPNode[] {
  const docNode = doc.xmlParts.document.find(
    (n) => getTagName(n) === "w:document",
  );
  if (!docNode) throw new Error("Cannot find w:document in FXP tree");
  const bodyNode = findChild(docNode, "w:body");
  if (!bodyNode) throw new Error("Cannot find w:body in FXP tree");
  return getChildren(bodyNode);
}

/**
 * Insert FXP nodes into the body before w:sectPr, or after a specific element.
 */
export function insertIntoBody(
  bodyChildren: FXPNode[],
  nodes: FXPNode[],
  afterXmlRef?: FXPNode,
): void {
  if (afterXmlRef) {
    // The element may sit inside a block-level content control: the new
    // nodes go next to it, in the array that really holds it.
    const holder = contentChildrenOf(bodyChildren).find((c) => c.node === afterXmlRef);
    if (holder) {
      holder.siblings.splice(holder.siblings.indexOf(afterXmlRef) + 1, 0, ...nodes);
      return;
    }
  }
  const sectPrIndex = bodyChildren.findIndex(
    (n) => getTagName(n) === "w:sectPr",
  );
  if (sectPrIndex !== -1) {
    bodyChildren.splice(sectPrIndex, 0, ...nodes);
  } else {
    bodyChildren.push(...nodes);
  }
}

/**
 * Compute the next available body paragraph ID index.
 */
export function getNextPIdx(doc: NormalizedDocument): number {
  let max = -1;
  for (const el of doc.body) {
    if (el.kind === "paragraph") {
      const num = parseInt(el.id.replace("p", ""), 10);
      if (!isNaN(num) && num > max) max = num;
    }
  }
  return max + 1;
}

/**
 * Find the FXP parent children array that directly contains a paragraph's xmlRef.
 * Searches w:body, w:tc (table cells, recursively), and header/footer roots,
 * reading through content controls: for a paragraph inside a block-level
 * control this is the children of its `w:sdtContent`.
 */
export function findParentChildren(
  doc: NormalizedDocument,
  xmlRef: FXPNode,
): FXPNode[] | undefined {
  const inBody = findInBlocks(getBodyChildren(doc), xmlRef);
  if (inBody) return inBody;

  // Check headers/footers
  for (const hf of [...doc.headers.values(), ...doc.footers.values()]) {
    const rootTag = hf.type === "header" ? "w:hdr" : "w:ftr";
    const rootNode = hf.xmlRef.find((n) => getTagName(n) === rootTag);
    if (rootNode) {
      const result = findInBlocks(getChildren(rootNode), xmlRef);
      if (result) return result;
    }
  }

  return undefined;
}

/**
 * Search block-level content (`children` of a body, cell, header or footer)
 * for `target`, descending into tables. Returns the array that holds it.
 */
function findInBlocks(children: FXPNode[], target: FXPNode): FXPNode[] | undefined {
  const blocks = contentChildrenOf(children);
  for (const block of blocks) {
    if (block.node === target) return block.siblings;
  }
  for (const block of blocks) {
    if (getTagName(block.node) !== "w:tbl") continue;
    for (const tr of contentChildrenOf(getChildren(block.node))) {
      if (getTagName(tr.node) !== "w:tr") continue;
      for (const tc of contentChildrenOf(getChildren(tr.node))) {
        if (getTagName(tc.node) !== "w:tc") continue;
        const result = findInBlocks(getChildren(tc.node), target);
        if (result) return result;
      }
    }
  }
  return undefined;
}

/**
 * Find the normalized paragraph array that contains a given paragraph ID.
 * Returns { array, index } for doc.body, hf.paragraphs, or cell.paragraphs
 * (searching nested tables recursively).
 */
export function findNormalizedContainer(
  doc: NormalizedDocument,
  paraId: string,
): { array: NormalizedContainerArray; index: number } | undefined {
  // Check doc.body
  const bodyIdx = doc.body.findIndex((el) => el.id === paraId);
  if (bodyIdx !== -1) return { array: doc.body, index: bodyIdx };

  // Check header/footer paragraphs
  for (const hf of [...doc.headers.values(), ...doc.footers.values()]) {
    const hfIdx = hf.paragraphs.findIndex((p) => p.id === paraId);
    if (hfIdx !== -1) return { array: hf.paragraphs, index: hfIdx };
  }

  // Check table cell paragraphs (body tables + hf tables), recursing into nested tables
  const allTables: NormalizedTable[] = [];
  for (const el of doc.body) {
    if (el.kind === "table") allTables.push(el);
  }
  for (const hf of [...doc.headers.values(), ...doc.footers.values()]) {
    allTables.push(...hf.tables);
  }

  for (const table of allTables) {
    const found = findInNormalizedTable(table, paraId);
    if (found) return found;
  }

  return undefined;
}

function findInNormalizedTable(
  table: NormalizedTable,
  paraId: string,
): { array: NormalizedContainerArray; index: number } | undefined {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const cellIdx = cell.paragraphs.findIndex((p) => p.id === paraId);
      if (cellIdx !== -1) return { array: cell.paragraphs, index: cellIdx };
      for (const nested of cell.nestedTables) {
        const found = findInNormalizedTable(nested, paraId);
        if (found) return found;
      }
    }
  }
  return undefined;
}

/**
 * Derive a new paragraph ID in the same namespace as the target paragraph.
 * E.g., target "hdr.default.p2" → "hdr.default.p3",
 *        target "tbl0.r1.c0.p0" → "tbl0.r1.c0.p1",
 *        target "p5" → "p6" (body namespace).
 *
 * Scans sibling paragraphs in the container to find the next available index.
 */
export function deriveNewParagraphId(
  targetId: string,
  containerArray: readonly { id: string }[],
): string {
  // Extract prefix: everything up to and including the last "p" before the trailing number
  const match = targetId.match(/^(.*p)(\d+)$/);
  if (!match) {
    // Fallback — shouldn't happen for valid paragraph IDs
    return `p${Date.now()}`;
  }
  const prefix = match[1]!; // e.g. "hdr.default.p", "tbl0.r1.c0.p", "p"

  // Find the max index among siblings with the same prefix
  let maxIdx = -1;
  for (const el of containerArray) {
    if (el.id.startsWith(prefix)) {
      const suffix = el.id.slice(prefix.length);
      const num = parseInt(suffix, 10);
      if (!isNaN(num) && num > maxIdx) maxIdx = num;
    }
  }

  return `${prefix}${maxIdx + 1}`;
}

/**
 * Re-normalize a paragraph in place, enriching listType.
 */
export function reNormalizeParagraph(doc: NormalizedDocument, para: NormalizedParagraph): void {
  const fresh = normalizeParagraph(para.xmlRef, para.id, para.control);
  para.text = fresh.text;
  para.runs = fresh.runs;
  para.controls = fresh.controls;
  para.style = fresh.style;
  para.listInfo = fresh.listInfo;
  para.revisions = fresh.revisions;
  para.paraId = fresh.paraId;
  para.hasPageBreak = fresh.hasPageBreak;
  para.pageBreakOffset = fresh.pageBreakOffset;
  if (doc.xmlParts.numbering) {
    enrichListType(para, doc.xmlParts.numbering);
  }
}
