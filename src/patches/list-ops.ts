/**
 * List operation executors for the 6 list patch types.
 *
 * Each executor mutates the FXP tree directly and updates the
 * NormalizedDocument accordingly.
 */

import type {
  NormalizedDocument,
  NormalizedParagraph,
  FXPNode,
  AddListPatch,
  AddListItemPatch,
  RemoveListItemPatch,
  ConvertToListPatch,
  RemoveListPatch,
  SetListPropertiesPatch,
} from "../types.js";
import { createNumberingDefinition } from "./numbering.js";
import { enrichListType } from "../normalizer/document.js";
import { normalizeParagraph } from "../normalizer/paragraph.js";
import {
  createElement,
  findChild,
  getChildren,
  getTagName,
  setAttr,
  cloneNode,
} from "../xml/helpers.js";
import { buildFieldRuns } from "./field-codes.js";
import { findParagraphById, findElementById } from "./validator.js";
import { createInsElement, createPPrChange, convertRunToDeleted, createDelElement, type RevisionContext } from "./revision.js";

// =============================================================================
// Helpers
// =============================================================================

/**
 * Build a w:p node with w:numPr for a list item.
 */
function buildListParagraphNode(
  text: string,
  numId: string,
  level: number,
  revCtx?: RevisionContext,
): FXPNode {
  const numPr = createElement("w:numPr", [
    createElement("w:ilvl", [], { "w:val": String(level) }),
    createElement("w:numId", [], { "w:val": numId }),
  ]);
  const pPr = createElement("w:pPr", [numPr]);

  const runs = buildFieldRuns(text);
  let children: FXPNode[];
  if (revCtx) {
    const insElement = createInsElement(revCtx, runs);
    children = [pPr, insElement];
  } else {
    children = [pPr, ...runs];
  }

  return { "w:p": children };
}

/**
 * Get the w:body children array from the document FXP tree.
 */
function getBodyChildren(doc: NormalizedDocument): FXPNode[] {
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
function insertIntoBody(
  bodyChildren: FXPNode[],
  nodes: FXPNode[],
  afterXmlRef?: FXPNode,
): void {
  if (afterXmlRef) {
    const idx = bodyChildren.indexOf(afterXmlRef);
    if (idx !== -1) {
      bodyChildren.splice(idx + 1, 0, ...nodes);
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
 * Compute the next available paragraph ID index.
 */
function getNextPIdx(doc: NormalizedDocument): number {
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
 * Find the FXP parent container of a paragraph's xmlRef.
 * Searches w:body, w:tc (table cells), and header/footer roots.
 */
function findParentChildren(
  doc: NormalizedDocument,
  xmlRef: FXPNode,
): FXPNode[] | undefined {
  // Check body
  const bodyChildren = getBodyChildren(doc);
  if (bodyChildren.includes(xmlRef)) return bodyChildren;

  // Check table cells in body
  for (const child of bodyChildren) {
    if (getTagName(child) === "w:tbl") {
      const result = findInTable(child, xmlRef);
      if (result) return result;
    }
  }

  // Check headers/footers
  for (const hf of [...doc.headers.values(), ...doc.footers.values()]) {
    const rootTag = hf.type === "header" ? "w:hdr" : "w:ftr";
    const rootNode = hf.xmlRef.find((n) => getTagName(n) === rootTag);
    if (rootNode) {
      const rootChildren = getChildren(rootNode);
      if (rootChildren.includes(xmlRef)) return rootChildren;
      for (const child of rootChildren) {
        if (getTagName(child) === "w:tbl") {
          const result = findInTable(child, xmlRef);
          if (result) return result;
        }
      }
    }
  }

  return undefined;
}

function findInTable(tblNode: FXPNode, target: FXPNode): FXPNode[] | undefined {
  for (const tr of getChildren(tblNode).filter((n) => getTagName(n) === "w:tr")) {
    for (const tc of getChildren(tr).filter((n) => getTagName(n) === "w:tc")) {
      const tcChildren = getChildren(tc);
      if (tcChildren.includes(target)) return tcChildren;
      // Check nested tables
      for (const child of tcChildren) {
        if (getTagName(child) === "w:tbl") {
          const result = findInTable(child, target);
          if (result) return result;
        }
      }
    }
  }
  return undefined;
}

/**
 * Find the normalized paragraph array that contains a given paragraph ID.
 * Returns { array, index } for doc.body, hf.paragraphs, or cell.paragraphs.
 */
function findNormalizedContainer(
  doc: NormalizedDocument,
  paraId: string,
): { array: Array<import("../types.js").NormalizedParagraph | import("../types.js").NormalizedTable | import("../types.js").NormalizedPageBreak>; index: number } | undefined {
  // Check doc.body
  const bodyIdx = doc.body.findIndex((el) => el.id === paraId);
  if (bodyIdx !== -1) return { array: doc.body, index: bodyIdx };

  // Check header/footer paragraphs
  for (const hf of [...doc.headers.values(), ...doc.footers.values()]) {
    const hfIdx = hf.paragraphs.findIndex((p) => p.id === paraId);
    if (hfIdx !== -1) return { array: hf.paragraphs as any, index: hfIdx };
  }

  // Check table cell paragraphs (body tables + hf tables)
  const allTables: import("../types.js").NormalizedTable[] = [];
  for (const el of doc.body) {
    if (el.kind === "table") allTables.push(el);
  }
  for (const hf of [...doc.headers.values(), ...doc.footers.values()]) {
    allTables.push(...hf.tables);
  }

  for (const table of allTables) {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        const cellIdx = cell.paragraphs.findIndex((p) => p.id === paraId);
        if (cellIdx !== -1) return { array: cell.paragraphs as any, index: cellIdx };
        // Check nested tables
        for (const nested of cell.nestedTables) {
          for (const nr of nested.rows) {
            for (const nc of nr.cells) {
              const ncIdx = nc.paragraphs.findIndex((p) => p.id === paraId);
              if (ncIdx !== -1) return { array: nc.paragraphs as any, index: ncIdx };
            }
          }
        }
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
function deriveNewParagraphId(
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
function reNormalize(doc: NormalizedDocument, para: NormalizedParagraph): void {
  const fresh = normalizeParagraph(para.xmlRef, para.id);
  para.text = fresh.text;
  para.runs = fresh.runs;
  para.style = fresh.style;
  para.listInfo = fresh.listInfo;
  para.revisions = fresh.revisions;
  if (doc.xmlParts.numbering) {
    enrichListType(para, doc.xmlParts.numbering);
  }
}

// =============================================================================
// Executors
// =============================================================================

export function executeAddList(
  doc: NormalizedDocument,
  patch: AddListPatch,
  revCtx?: RevisionContext,
): void {
  const numId = patch.numId ?? createNumberingDefinition(doc, patch.listType);
  const defaultLevel = patch.level ?? 0;
  const nodes: FXPNode[] = [];

  for (const item of patch.items) {
    const text = typeof item === "string" ? item : item.text;
    const level = typeof item === "string" ? defaultLevel : (item.level ?? defaultLevel);
    nodes.push(buildListParagraphNode(text, numId, level, revCtx));
  }

  if (patch.container) {
    // Insert into header/footer
    const map = patch.container.type === "header" ? doc.headers : doc.footers;
    for (const hf of map.values()) {
      if (hf.variant === patch.container.variant) {
        const rootTag = hf.type === "header" ? "w:hdr" : "w:ftr";
        const rootNode = hf.xmlRef.find((n) => getTagName(n) === rootTag);
        if (rootNode) {
          const rootChildren = getChildren(rootNode);
          rootChildren.push(...nodes);
          const prefix = patch.container.type === "header" ? "hdr" : "ftr";
          let nextPIdx = hf.paragraphs.length;
          for (const node of nodes) {
            const normalized = normalizeParagraph(
              node,
              `${prefix}.${patch.container.variant}.p${nextPIdx++}`,
            );
            if (doc.xmlParts.numbering) enrichListType(normalized, doc.xmlParts.numbering);
            hf.paragraphs.push(normalized);
          }
        }
        return;
      }
    }
    return;
  }

  const bodyChildren = getBodyChildren(doc);
  let afterXmlRef: FXPNode | undefined;
  if (patch.afterId) {
    const target = findElementById(patch.afterId, doc);
    if (target) afterXmlRef = target.xmlRef;
  }

  insertIntoBody(bodyChildren, nodes, afterXmlRef);

  // Normalize and add to doc.body
  let nextPIdx = getNextPIdx(doc);
  const insertIdx = patch.afterId
    ? doc.body.findIndex((el) => el.id === patch.afterId)
    : -1;

  for (let i = 0; i < nodes.length; i++) {
    const id = `p${nextPIdx++}`;
    const normalized = normalizeParagraph(nodes[i]!, id);
    if (doc.xmlParts.numbering) enrichListType(normalized, doc.xmlParts.numbering);
    if (insertIdx !== -1) {
      doc.body.splice(insertIdx + 1 + i, 0, normalized);
    } else {
      doc.body.push(normalized);
    }
  }
}

export function executeAddListItem(
  doc: NormalizedDocument,
  patch: AddListItemPatch,
  revCtx?: RevisionContext,
): void {
  const targetPara = findParagraphById(patch.target, doc)!;
  const numId = targetPara.listInfo!.numId;
  const level = patch.level ?? targetPara.listInfo!.level;

  const newNode = buildListParagraphNode(patch.text, numId, level, revCtx);

  // Find the parent FXP container
  const parentChildren = findParentChildren(doc, targetPara.xmlRef);
  if (!parentChildren) throw new Error("Cannot find parent container for target paragraph");

  const targetIdx = parentChildren.indexOf(targetPara.xmlRef);
  if (targetIdx === -1) throw new Error("Target xmlRef not found in parent");

  if (patch.position === "after") {
    parentChildren.splice(targetIdx + 1, 0, newNode);
  } else {
    parentChildren.splice(targetIdx, 0, newNode);
  }

  // Normalize and insert into the correct normalized container
  const container = findNormalizedContainer(doc, patch.target);
  const containerArray = container ? container.array : doc.body;
  const id = deriveNewParagraphId(patch.target, containerArray as readonly { id: string }[]);
  const normalized = normalizeParagraph(newNode, id);
  if (doc.xmlParts.numbering) enrichListType(normalized, doc.xmlParts.numbering);

  if (container) {
    if (patch.position === "after") {
      container.array.splice(container.index + 1, 0, normalized as any);
    } else {
      container.array.splice(container.index, 0, normalized as any);
    }
  } else {
    doc.body.push(normalized);
  }
}

export function executeRemoveListItem(
  doc: NormalizedDocument,
  patch: RemoveListItemPatch,
  revCtx?: RevisionContext,
): void {
  const targetPara = findParagraphById(patch.target, doc)!;

  if (revCtx) {
    // Mark runs as deleted instead of removing
    const children = getChildren(targetPara.xmlRef);
    // Store old pPr for pPrChange
    const oldPPr = findChild(targetPara.xmlRef, "w:pPr");
    if (oldPPr) {
      const pPrChange = createPPrChange(revCtx, oldPPr);
      const pPrChildren = getChildren(oldPPr);
      pPrChildren.push(pPrChange);
    }
    // Wrap runs in w:del — handle both direct w:r and w:r nested inside w:ins
    for (let i = children.length - 1; i >= 0; i--) {
      const child = children[i]!;
      const tag = getTagName(child);
      if (tag === "w:r") {
        // Clone the run and convert w:t → w:delText for valid OOXML
        const cloned = cloneNode(child);
        convertRunToDeleted(cloned);
        const delElement = createDelElement(revCtx, [cloned]);
        children[i] = delElement;
      } else if (tag === "w:ins") {
        // Inserted but not accepted text: removing the w:ins entirely is the
        // correct OOXML semantics (the content was never part of the accepted
        // document, so wrapping it in w:del would corrupt the file).
        children.splice(i, 1);
      }
    }
    reNormalize(doc, targetPara);
    return;
  }

  // Physical removal
  const parentChildren = findParentChildren(doc, targetPara.xmlRef);
  if (parentChildren) {
    const idx = parentChildren.indexOf(targetPara.xmlRef);
    if (idx !== -1) parentChildren.splice(idx, 1);
  }

  // Remove from the correct normalized container (body, hf.paragraphs, or cell.paragraphs)
  const container = findNormalizedContainer(doc, patch.target);
  if (container) {
    container.array.splice(container.index, 1);
  }
}

export function executeConvertToList(
  doc: NormalizedDocument,
  patch: ConvertToListPatch,
  revCtx?: RevisionContext,
): void {
  const numId = createNumberingDefinition(doc, patch.listType);
  const level = patch.level ?? 0;

  for (const target of patch.targets) {
    const para = findParagraphById(target, doc)!;
    const pNode = para.xmlRef;

    // Find or create w:pPr
    let pPr = findChild(pNode, "w:pPr");
    const pChildren = getChildren(pNode);

    if (revCtx && pPr) {
      const pPrChange = createPPrChange(revCtx, pPr);
      getChildren(pPr).push(pPrChange);
    }

    if (!pPr) {
      pPr = createElement("w:pPr", []);
      pChildren.unshift(pPr);
    }

    const pPrChildren = getChildren(pPr);

    // Remove existing w:numPr if present
    const existingIdx = pPrChildren.findIndex((c) => getTagName(c) === "w:numPr");
    if (existingIdx !== -1) pPrChildren.splice(existingIdx, 1);

    // Add new w:numPr
    const numPr = createElement("w:numPr", [
      createElement("w:ilvl", [], { "w:val": String(level) }),
      createElement("w:numId", [], { "w:val": numId }),
    ]);
    pPrChildren.unshift(numPr);

    reNormalize(doc, para);
  }
}

export function executeRemoveList(
  doc: NormalizedDocument,
  patch: RemoveListPatch,
  revCtx?: RevisionContext,
): void {
  for (const target of patch.targets) {
    const para = findParagraphById(target, doc)!;
    const pPr = findChild(para.xmlRef, "w:pPr");
    if (!pPr) continue;

    if (revCtx) {
      const pPrChange = createPPrChange(revCtx, pPr);
      getChildren(pPr).push(pPrChange);
    }

    const pPrChildren = getChildren(pPr);
    const numPrIdx = pPrChildren.findIndex((c) => getTagName(c) === "w:numPr");
    if (numPrIdx !== -1) pPrChildren.splice(numPrIdx, 1);

    reNormalize(doc, para);
  }
}

export function executeSetListProperties(
  doc: NormalizedDocument,
  patch: SetListPropertiesPatch,
  revCtx?: RevisionContext,
): void {
  const para = findParagraphById(patch.target, doc)!;
  const pPr = findChild(para.xmlRef, "w:pPr");
  if (!pPr) throw new Error("No w:pPr found on target paragraph");

  if (revCtx) {
    const pPrChange = createPPrChange(revCtx, pPr);
    getChildren(pPr).push(pPrChange);
  }

  const numPr = findChild(pPr, "w:numPr");
  if (!numPr) throw new Error("No w:numPr found on target paragraph");

  if (patch.level !== undefined) {
    const ilvl = findChild(numPr, "w:ilvl");
    if (ilvl) {
      setAttr(ilvl, "w:val", String(patch.level));
    } else {
      getChildren(numPr).unshift(
        createElement("w:ilvl", [], { "w:val": String(patch.level) }),
      );
    }
  }

  if (patch.listType !== undefined) {
    const newNumId = createNumberingDefinition(doc, patch.listType);
    const numIdNode = findChild(numPr, "w:numId");
    if (numIdNode) {
      setAttr(numIdNode, "w:val", newNumId);
    } else {
      getChildren(numPr).push(
        createElement("w:numId", [], { "w:val": newNumId }),
      );
    }
  }

  reNormalize(doc, para);
}
