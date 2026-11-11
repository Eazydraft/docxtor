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
  HeaderFooter,
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  NormalizedPageBreak,
  FXPNode,
} from "../types.js";
import { findChild, getChildren, getTagName } from "../xml/helpers.js";
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
 * Searches w:body, w:tc (table cells, recursively), and header/footer roots.
 */
export function findParentChildren(
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
  const fresh = normalizeParagraph(para.xmlRef, para.id);
  para.text = fresh.text;
  para.runs = fresh.runs;
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

/**
 * Resolve a `container: { type, variant, section? }` to one header or footer.
 *
 * Without `section`, the container must name exactly one part: when several
 * sections have a header or footer of that variant, the result is an error
 * that lists them. With `section`, that section must have one.
 */
export function resolveHeaderFooterContainer(
  doc: NormalizedDocument,
  container: { type: "header" | "footer"; variant: string; section?: number },
): { hf: HeaderFooter; rootNode: FXPNode } | { error: string } {
  const map = container.type === "header" ? doc.headers : doc.footers;
  const candidates = [...map.values()].filter((hf) => hf.variant === container.variant);
  if (candidates.length === 0) {
    return { error: `No ${container.type} with variant "${container.variant}" found` };
  }

  const sections = candidates.map((hf) => hf.sectionIndex);
  const list =
    sections.length > 1
      ? `sections ${sections.slice(0, -1).join(", ")} and ${sections[sections.length - 1]}`
      : `section ${sections[0]}`;

  let hf: HeaderFooter | undefined;
  if (container.section === undefined) {
    if (candidates.length > 1) {
      return {
        error: `${container.type} '${container.variant}' exists in ${list}; say which with \`section\``,
      };
    }
    hf = candidates[0];
  } else {
    hf = candidates.find((c) => c.sectionIndex === container.section);
    if (!hf) {
      return {
        error: `No ${container.type} with variant "${container.variant}" in section ${container.section}; it exists in ${list}`,
      };
    }
  }

  const rootTag = container.type === "header" ? "w:hdr" : "w:ftr";
  const rootNode = hf!.xmlRef.find((n) => getTagName(n) === rootTag);
  if (!rootNode) return { error: `${container.type} part has no ${rootTag} root` };
  return { hf: hf!, rootNode };
}
