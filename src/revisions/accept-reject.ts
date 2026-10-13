/**
 * Accept or reject every tracked change in a document.
 *
 * Walks document.xml, headers, footers, footnotes, endnotes and comments,
 * resolving each revision element according to the mode:
 *
 * | Element                              | Accept                         | Reject                            |
 * |--------------------------------------|--------------------------------|-----------------------------------|
 * | `w:ins` (runs)                       | unwrap                         | remove                            |
 * | `w:del` (runs)                       | remove                         | unwrap, w:delText → w:t           |
 * | `w:moveTo` / `w:moveFrom`            | as ins / as del                | as ins / as del                   |
 * | `w:*PrChange`                        | drop                           | restore old properties            |
 * | row `w:trPr/w:ins` / `w:del`         | drop marker / remove row       | remove row / drop marker          |
 * | cell `w:cellIns` / `w:cellDel`       | drop marker / remove cell      | remove cell / drop marker         |
 * | paragraph mark `w:ins`               | drop marker                    | merge paragraph into the next     |
 * | paragraph mark `w:del`               | merge paragraph into the next  | drop marker                       |
 *
 * With an `authors` filter only revisions by those authors are touched.
 * The caller must re-normalize the document afterwards (positional ids may
 * change; `paraId`s do not).
 */

import type { NormalizedDocument, FXPNode, RevisionFilterOptions } from "../types.js";
import {
  cloneNode,
  findChild,
  getAttr,
  getChildren,
  getTagName,
  needsSpacePreserve,
  setAttr,
  unescapeXmlText,
} from "../xml/helpers.js";

type Mode = "accept" | "reject";

interface WalkContext {
  mode: Mode;
  authors?: Set<string>;
  count: number;
}

/** Property-change elements and the siblings to keep when restoring the snapshot. */
const PROP_CHANGE_KEEP: Record<string, Set<string>> = {
  "w:rPrChange": new Set(["w:ins", "w:del", "w:moveFrom", "w:moveTo"]),
  "w:pPrChange": new Set(["w:rPr", "w:sectPr"]),
  "w:trPrChange": new Set(["w:ins", "w:del"]),
  "w:tcPrChange": new Set(["w:cellIns", "w:cellDel", "w:cellMerge"]),
  "w:tblPrChange": new Set(),
  "w:tblPrExChange": new Set(),
  "w:tblGridChange": new Set(),
  "w:sectPrChange": new Set(),
};

/** Parents whose `w:ins` / `w:del` children are markers owned by w:p / w:tr, not run containers. */
const MARKER_PARENTS = new Set(["w:rPr", "w:trPr", "w:tcPr"]);

const MOVE_RANGE_TAGS = new Set([
  "w:moveFromRangeStart",
  "w:moveFromRangeEnd",
  "w:moveToRangeStart",
  "w:moveToRangeEnd",
]);

/** Elements that may sit between two paragraphs without separating them. */
const TRANSPARENT_SIBLINGS = new Set([
  "w:bookmarkStart",
  "w:bookmarkEnd",
  "w:proofErr",
  "w:permStart",
  "w:permEnd",
  "w:commentRangeStart",
  "w:commentRangeEnd",
  "w:customXmlInsRangeStart",
  "w:customXmlInsRangeEnd",
  "w:customXmlDelRangeStart",
  "w:customXmlDelRangeEnd",
]);

// =============================================================================
// Public API
// =============================================================================

export function acceptAllRevisions(
  doc: NormalizedDocument,
  options?: RevisionFilterOptions,
): { accepted: number } {
  const ctx = makeContext("accept", options);
  walkAllParts(doc, ctx);
  return { accepted: ctx.count };
}

export function rejectAllRevisions(
  doc: NormalizedDocument,
  options?: RevisionFilterOptions,
): { rejected: number } {
  const ctx = makeContext("reject", options);
  walkAllParts(doc, ctx);
  return { rejected: ctx.count };
}

function makeContext(mode: Mode, options?: RevisionFilterOptions): WalkContext {
  return {
    mode,
    authors: options?.authors ? new Set(options.authors) : undefined,
    count: 0,
  };
}

function walkAllParts(doc: NormalizedDocument, ctx: WalkContext): void {
  processChildren(doc.xmlParts.document, undefined, ctx);
  for (const part of doc.xmlParts.headers.values()) processChildren(part, undefined, ctx);
  for (const part of doc.xmlParts.footers.values()) processChildren(part, undefined, ctx);
  for (const key of ["footnotes", "endnotes", "comments"] as const) {
    const part = doc.xmlParts[key];
    if (part) processChildren(part, undefined, ctx);
  }
}

function matches(node: FXPNode, ctx: WalkContext): boolean {
  if (!ctx.authors) return true;
  const author = getAttr(node, "w:author");
  return author !== undefined && ctx.authors.has(author);
}

// =============================================================================
// Walker
// =============================================================================

/**
 * Process a children array in place. Elements may be removed, unwrapped, or
 * merged, so the loop manages its own index.
 */
function processChildren(
  children: FXPNode[],
  parentTag: string | undefined,
  ctx: WalkContext,
): void {
  let i = 0;
  while (i < children.length) {
    const child = children[i]!;
    const tag = getTagName(child);
    if (!tag) {
      i++;
      continue;
    }

    const inMarkerParent = parentTag !== undefined && MARKER_PARENTS.has(parentTag);

    // --- Run-level insertions -------------------------------------------
    if ((tag === "w:ins" || tag === "w:moveTo") && !inMarkerParent) {
      if (matches(child, ctx)) {
        ctx.count++;
        if (ctx.mode === "accept") {
          children.splice(i, 1, ...getChildren(child)); // unwrap; reprocess from i
        } else {
          children.splice(i, 1);
        }
        continue;
      }
      processChildren(getChildren(child), tag, ctx);
      i++;
      continue;
    }

    // --- Run-level deletions --------------------------------------------
    if ((tag === "w:del" || tag === "w:moveFrom") && !inMarkerParent) {
      if (matches(child, ctx)) {
        ctx.count++;
        if (ctx.mode === "accept") {
          children.splice(i, 1);
        } else {
          const inner = getChildren(child);
          for (const node of inner) restoreDeletedContent(node);
          children.splice(i, 1, ...inner);
        }
        continue;
      }
      processChildren(getChildren(child), tag, ctx);
      i++;
      continue;
    }

    // --- Move range markers ---------------------------------------------
    if (MOVE_RANGE_TAGS.has(tag)) {
      if (matches(child, ctx)) {
        children.splice(i, 1);
        continue;
      }
      i++;
      continue;
    }

    // --- Property changes -----------------------------------------------
    if (tag in PROP_CHANGE_KEEP) {
      if (matches(child, ctx)) {
        ctx.count++;
        if (ctx.mode === "accept") {
          children.splice(i, 1);
        } else {
          restorePropertyChange(children, child, tag);
        }
        i = 0; // siblings may have been replaced
        continue;
      }
      i++;
      continue;
    }

    // --- Paragraphs: content first, then the paragraph mark ------------
    if (tag === "w:p") {
      processChildren(getChildren(child), tag, ctx);
      if (handleParagraphMark(children, i, child, ctx)) continue; // merged away
      i++;
      continue;
    }

    // --- Table rows -----------------------------------------------------
    if (tag === "w:tr") {
      if (handleRowMarkers(children, i, child, ctx)) continue; // row removed
      processChildren(getChildren(child), tag, ctx);
      i++;
      continue;
    }

    // --- Table cells ----------------------------------------------------
    if (tag === "w:tc") {
      if (handleCellMarkers(children, i, child, ctx)) continue; // cell removed
      processChildren(getChildren(child), tag, ctx);
      i++;
      continue;
    }

    // --- Everything else: recurse ---------------------------------------
    const grandChildren = getChildren(child);
    if (grandChildren.length > 0) processChildren(grandChildren, tag, ctx);
    i++;
  }
}

// =============================================================================
// Deleted content restoration (reject)
// =============================================================================

/**
 * Turn deleted run content back into live content: `w:delText` → `w:t`,
 * `w:delInstrText` → `w:instrText`. Recurses through inline containers.
 */
function restoreDeletedContent(node: FXPNode): void {
  const tag = getTagName(node);
  if (!tag) return;

  if (tag === "w:r") {
    const children = getChildren(node);
    for (let i = 0; i < children.length; i++) {
      const c = children[i]!;
      if ("w:delText" in c) {
        const replacement: FXPNode = { "w:t": c["w:delText"] as FXPNode[] };
        if (c[":@"]) replacement[":@"] = c[":@"];
        const text = textOf(replacement["w:t"] as FXPNode[]);
        if (needsSpacePreserve(text)) setAttr(replacement, "xml:space", "preserve");
        children[i] = replacement;
      } else if ("w:delInstrText" in c) {
        const replacement: FXPNode = { "w:instrText": c["w:delInstrText"] as FXPNode[] };
        if (c[":@"]) replacement[":@"] = c[":@"];
        children[i] = replacement;
      }
    }
    return;
  }

  for (const child of getChildren(node)) restoreDeletedContent(child);
}

function textOf(nodes: FXPNode[]): string {
  let text = "";
  for (const n of nodes) {
    if ("#text" in n) text += unescapeXmlText(String(n["#text"]));
  }
  return text;
}

// =============================================================================
// Property change restoration (reject)
// =============================================================================

/**
 * Replace the contents of a property container (`w:rPr`, `w:pPr`, ...) with
 * the snapshot stored inside its `*Change` child, keeping the siblings that
 * are not part of the snapshot (revision markers, paragraph-mark rPr, sectPr).
 */
function restorePropertyChange(
  containerChildren: FXPNode[],
  changeNode: FXPNode,
  changeTag: string,
): void {
  const keep = PROP_CHANGE_KEEP[changeTag] ?? new Set<string>();
  const snapshot = getChildren(changeNode).find((c) => getTagName(c) !== undefined);
  const snapshotChildren = snapshot ? getChildren(snapshot).map(cloneNode) : [];

  const kept = containerChildren.filter((c) => {
    const t = getTagName(c);
    return t !== undefined && keep.has(t) && c !== changeNode;
  });

  containerChildren.length = 0;
  if (changeTag === "w:rPrChange") {
    // Paragraph-mark markers come first in CT_ParaRPr
    containerChildren.push(...kept, ...snapshotChildren);
  } else {
    containerChildren.push(...snapshotChildren, ...kept);
  }
}

// =============================================================================
// Paragraph marks
// =============================================================================

/**
 * Resolve `w:pPr/w:rPr/w:ins|w:del` on a paragraph.
 * Returns true if the paragraph was merged into the next one (and removed).
 */
function handleParagraphMark(
  siblings: FXPNode[],
  index: number,
  pNode: FXPNode,
  ctx: WalkContext,
): boolean {
  const pPr = findChild(pNode, "w:pPr");
  if (!pPr) return false;
  const rPr = findChild(pPr, "w:rPr");
  if (!rPr) return false;
  const rPrChildren = getChildren(rPr);

  let removeParagraph = false;

  const insIdx = rPrChildren.findIndex((c) => getTagName(c) === "w:ins");
  if (insIdx !== -1 && matches(rPrChildren[insIdx]!, ctx)) {
    ctx.count++;
    rPrChildren.splice(insIdx, 1);
    if (ctx.mode === "reject") removeParagraph = true;
  }

  const delIdx = rPrChildren.findIndex((c) => getTagName(c) === "w:del");
  if (delIdx !== -1 && matches(rPrChildren[delIdx]!, ctx)) {
    ctx.count++;
    rPrChildren.splice(delIdx, 1);
    if (ctx.mode === "accept") removeParagraph = true;
  }

  // Tidy empty property containers
  if (rPrChildren.length === 0) {
    const pPrChildren = getChildren(pPr);
    pPrChildren.splice(pPrChildren.indexOf(rPr), 1);
    if (pPrChildren.length === 0) {
      const pChildren = getChildren(pNode);
      pChildren.splice(pChildren.indexOf(pPr), 1);
    }
  }

  if (!removeParagraph) return false;
  return mergeIntoNextParagraph(siblings, index, pNode);
}

/**
 * Merge a paragraph's content into the following paragraph and remove it.
 * If there is no following paragraph (end of body / cell, or a table comes
 * next) the paragraph is kept, as Word does.
 */
function mergeIntoNextParagraph(
  siblings: FXPNode[],
  index: number,
  pNode: FXPNode,
): boolean {
  // A section-break paragraph must survive or the section is lost
  const pPr = findChild(pNode, "w:pPr");
  if (pPr && findChild(pPr, "w:sectPr")) return false;

  let j = index + 1;
  while (j < siblings.length) {
    const t = getTagName(siblings[j]!);
    if (t === "w:p") break;
    // Whitespace text nodes and range markers do not separate paragraphs
    if (t === undefined || TRANSPARENT_SIBLINGS.has(t)) {
      j++;
      continue;
    }
    return false;
  }
  if (j >= siblings.length) return false;

  const next = siblings[j]!;
  const nextChildren = getChildren(next);
  const nextPPr = findChild(next, "w:pPr");
  const insertAt = nextPPr ? nextChildren.indexOf(nextPPr) + 1 : 0;

  const content = getChildren(pNode).filter((c) => getTagName(c) !== "w:pPr");
  nextChildren.splice(insertAt, 0, ...content);
  siblings.splice(index, 1);
  return true;
}

// =============================================================================
// Table rows and cells
// =============================================================================

/** Returns true if the row was removed. */
function handleRowMarkers(
  siblings: FXPNode[],
  index: number,
  trNode: FXPNode,
  ctx: WalkContext,
): boolean {
  const trPr = findChild(trNode, "w:trPr");
  if (!trPr) return false;
  const trPrChildren = getChildren(trPr);

  const insIdx = trPrChildren.findIndex((c) => getTagName(c) === "w:ins");
  if (insIdx !== -1 && matches(trPrChildren[insIdx]!, ctx)) {
    ctx.count++;
    if (ctx.mode === "reject") {
      siblings.splice(index, 1);
      return true;
    }
    trPrChildren.splice(insIdx, 1);
  }

  const delIdx = trPrChildren.findIndex((c) => getTagName(c) === "w:del");
  if (delIdx !== -1 && matches(trPrChildren[delIdx]!, ctx)) {
    ctx.count++;
    if (ctx.mode === "accept") {
      siblings.splice(index, 1);
      return true;
    }
    trPrChildren.splice(delIdx, 1);
  }

  return false;
}

/** Returns true if the cell was removed. */
function handleCellMarkers(
  siblings: FXPNode[],
  index: number,
  tcNode: FXPNode,
  ctx: WalkContext,
): boolean {
  const tcPr = findChild(tcNode, "w:tcPr");
  if (!tcPr) return false;
  const tcPrChildren = getChildren(tcPr);

  const cellIns = tcPrChildren.findIndex((c) => getTagName(c) === "w:cellIns");
  if (cellIns !== -1 && matches(tcPrChildren[cellIns]!, ctx)) {
    ctx.count++;
    if (ctx.mode === "reject") {
      siblings.splice(index, 1);
      return true;
    }
    tcPrChildren.splice(cellIns, 1);
  }

  const cellDel = tcPrChildren.findIndex((c) => getTagName(c) === "w:cellDel");
  if (cellDel !== -1 && matches(tcPrChildren[cellDel]!, ctx)) {
    ctx.count++;
    if (ctx.mode === "accept") {
      siblings.splice(index, 1);
      return true;
    }
    tcPrChildren.splice(cellDel, 1);
  }

  const cellMerge = tcPrChildren.findIndex((c) => getTagName(c) === "w:cellMerge");
  if (cellMerge !== -1 && matches(tcPrChildren[cellMerge]!, ctx)) {
    ctx.count++;
    tcPrChildren.splice(cellMerge, 1);
  }

  return false;
}
