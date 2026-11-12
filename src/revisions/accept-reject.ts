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
 * With an `authors` filter only revisions by those authors are touched; with
 * an `ids` filter only revisions with those `w:id`s.
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
  ids?: Set<string>;
  count: number;
  /**
   * Ids of `w:moveFromRangeStart` / `w:moveToRangeStart` markers that match
   * the author filter, collected per part. Range *end* markers carry only an
   * id, so they are matched through this set.
   */
  moveRangeIds: Set<string>;
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

const MOVE_RANGE_START_TAGS = new Set(["w:moveFromRangeStart", "w:moveToRangeStart"]);
const MOVE_RANGE_END_TAGS = new Set(["w:moveFromRangeEnd", "w:moveToRangeEnd"]);

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
    ids: options?.ids ? new Set(options.ids.map(String)) : undefined,
    count: 0,
    moveRangeIds: new Set(),
  };
}

function walkAllParts(doc: NormalizedDocument, ctx: WalkContext): void {
  const parts: FXPNode[][] = [
    doc.xmlParts.document,
    ...doc.xmlParts.headers.values(),
    ...doc.xmlParts.footers.values(),
  ];
  for (const key of ["footnotes", "endnotes", "comments"] as const) {
    const part = doc.xmlParts[key];
    if (part) parts.push(part);
  }
  for (const part of parts) {
    if (ctx.ids) expandMoveIds(part, ctx.ids);
    ctx.moveRangeIds = collectMatchingMoveRangeIds(part, ctx);
    processChildren(part, undefined, ctx);
  }
}

function matches(node: FXPNode, ctx: WalkContext): boolean {
  if (ctx.ids) {
    const id = getAttr(node, "w:id");
    if (id === undefined || !ctx.ids.has(id)) return false;
  }
  if (!ctx.authors) return true;
  const author = getAttr(node, "w:author");
  return author !== undefined && ctx.authors.has(author);
}

/**
 * A move is several elements with different ids: `w:moveFromRangeStart` / `w:moveToRangeStart`
 * markers (with matching `…RangeEnd` markers) around `w:moveFrom` / `w:moveTo` content. When an
 * `ids` filter names any part of a move, add the rest, including the other side of the move (the
 * "from" and "to" ranges share a `w:name`), so the move is resolved as a unit rather than leaving range
 * markers without content, content without its range, or one side of a move without the other.
 */
function expandMoveIds(nodes: FXPNode[], ids: Set<string>): void {
  const open: { from: string[]; to: string[] } = { from: [], to: [] };
  const groups: string[][] = []; // each: [range id, ...content ids inside it]
  const groupOf = new Map<string, string[]>();
  const byName = new Map<string, string[]>();
  const visit = (list: FXPNode[]): void => {
    for (const node of list) {
      const tag = getTagName(node);
      if (!tag) continue;
      const id = getAttr(node, "w:id");
      if (id !== undefined) {
        if (tag === "w:moveFromRangeStart" || tag === "w:moveToRangeStart") {
          const name = getAttr(node, "w:name");
          const g = (name !== undefined && byName.get(name)) || [];
          if (!g.length) groups.push(g);
          if (name !== undefined) byName.set(name, g);
          g.push(id);
          groupOf.set(id, g);
          (tag === "w:moveFromRangeStart" ? open.from : open.to).push(id);
        } else if (tag === "w:moveFromRangeEnd" || tag === "w:moveToRangeEnd") {
          const stack = tag === "w:moveFromRangeEnd" ? open.from : open.to;
          const at = stack.lastIndexOf(id);
          if (at !== -1) stack.splice(at, 1);
        } else if (tag === "w:moveFrom" || tag === "w:moveTo") {
          for (const rangeId of tag === "w:moveFrom" ? open.from : open.to) groupOf.get(rangeId)!.push(id);
        }
      }
      const children = getChildren(node);
      if (children.length > 0) visit(children);
    }
  };
  visit(nodes);
  for (const g of groups) if (g.some((id) => ids.has(id))) for (const id of g) ids.add(id);
}

/**
 * Pre-pass: ids of move range start markers that match the author filter,
 * so the author-less end markers can be removed together with them.
 */
function collectMatchingMoveRangeIds(nodes: FXPNode[], ctx: WalkContext): Set<string> {
  const ids = new Set<string>();
  const visit = (list: FXPNode[]): void => {
    for (const node of list) {
      const tag = getTagName(node);
      if (!tag) continue;
      if (MOVE_RANGE_START_TAGS.has(tag) && matches(node, ctx)) {
        const id = getAttr(node, "w:id");
        if (id !== undefined) ids.add(id);
      }
      const children = getChildren(node);
      if (children.length > 0) visit(children);
    }
  };
  visit(nodes);
  return ids;
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
    // Start markers carry w:author; end markers only an id, matched through
    // the ids collected from their start markers.
    if (MOVE_RANGE_START_TAGS.has(tag) || MOVE_RANGE_END_TAGS.has(tag)) {
      const id = getAttr(child, "w:id");
      const remove = MOVE_RANGE_START_TAGS.has(tag)
        ? matches(child, ctx)
        : id !== undefined && ctx.moveRangeIds.has(id);
      if (remove) {
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

    // --- Content controls around rows or cells --------------------------
    // Resolving a change can remove the only row or cell a control wraps.
    // An empty control is not valid in a table or a row: it goes too. The
    // control is recognised by what it held, not by its parent, so one
    // nested inside another control is cleaned up as well, and the outer
    // one after it if that leaves it empty in turn.
    if (tag === "w:sdt" && wrapsRowsOrCells(child)) {
      const content = findChild(child, "w:sdtContent");
      processChildren(getChildren(child), tag, ctx);
      const left = content ? getChildren(content).filter((c) => getTagName(c) !== undefined).length : 0;
      if (left === 0) {
        children.splice(i, 1);
        continue;
      }
      i++;
      continue;
    }

    // --- Everything else: recurse ---------------------------------------
    const grandChildren = getChildren(child);
    if (grandChildren.length > 0) processChildren(grandChildren, tag, ctx);
    i++;
  }
}

/**
 * True for a content control that wraps table rows or cells, directly or
 * through further controls.
 */
function wrapsRowsOrCells(sdt: FXPNode): boolean {
  const content = findChild(sdt, "w:sdtContent");
  if (!content) return false;
  return getChildren(content).some((c) => {
    const t = getTagName(c);
    return t === "w:tr" || t === "w:tc" || (t === "w:sdt" && wrapsRowsOrCells(c));
  });
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
 * Returns true if the paragraph was removed: merged into the next one, or,
 * with no next one and nothing left in it, dropped.
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
  let insertionRejected = false;

  const insIdx = rPrChildren.findIndex((c) => getTagName(c) === "w:ins");
  if (insIdx !== -1 && matches(rPrChildren[insIdx]!, ctx)) {
    ctx.count++;
    rPrChildren.splice(insIdx, 1);
    if (ctx.mode === "reject") {
      removeParagraph = true;
      insertionRejected = true;
    }
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
  // The paragraph goes. A content control in it whose content went with the
  // resolved change goes too, as it does when the paragraph is deleted
  // untracked; otherwise it would be merged into the next paragraph as an
  // empty field.
  pruneEmptyControls(getChildren(pNode));
  if (mergeIntoNextParagraph(siblings, index, pNode)) return true;
  // With no next paragraph a deleted mark keeps its paragraph, as in Word,
  // where the last paragraph mark cannot be deleted. A rejected insertion is
  // different: the paragraph was never part of the document.
  return insertionRejected && removeInsertedParagraph(siblings, index, pNode);
}

/**
 * Remove a paragraph whose inserted mark was rejected and that has no next
 * paragraph to merge into: the last paragraph of the body, a cell, a header
 * or a note, or one that a table follows. This is where a rejected
 * `insert_paragraph` after such a paragraph ends up; without this the
 * rejection would leave an empty paragraph the original did not have.
 *
 * Only a paragraph that holds nothing any more is removed, and only when
 * the paragraph before it stays: a cell must end with a paragraph, and two
 * tables with no paragraph between them would become one. A paragraph that
 * still holds something (text another author added, a bookmark, a comment's
 * range) is kept, since there is nowhere to move it.
 */
function removeInsertedParagraph(siblings: FXPNode[], index: number, pNode: FXPNode): boolean {
  const pPr = findChild(pNode, "w:pPr");
  // A section-break paragraph must survive or the section is lost
  if (pPr && findChild(pPr, "w:sectPr")) return false;
  if (getChildren(pNode).some((c) => c !== pPr && getTagName(c) !== undefined)) return false;

  let j = index - 1;
  while (j >= 0) {
    const t = getTagName(siblings[j]!);
    if (t === undefined || TRANSPARENT_SIBLINGS.has(t)) {
      j--;
      continue;
    }
    break;
  }
  if (j < 0 || getTagName(siblings[j]!) !== "w:p") return false;

  siblings.splice(index, 1);
  return true;
}

/**
 * Remove every inline content control under `children` that holds no
 * element any more, innermost first, so a control left holding only empty
 * controls goes as well.
 */
function pruneEmptyControls(children: FXPNode[]): void {
  for (let i = children.length - 1; i >= 0; i--) {
    const child = children[i]!;
    const tag = getTagName(child);
    if (tag === undefined || tag === "w:pPr") continue;
    if (tag === "w:sdt") {
      const content = findChild(child, "w:sdtContent");
      if (content) pruneEmptyControls(getChildren(content));
      if (!content || !getChildren(content).some((c) => getTagName(c) !== undefined)) {
        children.splice(i, 1);
      }
      continue;
    }
    if (tag !== "w:r") pruneEmptyControls(getChildren(child));
  }
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

  const cellMerge = tcPrChildren.find((c) => getTagName(c) === "w:cellMerge");
  if (cellMerge && matches(cellMerge, ctx)) {
    ctx.count++;
    if (ctx.mode === "reject") {
      restoreVerticalMerge(tcPrChildren, cellMerge); // may insert before the marker
    }
    tcPrChildren.splice(tcPrChildren.indexOf(cellMerge), 1);
  }

  return false;
}

/**
 * Rejecting a tracked vertical cell merge: put `w:vMerge` back to the state
 * recorded in `w:cellMerge/@w:vMergeOrig` ("rest" → restart, "cont" →
 * continue, absent → the cell was not merged, so drop `w:vMerge`).
 */
function restoreVerticalMerge(tcPrChildren: FXPNode[], cellMerge: FXPNode): void {
  const orig = getAttr(cellMerge, "w:vMergeOrig");
  const existingIdx = tcPrChildren.findIndex((c) => getTagName(c) === "w:vMerge");

  if (orig === undefined) {
    if (existingIdx !== -1) tcPrChildren.splice(existingIdx, 1);
    return;
  }

  const vMerge: FXPNode = { "w:vMerge": [] };
  if (orig === "rest") setAttr(vMerge, "w:val", "restart");

  if (existingIdx !== -1) {
    tcPrChildren[existingIdx] = vMerge;
  } else {
    // Keep schema order: w:vMerge precedes w:tcBorders and the cell markers
    const beforeIdx = tcPrChildren.findIndex((c) => {
      const t = getTagName(c);
      return t === "w:tcBorders" || t === "w:shd" || t === "w:noWrap" || t === "w:tcMar" ||
        t === "w:textDirection" || t === "w:tcFitText" || t === "w:vAlign" || t === "w:hideMark" ||
        t === "w:cellIns" || t === "w:cellDel" || t === "w:cellMerge" || t === "w:tcPrChange";
    });
    if (beforeIdx === -1) tcPrChildren.push(vMerge);
    else tcPrChildren.splice(beforeIdx, 0, vMerge);
  }
}
