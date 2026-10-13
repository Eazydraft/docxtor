/**
 * Whole-paragraph operations: insert_paragraph and delete_paragraph.
 *
 * insert_paragraph creates one or more sibling `w:p` nodes before or after a
 * target paragraph (in the body, a table cell, or a header/footer), optionally
 * inheriting the target's paragraph properties and run formatting. With
 * revision tracking, runs are wrapped in `w:ins` and the paragraph mark itself
 * carries `w:pPr/w:rPr/w:ins` so that accepting the change does not leave a
 * stray empty paragraph behind.
 *
 * delete_paragraph removes a paragraph outright, or with revision tracking
 * marks every run deleted and puts `w:del` on the paragraph mark — Word's
 * convention for a deleted paragraph.
 */

import type {
  NormalizedDocument,
  NormalizedParagraph,
  InsertParagraphPatch,
  DeleteParagraphPatch,
  FXPNode,
} from "../types.js";
import {
  cloneNode,
  createElement,
  findChild,
  getAttr,
  getChildren,
  getTagName,
} from "../xml/helpers.js";
import { buildFieldRuns } from "./field-codes.js";
import { normalizeParagraph } from "../normalizer/paragraph.js";
import { enrichListType } from "../normalizer/document.js";
import {
  createInsElement,
  createDelElement,
  convertRunToDeleted,
  nextRevisionId,
  type RevisionContext,
} from "./revision.js";
import {
  findParentChildren,
  findNormalizedContainer,
  deriveNewParagraphId,
  reNormalizeParagraph,
} from "./container-utils.js";
import { stampNewParagraphs, findPartRootForId } from "../normalizer/para-id.js";

// =============================================================================
// Style inheritance
// =============================================================================

/** w:pPr children that must not be copied onto a new paragraph. */
const PPR_EXCLUDE = new Set(["w:sectPr", "w:pPrChange", "w:rPr"]);
/** Paragraph-mark w:rPr children that are revision state, not formatting. */
const PARA_MARK_REVISION = new Set(["w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange"]);

/**
 * Clone the target's w:pPr for a new sibling paragraph: paragraph style,
 * numbering, indentation, spacing, and the paragraph-mark run properties —
 * minus section breaks and revision bookkeeping.
 */
function buildInheritedPPr(targetP: FXPNode): FXPNode | undefined {
  const pPr = findChild(targetP, "w:pPr");
  if (!pPr) return undefined;

  const children: FXPNode[] = [];
  let markRPr: FXPNode | undefined;

  for (const child of getChildren(pPr)) {
    const tag = getTagName(child);
    if (!tag) continue;
    if (tag === "w:rPr") {
      const kept = getChildren(child)
        .filter((c) => {
          const t = getTagName(c);
          return t !== undefined && !PARA_MARK_REVISION.has(t);
        })
        .map(cloneNode);
      if (kept.length > 0) markRPr = createElement("w:rPr", kept);
      continue;
    }
    if (PPR_EXCLUDE.has(tag)) continue;
    children.push(cloneNode(child));
  }

  // w:rPr is the last element of CT_PPr apart from sectPr/pPrChange
  if (markRPr) children.push(markRPr);
  if (children.length === 0) return undefined;
  return createElement("w:pPr", children);
}

/**
 * Run formatting for the new text: the target's last run's w:rPr (what Word
 * would give you if you pressed Enter at the end of the paragraph and typed).
 */
function buildInheritedRunRPr(target: NormalizedParagraph): FXPNode | undefined {
  const lastRun = target.runs[target.runs.length - 1];
  if (!lastRun) return undefined;
  const rPr = findChild(lastRun.xmlRef, "w:rPr");
  if (!rPr) return undefined;
  const kept = getChildren(rPr)
    .filter((c) => getTagName(c) !== "w:rPrChange")
    .map(cloneNode);
  if (kept.length === 0) return undefined;
  return createElement("w:rPr", kept);
}

// =============================================================================
// Paragraph mark revisions
// =============================================================================

/**
 * Add a `w:ins` or `w:del` marker to a paragraph's mark (`w:pPr/w:rPr`).
 * Creates w:pPr / w:rPr as needed, keeping schema order.
 */
export function addParagraphMarkRevision(
  pNode: FXPNode,
  kind: "ins" | "del",
  revCtx: RevisionContext,
): void {
  let pPr = findChild(pNode, "w:pPr");
  if (!pPr) {
    pPr = createElement("w:pPr", []);
    getChildren(pNode).unshift(pPr);
  }
  const pPrChildren = getChildren(pPr);

  let rPr = findChild(pPr, "w:rPr");
  if (!rPr) {
    rPr = createElement("w:rPr", []);
    const tailIdx = pPrChildren.findIndex((c) => {
      const t = getTagName(c);
      return t === "w:sectPr" || t === "w:pPrChange";
    });
    if (tailIdx === -1) pPrChildren.push(rPr);
    else pPrChildren.splice(tailIdx, 0, rPr);
  }

  const marker = createElement(`w:${kind}`, [], {
    "w:id": nextRevisionId(revCtx),
    "w:author": revCtx.author,
    "w:date": revCtx.date,
  });
  // w:ins / w:del are the first children of CT_ParaRPr
  getChildren(rPr).unshift(marker);
}

/** Find a paragraph-mark revision marker (`w:pPr/w:rPr/w:ins|w:del`). */
function findParagraphMarkMarker(pNode: FXPNode, kind: "ins" | "del"): FXPNode | undefined {
  const pPr = findChild(pNode, "w:pPr");
  if (!pPr) return undefined;
  const rPr = findChild(pPr, "w:rPr");
  if (!rPr) return undefined;
  return findChild(rPr, `w:${kind}`);
}

// =============================================================================
// insert_paragraph
// =============================================================================

/**
 * Build the new w:p nodes for an insert_paragraph patch.
 */
function buildParagraphNodes(
  target: NormalizedParagraph,
  patch: InsertParagraphPatch,
  revCtx?: RevisionContext,
): FXPNode[] {
  const inherit = patch.inheritStyle ?? true;
  const pPrTemplate = inherit ? buildInheritedPPr(target.xmlRef) : undefined;
  const runRPr = inherit ? buildInheritedRunRPr(target) : undefined;

  const lines = patch.text.split("\n");
  const nodes: FXPNode[] = [];

  for (const line of lines) {
    const children: FXPNode[] = [];
    if (pPrTemplate) children.push(cloneNode(pPrTemplate));

    const runs = buildFieldRuns(line, runRPr);
    if (revCtx) {
      if (runs.length > 0) children.push(createInsElement(revCtx, runs));
    } else {
      children.push(...runs);
    }

    const pNode: FXPNode = { "w:p": children };
    if (revCtx) addParagraphMarkRevision(pNode, "ins", revCtx);
    nodes.push(pNode);
  }

  return nodes;
}

/**
 * Execute an insert_paragraph patch.
 *
 * The target is resolved by the caller (positional or `para:` id); the new
 * paragraphs are inserted as siblings in the same container — body, table
 * cell, or header/footer — and added to the normalized view with the next
 * free positional ids so they are addressable later in the same batch.
 *
 * @returns the normalized paragraphs that were created
 */
export function executeInsertParagraph(
  doc: NormalizedDocument,
  target: NormalizedParagraph,
  patch: InsertParagraphPatch,
  revCtx?: RevisionContext,
): NormalizedParagraph[] {
  const parentChildren = findParentChildren(doc, target.xmlRef);
  if (!parentChildren) {
    throw new Error(`Cannot find parent container for paragraph "${target.id}"`);
  }
  const targetIdx = parentChildren.indexOf(target.xmlRef);
  if (targetIdx === -1) {
    throw new Error(`Paragraph "${target.id}" not found in its parent`);
  }

  const nodes = buildParagraphNodes(target, patch, revCtx);
  stampNewParagraphs(doc, nodes, findPartRootForId(doc, target.id));

  // XML insertion
  const xmlInsertAt = patch.position === "after" ? targetIdx + 1 : targetIdx;
  parentChildren.splice(xmlInsertAt, 0, ...nodes);

  // Normalized insertion
  const container = findNormalizedContainer(doc, target.id);
  if (!container) {
    throw new Error(`Paragraph "${target.id}" not found in normalized document`);
  }
  const normInsertAt = patch.position === "after" ? container.index + 1 : container.index;

  const created: NormalizedParagraph[] = [];
  for (let i = 0; i < nodes.length; i++) {
    const id = deriveNewParagraphId(target.id, container.array);
    const normalized = normalizeParagraph(nodes[i]!, id);
    if (doc.xmlParts.numbering) enrichListType(normalized, doc.xmlParts.numbering);
    container.array.splice(normInsertAt + i, 0, normalized);
    created.push(normalized);
  }

  return created;
}

// =============================================================================
// delete_paragraph
// =============================================================================

/** Inline containers whose runs should be marked deleted recursively. */
const INLINE_CONTAINERS = new Set([
  "w:hyperlink",
  "w:smartTag",
  "w:customXml",
  "w:dir",
  "w:bdo",
  "w:moveTo",
  "w:sdtContent",
  "w:fldSimple",
]);

/**
 * Mark every run among `children` as deleted, in place.
 *
 * - Consecutive `w:r` siblings are grouped into one `w:del`.
 * - `w:ins` by the same author is removed outright (it was never part of the
 *   accepted document); `w:ins` by another author gets `w:del` nested inside.
 * - Existing `w:del` / `w:moveFrom` content is left alone.
 * - Inline containers (hyperlinks, smart tags, content controls) are recursed.
 */
export function markRunsDeleted(children: FXPNode[], revCtx: RevisionContext): void {
  let i = 0;
  while (i < children.length) {
    const child = children[i]!;
    const tag = getTagName(child);

    if (tag === "w:r") {
      // Gather a group of consecutive runs
      let j = i;
      while (j < children.length && getTagName(children[j]!) === "w:r") j++;
      const group = children.slice(i, j);
      for (const run of group) convertRunToDeleted(run);
      children.splice(i, j - i, createDelElement(revCtx, group));
      i++;
      continue;
    }

    if (tag === "w:ins") {
      if (getAttr(child, "w:author") === revCtx.author) {
        children.splice(i, 1);
        continue;
      }
      markRunsDeleted(getChildren(child), revCtx);
      i++;
      continue;
    }

    if (tag === "w:sdt") {
      const content = findChild(child, "w:sdtContent");
      if (content) markRunsDeleted(getChildren(content), revCtx);
      i++;
      continue;
    }

    if (tag && INLINE_CONTAINERS.has(tag)) {
      markRunsDeleted(getChildren(child), revCtx);
    }
    // w:del, w:moveFrom, w:pPr, bookmarks, proofErr, comment ranges: untouched
    i++;
  }
}

/**
 * True if every run in the paragraph lives inside a `w:ins` by `author` and
 * the paragraph mark is also an insertion by `author` — i.e. the paragraph is
 * wholly our own unaccepted insertion and can simply be removed.
 */
function isWhollyOwnInsertion(pNode: FXPNode, author: string): boolean {
  const mark = findParagraphMarkMarker(pNode, "ins");
  if (!mark || getAttr(mark, "w:author") !== author) return false;

  const check = (children: FXPNode[]): boolean => {
    for (const child of children) {
      const tag = getTagName(child);
      if (tag === "w:r") return false;
      if (tag === "w:ins") {
        if (getAttr(child, "w:author") !== author) return false;
        continue;
      }
      if (tag === "w:del" || tag === "w:moveFrom") return false;
      if (tag === "w:sdt") {
        const content = findChild(child, "w:sdtContent");
        if (content && !check(getChildren(content))) return false;
        continue;
      }
      if (tag && INLINE_CONTAINERS.has(tag)) {
        if (!check(getChildren(child))) return false;
      }
    }
    return true;
  };
  return check(getChildren(pNode));
}

/**
 * Physically remove a paragraph from the XML and the normalized view.
 * If it is the only paragraph in a table cell, it is emptied instead
 * (a `w:tc` must contain at least one `w:p`).
 */
function removeParagraph(doc: NormalizedDocument, para: NormalizedParagraph): void {
  const parentChildren = findParentChildren(doc, para.xmlRef);
  if (!parentChildren) {
    throw new Error(`Cannot find parent container for paragraph "${para.id}"`);
  }
  const container = findNormalizedContainer(doc, para.id);
  if (!container) {
    throw new Error(`Paragraph "${para.id}" not found in normalized document`);
  }

  const isBody = container.array === doc.body;
  const isHeaderFooter = [...doc.headers.values(), ...doc.footers.values()]
    .some((hf) => hf.paragraphs === container.array);
  const isCell = !isBody && !isHeaderFooter;
  const paragraphCount = parentChildren.filter((c) => getTagName(c) === "w:p").length;

  if (isCell && paragraphCount === 1) {
    // Keep the cell valid: strip content, keep w:pPr
    const pChildren = getChildren(para.xmlRef);
    const pPr = findChild(para.xmlRef, "w:pPr");
    pChildren.length = 0;
    if (pPr) pChildren.push(pPr);
    reNormalizeParagraph(doc, para);
    return;
  }

  const idx = parentChildren.indexOf(para.xmlRef);
  if (idx !== -1) parentChildren.splice(idx, 1);
  container.array.splice(container.index, 1);
}

/**
 * Execute a delete_paragraph patch.
 *
 * With revision tracking: all runs are wrapped in `w:del` (w:t → w:delText),
 * and the paragraph mark gets `w:pPr/w:rPr/w:del`. Runs already inside a
 * `w:ins` by the same author are removed rather than double-marked; if the
 * whole paragraph is the author's own unaccepted insertion, it is removed.
 *
 * Without tracking: the `w:p` is removed.
 */
export function executeDeleteParagraph(
  doc: NormalizedDocument,
  target: NormalizedParagraph,
  _patch: DeleteParagraphPatch,
  revCtx?: RevisionContext,
): void {
  if (!revCtx) {
    removeParagraph(doc, target);
    return;
  }

  if (isWhollyOwnInsertion(target.xmlRef, revCtx.author)) {
    removeParagraph(doc, target);
    return;
  }

  markRunsDeleted(getChildren(target.xmlRef), revCtx);

  // Don't double-mark a mark that is already deleted
  if (!findParagraphMarkMarker(target.xmlRef, "del")) {
    addParagraphMarkRevision(target.xmlRef, "del", revCtx);
  }

  reNormalizeParagraph(doc, target);
}
