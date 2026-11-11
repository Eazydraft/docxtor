/**
 * Text operations: replace_text, insert_text, delete_text.
 *
 * These operate on a NormalizedParagraph by:
 * 1. Finding the target text in the normalized paragraph text
 * 2. Mapping to run(s) via the text-to-runs resolver
 * 3. Splitting runs if the match doesn't align with run boundaries
 * 4. Mutating the FXP w:t text nodes directly
 *
 * After execution, the FXP object reflects the edit, and the
 * paragraph should be re-normalized to update mappings.
 *
 * When a RevisionContext is provided, operations produce OOXML tracked
 * changes (w:ins, w:del) instead of direct mutations.
 */

import type {
  NormalizedParagraph,
  ReplaceTextPatch,
  InsertTextPatch,
  DeleteTextPatch,
  FXPNode,
  RunMapping,
} from "../types.js";
import { findOccurrence, resolveRunRanges } from "../mapping/text-to-runs.js";
import {
  splitRun,
  splitRunsForMatch,
  splitRunAtOffset,
  findRunParentChildren,
} from "./run-splitter.js";
import { readRunContent, writeRunContent } from "./run-content.js";
import {
  getChildren,
  setTextContent,
  getRunText,
  findTextNode,
  getTagName,
  setAttr,
  escapeXmlText,
  needsSpacePreserve,
  findChild,
  cloneNode,
  createElement,
  getAttr,
} from "../xml/helpers.js";
import type { RevisionContext } from "./revision.js";
import {
  createInsElement,
  createDelElement,
  convertRunToDeleted,
  createRevisionRun,
  nextRevisionId,
} from "./revision.js";
import {
  stripFormattingMarkers,
  parseFormattedText,
  type FormattedSegment,
} from "./markdown-parser.js";

/**
 * Remove a run from whichever parent holds it: the paragraph itself, or a
 * w:hyperlink, w:ins, w:del or content control inside it.
 */
function removeRunFromParagraph(paragraphNode: FXPNode, runNode: FXPNode): void {
  const parent = findRunParentChildren(paragraphNode, runNode);
  if (!parent) return;
  const idx = parent.indexOf(runNode);
  if (idx !== -1) parent.splice(idx, 1);
}

/**
 * Find the container element (w:ins, w:del, w:hyperlink, or the w:sdtContent
 * of a content control) that directly holds a run, or undefined if the run
 * is a direct child of the paragraph.
 */
function findRunContainer(
  paragraphNode: FXPNode,
  runNode: FXPNode,
): FXPNode | undefined {
  const search = (children: FXPNode[]): FXPNode | undefined => {
    for (const child of children) {
      const tag = getTagName(child);
      if (tag === "w:hyperlink" || tag === "w:ins" || tag === "w:del" || tag === "w:sdtContent") {
        if (getChildren(child).includes(runNode)) return child;
        const found = search(getChildren(child));
        if (found) return found;
      } else if (tag === "w:sdt") {
        const found = search(getChildren(child));
        if (found) return found;
      }
    }
    return undefined;
  };
  if (getChildren(paragraphNode).includes(runNode)) return undefined;
  return search(getChildren(paragraphNode));
}

// =============================================================================
// Replace text
// =============================================================================

/**
 * Execute a replace_text patch on a paragraph.
 *
 * Finds the target text, maps to runs, splits if needed, and replaces.
 * When revCtx is provided, wraps old text in w:del and new text in w:ins.
 */
export function executeReplaceText(
  paragraph: NormalizedParagraph,
  patch: ReplaceTextPatch,
  revCtx?: RevisionContext,
  markers: boolean = true,
): void {
  const occurrence = patch.occurrence ?? 1;
  const plainFind = stripFormattingMarkers(patch.find, markers);
  const matchStart = findOccurrence(paragraph.text, plainFind, occurrence);

  if (matchStart === -1) {
    throw new Error(
      `Text "${plainFind}" not found in paragraph "${paragraph.id}" (occurrence ${occurrence})`,
    );
  }

  const matchEnd = matchStart + plainFind.length;
  const ranges = resolveRunRanges(paragraph.runs, matchStart, matchEnd);

  if (ranges.length === 0) {
    throw new Error(
      `No runs found for text range [${matchStart}, ${matchEnd}] in paragraph "${paragraph.id}"`,
    );
  }

  if (revCtx) {
    executeReplaceTextRevision(paragraph, patch, ranges, revCtx, markers);
  } else {
    executeReplaceTextDirect(paragraph, patch, ranges, markers);
  }
}

/**
 * Direct (non-revision) replacement.
 *
 * If the replacement text contains GFM formatting markers (**bold**, _italic_,
 * ~~strike~~), it is parsed into segments and each segment becomes a separate
 * run with the appropriate w:rPr formatting, inheriting any existing formatting
 * from the matched run.
 */
function executeReplaceTextDirect(
  paragraph: NormalizedParagraph,
  patch: ReplaceTextPatch,
  ranges: ReturnType<typeof resolveRunRanges>,
  markers: boolean,
): void {
  const segments = parseFormattedText(patch.replace, markers);
  const hasFormatting = segments.some((s) => s.bold || s.italic || s.strikethrough || s.underline);
  const plainText = stripFormattingMarkers(patch.replace, markers);

  if (ranges.length === 1) {
    const range = ranges[0]!;
    const splitResult = splitRun(paragraph.xmlRef, range);

    if (!hasFormatting) {
      // Simple plain text replacement
      setRunText(splitResult.matched, plainText);
    } else {
      // Replace the matched run with formatted runs
      replaceRunWithFormattedRuns(paragraph.xmlRef, splitResult.matched, segments);
    }
  } else {
    const matchedRuns = splitRunsForMatch(paragraph.xmlRef, ranges);

    if (matchedRuns.length === 0) {
      throw new Error(
        `Failed to isolate matched runs in paragraph "${paragraph.id}"`,
      );
    }

    if (!hasFormatting) {
      setRunText(matchedRuns[0]!, plainText);
    } else {
      replaceRunWithFormattedRuns(paragraph.xmlRef, matchedRuns[0]!, segments);
    }

    // Remove the text of the other matched runs. A run that also holds a
    // zero-width child (a page break inside the match) keeps that child.
    for (let i = 1; i < matchedRuns.length; i++) {
      const content = readRunContent(matchedRuns[i]!);
      if (content.anchored.length > 0) {
        writeRunContent(matchedRuns[i]!, { ...content, text: "" });
      } else {
        removeRunFromParagraph(paragraph.xmlRef, matchedRuns[i]!);
      }
    }
  }
}

/**
 * Revision-tracked replacement:
 * 1. Split runs to isolate matched text
 * 2. Wrap matched runs in w:del (converting w:t → w:delText)
 * 3. Create new run(s) with replacement text wrapped in w:ins
 * 4. Insert w:ins immediately after w:del
 */
function executeReplaceTextRevision(
  paragraph: NormalizedParagraph,
  patch: ReplaceTextPatch,
  ranges: ReturnType<typeof resolveRunRanges>,
  revCtx: RevisionContext,
  markers: boolean,
): void {
  // Split runs to isolate the matched text
  let matchedRuns: FXPNode[];
  if (ranges.length === 1) {
    const range = ranges[0]!;
    const splitResult = splitRun(paragraph.xmlRef, range);
    matchedRuns = [splitResult.matched];
  } else {
    matchedRuns = splitRunsForMatch(paragraph.xmlRef, ranges);
  }

  if (matchedRuns.length === 0) {
    throw new Error(
      `Failed to isolate matched runs in paragraph "${paragraph.id}"`,
    );
  }

  // Get formatting from the first matched run for the replacement
  const rPrNode = findChild(matchedRuns[0]!, "w:rPr");

  // Create the replacement run(s); wrapped in w:ins below
  const segments = parseFormattedText(patch.replace, markers);
  const hasFormatting = segments.some((s) => s.bold || s.italic || s.strikethrough || s.underline);
  const plainText = stripFormattingMarkers(patch.replace, markers);

  let newRuns: FXPNode[] = [];
  if (segments.length > 0 && plainText.length > 0) {
    newRuns = hasFormatting
      ? createFormattedRuns(segments, rPrNode)
      : [createRevisionRun(plainText, rPrNode)];
  }
  // Mark the matched runs deleted (or drop them, where they were an
  // unaccepted insertion) and put the replacement where the match began.
  const place = detachMatchedRuns(paragraph.xmlRef, matchedRuns, revCtx);
  if (newRuns.length === 0) return;

  if (place.container) {
    // The match began inside an unaccepted insertion: the replacement must
    // land exactly where it was, not after the surrounding w:ins.
    insertIntoInsertion(paragraph.xmlRef, place.container, place.index, newRuns, revCtx);
  } else {
    place.parentChildren.splice(place.index, 0, createInsElement(revCtx, newRuns));
  }
}

/**
 * Where the content that replaces a match should go, after the match has
 * been detached: a position in a parent's children, or, when the match
 * began inside an unaccepted insertion that still holds other runs, a
 * position inside that `w:ins` container.
 */
interface DetachPlacement {
  parentChildren: FXPNode[];
  index: number;
  container?: FXPNode;
}

/**
 * Detach matched runs for a tracked edit, run by run, in whichever parent
 * each one lives.
 *
 * - A run inside an unaccepted `w:ins` is removed outright: it was never
 *   part of the accepted document, so there is nothing to mark deleted. A
 *   container left empty is removed too.
 * - Any other run is replaced by a `w:del` holding a copy of it; consecutive
 *   runs in the same parent share one `w:del`.
 *
 * A match may span both kinds — text that starts inside another author's
 * insertion and continues in plain text — which is why the parent of the
 * first run cannot stand in for the others.
 */
function detachMatchedRuns(
  paragraphNode: FXPNode,
  matchedRuns: FXPNode[],
  revCtx: RevisionContext,
): DetachPlacement {
  let placement: DetachPlacement | undefined;
  const groups: { parent: FXPNode[]; index: number; runs: FXPNode[] }[] = [];

  for (const run of matchedRuns) {
    const parent = findRunParentChildren(paragraphNode, run);
    if (!parent) throw new Error("Matched run not found in paragraph");
    const container = findRunContainer(paragraphNode, run);
    const insideIns = container != null && getTagName(container) === "w:ins";
    const index = parent.indexOf(run);
    parent.splice(index, 1);

    if (insideIns) {
      let containerKept = true;
      let containerParent: FXPNode[] | undefined;
      let containerIndex = -1;
      if (parent.length === 0) {
        containerParent = findRunParentChildren(paragraphNode, container);
        if (containerParent) {
          containerIndex = containerParent.indexOf(container);
          containerParent.splice(containerIndex, 1);
          containerKept = false;
        }
      }
      if (!placement) {
        placement = containerKept
          ? { parentChildren: parent, index, container }
          : { parentChildren: containerParent!, index: containerIndex };
      }
      continue;
    }

    const last = groups[groups.length - 1];
    if (last && last.parent === parent && last.index === index) {
      last.runs.push(run);
    } else {
      groups.push({ parent, index, runs: [run] });
    }
  }

  // Insert the w:del elements last group first, so earlier indices in the
  // same parent stay valid.
  for (let g = groups.length - 1; g >= 0; g--) {
    const group = groups[g]!;
    const deletedRuns = group.runs.map((run) => {
      const cloned = cloneNode(run);
      convertRunToDeleted(cloned);
      return cloned;
    });
    group.parent.splice(group.index, 0, createDelElement(revCtx, deletedRuns));
  }

  if (!placement) {
    // The match began in plain text: the replacement follows its w:del.
    const first = groups[0]!;
    placement = { parentChildren: first.parent, index: first.index + 1 };
  }
  return placement;
}

/**
 * Put `newRuns` at `index` among the children of an existing `w:ins`.
 *
 * - Same author: the runs join that insertion in place.
 * - Another author: the insertion is split around `index`; the new runs go
 *   between the halves in their own `w:ins`, so attribution and order are
 *   both kept. An empty half is dropped (and the whole container, if it
 *   ends up empty).
 */
function insertIntoInsertion(
  paragraphNode: FXPNode,
  container: FXPNode,
  index: number,
  newRuns: FXPNode[],
  revCtx: RevisionContext,
): void {
  const containerChildren = getChildren(container);
  const containerParent = findRunParentChildren(paragraphNode, container);
  if (!containerParent) {
    throw new Error("Insertion container not found in paragraph");
  }
  const containerIdx = containerParent.indexOf(container);

  if (newRuns.length === 0) {
    if (containerChildren.length === 0) containerParent.splice(containerIdx, 1);
    return;
  }

  if (getAttr(container, "w:author") === revCtx.author) {
    containerChildren.splice(index, 0, ...newRuns);
    return;
  }

  const tail = containerChildren.splice(index);
  const replacement: FXPNode[] = [];
  if (containerChildren.length > 0) replacement.push(container);
  replacement.push(createInsElement(revCtx, newRuns));
  if (tail.length > 0) {
    const second = createElement("w:ins", tail, {
      "w:id": nextRevisionId(revCtx),
      "w:author": getAttr(container, "w:author") ?? "",
      "w:date": getAttr(container, "w:date") ?? "",
    });
    replacement.push(second);
  }
  containerParent.splice(containerIdx, 1, ...replacement);
}

// =============================================================================
// Insert text
// =============================================================================

/**
 * Execute an insert_text patch on a paragraph.
 *
 * Strategy: find the insertion offset in the normalized text, then find
 * which run contains that offset. Insert the text directly into that run's
 * w:t node at the correct position.
 *
 * With revision tracking, the inserted text goes into a new run wrapped
 * in w:ins, and the target run is split if needed.
 */
export function executeInsertText(
  paragraph: NormalizedParagraph,
  patch: InsertTextPatch,
  revCtx?: RevisionContext,
  markers: boolean = true,
): void {
  const occurrence = patch.occurrence ?? 1;
  const plainAnchor = stripFormattingMarkers(patch.anchor, markers);
  const anchorStart = findOccurrence(paragraph.text, plainAnchor, occurrence);

  if (anchorStart === -1) {
    throw new Error(
      `Anchor text "${plainAnchor}" not found in paragraph "${paragraph.id}" (occurrence ${occurrence})`,
    );
  }

  // Determine the character offset where we want to insert
  const insertOffset =
    patch.position === "before"
      ? anchorStart
      : anchorStart + plainAnchor.length;

  if (revCtx) {
    executeInsertTextRevision(paragraph, patch, insertOffset, revCtx, markers);
  } else {
    executeInsertTextDirect(paragraph, patch, insertOffset, markers);
  }
}

/**
 * The run an insertion at `offset` goes into or next to: the first run whose
 * range reaches the offset, else the last run. Text inserted at the boundary
 * of two runs therefore joins the earlier one.
 *
 * The exception is a boundary where a content control starts or ends. There
 * the text must land on the anchor's side: inserted before an anchor, it
 * joins the run the anchor starts in, so text put before a word that follows
 * a control stays outside the control (and text put before a control's first
 * word goes inside it).
 */
export function findInsertTargetRun(
  paragraph: NormalizedParagraph,
  offset: number,
  position: "before" | "after" = "after",
): RunMapping | undefined {
  const runs = paragraph.runs;
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i]!;
    if (offset >= run.startOffset && offset <= run.endOffset) {
      const next = runs[i + 1];
      if (
        position === "before" &&
        offset === run.endOffset &&
        next !== undefined &&
        next.startOffset === offset &&
        next.control !== run.control
      ) {
        return next;
      }
      return run;
    }
  }
  return runs[runs.length - 1];
}

/**
 * Direct (non-revision) insertion — modifies the run's text in-place,
 * or creates multiple formatted runs if the text contains GFM markers.
 */
function executeInsertTextDirect(
  paragraph: NormalizedParagraph,
  patch: InsertTextPatch,
  insertOffset: number,
  markers: boolean,
): void {
  const targetRun = findInsertTargetRun(paragraph, insertOffset, patch.position);
  if (!targetRun) {
    throw new Error(
      `No runs found in paragraph "${paragraph.id}" for insertion`,
    );
  }

  const segments = parseFormattedText(patch.text, markers);
  const hasFormatting = segments.some((s) => s.bold || s.italic || s.strikethrough || s.underline);

  if (!hasFormatting) {
    // Simple plain text insertion
    const plainText = stripFormattingMarkers(patch.text, markers);
    const localInsertOffset = insertOffset - targetRun.startOffset;
    const currentText = getRunText(targetRun.xmlRef);
    const newText =
      currentText.slice(0, localInsertOffset) +
      plainText +
      currentText.slice(localInsertOffset);
    // A zero-width child at the insert point (a page break right after the
    // anchor) stays next to the anchor: before the new text when inserting
    // before the anchor, after it when inserting after the anchor.
    const side = sideOfChildrenAtInsert(patch);
    setRunText(targetRun.xmlRef, newText, (offset) =>
      offset < localInsertOffset || (offset === localInsertOffset && side === "before")
        ? offset
        : offset + plainText.length,
    );
  } else {
    // Formatted insertion: split the target run at the insert point and
    // insert formatted runs between the halves
    const localInsertOffset = insertOffset - targetRun.startOffset;
    const rPrNode = findChild(targetRun.xmlRef, "w:rPr");
    const formattedRuns = createFormattedRuns(segments, rPrNode);
    const split = splitRunAtOffset(
      paragraph.xmlRef,
      targetRun.xmlRef,
      localInsertOffset,
      sideOfChildrenAtInsert(patch),
    );
    split.parentChildren.splice(split.index, 0, ...formattedRuns);
  }
}

/**
 * Revision-tracked insertion:
 * 1. Find the target run containing the insertion offset
 * 2. Split the run at the insertion point if needed
 * 3. Create a new w:r wrapped in w:ins
 * 4. Insert between the split halves (or at boundary)
 */
function executeInsertTextRevision(
  paragraph: NormalizedParagraph,
  patch: InsertTextPatch,
  insertOffset: number,
  revCtx: RevisionContext,
  markers: boolean,
): void {
  const targetRun = findInsertTargetRun(paragraph, insertOffset, patch.position);
  if (!targetRun) {
    throw new Error(
      `No runs found in paragraph "${paragraph.id}" for insertion`,
    );
  }

  const localInsertOffset = insertOffset - targetRun.startOffset;

  // Get formatting from the target run to clone onto the new run
  const rPrNode = findChild(targetRun.xmlRef, "w:rPr");

  // Create the new run(s) with inserted text wrapped in w:ins
  const segments = parseFormattedText(patch.text, markers);
  const hasFormatting = segments.some((s) => s.bold || s.italic || s.strikethrough || s.underline);

  let insRuns: FXPNode[];
  if (hasFormatting) {
    insRuns = createFormattedRuns(segments, rPrNode);
  } else {
    const plainText = stripFormattingMarkers(patch.text, markers);
    insRuns = [createRevisionRun(plainText, rPrNode)];
  }
  const insElement = createInsElement(revCtx, insRuns);

  // Find the parent children array containing the target run
  const parentChildren = findRunParentChildren(paragraph.xmlRef, targetRun.xmlRef);
  if (!parentChildren) {
    throw new Error("Target run not found in paragraph");
  }

  const runIndex = parentChildren.indexOf(targetRun.xmlRef);

  // Determine if we need to split the run
  const atStart = localInsertOffset === 0;
  const atEnd = localInsertOffset === targetRun.text.length;

  if (atStart) {
    // Insert before the target run (no splitting needed)
    parentChildren.splice(runIndex, 0, insElement);
  } else if (atEnd) {
    // Insert after the target run (no splitting needed)
    parentChildren.splice(runIndex + 1, 0, insElement);
  } else {
    // Split the run at the insertion point; zero-width children keep their side
    const split = splitRunAtOffset(
      paragraph.xmlRef,
      targetRun.xmlRef,
      localInsertOffset,
      sideOfChildrenAtInsert(patch),
    );
    split.parentChildren.splice(split.index, 0, insElement);
  }
}

/**
 * Which side of inserted text a zero-width child at the insert point stays
 * on. Inserting after an anchor keeps a break that follows the anchor after
 * the new text; inserting before an anchor keeps a break that precedes the
 * anchor before it.
 */
function sideOfChildrenAtInsert(patch: InsertTextPatch): "before" | "after" {
  return patch.position === "after" ? "after" : "before";
}

/**
 * Create a w:r run node with text and optional formatting (for splitting).
 *
 * Handles \n in text by emitting w:br elements (matching the behavior
 * of getRunText which converts w:br → \n). Without this, splitting a
 * run containing line breaks would produce literal \n in w:t instead
 * of proper w:br elements.
 */
function createRunWithTextAndFormatting(
  text: string,
  rPrNode: FXPNode | undefined,
): FXPNode {
  const children: FXPNode[] = [];
  if (rPrNode) {
    children.push(cloneNode(rPrNode));
  }

  if (text.includes("\n")) {
    // Multi-line: interleave w:t and w:br
    const segments = text.split("\n");
    for (let i = 0; i < segments.length; i++) {
      const seg = segments[i]!;
      if (seg.length > 0 || i === 0) {
        const textNode = { "#text": escapeXmlText(seg) } as unknown as FXPNode;
        const wtNode: FXPNode = { "w:t": [textNode] };
        if (needsSpacePreserve(seg)) {
          setAttr(wtNode, "xml:space", "preserve");
        }
        children.push(wtNode);
      }
      if (i < segments.length - 1) {
        children.push({ "w:br": [] as FXPNode[] });
      }
    }
  } else {
    // Single line
    const textNode = { "#text": escapeXmlText(text) } as unknown as FXPNode;
    const wtNode: FXPNode = { "w:t": [textNode] };
    if (needsSpacePreserve(text)) {
      setAttr(wtNode, "xml:space", "preserve");
    }
    children.push(wtNode);
  }

  return { "w:r": children };
}

// =============================================================================
// Delete text
// =============================================================================

/**
 * Execute a delete_text patch on a paragraph.
 *
 * Without revisions: equivalent to replace_text with an empty replacement.
 * With revisions: wraps matched runs in w:del (text stays in XML but is marked as deleted).
 */
export function executeDeleteText(
  paragraph: NormalizedParagraph,
  patch: DeleteTextPatch,
  revCtx?: RevisionContext,
  markers: boolean = true,
): void {
  if (revCtx) {
    executeDeleteTextRevision(paragraph, patch, revCtx, markers);
  } else {
    executeReplaceText(
      paragraph,
      {
        type: "replace_text",
        target: patch.target,
        find: patch.find,
        replace: "",
        occurrence: patch.occurrence,
      },
      undefined,
      markers,
    );
  }
}

/**
 * Revision-tracked deletion:
 * 1. Split runs to isolate the matched text
 * 2. Wrap matched runs in w:del, converting w:t → w:delText
 * 3. Text stays in XML but is marked as deleted
 */
function executeDeleteTextRevision(
  paragraph: NormalizedParagraph,
  patch: DeleteTextPatch,
  revCtx: RevisionContext,
  markers: boolean,
): void {
  const occurrence = patch.occurrence ?? 1;
  const plainFind = stripFormattingMarkers(patch.find, markers);
  const matchStart = findOccurrence(paragraph.text, plainFind, occurrence);

  if (matchStart === -1) {
    throw new Error(
      `Text "${plainFind}" not found in paragraph "${paragraph.id}" (occurrence ${occurrence})`,
    );
  }

  const matchEnd = matchStart + plainFind.length;
  const ranges = resolveRunRanges(paragraph.runs, matchStart, matchEnd);

  if (ranges.length === 0) {
    throw new Error(
      `No runs found for text range [${matchStart}, ${matchEnd}] in paragraph "${paragraph.id}"`,
    );
  }

  // Split runs to isolate the matched text
  let matchedRuns: FXPNode[];
  if (ranges.length === 1) {
    const range = ranges[0]!;
    const splitResult = splitRun(paragraph.xmlRef, range);
    matchedRuns = [splitResult.matched];
  } else {
    matchedRuns = splitRunsForMatch(paragraph.xmlRef, ranges);
  }

  if (matchedRuns.length === 0) {
    throw new Error(
      `Failed to isolate matched runs in paragraph "${paragraph.id}"`,
    );
  }

  // Mark each matched run deleted in its own parent; runs that were an
  // unaccepted insertion are simply removed.
  detachMatchedRuns(paragraph.xmlRef, matchedRuns, revCtx);
}

// =============================================================================
// Formatted run creation from GFM markers
// =============================================================================

/**
 * Build an array of w:r FXP nodes from parsed formatted segments.
 *
 * Each segment becomes a run. Formatting flags from the segment (bold, italic,
 * strikethrough) are layered on top of any existing rPr inherited from the
 * base run.
 */
export function createFormattedRuns(
  segments: FormattedSegment[],
  baseRPr: FXPNode | undefined,
): FXPNode[] {
  const runs: FXPNode[] = [];

  for (const seg of segments) {
    if (seg.text.length === 0) continue;

    const run = createRunWithTextAndFormatting(seg.text, baseRPr);

    // Layer GFM formatting on top
    if (seg.bold || seg.italic || seg.strikethrough || seg.underline) {
      let rPr = findChild(run, "w:rPr");
      if (!rPr) {
        rPr = createElement("w:rPr", []);
        const children = getChildren(run);
        children.unshift(rPr);
      }
      const rPrChildren = getChildren(rPr);

      if (seg.bold && !findChild(rPr, "w:b")) {
        rPrChildren.push(createElement("w:b", []));
      }
      if (seg.italic && !findChild(rPr, "w:i")) {
        rPrChildren.push(createElement("w:i", []));
      }
      if (seg.strikethrough && !findChild(rPr, "w:strike")) {
        rPrChildren.push(createElement("w:strike", []));
      }
      if (seg.underline && !findChild(rPr, "w:u")) {
        rPrChildren.push(createElement("w:u", [], { "w:val": "single" }));
      }
    }

    runs.push(run);
  }

  return runs;
}

/**
 * Replace a single run node in a paragraph with multiple formatted runs.
 *
 * Finds the run in the paragraph's children (or nested containers) and
 * splices in the new runs.
 */
function replaceRunWithFormattedRuns(
  paragraphNode: FXPNode,
  runNode: FXPNode,
  segments: FormattedSegment[],
): void {
  const content = readRunContent(runNode);
  const newRuns = createFormattedRuns(segments, content.rPr);

  if (newRuns.length === 0) {
    // All segments were empty — just clear the run text
    setRunText(runNode, "");
    return;
  }

  // Zero-width children before the old text go in front of the first new
  // run's text; the rest follow the last run, so a page break survives.
  const leading = content.anchored.filter((a) => a.offset <= 0).map((a) => a.node);
  const trailing = content.anchored.filter((a) => a.offset > 0).map((a) => a.node);
  if (leading.length > 0) {
    const first = getChildren(newRuns[0]!);
    const at = first.length > 0 && "w:rPr" in first[0]! ? 1 : 0;
    first.splice(at, 0, ...leading);
  }
  if (trailing.length > 0) {
    getChildren(newRuns[newRuns.length - 1]!).push(...trailing);
  }

  // Find the run in parent's children and replace
  const parentChildren = findRunParentChildren(paragraphNode, runNode);
  if (parentChildren) {
    const idx = parentChildren.indexOf(runNode);
    if (idx !== -1) {
      parentChildren.splice(idx, 1, ...newRuns);
      return;
    }
  }

  // Fallback: just set text on the existing run (shouldn't happen)
  const plainText = segments.map((s) => s.text).join("");
  setRunText(runNode, plainText);
}

// =============================================================================
// Internal: run text manipulation
// =============================================================================

/**
 * Set the text content of a w:r run node.
 *
 * A run that holds exactly one w:t and nothing else of note has that text
 * node updated in place. Otherwise the run is rebuilt from its content:
 * `\n` becomes a soft `w:br`, an old soft break never survives next to the
 * new text, and zero-width children (page and column breaks, tabs,
 * rendering hints) are kept. `remap` says where each such child goes in the
 * new text, given its old offset; without it, a child that preceded the old
 * text stays in front and every other one follows the new text.
 */
function setRunText(
  runNode: FXPNode,
  text: string,
  remap?: (offset: number) => number,
): void {
  const content = readRunContent(runNode);
  const wtNode = findTextNode(runNode);
  const simple =
    wtNode !== undefined &&
    content.anchored.length === 0 &&
    !text.includes("\n") &&
    !content.text.includes("\n") &&
    getChildren(runNode).filter((c) => "w:t" in c).length === 1;

  if (simple) {
    const tChildren = (wtNode as Record<string, unknown>)["w:t"];
    if (Array.isArray(tChildren)) {
      const textNode = tChildren.find(
        (c: Record<string, unknown>) => "#text" in c,
      );
      if (textNode) {
        setTextContent(textNode as FXPNode, text);
      } else {
        tChildren.push({ "#text": escapeXmlText(text) });
      }
    }
    if (needsSpacePreserve(text)) {
      setAttr(wtNode, "xml:space", "preserve");
    }
    return;
  }

  const place = remap ?? ((offset: number) => (offset <= 0 ? 0 : text.length));
  writeRunContent(runNode, {
    rPr: content.rPr,
    text,
    anchored: content.anchored.map((a) => ({ node: a.node, offset: place(a.offset) })),
  });
}
