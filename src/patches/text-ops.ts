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
} from "../types.js";
import { findOccurrence, resolveRunRanges } from "../mapping/text-to-runs.js";
import { splitRun, splitRunsForMatch, findRunParentChildren } from "./run-splitter.js";
import {
  getChildren,
  setTextContent,
  removeChild,
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
 * Find the parent node that directly contains a run, and remove the run.
 * Handles runs that are direct children of w:p or nested inside
 * w:hyperlink, w:ins, or w:del containers.
 */
function removeRunFromParagraph(paragraphNode: FXPNode, runNode: FXPNode): void {
  // Try direct removal first
  if (removeChild(paragraphNode, runNode)) return;

  // Try inside container elements (w:hyperlink, w:ins, w:del)
  const children = getChildren(paragraphNode);
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:hyperlink" || tag === "w:ins" || tag === "w:del") {
      if (removeChild(child, runNode)) return;
      // Check nested containers
      const containerChildren = getChildren(child);
      for (const nested of containerChildren) {
        const nestedTag = getTagName(nested);
        if (nestedTag === "w:hyperlink" || nestedTag === "w:ins" || nestedTag === "w:del") {
          if (removeChild(nested, runNode)) return;
        }
      }
    }
  }
}

/**
 * Find the container element (w:ins, w:del, w:hyperlink) that directly holds
 * a run, or undefined if the run is a direct child of the paragraph.
 */
function findRunContainer(
  paragraphNode: FXPNode,
  runNode: FXPNode,
): FXPNode | undefined {
  const children = getChildren(paragraphNode);
  if (children.includes(runNode)) return undefined;

  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:hyperlink" || tag === "w:ins" || tag === "w:del") {
      const containerChildren = getChildren(child);
      if (containerChildren.includes(runNode)) return child;
      for (const nested of containerChildren) {
        const nestedTag = getTagName(nested);
        if (nestedTag === "w:hyperlink" || nestedTag === "w:ins" || nestedTag === "w:del") {
          if (getChildren(nested).includes(runNode)) return nested;
        }
      }
    }
  }
  return undefined;
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

    // Remove extra matched runs
    for (let i = 1; i < matchedRuns.length; i++) {
      removeRunFromParagraph(paragraph.xmlRef, matchedRuns[i]!);
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

  // Find the parent children array for the first matched run
  const parentChildren = findRunParentChildren(paragraph.xmlRef, matchedRuns[0]!);
  if (!parentChildren) {
    throw new Error("Matched run not found in paragraph");
  }

  // Check if the matched runs live inside a w:ins (unaccepted insertion).
  // If so, we must NOT wrap in w:del — just remove the runs from the w:ins
  // (deleting unaccepted content = removing the insertion, not creating a deletion).
  const container = findRunContainer(paragraph.xmlRef, matchedRuns[0]!);
  const insideIns = container != null && getTagName(container) === "w:ins";

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
  const insElement = newRuns.length > 0 ? createInsElement(revCtx, newRuns) : undefined;

  // Find the position of the first matched run in the parent
  const firstRunIndex = parentChildren.indexOf(matchedRuns[0]!);

  // Remove all matched runs from the parent
  for (const run of matchedRuns) {
    const idx = parentChildren.indexOf(run);
    if (idx !== -1) {
      parentChildren.splice(idx, 1);
    }
  }

  if (insideIns) {
    // The matched runs were an unaccepted insertion, so there is nothing to
    // mark deleted — they are simply gone. The replacement must still land
    // exactly where they were, not after the surrounding w:ins.
    insertIntoInsertion(paragraph.xmlRef, container!, firstRunIndex, newRuns, revCtx);
  } else {
    // Normal case: insert w:del and w:ins at the position of the first matched run
    // Clone matched runs and convert to deleted runs
    const deletedRuns: FXPNode[] = [];
    for (const run of matchedRuns) {
      const cloned = cloneNode(run);
      convertRunToDeleted(cloned);
      deletedRuns.push(cloned);
    }
    const delElement = createDelElement(revCtx, deletedRuns);

    if (insElement) {
      parentChildren.splice(firstRunIndex, 0, delElement, insElement);
    } else {
      parentChildren.splice(firstRunIndex, 0, delElement);
    }
  }
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
 * Direct (non-revision) insertion — modifies the run's text in-place,
 * or creates multiple formatted runs if the text contains GFM markers.
 */
function executeInsertTextDirect(
  paragraph: NormalizedParagraph,
  patch: InsertTextPatch,
  insertOffset: number,
  markers: boolean,
): void {
  let targetRun: typeof paragraph.runs[number] | undefined;

  for (const run of paragraph.runs) {
    if (insertOffset >= run.startOffset && insertOffset <= run.endOffset) {
      targetRun = run;
      break;
    }
  }

  if (!targetRun) {
    const lastRun = paragraph.runs[paragraph.runs.length - 1];
    if (!lastRun) {
      throw new Error(
        `No runs found in paragraph "${paragraph.id}" for insertion`,
      );
    }
    targetRun = lastRun;
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
    setRunText(targetRun.xmlRef, newText);
  } else {
    // Formatted insertion: split the target run at the insert point and
    // insert formatted runs between the halves
    const localInsertOffset = insertOffset - targetRun.startOffset;
    const rPrNode = findChild(targetRun.xmlRef, "w:rPr");
    const parentChildren = findRunParentChildren(paragraph.xmlRef, targetRun.xmlRef);
    if (!parentChildren) {
      throw new Error("Target run not found in paragraph");
    }

    const runIndex = parentChildren.indexOf(targetRun.xmlRef);
    const fullText = targetRun.text;
    const beforeText = fullText.slice(0, localInsertOffset);
    const afterText = fullText.slice(localInsertOffset);

    const newNodes: FXPNode[] = [];

    // Before part (if any)
    if (beforeText.length > 0) {
      newNodes.push(createRunWithTextAndFormatting(beforeText, rPrNode));
    }

    // Formatted inserted runs
    const formattedRuns = createFormattedRuns(segments, rPrNode);
    newNodes.push(...formattedRuns);

    // After part (if any)
    if (afterText.length > 0) {
      newNodes.push(createRunWithTextAndFormatting(afterText, rPrNode));
    }

    // Replace the original run
    parentChildren.splice(runIndex, 1, ...newNodes);
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
  let targetRun: typeof paragraph.runs[number] | undefined;

  for (const run of paragraph.runs) {
    if (insertOffset >= run.startOffset && insertOffset <= run.endOffset) {
      targetRun = run;
      break;
    }
  }

  if (!targetRun) {
    const lastRun = paragraph.runs[paragraph.runs.length - 1];
    if (!lastRun) {
      throw new Error(
        `No runs found in paragraph "${paragraph.id}" for insertion`,
      );
    }
    targetRun = lastRun;
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
    // Split the run at the insertion point
    const fullText = targetRun.text;
    const beforeText = fullText.slice(0, localInsertOffset);
    const afterText = fullText.slice(localInsertOffset);

    // Create two halves preserving formatting
    const beforeRun = createRunWithTextAndFormatting(beforeText, rPrNode);
    const afterRun = createRunWithTextAndFormatting(afterText, rPrNode);

    // Replace the original run with: beforeRun, insElement, afterRun
    parentChildren.splice(runIndex, 1, beforeRun, insElement, afterRun);
  }
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

  // Find the parent children array for the first matched run
  const parentChildren = findRunParentChildren(paragraph.xmlRef, matchedRuns[0]!);
  if (!parentChildren) {
    throw new Error("Matched run not found in paragraph");
  }

  // Check if the matched runs live inside a w:ins (unaccepted insertion).
  const container = findRunContainer(paragraph.xmlRef, matchedRuns[0]!);
  const insideIns = container != null && getTagName(container) === "w:ins";

  // Find the position of the first matched run in the parent
  const firstRunIndex = parentChildren.indexOf(matchedRuns[0]!);

  // Remove all matched runs from the parent
  for (const run of matchedRuns) {
    const idx = parentChildren.indexOf(run);
    if (idx !== -1) {
      parentChildren.splice(idx, 1);
    }
  }

  if (insideIns) {
    // Runs were inside w:ins — just remove them (and the w:ins if now empty).
    // No w:del needed: the content was never part of the accepted document.
    if (parentChildren.length === 0 && container) {
      const paraChildren = getChildren(paragraph.xmlRef);
      const containerIdx = paraChildren.indexOf(container);
      if (containerIdx !== -1) {
        paraChildren.splice(containerIdx, 1);
      }
    }
  } else {
    // Normal case: clone matched runs, wrap in w:del
    const deletedRuns: FXPNode[] = [];
    for (const run of matchedRuns) {
      const cloned = cloneNode(run);
      convertRunToDeleted(cloned);
      deletedRuns.push(cloned);
    }
    const delElement = createDelElement(revCtx, deletedRuns);
    parentChildren.splice(firstRunIndex, 0, delElement);
  }
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
  const rPrNode = findChild(runNode, "w:rPr");
  const newRuns = createFormattedRuns(segments, rPrNode);

  if (newRuns.length === 0) {
    // All segments were empty — just clear the run text
    setRunText(runNode, "");
    return;
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
 * Finds or creates the w:t element and sets its text.
 * Escapes XML special characters and manages xml:space="preserve".
 *
 * If the text contains newlines (\n), they are converted to w:br
 * (soft line break) elements within the run. This allows replacement
 * text to contain line breaks. A run that already holds a soft break
 * (or several w:t elements) is rebuilt as well, so the old break never
 * survives next to the new text.
 */
function setRunText(runNode: FXPNode, text: string): void {
  if (text.includes("\n") || runHasMixedContent(runNode)) {
    // Rebuild the run's content children with w:t/w:br pairs
    setRunTextWithBreaks(runNode, text);
    return;
  }

  const wtNode = findTextNode(runNode);
  if (wtNode) {
    const tChildren = (wtNode as Record<string, unknown>)["w:t"];
    if (Array.isArray(tChildren)) {
      // Find or create the text node
      const textNode = tChildren.find(
        (c: Record<string, unknown>) => "#text" in c,
      );
      if (textNode) {
        setTextContent(textNode as FXPNode, text);
      } else {
        tChildren.push({ "#text": escapeXmlText(text) });
      }
    }

    // Update xml:space="preserve" if needed
    if (needsSpacePreserve(text)) {
      setAttr(wtNode, "xml:space", "preserve");
    }
  } else {
    // No w:t exists — create one
    const children = getChildren(runNode);
    const newWt: FXPNode = {
      "w:t": [{ "#text": escapeXmlText(text) } as unknown as FXPNode],
    };
    if (needsSpacePreserve(text)) {
      setAttr(newWt, "xml:space", "preserve");
    }
    children.push(newWt);
  }
}

/**
 * True if the run's text cannot be updated by editing a single w:t: it holds
 * a soft line break, or more than one w:t element.
 */
function runHasMixedContent(runNode: FXPNode): boolean {
  let textNodes = 0;
  for (const child of getChildren(runNode)) {
    if ("w:br" in child) {
      const brType = getAttr(child, "w:type");
      if (!brType || brType === "textWrapping") return true;
    } else if ("w:t" in child) {
      textNodes++;
    }
  }
  return textNodes > 1;
}

/**
 * Replace a run's text content with text containing line breaks.
 *
 * Splits text on \n and interleaves w:t and w:br elements.
 * Preserves the run's w:rPr (and any other non-text children such as
 * page breaks or tabs, which keep their position relative to the start).
 */
function setRunTextWithBreaks(runNode: FXPNode, text: string): void {
  const children = getChildren(runNode);

  // Keep w:rPr (formatting) and non-text children such as page breaks;
  // drop w:t and soft breaks, which the new text replaces
  const kept = children.filter((c) => {
    if ("w:t" in c) return false;
    if ("w:br" in c) {
      const brType = getAttr(c, "w:type");
      return brType !== undefined && brType !== "textWrapping";
    }
    return true;
  });
  children.length = 0;
  children.push(...kept);

  const segments = text.split("\n");
  for (let i = 0; i < segments.length; i++) {
    const seg = segments[i]!;
    if (seg.length > 0 || i === 0) {
      // Add w:t element for this segment
      const wt: FXPNode = {
        "w:t": [{ "#text": escapeXmlText(seg) } as unknown as FXPNode],
      };
      if (needsSpacePreserve(seg)) {
        setAttr(wt, "xml:space", "preserve");
      }
      children.push(wt);
    }
    // Add w:br between segments (not after the last one)
    if (i < segments.length - 1) {
      children.push({ "w:br": [] as FXPNode[] });
    }
  }
}
