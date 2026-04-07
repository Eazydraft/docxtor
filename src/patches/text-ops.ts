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
 */

import type {
  NormalizedParagraph,
  ReplaceTextPatch,
  InsertTextPatch,
  DeleteTextPatch,
  FXPNode,
} from "../types.js";
import { findOccurrence, resolveRunRanges } from "../mapping/text-to-runs.js";
import { splitRun, splitRunsForMatch } from "./run-splitter.js";
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
} from "../xml/helpers.js";

/**
 * Find the parent node that directly contains a run, and remove the run.
 * Handles runs that are direct children of w:p or nested inside w:hyperlink.
 */
function removeRunFromParagraph(paragraphNode: FXPNode, runNode: FXPNode): void {
  // Try direct removal first
  if (removeChild(paragraphNode, runNode)) return;

  // Try inside hyperlink elements
  const children = getChildren(paragraphNode);
  for (const child of children) {
    if (getTagName(child) === "w:hyperlink") {
      if (removeChild(child, runNode)) return;
    }
  }
}

/**
 * Execute a replace_text patch on a paragraph.
 *
 * Finds the target text, maps to runs, splits if needed, and replaces.
 */
export function executeReplaceText(
  paragraph: NormalizedParagraph,
  patch: ReplaceTextPatch,
): void {
  const occurrence = patch.occurrence ?? 1;
  const matchStart = findOccurrence(paragraph.text, patch.find, occurrence);

  if (matchStart === -1) {
    throw new Error(
      `Text "${patch.find}" not found in paragraph "${paragraph.id}" (occurrence ${occurrence})`,
    );
  }

  const matchEnd = matchStart + patch.find.length;
  const ranges = resolveRunRanges(paragraph.runs, matchStart, matchEnd);

  if (ranges.length === 0) {
    throw new Error(
      `No runs found for text range [${matchStart}, ${matchEnd}] in paragraph "${paragraph.id}"`,
    );
  }

  if (ranges.length === 1) {
    // Single run match — split if needed, then replace text
    const range = ranges[0]!;
    const splitResult = splitRun(paragraph.xmlRef, range);

    // Set the text on the matched run
    setRunText(splitResult.matched, patch.replace);
  } else {
    // Multi-run match — split edges, then consolidate
    const matchedRuns = splitRunsForMatch(paragraph.xmlRef, ranges);

    if (matchedRuns.length === 0) {
      throw new Error(
        `Failed to isolate matched runs in paragraph "${paragraph.id}"`,
      );
    }

    // Put all replacement text in the first matched run
    setRunText(matchedRuns[0]!, patch.replace);

    // Remove the remaining matched runs (they're now empty / redundant)
    for (let i = 1; i < matchedRuns.length; i++) {
      removeRunFromParagraph(paragraph.xmlRef, matchedRuns[i]!);
    }
  }
}

/**
 * Execute an insert_text patch on a paragraph.
 *
 * Strategy: find the insertion offset in the normalized text, then find
 * which run contains that offset. Insert the text directly into that run's
 * w:t node at the correct position.
 *
 * If the offset falls between two runs, we append to the preceding run
 * or prepend to the following run.
 */
export function executeInsertText(
  paragraph: NormalizedParagraph,
  patch: InsertTextPatch,
): void {
  const occurrence = patch.occurrence ?? 1;
  const anchorStart = findOccurrence(paragraph.text, patch.anchor, occurrence);

  if (anchorStart === -1) {
    throw new Error(
      `Anchor text "${patch.anchor}" not found in paragraph "${paragraph.id}" (occurrence ${occurrence})`,
    );
  }

  // Determine the character offset where we want to insert
  const insertOffset =
    patch.position === "before"
      ? anchorStart
      : anchorStart + patch.anchor.length;

  // Find the run that contains the insertion point.
  // We look for the run whose [startOffset, endOffset) range includes insertOffset,
  // or the run boundary where insertOffset sits exactly at a run edge.
  let targetRun: typeof paragraph.runs[number] | undefined;
  let localInsertOffset: number;

  for (const run of paragraph.runs) {
    if (insertOffset >= run.startOffset && insertOffset <= run.endOffset) {
      targetRun = run;
      break;
    }
  }

  if (!targetRun) {
    // Insertion is past all runs (shouldn't happen for valid anchors)
    // but handle gracefully by appending to the last run
    const lastRun = paragraph.runs[paragraph.runs.length - 1];
    if (!lastRun) {
      throw new Error(
        `No runs found in paragraph "${paragraph.id}" for insertion`,
      );
    }
    targetRun = lastRun;
  }

  localInsertOffset = insertOffset - targetRun.startOffset;

  // Build the new text for this run
  const currentText = getRunText(targetRun.xmlRef);
  const newText =
    currentText.slice(0, localInsertOffset) +
    patch.text +
    currentText.slice(localInsertOffset);

  setRunText(targetRun.xmlRef, newText);
}

/**
 * Execute a delete_text patch on a paragraph.
 *
 * This is equivalent to replace_text with an empty replacement.
 */
export function executeDeleteText(
  paragraph: NormalizedParagraph,
  patch: DeleteTextPatch,
): void {
  executeReplaceText(paragraph, {
    type: "replace_text",
    target: patch.target,
    find: patch.find,
    replace: "",
    occurrence: patch.occurrence,
  });
}

/**
 * Set the text content of a w:r run node.
 * Finds or creates the w:t element and sets its text.
 * Escapes XML special characters and manages xml:space="preserve".
 *
 * If the text contains newlines (\n), they are converted to w:br
 * (soft line break) elements within the run. This allows replacement
 * text to contain line breaks.
 */
function setRunText(runNode: FXPNode, text: string): void {
  if (text.includes("\n")) {
    // Multi-line: rebuild the run's content children with w:t/w:br pairs
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
 * Replace a run's text content with text containing line breaks.
 *
 * Splits text on \n and interleaves w:t and w:br elements.
 * Preserves the run's w:rPr formatting.
 */
function setRunTextWithBreaks(runNode: FXPNode, text: string): void {
  const children = getChildren(runNode);

  // Keep w:rPr (formatting), remove everything else (w:t, w:br, etc.)
  const rPrNodes = children.filter((c) => "w:rPr" in c);
  children.length = 0;
  children.push(...rPrNodes);

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
