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
  setAttr,
} from "../xml/helpers.js";

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
      removeChild(paragraph.xmlRef, matchedRuns[i]!);
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
 */
function setRunText(runNode: FXPNode, text: string): void {
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
        tChildren.push({ "#text": text });
      }
    }

    // Update xml:space="preserve" if needed
    if (text !== text.trim() || text.includes("  ")) {
      setAttr(wtNode, "xml:space", "preserve");
    }
  } else {
    // No w:t exists — create one
    const children = getChildren(runNode);
    const newWt: FXPNode = { "w:t": [{ "#text": text } as unknown as FXPNode] };
    if (text !== text.trim() || text.includes("  ")) {
      setAttr(newWt, "xml:space", "preserve");
    }
    children.push(newWt);
  }

  // If text is empty, remove the run entirely? No — let the caller decide.
  // Empty runs are valid XML, and the caller may want to keep formatting.
}
