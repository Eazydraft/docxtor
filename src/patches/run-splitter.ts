/**
 * Run splitter: the critical piece of surgical editing.
 *
 * When a text match doesn't align with run boundaries, we need to split
 * runs so the edit can be applied precisely. For example:
 *
 *   Run 1: "This agreement is va"  (plain)
 *   Run 2: "lid and enforceable"   (bold)
 *
 * To replace "valid" → "enforceable", we need to:
 *   1. Split run 1 at offset 18: "This agreement is " | "va"
 *   2. Split run 2 at offset 3:  "lid" | " and enforceable"
 *   3. Now "va" + "lid" are isolated runs that can be replaced
 *
 * The split preserves formatting (w:rPr) — new runs clone the original's
 * properties. This is what ensures minimal redlines and formatting preservation.
 */

import type { FXPNode } from "../types.js";
import type { ResolvedRunRange } from "../mapping/text-to-runs.js";
import {
  findChild,
  cloneNode,
  getChildren,
  getTagName,
  setAttr,
  needsSpacePreserve,
  escapeXmlText,
} from "../xml/helpers.js";

/**
 * Result of splitting a run.
 */
export interface SplitResult {
  /** The part before the matched range (may be undefined if start-aligned). */
  before?: FXPNode;
  /** The matched part (the run covering the matched text). */
  matched: FXPNode;
  /** The part after the matched range (may be undefined if end-aligned). */
  after?: FXPNode;
}

/**
 * Find the parent children array that directly contains a run node.
 *
 * Runs can be direct children of w:p, or nested inside w:hyperlink,
 * w:ins, or w:del. We need the actual parent's children array to do
 * splice operations.
 */
export function findRunParentChildren(
  paragraphNode: FXPNode,
  runNode: FXPNode,
): FXPNode[] | undefined {
  // Check direct children first
  const children = getChildren(paragraphNode);
  if (children.includes(runNode)) {
    return children;
  }

  // Check inside container elements (w:hyperlink, w:ins, w:del)
  return findRunInContainers(children, runNode);
}

/**
 * Recursively search container elements for a run node.
 */
function findRunInContainers(
  children: FXPNode[],
  runNode: FXPNode,
): FXPNode[] | undefined {
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:hyperlink" || tag === "w:ins" || tag === "w:del") {
      const containerChildren = getChildren(child);
      if (containerChildren.includes(runNode)) {
        return containerChildren;
      }
      // Recurse deeper (e.g., w:ins inside w:hyperlink)
      const found = findRunInContainers(containerChildren, runNode);
      if (found) return found;
    }
  }
  return undefined;
}

/**
 * Split a run at the boundaries defined by a ResolvedRunRange.
 *
 * This mutates the parent paragraph's children array:
 * - The original run is replaced with up to 3 runs (before, matched, after)
 * - All new runs clone the original's w:rPr formatting
 *
 * @param paragraphNode - The w:p FXP node containing the run
 * @param range - The resolved range within the run to isolate
 * @returns The SplitResult with references to the new nodes
 */
export function splitRun(
  paragraphNode: FXPNode,
  range: ResolvedRunRange,
): SplitResult {
  // If the range covers the full run, no splitting needed
  if (range.isFullRun) {
    return { matched: range.run.xmlRef };
  }

  const runNode = range.run.xmlRef;
  const fullText = range.run.text;
  const parentChildren = findRunParentChildren(paragraphNode, runNode);

  if (!parentChildren) {
    throw new Error("Run not found in paragraph children or hyperlink children");
  }

  const runIndex = parentChildren.indexOf(runNode);

  // Extract the w:rPr node (formatting) from the original run
  const rPrNode = findChild(runNode, "w:rPr");

  // Calculate the three text segments
  const beforeText = fullText.slice(0, range.localStart);
  const matchedText = fullText.slice(range.localStart, range.localEnd);
  const afterText = fullText.slice(range.localEnd);

  const result: SplitResult = {
    matched: undefined!,
  };

  // Build the replacement runs
  const newRuns: FXPNode[] = [];

  if (beforeText.length > 0) {
    const beforeRun = createRunWithText(beforeText, rPrNode);
    result.before = beforeRun;
    newRuns.push(beforeRun);
  }

  const matchedRun = createRunWithText(matchedText, rPrNode);
  result.matched = matchedRun;
  newRuns.push(matchedRun);

  if (afterText.length > 0) {
    const afterRun = createRunWithText(afterText, rPrNode);
    result.after = afterRun;
    newRuns.push(afterRun);
  }

  // Replace the original run with the new runs in the parent
  parentChildren.splice(runIndex, 1, ...newRuns);

  return result;
}

/**
 * Split runs across multiple ResolvedRunRanges to isolate a text match.
 *
 * For a match that spans multiple runs, this:
 * 1. Splits the first run at the match start (if not start-aligned)
 * 2. Keeps middle runs intact (they're fully covered)
 * 3. Splits the last run at the match end (if not end-aligned)
 *
 * @returns Array of FXP run nodes that together contain exactly the matched text
 */
export function splitRunsForMatch(
  paragraphNode: FXPNode,
  ranges: ResolvedRunRange[],
): FXPNode[] {
  if (ranges.length === 0) return [];

  if (ranges.length === 1) {
    const result = splitRun(paragraphNode, ranges[0]!);
    return [result.matched];
  }

  // Multi-run match: split from last to first to preserve indices
  const matchedRuns: FXPNode[] = [];

  // Process in reverse order to avoid index invalidation
  // But collect results in forward order
  const results: { index: number; split: SplitResult }[] = [];

  // Split last run first (if needed)
  const lastRange = ranges[ranges.length - 1]!;
  if (!lastRange.isEndAligned) {
    const split = splitRun(paragraphNode, lastRange);
    results.push({ index: ranges.length - 1, split });
  }

  // Split first run (if needed) — must be done after last to avoid index issues
  // Actually, since we're working on different runs, order doesn't matter
  // as long as we split from right to left in the children array
  const firstRange = ranges[0]!;
  if (!firstRange.isStartAligned) {
    const split = splitRun(paragraphNode, firstRange);
    results.push({ index: 0, split });
  }

  // Now collect all matched runs in order
  // After splitting, we need to re-find the runs
  for (let i = 0; i < ranges.length; i++) {
    const range = ranges[i]!;
    const result = results.find((r) => r.index === i);

    if (result) {
      matchedRuns.push(result.split.matched);
    } else {
      // Middle run or already-aligned run: use original xmlRef
      // Check the run still exists in its parent (splits may have shifted)
      const parentChildren = findRunParentChildren(paragraphNode, range.run.xmlRef);
      if (parentChildren && parentChildren.includes(range.run.xmlRef)) {
        matchedRuns.push(range.run.xmlRef);
      }
    }
  }

  return matchedRuns;
}

/**
 * Create a new w:r run node with the given text and optional formatting.
 */
function createRunWithText(
  text: string,
  rPrNode: FXPNode | undefined,
): FXPNode {
  const children: FXPNode[] = [];

  // Clone formatting if present
  if (rPrNode) {
    children.push(cloneNode(rPrNode));
  }

  // Create w:t element with text
  const textNode = { "#text": escapeXmlText(text) } as unknown as FXPNode;
  const wtNode: FXPNode = { "w:t": [textNode] };

  // Add xml:space="preserve" if text has leading/trailing whitespace
  if (needsSpacePreserve(text)) {
    setAttr(wtNode, "xml:space", "preserve");
  }

  children.push(wtNode);

  return { "w:r": children };
}
