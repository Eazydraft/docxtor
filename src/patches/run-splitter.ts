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
import { getChildren, getTagName } from "../xml/helpers.js";
import {
  readRunContent,
  partitionContent,
  createRunFromContent,
  hasContent,
} from "./run-content.js";

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
 * - Zero-width children (page and column breaks, tabs, rendering hints,
 *   drawings) stay at the character position they had: a break just before
 *   the match lands in the before run, one just after it in the after run,
 *   so a partial replacement never drops them
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
  const parentChildren = findRunParentChildren(paragraphNode, runNode);

  if (!parentChildren) {
    throw new Error("Run not found in paragraph children or hyperlink children");
  }

  const runIndex = parentChildren.indexOf(runNode);
  const parts = partitionContent(readRunContent(runNode), range.localStart, range.localEnd);

  const result: SplitResult = {
    matched: undefined!,
  };

  // Build the replacement runs
  const newRuns: FXPNode[] = [];

  if (hasContent(parts.before)) {
    const beforeRun = createRunFromContent(parts.before);
    result.before = beforeRun;
    newRuns.push(beforeRun);
  }

  const matchedRun = createRunFromContent(parts.matched);
  result.matched = matchedRun;
  newRuns.push(matchedRun);

  if (hasContent(parts.after)) {
    const afterRun = createRunFromContent(parts.after);
    result.after = afterRun;
    newRuns.push(afterRun);
  }

  // Replace the original run with the new runs in the parent
  parentChildren.splice(runIndex, 1, ...newRuns);

  return result;
}

/**
 * Split a run in two at a character offset, for an insertion between the
 * halves. A zero-width child sitting exactly at the offset goes with the
 * half named by `childrenAtOffset`: text inserted before an anchor that
 * follows a page break lands after the break ("before"), text inserted
 * after an anchor that precedes a break lands before it ("after").
 *
 * The original run is replaced in its parent by the halves that have any
 * content; the returned index is where the inserted node(s) should go.
 */
export function splitRunAtOffset(
  paragraphNode: FXPNode,
  runNode: FXPNode,
  offset: number,
  childrenAtOffset: "before" | "after" = "before",
): { parentChildren: FXPNode[]; index: number; before?: FXPNode; after?: FXPNode } {
  const parentChildren = findRunParentChildren(paragraphNode, runNode);
  if (!parentChildren) {
    throw new Error("Run not found in paragraph children or hyperlink children");
  }
  const runIndex = parentChildren.indexOf(runNode);
  const parts = partitionContent(readRunContent(runNode), offset, offset, childrenAtOffset);

  const newRuns: FXPNode[] = [];
  const result: { parentChildren: FXPNode[]; index: number; before?: FXPNode; after?: FXPNode } = {
    parentChildren,
    index: runIndex,
  };
  if (hasContent(parts.before)) {
    result.before = createRunFromContent(parts.before);
    newRuns.push(result.before);
    result.index = runIndex + 1;
  }
  if (hasContent(parts.after)) {
    result.after = createRunFromContent(parts.after);
    newRuns.push(result.after);
  }
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
