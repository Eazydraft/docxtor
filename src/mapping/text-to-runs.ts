/**
 * Text-to-runs mapping: given a character range in the normalized
 * paragraph text, resolves which RunMapping(s) are affected and
 * computes sub-run offsets.
 *
 * This is the bridge between "find text in paragraph" and
 * "mutate the right XML runs."
 */

import type { RunMapping } from "../types.js";

/**
 * A resolved range within a specific run.
 * Tells you exactly which portion of a run is affected by a text match.
 */
export interface ResolvedRunRange {
  /** The RunMapping this range falls in. */
  run: RunMapping;
  /** Start offset within this run's text (0-based, inclusive). */
  localStart: number;
  /** End offset within this run's text (0-based, exclusive). */
  localEnd: number;
  /** Whether this range covers the entire run text. */
  isFullRun: boolean;
  /** Whether this range starts at the beginning of the run. */
  isStartAligned: boolean;
  /** Whether this range ends at the end of the run. */
  isEndAligned: boolean;
}

/**
 * Find the Nth occurrence of a substring in a paragraph's text.
 *
 * @param paragraphText - The full normalized paragraph text
 * @param find - The text to search for
 * @param occurrence - Which occurrence (1-based). Default: 1.
 * @returns The start index of the match, or -1 if not found.
 */
export function findOccurrence(
  paragraphText: string,
  find: string,
  occurrence: number = 1,
): number {
  if (find.length === 0) return -1;

  let count = 0;
  let startIdx = 0;

  while (startIdx <= paragraphText.length - find.length) {
    const idx = paragraphText.indexOf(find, startIdx);
    if (idx === -1) return -1;

    count++;
    if (count === occurrence) return idx;

    startIdx = idx + 1;
  }

  return -1;
}

/**
 * Resolve a character range to the affected RunMapping(s).
 *
 * Given a start and end offset in the paragraph's normalized text,
 * returns the list of runs that overlap with this range, along with
 * the local offsets within each run.
 *
 * @param runs - The paragraph's RunMapping array
 * @param rangeStart - Start offset in paragraph text (inclusive)
 * @param rangeEnd - End offset in paragraph text (exclusive)
 */
export function resolveRunRanges(
  runs: RunMapping[],
  rangeStart: number,
  rangeEnd: number,
): ResolvedRunRange[] {
  const results: ResolvedRunRange[] = [];

  for (const run of runs) {
    // Check if this run overlaps with the range
    if (run.endOffset <= rangeStart || run.startOffset >= rangeEnd) {
      continue; // No overlap
    }

    // Calculate local offsets within this run
    const localStart = Math.max(0, rangeStart - run.startOffset);
    const localEnd = Math.min(run.text.length, rangeEnd - run.startOffset);

    results.push({
      run,
      localStart,
      localEnd,
      isFullRun: localStart === 0 && localEnd === run.text.length,
      isStartAligned: localStart === 0,
      isEndAligned: localEnd === run.text.length,
    });
  }

  return results;
}

/**
 * Find all occurrences of a substring in a paragraph's text.
 *
 * @param paragraphText - The full normalized paragraph text
 * @param find - The text to search for
 * @returns Array of start indices
 */
export function findAllOccurrences(
  paragraphText: string,
  find: string,
): number[] {
  if (find.length === 0) return [];

  const indices: number[] = [];
  let startIdx = 0;

  while (startIdx <= paragraphText.length - find.length) {
    const idx = paragraphText.indexOf(find, startIdx);
    if (idx === -1) break;
    indices.push(idx);
    startIdx = idx + 1;
  }

  return indices;
}
