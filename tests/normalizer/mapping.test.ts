import { describe, it, expect } from "vitest";
import {
  findOccurrence,
  findAllOccurrences,
  resolveRunRanges,
} from "../../src/mapping/text-to-runs.js";
import type { RunMapping } from "../../src/types.js";

// Helper to create mock RunMappings
function mockRuns(texts: string[]): RunMapping[] {
  const runs: RunMapping[] = [];
  let offset = 0;
  for (const text of texts) {
    runs.push({
      text,
      startOffset: offset,
      endOffset: offset + text.length,
      xmlRef: {} as any,
      textXmlRef: {} as any,
      formatting: {},
    });
    offset += text.length;
  }
  return runs;
}

describe("findOccurrence", () => {
  it("should find the first occurrence", () => {
    expect(findOccurrence("hello world hello", "hello", 1)).toBe(0);
  });

  it("should find the second occurrence", () => {
    expect(findOccurrence("hello world hello", "hello", 2)).toBe(12);
  });

  it("should return -1 when occurrence not found", () => {
    expect(findOccurrence("hello world", "hello", 2)).toBe(-1);
  });

  it("should return -1 for empty search string", () => {
    expect(findOccurrence("hello", "", 1)).toBe(-1);
  });

  it("should return -1 when text not found", () => {
    expect(findOccurrence("hello world", "xyz", 1)).toBe(-1);
  });

  it("should default to first occurrence", () => {
    expect(findOccurrence("abc abc", "abc")).toBe(0);
  });
});

describe("findAllOccurrences", () => {
  it("should find all occurrences", () => {
    expect(findAllOccurrences("abcabcabc", "abc")).toEqual([0, 3, 6]);
  });

  it("should handle overlapping matches", () => {
    expect(findAllOccurrences("aaa", "aa")).toEqual([0, 1]);
  });

  it("should return empty for no matches", () => {
    expect(findAllOccurrences("hello", "xyz")).toEqual([]);
  });

  it("should return empty for empty search", () => {
    expect(findAllOccurrences("hello", "")).toEqual([]);
  });
});

describe("resolveRunRanges", () => {
  it("should resolve a range within a single run", () => {
    const runs = mockRuns(["Hello World"]);
    const ranges = resolveRunRanges(runs, 6, 11); // "World"

    expect(ranges.length).toBe(1);
    expect(ranges[0]!.run).toBe(runs[0]);
    expect(ranges[0]!.localStart).toBe(6);
    expect(ranges[0]!.localEnd).toBe(11);
    expect(ranges[0]!.isFullRun).toBe(false);
    expect(ranges[0]!.isStartAligned).toBe(false);
    expect(ranges[0]!.isEndAligned).toBe(true);
  });

  it("should resolve a range spanning two runs", () => {
    const runs = mockRuns(["This is va", "lid text"]);
    // "valid" spans runs[0][8..10] + runs[1][0..3]
    const ranges = resolveRunRanges(runs, 8, 13); // "valid"

    expect(ranges.length).toBe(2);

    // First run: "va"
    expect(ranges[0]!.run).toBe(runs[0]);
    expect(ranges[0]!.localStart).toBe(8);
    expect(ranges[0]!.localEnd).toBe(10);
    expect(ranges[0]!.isEndAligned).toBe(true);

    // Second run: "lid"
    expect(ranges[1]!.run).toBe(runs[1]);
    expect(ranges[1]!.localStart).toBe(0);
    expect(ranges[1]!.localEnd).toBe(3);
    expect(ranges[1]!.isStartAligned).toBe(true);
  });

  it("should resolve a range that covers an entire run", () => {
    const runs = mockRuns(["start ", "middle", " end"]);
    const ranges = resolveRunRanges(runs, 6, 12); // "middle"

    expect(ranges.length).toBe(1);
    expect(ranges[0]!.isFullRun).toBe(true);
    expect(ranges[0]!.isStartAligned).toBe(true);
    expect(ranges[0]!.isEndAligned).toBe(true);
  });

  it("should resolve a range spanning three runs", () => {
    const runs = mockRuns(["aa", "bb", "cc"]);
    const ranges = resolveRunRanges(runs, 1, 5); // "abbc"

    expect(ranges.length).toBe(3);
    expect(ranges[0]!.localStart).toBe(1); // "a" from first run
    expect(ranges[0]!.localEnd).toBe(2);
    expect(ranges[1]!.localStart).toBe(0); // "bb" - full middle run
    expect(ranges[1]!.localEnd).toBe(2);
    expect(ranges[1]!.isFullRun).toBe(true);
    expect(ranges[2]!.localStart).toBe(0); // "c" from last run
    expect(ranges[2]!.localEnd).toBe(1);
  });

  it("should return empty for range outside all runs", () => {
    const runs = mockRuns(["hello"]);
    const ranges = resolveRunRanges(runs, 10, 15);
    expect(ranges.length).toBe(0);
  });

  it("should handle start-aligned single run", () => {
    const runs = mockRuns(["Hello World"]);
    const ranges = resolveRunRanges(runs, 0, 5); // "Hello"

    expect(ranges.length).toBe(1);
    expect(ranges[0]!.isStartAligned).toBe(true);
    expect(ranges[0]!.isEndAligned).toBe(false);
  });
});
