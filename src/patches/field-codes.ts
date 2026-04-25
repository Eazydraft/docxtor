/**
 * Field code placeholder support for OOXML.
 *
 * Supports {{FIELD}} syntax in text content. When encountered, emits
 * proper w:fldChar / w:instrText XML instead of plain text.
 *
 * Supported fields:
 * - {{PAGE}}     → PAGE      (current page number)
 * - {{NUMPAGES}} → NUMPAGES  (total page count)
 * - {{DATE}}     → DATE      (current date)
 * - {{TIME}}     → TIME      (current time)
 * - {{FILENAME}} → FILENAME  (document filename)
 */

import type { FXPNode } from "../types.js";
import {
  createElement,
  createRun,
  getChildren,
  setAttr,
} from "../xml/helpers.js";
import type { FormattedSegment } from "./markdown-parser.js";
import { parseFormattedText } from "./markdown-parser.js";
import { createFormattedRuns } from "./text-ops.js";

/** Supported OOXML field codes. */
const SUPPORTED_FIELDS = new Set([
  "PAGE",
  "NUMPAGES",
  "DATE",
  "TIME",
  "FILENAME",
]);

/** A parsed segment: either plain text or a field code. */
export type TextSegment =
  | { kind: "text"; text: string }
  | { kind: "field"; fieldCode: string };

/**
 * Parse text containing {{FIELD}} placeholders into segments.
 *
 * Example: "Page {{PAGE}} of {{NUMPAGES}}" →
 *   [{ kind: "text", text: "Page " },
 *    { kind: "field", fieldCode: "PAGE" },
 *    { kind: "text", text: " of " },
 *    { kind: "field", fieldCode: "NUMPAGES" }]
 */
export function parseFieldPlaceholders(text: string): TextSegment[] {
  const segments: TextSegment[] = [];
  const regex = /\{\{(\w+)\}\}/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = regex.exec(text)) !== null) {
    const fieldName = match[1]!;

    // Only treat as field if it's a supported field code
    if (!SUPPORTED_FIELDS.has(fieldName)) {
      continue;
    }

    // Add preceding plain text
    if (match.index > lastIndex) {
      segments.push({ kind: "text", text: text.slice(lastIndex, match.index) });
    }

    segments.push({ kind: "field", fieldCode: fieldName });
    lastIndex = match.index + match[0].length;
  }

  // Add trailing text
  if (lastIndex < text.length) {
    segments.push({ kind: "text", text: text.slice(lastIndex) });
  }

  // If no fields were found, return a single text segment
  if (segments.length === 0) {
    segments.push({ kind: "text", text });
  }

  return segments;
}

/**
 * Check if text contains any field placeholders.
 */
export function hasFieldPlaceholders(text: string): boolean {
  const regex = /\{\{(\w+)\}\}/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(text)) !== null) {
    if (SUPPORTED_FIELDS.has(match[1]!)) return true;
  }
  return false;
}

/**
 * Build FXP run nodes for a text string that may contain field placeholders
 * and/or inline markdown formatting.
 *
 * Formatting state is tracked across field boundaries so that markers like
 * `**{{PAGE}}**` correctly apply bold to the field runs rather than emitting
 * literal `**` characters.
 *
 * Algorithm:
 * 1. Parse field placeholders to get text/field segments.
 * 2. Concatenate just the text parts and run the GFM parser on that combined
 *    string to get a flat list of FormattedSegments with correct formatting.
 * 3. Walk the original segments; for text segments consume characters from the
 *    formatted list (which already has markers stripped and flags set). For
 *    field segments, snapshot the currently-active formatting and apply it to
 *    the field code runs.
 *
 * @param text - Text possibly containing {{FIELD}} placeholders and/or formatting markers
 * @param rPrNode - Optional run properties to clone onto each run
 * @returns Array of w:r FXP nodes
 */
export function buildFieldRuns(text: string, rPrNode?: FXPNode): FXPNode[] {
  const segments = parseFieldPlaceholders(text);

  // Fast path: no field placeholders — just parse formatting directly
  if (segments.length === 1 && segments[0]!.kind === "text") {
    const formatted = parseFormattedText(segments[0]!.text);
    const hasFormatting = formatted.some(
      (s) => s.bold || s.italic || s.strikethrough || s.underline,
    );
    if (hasFormatting) {
      return createFormattedRuns(formatted, rPrNode);
    }
    // Still strip markers (though parseFormattedText already strips them
    // for well-formed input, this handles edge cases)
    const plainText = formatted.map((s) => s.text).join("");
    return plainText.length > 0 ? [createRun(plainText, rPrNode)] : [];
  }

  // Concatenate all text segments (preserving order) and parse formatting
  // on the combined string so state carries across field boundaries.
  const combinedText = segments
    .filter((s): s is { kind: "text"; text: string } => s.kind === "text")
    .map((s) => s.text)
    .join("");

  const formattedSegs = parseFormattedText(combinedText);

  // Build a flat character-level formatting map from the parsed segments.
  // We'll consume characters from this as we walk through the original segments.
  interface CharFmt {
    bold: boolean;
    italic: boolean;
    strikethrough: boolean;
    underline: boolean;
  }
  const charFmts: CharFmt[] = [];
  for (const seg of formattedSegs) {
    const fmt: CharFmt = {
      bold: !!seg.bold,
      italic: !!seg.italic,
      strikethrough: !!seg.strikethrough,
      underline: !!seg.underline,
    };
    for (let i = 0; i < seg.text.length; i++) {
      charFmts.push(fmt);
    }
  }

  const runs: FXPNode[] = [];
  let charIdx = 0; // position in charFmts

  // Track the "last seen" formatting for field segments
  let activeFmt: CharFmt = {
    bold: false,
    italic: false,
    strikethrough: false,
    underline: false,
  };

  for (const segment of segments) {
    if (segment.kind === "text") {
      // Consume len(stripped text) characters from charFmts.
      // The stripped text length == number of chars parseFormattedText produced
      // for this segment's input. We need to figure out how many chars from
      // charFmts belong to this text segment.
      const segFormatted = parseFormattedText(segment.text);
      const segPlainLen = segFormatted.reduce((n, s) => n + s.text.length, 0);

      if (segPlainLen > 0) {
        // Re-group the character formatting back into contiguous runs
        const subFmts = charFmts.slice(charIdx, charIdx + segPlainLen);
        const segPlainText = segFormatted.map((s) => s.text).join("");

        // Build FormattedSegment[] from chars
        const grouped: FormattedSegment[] = [];
        let runStart = 0;
        for (let i = 1; i <= subFmts.length; i++) {
          if (
            i === subFmts.length ||
            subFmts[i]!.bold !== subFmts[runStart]!.bold ||
            subFmts[i]!.italic !== subFmts[runStart]!.italic ||
            subFmts[i]!.strikethrough !== subFmts[runStart]!.strikethrough ||
            subFmts[i]!.underline !== subFmts[runStart]!.underline
          ) {
            const fmt = subFmts[runStart]!;
            const seg: FormattedSegment = {
              text: segPlainText.slice(runStart, i),
            };
            if (fmt.bold) seg.bold = true;
            if (fmt.italic) seg.italic = true;
            if (fmt.strikethrough) seg.strikethrough = true;
            if (fmt.underline) seg.underline = true;
            grouped.push(seg);
            runStart = i;
          }
        }

        const hasFormatting = grouped.some(
          (s) => s.bold || s.italic || s.strikethrough || s.underline,
        );
        if (hasFormatting) {
          runs.push(...createFormattedRuns(grouped, rPrNode));
        } else {
          runs.push(createRun(segPlainText, rPrNode));
        }

        // Update activeFmt to the formatting at the end of this text segment
        if (subFmts.length > 0) {
          activeFmt = subFmts[subFmts.length - 1]!;
        }
      }

      charIdx += segPlainLen;
    } else {
      // Field code: apply active formatting to the field runs
      const fieldRPr = buildFormattedRPr(rPrNode, activeFmt);
      runs.push(...buildFieldCodeRuns(segment.fieldCode, fieldRPr));
    }
  }

  return runs;
}

/**
 * Build the w:fldChar / w:instrText run sequence for a field code.
 */
function buildFieldCodeRuns(fieldCode: string, rPrNode?: FXPNode): FXPNode[] {
  const cloneRPr = (): FXPNode[] => {
    if (!rPrNode) return [];
    return [JSON.parse(JSON.stringify(rPrNode)) as FXPNode];
  };

  // Begin run
  const beginRun: FXPNode = {
    "w:r": [
      ...cloneRPr(),
      createElement("w:fldChar", [], { "w:fldCharType": "begin" }),
    ],
  };

  // Instruction run
  const instrText = ` ${fieldCode} `;
  const instrTextNode: FXPNode = {
    "w:instrText": [{ "#text": instrText } as unknown as FXPNode],
  };
  setAttr(instrTextNode, "xml:space", "preserve");

  const instrRun: FXPNode = {
    "w:r": [...cloneRPr(), instrTextNode],
  };

  // Separate run
  const separateRun: FXPNode = {
    "w:r": [
      ...cloneRPr(),
      createElement("w:fldChar", [], { "w:fldCharType": "separate" }),
    ],
  };

  // End run
  const endRun: FXPNode = {
    "w:r": [
      ...cloneRPr(),
      createElement("w:fldChar", [], { "w:fldCharType": "end" }),
    ],
  };

  return [beginRun, instrRun, separateRun, endRun];
}

/**
 * Build a w:rPr node that merges the base rPrNode with GFM formatting flags.
 * Returns undefined if neither base nor flags produce any properties.
 */
function buildFormattedRPr(
  baseRPr: FXPNode | undefined,
  fmt: { bold: boolean; italic: boolean; strikethrough: boolean; underline: boolean },
): FXPNode | undefined {
  const needsFormatting = fmt.bold || fmt.italic || fmt.strikethrough || fmt.underline;
  if (!baseRPr && !needsFormatting) return undefined;

  // Clone the base rPr (or create a fresh one)
  const rPr: FXPNode = baseRPr
    ? (JSON.parse(JSON.stringify(baseRPr)) as FXPNode)
    : createElement("w:rPr", []);

  if (!needsFormatting) return rPr;

  const children = getChildren(rPr);

  if (fmt.bold && !children.some((c) => "w:b" in c)) {
    children.push(createElement("w:b", []));
  }
  if (fmt.italic && !children.some((c) => "w:i" in c)) {
    children.push(createElement("w:i", []));
  }
  if (fmt.strikethrough && !children.some((c) => "w:strike" in c)) {
    children.push(createElement("w:strike", []));
  }
  if (fmt.underline && !children.some((c) => "w:u" in c)) {
    children.push(createElement("w:u", [], { "w:val": "single" }));
  }

  return rPr;
}
