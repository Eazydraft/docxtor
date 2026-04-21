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
  setAttr,
} from "../xml/helpers.js";

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
 * Build FXP run nodes for a text string that may contain field placeholders.
 *
 * Plain text segments produce normal w:r runs.
 * Field segments produce the standard OOXML field code pattern:
 *   <w:r><w:fldChar w:fldCharType="begin"/></w:r>
 *   <w:r><w:instrText> FIELDCODE </w:instrText></w:r>
 *   <w:r><w:fldChar w:fldCharType="separate"/></w:r>
 *   <w:r><w:fldChar w:fldCharType="end"/></w:r>
 *
 * @param text - Text possibly containing {{FIELD}} placeholders
 * @param rPrNode - Optional run properties to clone onto each run
 * @returns Array of w:r FXP nodes
 */
export function buildFieldRuns(
  text: string,
  rPrNode?: FXPNode,
): FXPNode[] {
  const segments = parseFieldPlaceholders(text);
  const runs: FXPNode[] = [];

  for (const segment of segments) {
    if (segment.kind === "text") {
      if (segment.text.length > 0) {
        runs.push(createRun(segment.text, rPrNode));
      }
    } else {
      // Field code: emit begin/instrText/separate/end pattern
      runs.push(...buildFieldCodeRuns(segment.fieldCode, rPrNode));
    }
  }

  return runs;
}

/**
 * Build the w:fldChar / w:instrText run sequence for a field code.
 */
function buildFieldCodeRuns(
  fieldCode: string,
  rPrNode?: FXPNode,
): FXPNode[] {
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
    "w:instrText": [
      { "#text": instrText } as unknown as FXPNode,
    ],
  };
  setAttr(instrTextNode, "xml:space", "preserve");

  const instrRun: FXPNode = {
    "w:r": [
      ...cloneRPr(),
      instrTextNode,
    ],
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
