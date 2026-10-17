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

/** Stands in for a field while inline markers are parsed around it. */
const FIELD_PLACEHOLDER = "\uFFFC";

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
 * 2. Replace every field with a placeholder character and run the GFM parser
 *    on the whole string, so markers open and close across fields and the
 *    placeholder (a non-word character) keeps `_{{PAGE}}_` a valid italic.
 * 3. Walk the formatted segments: text becomes runs with that formatting,
 *    each placeholder becomes the field-code run sequence carrying the
 *    formatting active at that point.
 *
 * @param text - Text possibly containing {{FIELD}} placeholders and/or formatting markers
 * @param rPrNode - Optional run properties to clone onto each run
 * @param markers - Parse inline formatting markers and escapes (default true);
 *   with `false` the text is written verbatim
 * @returns Array of w:r FXP nodes
 */
export function buildFieldRuns(text: string, rPrNode?: FXPNode, markers: boolean = true): FXPNode[] {
  const segments = parseFieldPlaceholders(text);

  // Fast path: no field placeholders — just parse formatting directly
  if (segments.length === 1 && segments[0]!.kind === "text") {
    return emitTextRuns(parseFormattedText(segments[0]!.text, markers), rPrNode);
  }

  const fieldCodes = segments
    .filter((s): s is { kind: "field"; fieldCode: string } => s.kind === "field")
    .map((s) => s.fieldCode);
  const combined = segments
    .map((s) => (s.kind === "text" ? s.text : FIELD_PLACEHOLDER))
    .join("");
  const formatted = parseFormattedText(combined, markers);

  const runs: FXPNode[] = [];
  let pending: FormattedSegment[] = [];
  let fieldIdx = 0;
  const flushPending = (): void => {
    if (pending.length > 0) {
      runs.push(...emitTextRuns(pending, rPrNode));
      pending = [];
    }
  };

  for (const seg of formatted) {
    const parts = seg.text.split(FIELD_PLACEHOLDER);
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i]!;
      if (part.length > 0) pending.push({ ...seg, text: part });
      if (i < parts.length - 1) {
        flushPending();
        const fieldRPr = buildFormattedRPr(rPrNode, {
          bold: !!seg.bold,
          italic: !!seg.italic,
          strikethrough: !!seg.strikethrough,
          underline: !!seg.underline,
        });
        runs.push(...buildFieldCodeRuns(fieldCodes[fieldIdx++]!, fieldRPr));
      }
    }
  }
  flushPending();

  return runs;
}

/**
 * Turn formatted segments into runs: one formatted run per segment when any
 * segment carries formatting, otherwise a single plain run.
 */
function emitTextRuns(segments: FormattedSegment[], rPrNode?: FXPNode): FXPNode[] {
  const hasFormatting = segments.some(
    (s) => s.bold || s.italic || s.strikethrough || s.underline,
  );
  if (hasFormatting) {
    return createFormattedRuns(segments, rPrNode);
  }
  const plainText = segments.map((s) => s.text).join("");
  return plainText.length > 0 ? [createRun(plainText, rPrNode)] : [];
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
