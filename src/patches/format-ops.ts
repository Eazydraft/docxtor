/**
 * Formatting operations: apply_formatting patch execution.
 *
 * Manipulates inline w:rPr (run properties) on targeted runs.
 * Never touches w:pPr, styles.xml, or any global definitions.
 *
 * The approach:
 * 1. If `find` is provided: locate the substring via findOccurrence + resolveRunRanges,
 *    split runs to isolate the match, then apply formatting to isolated runs.
 * 2. If `find` is omitted: apply formatting to all runs in the paragraph.
 *
 * Formatting semantics:
 * - `true` / value → ensure the property is present (add/update w:rPr child)
 * - `false` / `null` → remove the property from w:rPr
 * - Property not mentioned → leave unchanged
 */

import type {
  NormalizedParagraph,
  ApplyFormattingPatch,
  FormattingUpdate,
  FXPNode,
} from "../types.js";
import { findOccurrence, resolveRunRanges } from "../mapping/text-to-runs.js";
import { splitRunsForMatch } from "./run-splitter.js";
import {
  findChild,
  getChildren,
  createElement,
  setAttr,
  removeChild,
} from "../xml/helpers.js";

// =============================================================================
// Public API
// =============================================================================

/**
 * Execute an apply_formatting patch on a paragraph.
 */
export function executeApplyFormatting(
  paragraph: NormalizedParagraph,
  patch: ApplyFormattingPatch,
): void {
  if (patch.find !== undefined && patch.find.length > 0) {
    applyFormattingToSubstring(paragraph, patch.find, patch.occurrence ?? 1, patch.formatting);
  } else {
    applyFormattingToAllRuns(paragraph, patch.formatting);
  }
}

// =============================================================================
// Internal: target selection
// =============================================================================

/**
 * Apply formatting to a substring within a paragraph.
 * Isolates the matched runs via splitting, then applies formatting.
 */
function applyFormattingToSubstring(
  paragraph: NormalizedParagraph,
  find: string,
  occurrence: number,
  formatting: FormattingUpdate,
): void {
  const matchStart = findOccurrence(paragraph.text, find, occurrence);
  if (matchStart === -1) {
    throw new Error(
      `Text "${find}" not found in paragraph "${paragraph.id}" (occurrence ${occurrence})`,
    );
  }

  const matchEnd = matchStart + find.length;
  const ranges = resolveRunRanges(paragraph.runs, matchStart, matchEnd);

  if (ranges.length === 0) {
    throw new Error(
      `No runs found for text range [${matchStart}, ${matchEnd}] in paragraph "${paragraph.id}"`,
    );
  }

  // Split runs to isolate the matched text, then apply formatting to each
  const matchedRunNodes = splitRunsForMatch(paragraph.xmlRef, ranges);
  for (const runNode of matchedRunNodes) {
    applyFormattingToRunNode(runNode, formatting);
  }
}

/**
 * Apply formatting to all runs in a paragraph.
 */
function applyFormattingToAllRuns(
  paragraph: NormalizedParagraph,
  formatting: FormattingUpdate,
): void {
  for (const run of paragraph.runs) {
    applyFormattingToRunNode(run.xmlRef, formatting);
  }
}

// =============================================================================
// Internal: w:rPr manipulation
// =============================================================================

/**
 * Apply a FormattingUpdate to a single w:r run node's w:rPr.
 */
function applyFormattingToRunNode(
  runNode: FXPNode,
  formatting: FormattingUpdate,
): void {
  // Only create w:rPr if we're actually adding properties
  const hasAdds = Object.entries(formatting).some(
    ([, v]) => v !== null && v !== false && v !== undefined,
  );

  let rPr = findChild(runNode, "w:rPr");

  // If we only have removes and there's no rPr, nothing to do
  if (!rPr && !hasAdds) return;

  // Create rPr if needed (must be first child of w:r)
  if (!rPr) {
    rPr = createElement("w:rPr");
    const children = getChildren(runNode);
    children.unshift(rPr);
  }

  // Apply each formatting property
  if (formatting.bold !== undefined) {
    applyBooleanProp(rPr, "w:b", formatting.bold);
  }
  if (formatting.italic !== undefined) {
    applyBooleanProp(rPr, "w:i", formatting.italic);
  }
  if (formatting.strikethrough !== undefined) {
    applyBooleanProp(rPr, "w:strike", formatting.strikethrough);
  }
  if (formatting.underline !== undefined) {
    applyAttrProp(rPr, "w:u", "w:val", formatting.underline);
  }
  if (formatting.fontSize !== undefined) {
    applyAttrProp(rPr, "w:sz", "w:val", formatting.fontSize !== null ? String(formatting.fontSize) : null);
  }
  if (formatting.fontFamily !== undefined) {
    applyFontFamily(rPr, formatting.fontFamily);
  }
  if (formatting.color !== undefined) {
    applyAttrProp(rPr, "w:color", "w:val", formatting.color);
  }
  if (formatting.highlight !== undefined) {
    applyAttrProp(rPr, "w:highlight", "w:val", formatting.highlight);
  }
}

/**
 * Set or remove a boolean property element (e.g., w:b, w:i, w:strike).
 * - `true` → ensure the element exists (no attributes needed)
 * - `false` / `null` → remove the element if present
 */
function applyBooleanProp(
  rPr: FXPNode,
  tagName: string,
  value: boolean | null,
): void {
  const existing = findChild(rPr, tagName);

  if (value === true) {
    if (!existing) {
      getChildren(rPr).push(createElement(tagName));
    }
  } else {
    // false or null → remove
    if (existing) {
      removeChild(rPr, existing);
    }
  }
}

/**
 * Set or remove a property element that carries its value in an attribute.
 * e.g., <w:u w:val="single"/>, <w:sz w:val="24"/>, <w:color w:val="FF0000"/>
 *
 * - non-null string → create/update element with the attribute
 * - `null` → remove the element
 */
function applyAttrProp(
  rPr: FXPNode,
  tagName: string,
  attrName: string,
  value: string | null,
): void {
  const existing = findChild(rPr, tagName);

  if (value !== null) {
    if (existing) {
      setAttr(existing, attrName, value);
    } else {
      const el = createElement(tagName, [], { [attrName]: value });
      getChildren(rPr).push(el);
    }
  } else {
    // null → remove
    if (existing) {
      removeChild(rPr, existing);
    }
  }
}

/**
 * Set or remove font family.
 * In OOXML, font family is set via w:rFonts with multiple attributes:
 *   <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
 *
 * - non-null string → create/update w:rFonts with w:ascii and w:hAnsi
 * - `null` → remove w:rFonts entirely
 */
function applyFontFamily(rPr: FXPNode, value: string | null): void {
  const existing = findChild(rPr, "w:rFonts");

  if (value !== null) {
    if (existing) {
      setAttr(existing, "w:ascii", value);
      setAttr(existing, "w:hAnsi", value);
    } else {
      const el = createElement("w:rFonts", [], {
        "w:ascii": value,
        "w:hAnsi": value,
      });
      getChildren(rPr).push(el);
    }
  } else {
    if (existing) {
      removeChild(rPr, existing);
    }
  }
}
