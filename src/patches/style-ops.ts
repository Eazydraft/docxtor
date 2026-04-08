/**
 * Style operations: apply_style patch execution.
 *
 * Re-points paragraphs or runs to pre-existing styles in styles.xml.
 * Never creates, modifies, or deletes style definitions.
 *
 * Two modes:
 * 1. Paragraph style (no `find`): Sets w:pPr > w:pStyle on the paragraph.
 *    The style must be type "paragraph" in styles.xml.
 *
 * 2. Character style (with `find`): Isolates matched runs via splitting,
 *    then sets w:rPr > w:rStyle on each matched run.
 *    The style must be type "character" in styles.xml.
 */

import type {
  NormalizedParagraph,
  NormalizedDocument,
  ApplyStylePatch,
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
  cloneNode,
} from "../xml/helpers.js";
import type { RevisionContext } from "./revision.js";
import { createPPrChange, createRPrChange } from "./revision.js";

// =============================================================================
// Public API
// =============================================================================

/**
 * Execute an apply_style patch on a paragraph.
 *
 * Validates that the style exists and is the correct type before applying.
 */
export function executeApplyStyle(
  paragraph: NormalizedParagraph,
  patch: ApplyStylePatch,
  doc: NormalizedDocument,
  revCtx?: RevisionContext,
): void {
  const styleInfo = doc.styles.find((s) => s.id === patch.style);
  if (!styleInfo) {
    throw new Error(
      `Style "${patch.style}" not found in document styles.xml`,
    );
  }

  if (patch.find !== undefined && patch.find.length > 0) {
    // Character style mode
    if (styleInfo.type !== "character") {
      throw new Error(
        `Style "${patch.style}" is type "${styleInfo.type}", but character style is required when "find" is specified`,
      );
    }
    applyCharacterStyle(paragraph, patch.find, patch.occurrence ?? 1, patch.style, revCtx);
  } else {
    // Paragraph style mode
    if (styleInfo.type !== "paragraph") {
      throw new Error(
        `Style "${patch.style}" is type "${styleInfo.type}", but paragraph style is required when "find" is omitted`,
      );
    }
    applyParagraphStyle(paragraph, patch.style, revCtx);
  }
}

// =============================================================================
// Internal: paragraph style
// =============================================================================

/**
 * Apply a paragraph style by setting w:pPr > w:pStyle.
 *
 * If w:pPr doesn't exist, it's created as the first child of w:p.
 * If w:pStyle already exists, its w:val is updated.
 * If w:pStyle doesn't exist, it's created as the first child of w:pPr.
 *
 * When revCtx is provided, adds w:pPrChange with the old pPr snapshot.
 */
function applyParagraphStyle(
  paragraph: NormalizedParagraph,
  styleId: string,
  revCtx?: RevisionContext,
): void {
  const paraNode = paragraph.xmlRef;

  // Snapshot old pPr before modification (for revision tracking)
  let pPr = findChild(paraNode, "w:pPr");
  const oldPPr = revCtx && pPr ? cloneNode(pPr) : undefined;

  // Ensure w:pPr exists (must be first child of w:p)
  if (!pPr) {
    pPr = createElement("w:pPr");
    const children = getChildren(paraNode);
    children.unshift(pPr);
  }

  // Set or update w:pStyle
  let pStyle = findChild(pPr, "w:pStyle");
  if (pStyle) {
    setAttr(pStyle, "w:val", styleId);
  } else {
    pStyle = createElement("w:pStyle", [], { "w:val": styleId });
    // w:pStyle should be first child of w:pPr
    getChildren(pPr).unshift(pStyle);
  }

  // Add w:pPrChange with the old snapshot (revision tracking)
  if (revCtx) {
    // Remove any existing w:pPrChange first
    const existingChange = findChild(pPr, "w:pPrChange");
    if (existingChange) {
      removeChild(pPr, existingChange);
    }

    const oldPPrForChange = oldPPr ?? createElement("w:pPr");
    const pPrChange = createPPrChange(revCtx, oldPPrForChange);
    getChildren(pPr).push(pPrChange);
  }
}

// =============================================================================
// Internal: character style
// =============================================================================

/**
 * Apply a character style to a substring by setting w:rPr > w:rStyle
 * on each matched run.
 *
 * Uses the same find/split machinery as formatting patches.
 * When revCtx is provided, adds w:rPrChange with the old rPr snapshot.
 */
function applyCharacterStyle(
  paragraph: NormalizedParagraph,
  find: string,
  occurrence: number,
  styleId: string,
  revCtx?: RevisionContext,
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

  // Split runs to isolate the matched text
  const matchedRunNodes = splitRunsForMatch(paragraph.xmlRef, ranges);

  // Apply w:rStyle to each matched run
  for (const runNode of matchedRunNodes) {
    applyRStyleToRunNode(runNode, styleId, revCtx);
  }
}

/**
 * Set w:rPr > w:rStyle on a single run node.
 *
 * Creates w:rPr if needed. If w:rStyle already exists, updates it.
 * Otherwise creates it as the first child of w:rPr.
 * When revCtx is provided, adds w:rPrChange with the old rPr snapshot.
 */
function applyRStyleToRunNode(
  runNode: FXPNode,
  styleId: string,
  revCtx?: RevisionContext,
): void {
  // Snapshot old rPr before modification (for revision tracking)
  let rPr = findChild(runNode, "w:rPr");
  const oldRPr = revCtx && rPr ? cloneNode(rPr) : undefined;

  // Ensure w:rPr exists (must be first child of w:r)
  if (!rPr) {
    rPr = createElement("w:rPr");
    const children = getChildren(runNode);
    children.unshift(rPr);
  }

  // Remove any existing w:rStyle first
  const existingRStyle = findChild(rPr, "w:rStyle");
  if (existingRStyle) {
    removeChild(rPr, existingRStyle);
  }

  // Create new w:rStyle as first child of w:rPr
  const rStyle = createElement("w:rStyle", [], { "w:val": styleId });
  getChildren(rPr).unshift(rStyle);

  // Add w:rPrChange with the old snapshot (revision tracking)
  if (revCtx) {
    // Remove any existing w:rPrChange first
    const existingChange = findChild(rPr, "w:rPrChange");
    if (existingChange) {
      removeChild(rPr, existingChange);
    }

    const oldRPrForChange = oldRPr ?? createElement("w:rPr");
    const rPrChange = createRPrChange(revCtx, oldRPrForChange);
    getChildren(rPr).push(rPrChange);
  }
}
