/**
 * Patch engine: executes patches sequentially against a NormalizedDocument.
 *
 * After each patch:
 * 1. The FXP object is mutated directly (via xmlRef pointers)
 * 2. The affected paragraph is re-normalized to update text + run mappings
 *
 * This is the "Option A: re-normalize after each patch" approach — simple
 * and correct. Can be optimized to incremental updates later.
 */

import type {
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  Patch,
  PatchResult,
  PatchOptions,
  AppendTextPatch,
  FXPNode,
} from "../types.js";
import {
  validatePatch,
  findParagraphById,
  findTableById,
} from "./validator.js";
import {
  executeReplaceText,
  executeInsertText,
  executeDeleteText,
} from "./text-ops.js";
import {
  executeAddTableRow,
  executeDeleteTableRow,
  executeReplaceCellText,
} from "./table-ops.js";
import { executeApplyFormatting } from "./format-ops.js";
import { executeApplyStyle } from "./style-ops.js";
import { normalizeParagraph } from "../normalizer/paragraph.js";
import { normalizeTable } from "../normalizer/table.js";
import { createRevisionContext, type RevisionContext } from "./revision.js";
import {
  findChild,
  getChildren,
  getTagName,
  createRun,
} from "../xml/helpers.js";
import { createInsElement, createRevisionRun } from "./revision.js";

/**
 * Apply a list of patches to a NormalizedDocument.
 *
 * Patches are applied sequentially in order. After each patch,
 * the affected elements are re-normalized so subsequent patches
 * operate on updated text/mappings.
 *
 * @param options - Optional patch options (e.g., revision tracking)
 * @returns Array of PatchResult for each patch
 */
export function applyPatches(
  doc: NormalizedDocument,
  patches: Patch[],
  options?: PatchOptions,
): PatchResult[] {
  const results: PatchResult[] = [];

  // Create revision context if tracking revisions
  const revCtx = options?.trackRevisions
    ? createRevisionContext(doc, options)
    : undefined;

  for (const patch of patches) {
    const result = applySinglePatch(doc, patch, revCtx);
    results.push(result);
  }

  return results;
}

/**
 * Apply a single patch and re-normalize affected elements.
 */
function applySinglePatch(
  doc: NormalizedDocument,
  patch: Patch,
  revCtx?: RevisionContext,
): PatchResult {
  // Validate
  const validation = validatePatch(patch, doc);
  if (!validation.valid) {
    return { success: false, patch, error: validation.error };
  }

  try {
    switch (patch.type) {
      case "replace_text": {
        const para = findParagraphById(patch.target, doc)!;
        executeReplaceText(para, patch, revCtx);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "insert_text": {
        const para = findParagraphById(patch.target, doc)!;
        executeInsertText(para, patch, revCtx);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "delete_text": {
        const para = findParagraphById(patch.target, doc)!;
        executeDeleteText(para, patch, revCtx);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "replace_cell_text": {
        const para = findParagraphById(patch.target, doc)!;
        executeReplaceCellText(para, patch, revCtx);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "add_table_row": {
        const table = findTableById(patch.target, doc)!;
        executeAddTableRow(table, patch, revCtx);
        reNormalizeTable(doc, table);
        break;
      }
      case "delete_table_row": {
        const table = findTableById(patch.target, doc)!;
        executeDeleteTableRow(table, patch, revCtx);
        reNormalizeTable(doc, table);
        break;
      }
      case "apply_formatting": {
        const para = findParagraphById(patch.target, doc)!;
        executeApplyFormatting(para, patch, revCtx);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "apply_style": {
        const para = findParagraphById(patch.target, doc)!;
        executeApplyStyle(para, patch, doc, revCtx);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "append_text": {
        executeAppendText(doc, patch, revCtx);
        break;
      }
    }

    return { success: true, patch };
  } catch (err) {
    return {
      success: false,
      patch,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

/**
 * Re-normalize a paragraph after a patch mutated its XML.
 * Updates the paragraph's text, runs, and mappings in-place.
 */
function reNormalizeParagraph(
  _doc: NormalizedDocument,
  paragraph: NormalizedParagraph,
): void {
  const fresh = normalizeParagraph(paragraph.xmlRef, paragraph.id);

  // Update in-place
  paragraph.text = fresh.text;
  paragraph.runs = fresh.runs;
  paragraph.style = fresh.style;
  paragraph.listInfo = fresh.listInfo;
  paragraph.revisions = fresh.revisions;
}

/**
 * Re-normalize a table after a structural change (row add/delete).
 * This replaces the table's rows/cells with fresh normalized data.
 */
function reNormalizeTable(
  _doc: NormalizedDocument,
  table: NormalizedTable,
): void {
  const fresh = normalizeTable(table.xmlRef, table.id);

  // Update in-place
  table.rows = fresh.rows;
  table.style = fresh.style;
}

// =============================================================================
// Append text (document-level operation)
// =============================================================================

/**
 * Execute an append_text patch: creates a new w:p element at the end
 * of the document body and adds it to the normalized body array.
 *
 * Inserts before w:sectPr if present so section properties stay last.
 *
 * With revision tracking, the new paragraph's run is wrapped in w:ins.
 */
function executeAppendText(
  doc: NormalizedDocument,
  patch: AppendTextPatch,
  revCtx?: RevisionContext,
): void {
  // Find w:body in the FXP tree
  const docNode = doc.xmlParts.document.find(
    (n) => getTagName(n) === "w:document",
  );
  if (!docNode) {
    throw new Error("Cannot find w:document in FXP tree");
  }

  const bodyNode = findChild(docNode, "w:body");
  if (!bodyNode) {
    throw new Error("Cannot find w:body in FXP tree");
  }

  const bodyChildren = getChildren(bodyNode);

  // Build the new w:p FXP node
  let paragraphChildren: FXPNode[];

  if (revCtx) {
    const run = createRevisionRun(patch.text, undefined);
    const insElement = createInsElement(revCtx, [run]);
    paragraphChildren = [insElement];
  } else {
    const run = createRun(patch.text);
    paragraphChildren = [run];
  }

  const newParagraph: FXPNode = { "w:p": paragraphChildren };

  // Insert before w:sectPr if present, otherwise append at the end
  const sectPrIndex = bodyChildren.findIndex(
    (n) => getTagName(n) === "w:sectPr",
  );
  if (sectPrIndex !== -1) {
    bodyChildren.splice(sectPrIndex, 0, newParagraph);
  } else {
    bodyChildren.push(newParagraph);
  }

  // Compute the next paragraph ID
  let maxPIdx = -1;
  for (const el of doc.body) {
    if (el.kind === "paragraph") {
      const num = parseInt(el.id.replace("p", ""), 10);
      if (!isNaN(num) && num > maxPIdx) maxPIdx = num;
    }
  }
  const newId = `p${maxPIdx + 1}`;

  // Normalize the new paragraph and add it to the body
  const normalized = normalizeParagraph(newParagraph, newId);
  doc.body.push(normalized);
}
