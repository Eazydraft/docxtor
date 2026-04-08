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

/**
 * Apply a list of patches to a NormalizedDocument.
 *
 * Patches are applied sequentially in order. After each patch,
 * the affected elements are re-normalized so subsequent patches
 * operate on updated text/mappings.
 *
 * @returns Array of PatchResult for each patch
 */
export function applyPatches(
  doc: NormalizedDocument,
  patches: Patch[],
): PatchResult[] {
  const results: PatchResult[] = [];

  for (const patch of patches) {
    const result = applySinglePatch(doc, patch);
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
        executeReplaceText(para, patch);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "insert_text": {
        const para = findParagraphById(patch.target, doc)!;
        executeInsertText(para, patch);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "delete_text": {
        const para = findParagraphById(patch.target, doc)!;
        executeDeleteText(para, patch);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "replace_cell_text": {
        const para = findParagraphById(patch.target, doc)!;
        executeReplaceCellText(para, patch);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "add_table_row": {
        const table = findTableById(patch.target, doc)!;
        executeAddTableRow(table, patch);
        reNormalizeTable(doc, table);
        break;
      }
      case "delete_table_row": {
        const table = findTableById(patch.target, doc)!;
        executeDeleteTableRow(table, patch);
        reNormalizeTable(doc, table);
        break;
      }
      case "apply_formatting": {
        const para = findParagraphById(patch.target, doc)!;
        executeApplyFormatting(para, patch);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "apply_style": {
        const para = findParagraphById(patch.target, doc)!;
        executeApplyStyle(para, patch, doc);
        reNormalizeParagraph(doc, para);
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
