/**
 * Table operations: add_row, delete_row, replace_cell_text.
 *
 * Table edits work by:
 * - add_row: clone an existing row (preserving cell styles), replace text
 * - delete_row: remove the w:tr node from the w:tbl
 * - replace_cell_text: delegate to text-ops on the paragraph inside the cell
 */

import type {
  NormalizedTable,
  NormalizedParagraph,
  AddTableRowPatch,
  DeleteTableRowPatch,
  ReplaceCellTextPatch,
} from "../types.js";
import {
  cloneNode,
  findChildren,
  getChildren,
  insertAfter,
  removeChild,
  findChild,
  createElement,
  escapeXmlText,
  needsSpacePreserve,
} from "../xml/helpers.js";
import { executeReplaceText } from "./text-ops.js";
import type { RevisionContext } from "./revision.js";
import { createTrPrIns, createTrPrDel } from "./revision.js";

/**
 * Execute an add_table_row patch.
 *
 * Clones an existing row to preserve styling, then replaces cell text.
 */
export function executeAddTableRow(
  table: NormalizedTable,
  patch: AddTableRowPatch,
  revCtx?: RevisionContext,
): void {
  if (table.rows.length === 0) {
    throw new Error(`Table "${table.id}" has no rows to clone from`);
  }

  if (patch.afterRow < -1 || patch.afterRow >= table.rows.length) {
    throw new Error(
      `Invalid afterRow ${patch.afterRow} for table "${table.id}" with ${table.rows.length} rows`,
    );
  }

  // Pick a row to clone (the row we're inserting after, or the first row)
  const sourceRowIdx = patch.afterRow >= 0 ? patch.afterRow : 0;
  const sourceRow = table.rows[sourceRowIdx]!;

  // Verify cell count matches
  if (patch.cells.length !== sourceRow.cells.length) {
    throw new Error(
      `Cell count mismatch: patch has ${patch.cells.length} cells, ` +
        `but row has ${sourceRow.cells.length} cells`,
    );
  }

  // Clone the source row
  const newRowNode = cloneNode(sourceRow.xmlRef);

  // Replace text in each cell of the cloned row
  const tcNodes = findChildren(newRowNode, "w:tc");
  for (let i = 0; i < tcNodes.length; i++) {
    const tcNode = tcNodes[i]!;
    const cellText = patch.cells[i]!;

    // Find the first paragraph in the cell
    const pNode = findChild(tcNode, "w:p");
    if (pNode) {
      // Clear existing runs and add new text
      const pChildren = getChildren(pNode);

      // Remove all existing w:r runs
      const runs = pChildren.filter((c) => "w:r" in c);
      for (const run of runs) {
        const idx = pChildren.indexOf(run);
        if (idx !== -1) pChildren.splice(idx, 1);
      }

      // Add a new run with the cell text
      const wtNode: Record<string, unknown> = {
        "w:t": [{ "#text": escapeXmlText(cellText) } as any],
      };
      if (needsSpacePreserve(cellText)) {
        if (!wtNode[":@"]) wtNode[":@"] = {};
        (wtNode[":@"] as Record<string, string>)["@_xml:space"] = "preserve";
      }
      const newRun = {
        "w:r": [wtNode as any],
      };
      pChildren.push(newRun);
    }
  }

  // Insert the new row into the table XML
  const tableChildren = getChildren(table.xmlRef);
  if (patch.afterRow === -1) {
    // Insert at the beginning, after tblPr and tblGrid
    const firstTr = tableChildren.find((c) => "w:tr" in c);
    if (firstTr) {
      const idx = tableChildren.indexOf(firstTr);
      tableChildren.splice(idx, 0, newRowNode);
    } else {
      tableChildren.push(newRowNode);
    }
  } else {
    insertAfter(table.xmlRef, sourceRow.xmlRef, newRowNode);
  }

  // If tracking revisions, add w:ins to the new row's w:trPr
  if (revCtx) {
    const rowChildren = getChildren(newRowNode);
    let trPr = rowChildren.find((c) => "w:trPr" in c);
    if (!trPr) {
      // Create w:trPr and insert at the beginning
      trPr = createElement("w:trPr", []);
      rowChildren.unshift(trPr);
    }
    const trPrChildren = getChildren(trPr);
    trPrChildren.push(createTrPrIns(revCtx));
  }
}

/**
 * Execute a delete_table_row patch.
 */
export function executeDeleteTableRow(
  table: NormalizedTable,
  patch: DeleteTableRowPatch,
  revCtx?: RevisionContext,
): void {
  if (patch.row < 0 || patch.row >= table.rows.length) {
    throw new Error(
      `Invalid row index ${patch.row} for table "${table.id}" with ${table.rows.length} rows`,
    );
  }

  const row = table.rows[patch.row]!;

  if (revCtx) {
    // With revision tracking: keep the row, add w:del to w:trPr
    const rowChildren = getChildren(row.xmlRef);
    let trPr = rowChildren.find((c) => "w:trPr" in c);
    if (!trPr) {
      trPr = createElement("w:trPr", []);
      rowChildren.unshift(trPr);
    }
    const trPrChildren = getChildren(trPr);
    trPrChildren.push(createTrPrDel(revCtx));
  } else {
    // Without revision tracking: remove the row from XML
    removeChild(table.xmlRef, row.xmlRef);
  }
}

/**
 * Execute a replace_cell_text patch.
 *
 * This delegates to the text replace operation on the targeted paragraph.
 */
export function executeReplaceCellText(
  paragraph: NormalizedParagraph,
  patch: ReplaceCellTextPatch,
  revCtx?: RevisionContext,
): void {
  executeReplaceText(paragraph, {
    type: "replace_text",
    target: patch.target,
    find: patch.find,
    replace: patch.replace,
    occurrence: patch.occurrence,
  }, revCtx);
}
