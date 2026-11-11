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
  FXPNode,
} from "../types.js";
import {
  cloneNode,
  findChildren,
  getChildren,
  findChild,
  contentChildren,
  sdtContentChildren,
  createElement,
  createRun,
  getTagName,
  RUN_CONTAINER_TAGS,
} from "../xml/helpers.js";
import { readControlInfo } from "../normalizer/control.js";
import { executeReplaceText } from "./text-ops.js";
import type { RevisionContext } from "./revision.js";
import { createTrPrIns, createTrPrDel } from "./revision.js";

/**
 * Execute an add_table_row patch.
 *
 * Clones an existing row to preserve styling, then replaces cell text.
 *
 * @returns the new w:tr node (so the caller can re-stamp paragraph ids
 *          that were duplicated by cloning)
 */
export function executeAddTableRow(
  table: NormalizedTable,
  patch: AddTableRowPatch,
  revCtx?: RevisionContext,
): FXPNode {
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

  // Clone the source row. A content control in the row (around a cell, a
  // paragraph or some runs) is not cloned with it: the copy keeps what the
  // control held and drops the wrapper, so no control id appears twice.
  const newRowNode = cloneNode(sourceRow.xmlRef);
  unwrapControls(newRowNode);

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

      // Remove all existing w:r runs (including those wrapped in w:ins from tracked changes)
      const toRemove = pChildren.filter(
        (c) => "w:r" in c || "w:ins" in c || "w:del" in c,
      );
      for (const node of toRemove) {
        const idx = pChildren.indexOf(node);
        if (idx !== -1) pChildren.splice(idx, 1);
      }

      // Add a new run with the cell text
      pChildren.push(createRun(cellText));
    }
  }

  // Insert the new row into the table XML, next to the source row in the
  // array that really holds it (a row can sit inside a content control).
  const rows = contentChildren(table.xmlRef).filter((c) => "w:tr" in c.node);
  if (patch.afterRow === -1) {
    // Insert at the beginning, after tblPr and tblGrid
    const first = rows[0];
    if (first) {
      first.siblings.splice(first.siblings.indexOf(first.node), 0, newRowNode);
    } else {
      getChildren(table.xmlRef).push(newRowNode);
    }
  } else {
    const source = rows.find((c) => c.node === sourceRow.xmlRef);
    if (!source) throw new Error("Reference child not found in parent");
    source.siblings.splice(source.siblings.indexOf(source.node) + 1, 0, newRowNode);
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

  return newRowNode;
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
    const holder = contentChildren(table.xmlRef).find((c) => c.node === row.xmlRef);
    if (holder) {
      holder.siblings.splice(holder.siblings.indexOf(row.xmlRef), 1);
      // A control that wrapped only this row would be left empty in the
      // table: it goes with the row (validation refuses a locked one).
      removeEmptyWrappers(getChildren(table.xmlRef), holder.sdts);
    }
  }
}

/**
 * Remove the wrappers in `sdts` (outermost first) that no longer hold any
 * element, innermost first, from the tree under `children`.
 */
function removeEmptyWrappers(children: FXPNode[], sdts: FXPNode[]): void {
  for (let i = sdts.length - 1; i >= 0; i--) {
    const sdt = sdts[i]!;
    if (sdtContentChildren(sdt).some((c) => !("#text" in c))) return;
    const parent = i === 0 ? children : sdtContentChildren(sdts[i - 1]!);
    const at = parent.indexOf(sdt);
    if (at === -1) return;
    parent.splice(at, 1);
  }
}

/**
 * Replace every content control under `node` by its content, in place.
 * Used on a cloned row: the clone must not repeat the source's control ids.
 * A control showing its prompt contributes structure only (its cell, its
 * paragraphs): the prompt is not document text, so it is not copied.
 */
function unwrapControls(node: FXPNode): void {
  const children = getChildren(node);
  for (let i = 0; i < children.length; i++) {
    const child = children[i]!;
    if ("w:sdt" in child) {
      const content = sdtContentChildren(child);
      if (readControlInfo(child).placeholder) removeRuns(content);
      children.splice(i, 1, ...content);
      i--; // re-examine what took its place (a nested control, say)
      continue;
    }
    if (!("#text" in child)) unwrapControls(child);
  }
}

/** Remove every run under `children`, keeping paragraphs, cells and their properties. */
function removeRuns(children: FXPNode[]): void {
  for (let i = children.length - 1; i >= 0; i--) {
    const child = children[i]!;
    const tag = getTagName(child);
    if (tag === undefined) continue;
    if (RUN_CONTAINER_TAGS.has(tag)) {
      children.splice(i, 1);
    } else if (tag === "w:sdt") {
      removeRuns(sdtContentChildren(child));
    } else if (!tag.endsWith("Pr")) {
      removeRuns(getChildren(child));
    }
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
  markers: boolean = true,
): void {
  executeReplaceText(paragraph, {
    type: "replace_text",
    target: patch.target,
    find: patch.find,
    replace: patch.replace,
    occurrence: patch.occurrence,
  }, revCtx, markers);
}

// =============================================================================
// Build a w:tbl FXP node from scratch
// =============================================================================

/**
 * Build a complete w:tbl FXP node with TableGrid style and single borders.
 *
 * @param headers - Optional header row cells
 * @param rows - Data rows (each sub-array is one row of cell values)
 * @returns A w:tbl FXP node ready for insertion into w:body
 */
export function buildTableNode(
  headers: string[] | undefined,
  rows: string[][],
): FXPNode {
  const colCount =
    headers?.length ?? (rows.length > 0 ? rows[0]!.length : 0);

  // w:tblPr — table properties with TableGrid style and borders
  const borderAttrs = { "w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto" };
  const tblBorders = createElement("w:tblBorders", [
    createElement("w:top", [], borderAttrs),
    createElement("w:left", [], borderAttrs),
    createElement("w:bottom", [], borderAttrs),
    createElement("w:right", [], borderAttrs),
    createElement("w:insideH", [], borderAttrs),
    createElement("w:insideV", [], borderAttrs),
  ]);

  const tblPr = createElement("w:tblPr", [
    createElement("w:tblStyle", [], { "w:val": "TableGrid" }),
    createElement("w:tblW", [], { "w:w": "0", "w:type": "auto" }),
    tblBorders,
    createElement("w:tblLook", [], {
      "w:val": "04A0",
      "w:firstRow": "1",
      "w:lastRow": "0",
      "w:firstColumn": "1",
      "w:lastColumn": "0",
      "w:noHBand": "0",
      "w:noVBand": "1",
    }),
  ]);

  // w:tblGrid — one gridCol per column
  const gridCols: FXPNode[] = [];
  for (let i = 0; i < colCount; i++) {
    gridCols.push(createElement("w:gridCol", []));
  }
  const tblGrid = createElement("w:tblGrid", gridCols);

  // Build rows
  const allRows: FXPNode[] = [];

  if (headers) {
    allRows.push(buildTableRowNode(headers));
  }

  for (const row of rows) {
    allRows.push(buildTableRowNode(row));
  }

  return createElement("w:tbl", [tblPr, tblGrid, ...allRows]);
}

/**
 * Build a single w:tr node from cell text values.
 */
function buildTableRowNode(cells: string[]): FXPNode {
  const tcNodes: FXPNode[] = [];
  for (const cellText of cells) {
    const run = createRun(cellText);
    const para: FXPNode = { "w:p": [run] };
    const tc = createElement("w:tc", [
      createElement("w:tcPr", [
        createElement("w:tcW", [], { "w:w": "0", "w:type": "auto" }),
      ]),
      para,
    ]);
    tcNodes.push(tc);
  }
  return createElement("w:tr", tcNodes);
}
