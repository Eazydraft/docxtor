/**
 * Table normalizer: walks a w:tbl FXP node and produces a
 * NormalizedTable with rows, cells, and inner paragraphs.
 *
 * Tables are containers — their text lives in paragraphs inside cells.
 * We normalize the structure and delegate paragraph normalization to
 * the paragraph normalizer.
 *
 * Supports nested tables: if a table cell contains another w:tbl,
 * it is recursively normalized and included in the cell's nestedTables.
 */

import type {
  FXPNode,
  NormalizedTable,
  NormalizedTableRow,
  NormalizedTableCell,
  NormalizedParagraph,
  ControlInfo,
} from "../types.js";
import {
  findChild,
  contentChildren,
  getAttr,
} from "../xml/helpers.js";
import { normalizeParagraph } from "./paragraph.js";
import { ControlResolver } from "./control.js";

/**
 * Normalize a w:tbl FXP node into a NormalizedTable.
 *
 * @param tableNode - The w:tbl FXP node
 * @param id - The stable ID (e.g., "tbl0")
 * @param control - The innermost content control wrapped around the table
 * @param resolver - Shared across one part so a control is read once
 */
export function normalizeTable(
  tableNode: FXPNode,
  id: string,
  control?: ControlInfo,
  resolver: ControlResolver = new ControlResolver(),
): NormalizedTable {
  const rows: NormalizedTableRow[] = [];
  // Rows, cells and cell content are read through content controls: a
  // wrapped cell is still a cell, a wrapped paragraph still a paragraph.
  const trNodes = contentChildren(tableNode).filter((c) => "w:tr" in c.node);

  for (let rowIdx = 0; rowIdx < trNodes.length; rowIdx++) {
    const tr = trNodes[rowIdx]!;
    const rowControl = resolver.resolve(tr.sdts, control);
    const cells: NormalizedTableCell[] = [];
    const tcNodes = contentChildren(tr.node).filter((c) => "w:tc" in c.node);

    for (let colIdx = 0; colIdx < tcNodes.length; colIdx++) {
      const tc = tcNodes[colIdx]!;
      const cellControl = resolver.resolve(tc.sdts, rowControl);
      const normalizedParas: NormalizedParagraph[] = [];
      const nestedTables: NormalizedTable[] = [];

      for (const item of contentChildren(tc.node)) {
        const inner = resolver.resolve(item.sdts, cellControl);
        if ("w:p" in item.node) {
          normalizedParas.push(
            normalizeParagraph(
              item.node,
              `${id}.r${rowIdx}.c${colIdx}.p${normalizedParas.length}`,
              inner,
            ),
          );
        } else if ("w:tbl" in item.node) {
          // Nested tables within this cell
          nestedTables.push(
            normalizeTable(
              item.node,
              `${id}.r${rowIdx}.c${colIdx}.tbl${nestedTables.length}`,
              inner,
              resolver,
            ),
          );
        }
      }

      const cell: NormalizedTableCell = {
        rowIndex: rowIdx,
        colIndex: colIdx,
        paragraphs: normalizedParas,
        nestedTables,
        xmlRef: tc.node,
      };
      if (tc.sdts.length > 0 && cellControl) cell.control = cellControl;
      cells.push(cell);
    }

    const row: NormalizedTableRow = {
      index: rowIdx,
      cells,
      xmlRef: tr.node,
    };
    if (tr.sdts.length > 0 && rowControl) row.control = rowControl;
    rows.push(row);
  }

  // Extract table style
  const tblPr = findChild(tableNode, "w:tblPr");
  const tblStyle = tblPr ? findChild(tblPr, "w:tblStyle") : undefined;
  const style = tblStyle ? getAttr(tblStyle, "w:val") : undefined;

  const table: NormalizedTable = {
    kind: "table",
    id,
    rows,
    xmlRef: tableNode,
    style,
  };
  if (control) table.control = control;
  return table;
}
