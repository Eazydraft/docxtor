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
} from "../types.js";
import {
  findChild,
  findChildren,
  getAttr,
} from "../xml/helpers.js";
import { normalizeParagraph } from "./paragraph.js";

/**
 * Normalize a w:tbl FXP node into a NormalizedTable.
 *
 * @param tableNode - The w:tbl FXP node
 * @param id - The stable ID (e.g., "tbl0")
 */
export function normalizeTable(
  tableNode: FXPNode,
  id: string,
): NormalizedTable {
  const rows: NormalizedTableRow[] = [];
  const trNodes = findChildren(tableNode, "w:tr");

  for (let rowIdx = 0; rowIdx < trNodes.length; rowIdx++) {
    const trNode = trNodes[rowIdx]!;
    const cells: NormalizedTableCell[] = [];
    const tcNodes = findChildren(trNode, "w:tc");

    for (let colIdx = 0; colIdx < tcNodes.length; colIdx++) {
      const tcNode = tcNodes[colIdx]!;
      const paragraphs = findChildren(tcNode, "w:p");
      const normalizedParas = paragraphs.map((pNode, pIdx) =>
        normalizeParagraph(
          pNode,
          `${id}.r${rowIdx}.c${colIdx}.p${pIdx}`,
        ),
      );

      // Normalize nested tables within this cell
      const nestedTblNodes = findChildren(tcNode, "w:tbl");
      const nestedTables = nestedTblNodes.map((tblNode, tblIdx) =>
        normalizeTable(
          tblNode,
          `${id}.r${rowIdx}.c${colIdx}.tbl${tblIdx}`,
        ),
      );

      cells.push({
        rowIndex: rowIdx,
        colIndex: colIdx,
        paragraphs: normalizedParas,
        nestedTables,
        xmlRef: tcNode,
      });
    }

    rows.push({
      index: rowIdx,
      cells,
      xmlRef: trNode,
    });
  }

  // Extract table style
  const tblPr = findChild(tableNode, "w:tblPr");
  const tblStyle = tblPr ? findChild(tblPr, "w:tblStyle") : undefined;
  const style = tblStyle ? getAttr(tblStyle, "w:val") : undefined;

  return {
    kind: "table",
    id,
    rows,
    xmlRef: tableNode,
    style,
  };
}
