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
  AddTablePatch,
  FXPNode,
} from "../types.js";
import {
  validatePatch,
  findParagraphById,
  findTableById,
  findElementById,
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
  buildTableNode,
} from "./table-ops.js";
import { executeApplyFormatting } from "./format-ops.js";
import { executeApplyStyle } from "./style-ops.js";
import { normalizeParagraph } from "../normalizer/paragraph.js";
import { normalizeTable } from "../normalizer/table.js";
import { createRevisionContext, type RevisionContext } from "./revision.js";
import {
  findChild,
  findChildren,
  getChildren,
  getTagName,
  createElement,
} from "../xml/helpers.js";
import { createInsElement, createTrPrIns } from "./revision.js";
import {
  executeAddHeader,
  executeAddFooter,
  executeRemoveHeader,
  executeRemoveFooter,
} from "./header-footer-ops.js";
import { buildFieldRuns } from "./field-codes.js";

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
      case "add_table": {
        executeAddTablePatch(doc, patch, revCtx);
        break;
      }
      case "add_header": {
        executeAddHeader(doc, patch, revCtx);
        break;
      }
      case "add_footer": {
        executeAddFooter(doc, patch, revCtx);
        break;
      }
      case "remove_header": {
        executeRemoveHeader(doc, patch);
        break;
      }
      case "remove_footer": {
        executeRemoveFooter(doc, patch);
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
 * Get the w:body node and its children from the document FXP tree.
 */
function getBodyNode(doc: NormalizedDocument): {
  bodyNode: FXPNode;
  bodyChildren: FXPNode[];
} {
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
  return { bodyNode, bodyChildren: getChildren(bodyNode) };
}

/**
 * Insert FXP nodes into the body before w:sectPr, or after a specific element.
 * Returns the insertion index.
 */
function insertIntoBody(
  bodyChildren: FXPNode[],
  nodes: FXPNode[],
  afterXmlRef?: FXPNode,
): void {
  if (afterXmlRef) {
    const idx = bodyChildren.indexOf(afterXmlRef);
    if (idx !== -1) {
      bodyChildren.splice(idx + 1, 0, ...nodes);
      return;
    }
  }
  // Default: insert before w:sectPr if present, otherwise append
  const sectPrIndex = bodyChildren.findIndex(
    (n) => getTagName(n) === "w:sectPr",
  );
  if (sectPrIndex !== -1) {
    bodyChildren.splice(sectPrIndex, 0, ...nodes);
  } else {
    bodyChildren.push(...nodes);
  }
}

/**
 * Compute the next available paragraph and table IDs.
 */
function getNextIds(doc: NormalizedDocument): {
  nextPIdx: number;
  nextTblIdx: number;
} {
  let maxPIdx = -1;
  let maxTblIdx = -1;
  for (const el of doc.body) {
    if (el.kind === "paragraph") {
      const num = parseInt(el.id.replace("p", ""), 10);
      if (!isNaN(num) && num > maxPIdx) maxPIdx = num;
    } else if (el.kind === "table") {
      const num = parseInt(el.id.replace("tbl", ""), 10);
      if (!isNaN(num) && num > maxTblIdx) maxTblIdx = num;
    }
  }
  return { nextPIdx: maxPIdx + 1, nextTblIdx: maxTblIdx + 1 };
}

// =============================================================================
// Markdown table parsing
// =============================================================================

/** A contiguous block of text lines or a markdown table. */
type TextSegment =
  | { kind: "text"; lines: string[] }
  | { kind: "table"; headers: string[]; rows: string[][] };

/**
 * Parse text into segments of plain lines and markdown tables.
 *
 * A markdown table is detected as:
 * - A header row: `| col1 | col2 | ...`
 * - A separator row: `| --- | --- | ...` (dashes, colons, spaces)
 * - Zero or more data rows: `| val1 | val2 | ...`
 */
function parseTextSegments(text: string): TextSegment[] {
  const lines = text.split("\n");
  const segments: TextSegment[] = [];
  let i = 0;

  while (i < lines.length) {
    // Try to detect a markdown table starting at line i
    if (
      i + 1 < lines.length &&
      isMarkdownTableRow(lines[i]!) &&
      isMarkdownSeparator(lines[i + 1]!)
    ) {
      const headerCells = parseMarkdownRow(lines[i]!);
      // Skip header and separator
      i += 2;
      const dataRows: string[][] = [];
      while (i < lines.length && isMarkdownTableRow(lines[i]!)) {
        const cells = parseMarkdownRow(lines[i]!);
        // Pad or trim to match header count
        while (cells.length < headerCells.length) cells.push("");
        if (cells.length > headerCells.length)
          cells.length = headerCells.length;
        dataRows.push(cells);
        i++;
      }
      segments.push({ kind: "table", headers: headerCells, rows: dataRows });
    } else {
      // Plain text line
      const last = segments[segments.length - 1];
      if (last && last.kind === "text") {
        last.lines.push(lines[i]!);
      } else {
        segments.push({ kind: "text", lines: [lines[i]!] });
      }
      i++;
    }
  }

  return segments;
}

function isMarkdownTableRow(line: string): boolean {
  const trimmed = line.trim();
  return trimmed.startsWith("|") && trimmed.endsWith("|") && trimmed.length > 1;
}

function isMarkdownSeparator(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed.startsWith("|") || !trimmed.endsWith("|")) return false;
  // Each cell should be dashes (with optional colons for alignment)
  const cells = trimmed.slice(1, -1).split("|");
  return cells.every((cell) => /^\s*:?-+:?\s*$/.test(cell));
}

function parseMarkdownRow(line: string): string[] {
  const trimmed = line.trim();
  // Remove leading/trailing pipes and split
  const inner = trimmed.startsWith("|") ? trimmed.slice(1) : trimmed;
  const content = inner.endsWith("|") ? inner.slice(0, -1) : inner;
  return content.split("|").map((cell) => cell.trim());
}

// =============================================================================
// Revision tracking for new tables
// =============================================================================

/**
 * Mark all rows in a table node as inserted revisions.
 * Adds w:ins to w:trPr for each w:tr in the table.
 */
function markTableRowsAsInserted(
  tableNode: FXPNode,
  revCtx: RevisionContext,
): void {
  const trNodes = findChildren(tableNode, "w:tr");
  for (const trNode of trNodes) {
    const rowChildren = getChildren(trNode);
    let trPr = rowChildren.find((c) => "w:trPr" in c);
    if (!trPr) {
      trPr = createElement("w:trPr", []);
      rowChildren.unshift(trPr);
    }
    const trPrChildren = getChildren(trPr);
    trPrChildren.push(createTrPrIns(revCtx));
  }
}

// =============================================================================
// Execute append_text
// =============================================================================

/**
 * Execute an append_text patch: creates new w:p and/or w:tbl elements at the
 * end of the document body (or a header/footer if container is specified).
 *
 * Supports markdown table syntax — pipe-delimited rows with a separator line
 * are automatically converted to OOXML tables.
 * Supports {{FIELD}} placeholders (e.g., {{PAGE}}, {{NUMPAGES}}).
 *
 * Inserts before w:sectPr if present so section properties stay last.
 * With revision tracking, new paragraph runs are wrapped in w:ins.
 */
function executeAppendText(
  doc: NormalizedDocument,
  patch: AppendTextPatch,
  revCtx?: RevisionContext,
): void {
  if (patch.container) {
    // Target a header/footer container
    executeAppendTextToContainer(doc, patch, revCtx);
    return;
  }

  const { bodyChildren } = getBodyNode(doc);
  const segments = parseTextSegments(patch.text);
  const newNodes: FXPNode[] = [];

  for (const segment of segments) {
    if (segment.kind === "text") {
      for (const line of segment.lines) {
        const runs = buildFieldRuns(line);
        if (revCtx) {
          const insElement = createInsElement(revCtx, runs);
          newNodes.push({ "w:p": [insElement] });
        } else {
          newNodes.push({ "w:p": runs });
        }
      }
    } else {
      // Table segment
      const tblNode = buildTableNode(segment.headers, segment.rows);
      if (revCtx) {
        markTableRowsAsInserted(tblNode, revCtx);
      }
      newNodes.push(tblNode);
    }
  }

  // Insert all new nodes before w:sectPr
  insertIntoBody(bodyChildren, newNodes);

  // Normalize and add to doc.body
  let { nextPIdx, nextTblIdx } = getNextIds(doc);
  for (const node of newNodes) {
    if ("w:tbl" in node) {
      const id = `tbl${nextTblIdx++}`;
      const normalized = normalizeTable(node, id);
      doc.body.push(normalized);
    } else {
      const id = `p${nextPIdx++}`;
      const normalized = normalizeParagraph(node, id);
      doc.body.push(normalized);
    }
  }
}

// =============================================================================
// Execute add_table
// =============================================================================

/**
 * Execute an add_table patch: creates a new w:tbl element and inserts it
 * into the document body.
 */
function executeAddTablePatch(
  doc: NormalizedDocument,
  patch: AddTablePatch,
  revCtx?: RevisionContext,
): void {
  if (patch.container) {
    if (patch.afterId) {
      throw new Error(`"afterId" is not supported when "container" is set`);
    }
    executeAddTableToContainer(doc, patch, revCtx);
    return;
  }

  const { bodyChildren } = getBodyNode(doc);
  const tableNode = buildTableNode(patch.headers, patch.rows);

  if (revCtx) {
    markTableRowsAsInserted(tableNode, revCtx);
  }

  // Find the afterId target's xmlRef if specified
  let afterXmlRef: FXPNode | undefined;
  if (patch.afterId) {
    const target = findElementById(patch.afterId, doc);
    if (target) {
      afterXmlRef = target.xmlRef;
    }
  }

  insertIntoBody(bodyChildren, [tableNode], afterXmlRef);

  // Normalize and add to doc.body
  const { nextTblIdx } = getNextIds(doc);
  const id = `tbl${nextTblIdx}`;
  const normalized = normalizeTable(tableNode, id);

  // Insert into body array at correct position (after the target element)
  if (patch.afterId) {
    const targetIdx = doc.body.findIndex((el) => el.id === patch.afterId);
    if (targetIdx !== -1) {
      doc.body.splice(targetIdx + 1, 0, normalized);
      return;
    }
  }
  doc.body.push(normalized);
}

// =============================================================================
// Container helpers (header/footer targeting for append_text and add_table)
// =============================================================================

/**
 * Find a header/footer's root FXP node by type and variant.
 */
function findContainerRoot(
  doc: NormalizedDocument,
  containerType: "header" | "footer",
  variant: string,
): { rootNode: FXPNode; hf: import("../types.js").HeaderFooter } | undefined {
  const map = containerType === "header" ? doc.headers : doc.footers;
  for (const hf of map.values()) {
    if (hf.variant === variant) {
      const rootTag = containerType === "header" ? "w:hdr" : "w:ftr";
      const rootNode = hf.xmlRef.find((n) => getTagName(n) === rootTag);
      if (rootNode) return { rootNode, hf };
    }
  }
  return undefined;
}

/**
 * Append text to a header/footer container.
 */
function executeAppendTextToContainer(
  doc: NormalizedDocument,
  patch: AppendTextPatch,
  revCtx?: RevisionContext,
): void {
  const container = patch.container!;
  const result = findContainerRoot(doc, container.type, container.variant);
  if (!result) {
    throw new Error(
      `${container.type} with variant "${container.variant}" not found`,
    );
  }

  const { rootNode, hf } = result;
  const rootChildren = getChildren(rootNode);
  const prefix = container.type === "header" ? "hdr" : "ftr";

  const segments = parseTextSegments(patch.text);
  let nextPIdx = hf.paragraphs.length;

  for (const segment of segments) {
    if (segment.kind === "text") {
      for (const line of segment.lines) {
        let pNode: FXPNode;
        const runs = buildFieldRuns(line);
        if (revCtx) {
          const insElement = createInsElement(revCtx, runs);
          pNode = { "w:p": [insElement] };
        } else {
          pNode = { "w:p": runs };
        }
        rootChildren.push(pNode);
        const normalized = normalizeParagraph(
          pNode,
          `${prefix}.${container.variant}.p${nextPIdx++}`,
        );
        hf.paragraphs.push(normalized);
      }
    } else {
      // Table segment — add to container XML
      const tblNode = buildTableNode(segment.headers, segment.rows);
      if (revCtx) {
        markTableRowsAsInserted(tblNode, revCtx);
      }
      rootChildren.push(tblNode);

      // Normalize and track
      const nextTblIdx = hf.tables.length;
      const tblId = `${prefix}.${container.variant}.tbl${nextTblIdx}`;
      const normalizedTbl = normalizeTable(tblNode, tblId);
      hf.tables.push(normalizedTbl);
    }
  }
}

/**
 * Execute add_table targeting a header/footer container.
 */
function executeAddTableToContainer(
  doc: NormalizedDocument,
  patch: AddTablePatch,
  revCtx?: RevisionContext,
): void {
  const container = patch.container!;
  const result = findContainerRoot(doc, container.type, container.variant);
  if (!result) {
    throw new Error(
      `${container.type} with variant "${container.variant}" not found`,
    );
  }

  const { rootNode, hf } = result;
  const rootChildren = getChildren(rootNode);
  const tableNode = buildTableNode(patch.headers, patch.rows);

  if (revCtx) {
    markTableRowsAsInserted(tableNode, revCtx);
  }

  rootChildren.push(tableNode);

  // Normalize and track
  const prefix = container.type === "header" ? "hdr" : "ftr";
  const nextTblIdx = hf.tables.length;
  const id = `${prefix}.${container.variant}.tbl${nextTblIdx}`;
  const normalized = normalizeTable(tableNode, id);
  hf.tables.push(normalized);
}
