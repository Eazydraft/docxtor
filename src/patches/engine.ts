/**
 * Patch engine: executes patches sequentially against a NormalizedDocument.
 *
 * After each patch:
 * 1. The FXP object is mutated directly (via xmlRef pointers)
 * 2. The affected paragraph is re-normalized to update text + run mappings
 *
 * This is the "Option A: re-normalize after each patch" approach — simple
 * and correct. Can be optimized to incremental updates later.
 *
 * With `atomic: true`, the whole batch runs against a deep clone of the
 * document; the clone replaces the document only if every patch succeeds.
 */

import type {
  NormalizedDocument,
  NormalizedTable,
  Patch,
  PatchResult,
  PatchOptions,
  AppendTextPatch,
  AddTablePatch,
  DocxXmlParts,
  FXPNode,
  ControlInfo,
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
import { enrichListType, reNormalizeDocument } from "../normalizer/document.js";
import { createRevisionContext, revisionTimestamp, type RevisionContext } from "./revision.js";
import { cloneXmlPart } from "../xml/parser.js";
import {
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
import {
  executeAddList,
  executeAddListItem,
  executeRemoveListItem,
  executeConvertToList,
  executeRemoveList,
  executeSetListProperties,
} from "./list-ops.js";
import { createNumberingDefinition } from "./numbering.js";
import { parseTextSegments } from "./text-segments.js";
import {
  getBodyChildren,
  insertIntoBody,
  reNormalizeParagraph,
} from "./container-utils.js";
import { executeInsertParagraph, executeDeleteParagraph } from "./paragraph-ops.js";
import { executeAddComment } from "./comment-ops.js";
import {
  planControlEdit,
  applyControlPlan,
  syncControlValue,
  executeSetCheckbox,
} from "./control-ops.js";
import { stampNewParagraphs, findPartRootForId } from "../normalizer/para-id.js";

/**
 * Apply a list of patches to a NormalizedDocument.
 *
 * Patches are applied sequentially in order. After each patch,
 * the affected elements are re-normalized so subsequent patches
 * operate on updated text/mappings.
 *
 * @param options - Optional patch options (e.g., revision tracking, atomic)
 * @returns Array of PatchResult for each patch
 */
export function applyPatches(
  doc: NormalizedDocument,
  patches: Patch[],
  options?: PatchOptions,
): PatchResult[] {
  if (options?.atomic) {
    return applyPatchesAtomic(doc, patches, options);
  }
  return applyPatchesSequential(doc, patches, options);
}

function applyPatchesSequential(
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
    const result = applySinglePatch(doc, patch, revCtx, options);
    results.push(result);
  }

  return results;
}

/**
 * All-or-nothing: apply every patch to a deep clone; swap the clone in only
 * if every patch succeeded. The results describe every patch either way.
 */
function applyPatchesAtomic(
  doc: NormalizedDocument,
  patches: Patch[],
  options: PatchOptions,
): PatchResult[] {
  const clone = cloneDocument(doc);
  const results = applyPatchesSequential(clone, patches, { ...options, atomic: false });

  if (results.every((r) => r.success)) {
    doc.body = clone.body;
    doc.headers = clone.headers;
    doc.footers = clone.footers;
    doc.xmlParts = clone.xmlParts;
  }

  return results;
}

/**
 * Deep-clone a document's XML parts and rebuild the normalized view over
 * the clone, so no xmlRef points back into the original.
 */
export function cloneDocument(doc: NormalizedDocument): NormalizedDocument {
  const xmlParts = {} as DocxXmlParts;
  for (const [key, value] of Object.entries(doc.xmlParts)) {
    if (value === undefined) continue;
    if (value instanceof Map) {
      const copy = new Map<string, FXPNode[]>();
      for (const [k, v] of value) copy.set(k, cloneXmlPart(v));
      xmlParts[key] = copy;
    } else {
      xmlParts[key] = cloneXmlPart(value);
    }
  }

  const clone: NormalizedDocument = {
    body: [],
    headers: new Map(),
    footers: new Map(),
    xmlParts,
    styles: doc.styles,
    loadOptions: doc.loadOptions,
  };
  reNormalizeDocument(clone);
  return clone;
}

/**
 * Whether inline formatting markers apply to a patch: the patch's own
 * `markers` field wins, then `PatchOptions.inlineMarkers`, then `true`.
 */
function markersFor(patch: Patch, options?: PatchOptions): boolean {
  const own = (patch as { markers?: boolean }).markers;
  return own ?? options?.inlineMarkers ?? true;
}

/**
 * Apply a single patch and re-normalize affected elements.
 */
function applySinglePatch(
  doc: NormalizedDocument,
  patch: Patch,
  revCtx?: RevisionContext,
  options?: PatchOptions,
): PatchResult {
  const markers = markersFor(patch, options);

  // Validate
  const validation = validatePatch(patch, doc, markers);
  if (!validation.valid) {
    return { success: false, patch, error: validation.error };
  }

  try {
    switch (patch.type) {
      case "replace_text":
      case "insert_text":
      case "delete_text":
      case "replace_cell_text": {
        const para = findParagraphById(patch.target, doc)!;
        // A match in a content control may fill it, flip a checkbox, or be
        // an ordinary edit after which the control's stored value follows.
        const plan = planControlEdit(para, patch, markers);
        if (plan && plan.mode !== "edit") {
          applyControlPlan(para, plan, revCtx, markers);
        } else {
          if (patch.type === "replace_text") executeReplaceText(para, patch, revCtx, markers);
          else if (patch.type === "insert_text") executeInsertText(para, patch, revCtx, markers);
          else if (patch.type === "delete_text") executeDeleteText(para, patch, revCtx, markers);
          else executeReplaceCellText(para, patch, revCtx, markers);
          if (plan) syncControlValue(plan.control, plan.text);
        }
        reNormalizeParagraph(doc, para);
        break;
      }
      case "set_checkbox": {
        const para = findParagraphById(patch.target, doc)!;
        executeSetCheckbox(para, patch, revCtx);
        reNormalizeParagraph(doc, para);
        break;
      }
      case "insert_paragraph": {
        const para = findParagraphById(patch.target, doc)!;
        executeInsertParagraph(doc, para, patch, revCtx, markers);
        break;
      }
      case "delete_paragraph": {
        const para = findParagraphById(patch.target, doc)!;
        executeDeleteParagraph(doc, para, patch, revCtx);
        break;
      }
      case "add_comment": {
        const para = findParagraphById(patch.target, doc)!;
        executeAddComment(doc, para, patch, {
          defaultAuthor: options?.revisionAuthor ?? "docxtor",
          date: options?.revisionDate ?? revisionTimestamp(),
          markers,
        });
        reNormalizeParagraph(doc, para);
        break;
      }
      case "add_table_row": {
        const table = findTableById(patch.target, doc)!;
        const newRow = executeAddTableRow(table, patch, revCtx);
        // The cloned row carries the source row's paraIds — give it fresh ones
        stampNewParagraphs(doc, [newRow], findPartRootForId(doc, table.id), true);
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
        executeAppendText(doc, patch, revCtx, markers);
        break;
      }
      case "add_table": {
        executeAddTablePatch(doc, patch, revCtx);
        break;
      }
      case "add_header": {
        executeAddHeader(doc, patch, revCtx, markers);
        break;
      }
      case "add_footer": {
        executeAddFooter(doc, patch, revCtx, markers);
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
      case "add_list": {
        executeAddList(doc, patch, revCtx, markers);
        break;
      }
      case "add_list_item": {
        executeAddListItem(doc, patch, revCtx, markers);
        break;
      }
      case "remove_list_item": {
        executeRemoveListItem(doc, patch, revCtx);
        break;
      }
      case "convert_to_list": {
        executeConvertToList(doc, patch, revCtx);
        break;
      }
      case "remove_list": {
        executeRemoveList(doc, patch, revCtx);
        break;
      }
      case "set_list_properties": {
        executeSetListProperties(doc, patch, revCtx);
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
 * Re-normalize a table after a structural change (row add/delete).
 * This replaces the table's rows/cells with fresh normalized data.
 */
function reNormalizeTable(
  doc: NormalizedDocument,
  table: NormalizedTable,
): void {
  const fresh = normalizeTable(table.xmlRef, table.id, table.control);

  // Update in-place
  table.rows = fresh.rows;
  table.style = fresh.style;

  if (doc.xmlParts.numbering) {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        for (const para of cell.paragraphs) {
          enrichListType(para, doc.xmlParts.numbering);
        }
      }
    }
  }
}

// =============================================================================
// Append text (document-level operation)
// =============================================================================

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
  revCtx: RevisionContext | undefined,
  markers: boolean,
): void {
  if (patch.container) {
    // Target a header/footer container
    executeAppendTextToContainer(doc, patch, revCtx, markers);
    return;
  }

  const bodyChildren = getBodyChildren(doc);
  const segments = parseTextSegments(patch.text);
  const newNodes: FXPNode[] = [];

  // Track list context for numId reuse across table segments
  let lastListNumId: string | undefined;
  let lastListType: "bullet" | "ordered" | undefined;

  for (const segment of segments) {
    if (segment.kind === "text") {
      // A non-blank text segment resets list context
      if (segment.lines.some((l) => l.trim().length > 0)) {
        lastListNumId = undefined;
        lastListType = undefined;
      }
      for (const line of segment.lines) {
        const runs = buildFieldRuns(line, undefined, markers);
        if (revCtx) {
          const insElement = createInsElement(revCtx, runs);
          newNodes.push({ "w:p": [insElement] });
        } else {
          newNodes.push({ "w:p": runs });
        }
      }
    } else if (segment.kind === "table") {
      const tblNode = buildTableNode(segment.headers, segment.rows);
      if (revCtx) {
        markTableRowsAsInserted(tblNode, revCtx);
      }
      newNodes.push(tblNode);
      // Don't reset list context — lists continue across tables
    } else {
      // List segment
      let numId: string;
      if (lastListNumId && lastListType === segment.listType) {
        numId = lastListNumId;
      } else {
        numId = createNumberingDefinition(doc, segment.listType);
      }
      lastListNumId = numId;
      lastListType = segment.listType;

      for (const item of segment.items) {
        const numPr = createElement("w:numPr", [
          createElement("w:ilvl", [], { "w:val": String(item.level) }),
          createElement("w:numId", [], { "w:val": numId }),
        ]);
        const pPr = createElement("w:pPr", [numPr]);
        const runs = buildFieldRuns(item.text, undefined, markers);
        if (revCtx) {
          const insElement = createInsElement(revCtx, runs);
          newNodes.push({ "w:p": [pPr, insElement] });
        } else {
          newNodes.push({ "w:p": [pPr, ...runs] });
        }
      }
    }
  }

  stampNewParagraphs(doc, newNodes, findPartRootForId(doc, "body"));

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
      if (doc.xmlParts.numbering) {
        enrichListType(normalized, doc.xmlParts.numbering);
      }
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

  const bodyChildren = getBodyChildren(doc);
  const tableNode = buildTableNode(patch.headers, patch.rows);

  if (revCtx) {
    markTableRowsAsInserted(tableNode, revCtx);
  }

  stampNewParagraphs(doc, [tableNode], findPartRootForId(doc, "body"));

  // Find the afterId target's xmlRef if specified
  let afterXmlRef: FXPNode | undefined;
  let afterElementId: string | undefined;
  let afterControl: ControlInfo | undefined;
  if (patch.afterId) {
    const target = findElementById(patch.afterId, doc);
    if (target) {
      afterXmlRef = target.xmlRef;
      afterElementId = target.id;
      // Next to an element inside a block-level control, the table is inside it too
      if (target.kind !== "page_break") afterControl = target.control;
    }
  }

  insertIntoBody(bodyChildren, [tableNode], afterXmlRef);

  // Normalize and add to doc.body
  const { nextTblIdx } = getNextIds(doc);
  const id = `tbl${nextTblIdx}`;
  const normalized = normalizeTable(tableNode, id, afterControl);

  // Insert into body array at correct position (after the target element)
  if (afterElementId !== undefined) {
    const targetIdx = doc.body.findIndex((el) => el.id === afterElementId);
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
  revCtx: RevisionContext | undefined,
  markers: boolean,
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

  // Track list context for numId reuse across segments (same as body path)
  let lastListNumId: string | undefined;
  let lastListType: "bullet" | "ordered" | undefined;

  for (const segment of segments) {
    if (segment.kind === "text") {
      // A non-blank text segment resets list context
      if (segment.lines.some((l) => l.trim().length > 0)) {
        lastListNumId = undefined;
        lastListType = undefined;
      }
      for (const line of segment.lines) {
        let pNode: FXPNode;
        const runs = buildFieldRuns(line, undefined, markers);
        if (revCtx) {
          const insElement = createInsElement(revCtx, runs);
          pNode = { "w:p": [insElement] };
        } else {
          pNode = { "w:p": runs };
        }
        stampNewParagraphs(doc, [pNode], rootNode);
        rootChildren.push(pNode);
        const normalized = normalizeParagraph(
          pNode,
          `${prefix}.${container.variant}.p${nextPIdx++}`,
        );
        hf.paragraphs.push(normalized);
      }
    } else if (segment.kind === "table") {
      // Table segment — add to container XML
      const tblNode = buildTableNode(segment.headers, segment.rows);
      if (revCtx) {
        markTableRowsAsInserted(tblNode, revCtx);
      }
      stampNewParagraphs(doc, [tblNode], rootNode);
      rootChildren.push(tblNode);

      // Normalize and track
      const nextTblIdx = hf.tables.length;
      const tblId = `${prefix}.${container.variant}.tbl${nextTblIdx}`;
      const normalizedTbl = normalizeTable(tblNode, tblId);
      hf.tables.push(normalizedTbl);
    } else {
      // List segment in container — reuse numId if same list type
      let numId: string;
      if (lastListNumId && lastListType === segment.listType) {
        numId = lastListNumId;
      } else {
        numId = createNumberingDefinition(doc, segment.listType);
      }
      lastListNumId = numId;
      lastListType = segment.listType;
      for (const item of segment.items) {
        const numPr = createElement("w:numPr", [
          createElement("w:ilvl", [], { "w:val": String(item.level) }),
          createElement("w:numId", [], { "w:val": numId }),
        ]);
        const pPr = createElement("w:pPr", [numPr]);
        const runs = buildFieldRuns(item.text, undefined, markers);
        let pNode: FXPNode;
        if (revCtx) {
          const insElement = createInsElement(revCtx, runs);
          pNode = { "w:p": [pPr, insElement] };
        } else {
          pNode = { "w:p": [pPr, ...runs] };
        }
        stampNewParagraphs(doc, [pNode], rootNode);
        rootChildren.push(pNode);
        const normalized = normalizeParagraph(
          pNode,
          `${prefix}.${container.variant}.p${nextPIdx++}`,
        );
        if (doc.xmlParts.numbering) {
          enrichListType(normalized, doc.xmlParts.numbering);
        }
        hf.paragraphs.push(normalized);
      }
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

  stampNewParagraphs(doc, [tableNode], rootNode);
  rootChildren.push(tableNode);

  // Normalize and track
  const prefix = container.type === "header" ? "hdr" : "ftr";
  const nextTblIdx = hf.tables.length;
  const id = `${prefix}.${container.variant}.tbl${nextTblIdx}`;
  const normalized = normalizeTable(tableNode, id);
  hf.tables.push(normalized);
}
