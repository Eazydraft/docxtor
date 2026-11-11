/**
 * Patch validator: validates patch structure before execution.
 */

import type {
  Patch,
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  BodyElement,
} from "../types.js";
import { getChildren, getTagName, getAttr } from "../xml/helpers.js";
import { hasBlockMarkdown } from "./text-segments.js";
import { findOccurrence, resolveRunRanges } from "../mapping/text-to-runs.js";
import {
  ControlEditError,
  planControlEdit,
  checkRunsEditable,
  checkBlockInsert,
  checkParagraphDelete,
  findCheckbox,
} from "./control-ops.js";

export interface ValidationResult {
  valid: boolean;
  error?: string;
}

/** Prefix for native paragraph identity targets: `para:4F2A1C7B`. */
export const PARA_ID_PREFIX = "para:";

/**
 * Validate a single patch against the document.
 *
 * @param markers - Whether inline formatting markers apply to the patch's
 *                  text fields (see `PatchOptions.inlineMarkers`). Default: true.
 */
export function validatePatch(
  patch: Patch,
  doc: NormalizedDocument,
  markers: boolean = true,
): ValidationResult {
  const result = validateStructure(patch, doc);
  if (!result.valid) return result;
  // What a content control allows: a locked, data-bound or typed control
  // refuses some edits, with a reason the caller can act on.
  try {
    checkControls(patch, doc, markers);
  } catch (err) {
    if (err instanceof ControlEditError) return { valid: false, error: err.message };
    throw err;
  }
  return result;
}

/** Runs of `para` covered by the given occurrence of `find`, or undefined when absent. */
function matchedRuns(para: NormalizedParagraph, find: string | undefined, occurrence: number | undefined) {
  if (find === undefined || find.length === 0) return undefined;
  const start = findOccurrence(para.text, find, occurrence ?? 1);
  if (start === -1) return undefined;
  return resolveRunRanges(para.runs, start, start + find.length).map((r) => r.run);
}

function checkControls(patch: Patch, doc: NormalizedDocument, markers: boolean): void {
  switch (patch.type) {
    case "replace_text":
    case "insert_text":
    case "delete_text":
    case "replace_cell_text":
      planControlEdit(findParagraphById(patch.target, doc)!, patch, markers);
      return;
    case "set_checkbox":
      findCheckbox(findParagraphById(patch.target, doc)!, patch);
      return;
    case "apply_formatting":
    case "apply_style": {
      const para = findParagraphById(patch.target, doc)!;
      // A paragraph style changes no run, so only a control around the
      // paragraph matters; anything else is checked run by run.
      if (patch.type === "apply_style" && patch.find === undefined) checkRunsEditable(para, []);
      else checkRunsEditable(para, matchedRuns(para, patch.find, patch.occurrence));
      return;
    }
    case "insert_paragraph":
    case "add_list_item":
      checkBlockInsert(findParagraphById(patch.target, doc)!);
      return;
    case "delete_paragraph":
    case "remove_list_item":
      checkParagraphDelete(findParagraphById(patch.target, doc)!);
      return;
    case "add_table":
    case "add_list": {
      const after = patch.afterId && !patch.container ? findElementById(patch.afterId, doc) : undefined;
      if (after?.kind === "paragraph") checkBlockInsert(after);
      return;
    }
    default:
      return;
  }
}

function validateStructure(
  patch: Patch,
  doc: NormalizedDocument,
): ValidationResult {
  switch (patch.type) {
    case "insert_paragraph": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return { valid: false, error: `Paragraph "${patch.target}" not found` };
      }
      if (!patch.text || patch.text.length === 0) {
        return { valid: false, error: "Empty paragraph text" };
      }
      if (!["before", "after"].includes(patch.position)) {
        return { valid: false, error: `Invalid position: ${patch.position}` };
      }
      if (hasBlockMarkdown(patch.text)) {
        return {
          valid: false,
          error:
            "Markdown tables and lists are not supported in insert_paragraph; use add_table or add_list with afterId",
        };
      }
      return { valid: true };
    }
    case "delete_paragraph": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return { valid: false, error: `Paragraph "${patch.target}" not found` };
      }
      return { valid: true };
    }
    case "add_comment": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return { valid: false, error: `Paragraph "${patch.target}" not found` };
      }
      if (!patch.text || patch.text.length === 0) {
        return { valid: false, error: "Empty comment text" };
      }
      if (patch.find !== undefined && patch.find.length === 0) {
        return { valid: false, error: "Empty find string" };
      }
      if (patch.find !== undefined && patch.occurrence !== undefined) {
        if (patch.occurrence < 1 || !Number.isInteger(patch.occurrence)) {
          return {
            valid: false,
            error: `Invalid occurrence: ${patch.occurrence} (must be a positive integer)`,
          };
        }
      }
      return { valid: true };
    }
    case "set_checkbox": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return { valid: false, error: `Paragraph "${patch.target}" not found` };
      }
      if (typeof patch.checked !== "boolean") {
        return { valid: false, error: "checked must be true or false" };
      }
      if (patch.occurrence !== undefined && (patch.occurrence < 1 || !Number.isInteger(patch.occurrence))) {
        return {
          valid: false,
          error: `Invalid occurrence: ${patch.occurrence} (must be a positive integer)`,
        };
      }
      return { valid: true };
    }
    case "replace_text":
    case "insert_text":
    case "delete_text":
    case "replace_cell_text": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return {
          valid: false,
          error: `Paragraph "${patch.target}" not found`,
        };
      }
      if ("find" in patch && patch.find.length === 0) {
        return { valid: false, error: "Empty find string" };
      }
      if (patch.type === "insert_text" && patch.text.length === 0) {
        return { valid: false, error: "Empty insert text" };
      }
      if ("occurrence" in patch && patch.occurrence !== undefined) {
        if (patch.occurrence < 1 || !Number.isInteger(patch.occurrence)) {
          return {
            valid: false,
            error: `Invalid occurrence: ${patch.occurrence} (must be a positive integer)`,
          };
        }
      }
      return { valid: true };
    }
    case "add_table_row": {
      const table = findTableById(patch.target, doc);
      if (!table) {
        return {
          valid: false,
          error: `Table "${patch.target}" not found`,
        };
      }
      if (patch.cells.length === 0) {
        return { valid: false, error: "Empty cells array" };
      }
      if (patch.afterRow < -1 || patch.afterRow >= table.rows.length) {
        return {
          valid: false,
          error: `Invalid afterRow ${patch.afterRow} for table "${patch.target}" with ${table.rows.length} rows`,
        };
      }
      return { valid: true };
    }
    case "delete_table_row": {
      const table = findTableById(patch.target, doc);
      if (!table) {
        return {
          valid: false,
          error: `Table "${patch.target}" not found`,
        };
      }
      if (patch.row < 0 || patch.row >= table.rows.length) {
        return {
          valid: false,
          error: `Invalid row index ${patch.row} for table "${patch.target}" with ${table.rows.length} rows`,
        };
      }
      return { valid: true };
    }
    case "apply_formatting": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return {
          valid: false,
          error: `Paragraph "${patch.target}" not found`,
        };
      }
      if (patch.find !== undefined && patch.find.length === 0) {
        return { valid: false, error: "Empty find string" };
      }
      // Only validate occurrence when find is provided (it's ignored otherwise)
      if (patch.find !== undefined && patch.find.length > 0) {
        if (patch.occurrence !== undefined) {
          if (patch.occurrence < 1 || !Number.isInteger(patch.occurrence)) {
            return {
              valid: false,
              error: `Invalid occurrence: ${patch.occurrence} (must be a positive integer)`,
            };
          }
        }
      }
      // Must have at least one formatting property with a meaningful value
      // (undefined values are ignored, so { bold: undefined } is effectively empty)
      if (
        !patch.formatting ||
        !Object.values(patch.formatting).some((v) => v !== undefined)
      ) {
        return { valid: false, error: "Empty formatting object" };
      }
      return { valid: true };
    }
    case "apply_style": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return {
          valid: false,
          error: `Paragraph "${patch.target}" not found`,
        };
      }
      if (patch.find !== undefined && patch.find.length === 0) {
        return { valid: false, error: "Empty find string" };
      }
      // Only validate occurrence when find is provided (it's ignored otherwise)
      if (patch.find !== undefined && patch.find.length > 0) {
        if (patch.occurrence !== undefined) {
          if (patch.occurrence < 1 || !Number.isInteger(patch.occurrence)) {
            return {
              valid: false,
              error: `Invalid occurrence: ${patch.occurrence} (must be a positive integer)`,
            };
          }
        }
      }
      if (!patch.style || patch.style.length === 0) {
        return { valid: false, error: "Empty style ID" };
      }
      // Validate style exists and is correct type
      const styleInfo = doc.styles.find((s) => s.id === patch.style);
      if (!styleInfo) {
        return {
          valid: false,
          error: `Style "${patch.style}" not found in document styles`,
        };
      }
      if (patch.find !== undefined && patch.find.length > 0) {
        // Character style mode
        if (styleInfo.type !== "character") {
          return {
            valid: false,
            error: `Style "${patch.style}" is type "${styleInfo.type}", expected "character" when "find" is specified`,
          };
        }
      } else {
        // Paragraph style mode
        if (styleInfo.type !== "paragraph") {
          return {
            valid: false,
            error: `Style "${patch.style}" is type "${styleInfo.type}", expected "paragraph" when "find" is omitted`,
          };
        }
      }
      return { valid: true };
    }
    case "append_text": {
      if (!patch.text || patch.text.length === 0) {
        return { valid: false, error: "Empty append text" };
      }
      if (patch.container) {
        const containerValidation = validateContainerTarget(patch.container, doc);
        if (!containerValidation.valid) {
          return containerValidation;
        }
      }
      return { valid: true };
    }
    case "add_table": {
      if (!patch.rows || patch.rows.length === 0) {
        if (!patch.headers || patch.headers.length === 0) {
          return { valid: false, error: "Table must have at least one row or headers" };
        }
      }
      // Determine column count from headers or first row
      const colCount = patch.headers?.length ?? (patch.rows.length > 0 ? patch.rows[0]!.length : 0);
      if (colCount === 0) {
        return { valid: false, error: "Table must have at least one column" };
      }
      // All rows must have the same column count
      for (let i = 0; i < patch.rows.length; i++) {
        if (patch.rows[i]!.length !== colCount) {
          return {
            valid: false,
            error: `Row ${i} has ${patch.rows[i]!.length} cells, expected ${colCount}`,
          };
        }
      }
      // If headers provided, must match column count
      if (patch.headers && patch.rows.length > 0 && patch.headers.length !== patch.rows[0]!.length) {
        return {
          valid: false,
          error: `Headers have ${patch.headers.length} cells, but rows have ${patch.rows[0]!.length}`,
        };
      }
      // Validate afterId target exists if provided
      if (patch.afterId) {
        if (patch.container) {
          return {
            valid: false,
            error: `"afterId" is not supported when "container" is set — container-targeted tables are appended to the end of the header/footer`,
          };
        }
        const found = findElementById(patch.afterId, doc);
        if (!found) {
          return {
            valid: false,
            error: `Element "${patch.afterId}" not found`,
          };
        }
      }
      // Validate container if provided
      if (patch.container) {
        const containerValidation = validateContainerTarget(patch.container, doc);
        if (!containerValidation.valid) {
          return containerValidation;
        }
      }
      return { valid: true };
    }
    case "add_header":
    case "add_footer": {
      if (!["default", "first", "even"].includes(patch.variant)) {
        return { valid: false, error: `Invalid variant: ${patch.variant}` };
      }
      if (!patch.paragraphs || patch.paragraphs.length === 0) {
        return { valid: false, error: "At least one paragraph is required" };
      }
      return { valid: true };
    }
    case "remove_header":
    case "remove_footer": {
      if (!["default", "first", "even"].includes(patch.variant)) {
        return { valid: false, error: `Invalid variant: ${patch.variant}` };
      }
      return { valid: true };
    }

    // =========================================================================
    // List operations
    // =========================================================================

    case "add_list": {
      const validTypes = ["bullet", "ordered", "multilevel"];
      if (!validTypes.includes(patch.listType)) {
        return { valid: false, error: `Invalid listType: ${patch.listType}` };
      }
      if (!patch.items || patch.items.length === 0) {
        return { valid: false, error: "items must be non-empty" };
      }
      if (patch.afterId && patch.container) {
        return {
          valid: false,
          error: `"afterId" is not supported when "container" is set`,
        };
      }
      if (patch.afterId) {
        const found = findElementById(patch.afterId, doc);
        if (!found) {
          return { valid: false, error: `Element "${patch.afterId}" not found` };
        }
      }
      if (patch.container) {
        const containerValidation = validateContainerTarget(patch.container, doc);
        if (!containerValidation.valid) return containerValidation;
      }
      if (patch.numId !== undefined) {
        // Validate that the numId exists in numbering.xml
        if (!doc.xmlParts.numbering) {
          return { valid: false, error: `numId "${patch.numId}" specified but document has no numbering.xml` };
        }
        const root = doc.xmlParts.numbering.find((n) => getTagName(n) === "w:numbering");
        const numExists = root
          ? getChildren(root).some(
              (c) => getTagName(c) === "w:num" && getAttr(c, "w:numId") === patch.numId,
            )
          : false;
        if (!numExists) {
          return { valid: false, error: `numId "${patch.numId}" not found in numbering.xml` };
        }
      }
      return { valid: true };
    }

    case "add_list_item": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return { valid: false, error: `Paragraph "${patch.target}" not found` };
      }
      if (!para.listInfo) {
        return { valid: false, error: `Paragraph "${patch.target}" is not a list item` };
      }
      if (!patch.text && patch.text !== "") {
        return { valid: false, error: "text is required" };
      }
      if (!["before", "after"].includes(patch.position)) {
        return { valid: false, error: `Invalid position: ${patch.position}` };
      }
      if (patch.level !== undefined && patch.level < 0) {
        return { valid: false, error: "level must be >= 0" };
      }
      return { valid: true };
    }

    case "remove_list_item": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return { valid: false, error: `Paragraph "${patch.target}" not found` };
      }
      if (!para.listInfo) {
        return { valid: false, error: `Paragraph "${patch.target}" is not a list item` };
      }
      return { valid: true };
    }

    case "convert_to_list": {
      const validTypes = ["bullet", "ordered", "multilevel"];
      if (!validTypes.includes(patch.listType)) {
        return { valid: false, error: `Invalid listType: ${patch.listType}` };
      }
      if (!patch.targets || patch.targets.length === 0) {
        return { valid: false, error: "targets must be non-empty" };
      }
      for (const target of patch.targets) {
        const para = findParagraphById(target, doc);
        if (!para) {
          return { valid: false, error: `Paragraph "${target}" not found` };
        }
      }
      return { valid: true };
    }

    case "remove_list": {
      if (!patch.targets || patch.targets.length === 0) {
        return { valid: false, error: "targets must be non-empty" };
      }
      for (const target of patch.targets) {
        const para = findParagraphById(target, doc);
        if (!para) {
          return { valid: false, error: `Paragraph "${target}" not found` };
        }
        if (!para.listInfo) {
          return { valid: false, error: `Paragraph "${target}" is not a list item` };
        }
      }
      return { valid: true };
    }

    case "set_list_properties": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return { valid: false, error: `Paragraph "${patch.target}" not found` };
      }
      if (!para.listInfo) {
        return { valid: false, error: `Paragraph "${patch.target}" is not a list item` };
      }
      if (patch.listType === undefined && patch.level === undefined) {
        return { valid: false, error: "At least one of listType or level must be provided" };
      }
      if (patch.listType !== undefined && !["bullet", "ordered", "multilevel"].includes(patch.listType)) {
        return { valid: false, error: `Invalid listType: ${patch.listType}` };
      }
      if (patch.level !== undefined && patch.level < 0) {
        return { valid: false, error: "level must be >= 0" };
      }
      return { valid: true };
    }

    default:
      return {
        valid: false,
        error: `Unknown patch type: ${(patch as Patch).type}`,
      };
  }
}

function validateContainerTarget(
  container: { type: string; variant: string },
  doc: NormalizedDocument,
): ValidationResult {
  if (!["header", "footer"].includes(container.type)) {
    return { valid: false, error: `Invalid container type: ${container.type}` };
  }
  if (!["default", "first", "even"].includes(container.variant)) {
    return { valid: false, error: `Invalid container variant: ${container.variant}` };
  }

  const parts = container.type === "header" ? doc.headers : doc.footers;
  const exists = [...parts.values()].some((part) => part.variant === container.variant);
  if (!exists) {
    return {
      valid: false,
      error: `No ${container.type} with variant "${container.variant}" found`,
    };
  }

  return { valid: true };
}

/**
 * Find a paragraph by its ID anywhere in the document
 * (body, tables, headers, footers).
 *
 * Accepts either a positional id (`p12`, `tbl0.r1.c0.p0`, `hdr.default.p0`)
 * or a native id in the form `para:4F2A1C7B`.
 */
export function findParagraphById(
  id: string,
  doc: NormalizedDocument,
): NormalizedParagraph | undefined {
  if (id.startsWith(PARA_ID_PREFIX)) {
    return findParagraphByParaId(id.slice(PARA_ID_PREFIX.length), doc);
  }
  return findParagraph(doc, (p) => p.id === id);
}

/**
 * Find a paragraph by its native `w14:paraId` anywhere in the document.
 * Resolution order: body (paragraphs and table cells in document order),
 * headers, footers. Case-insensitive.
 */
export function findParagraphByParaId(
  paraId: string,
  doc: NormalizedDocument,
): NormalizedParagraph | undefined {
  const wanted = paraId.toUpperCase();
  return findParagraph(doc, (p) => p.paraId === wanted);
}

function findParagraph(
  doc: NormalizedDocument,
  predicate: (p: NormalizedParagraph) => boolean,
): NormalizedParagraph | undefined {
  // Search body elements
  for (const element of doc.body) {
    const found = searchParagraphInElement(predicate, element);
    if (found) return found;
  }

  // Search headers (paragraphs, then tables)
  for (const header of doc.headers.values()) {
    for (const para of header.paragraphs) {
      if (predicate(para)) return para;
    }
    for (const table of header.tables) {
      const found = searchParagraphInTable(predicate, table);
      if (found) return found;
    }
  }

  // Search footers
  for (const footer of doc.footers.values()) {
    for (const para of footer.paragraphs) {
      if (predicate(para)) return para;
    }
    for (const table of footer.tables) {
      const found = searchParagraphInTable(predicate, table);
      if (found) return found;
    }
  }

  return undefined;
}

/**
 * Find a table by its ID in the document body, headers, and footers (including nested tables).
 */
export function findTableById(
  id: string,
  doc: NormalizedDocument,
): NormalizedTable | undefined {
  for (const element of doc.body) {
    if (element.kind === "table") {
      const found = searchTableInTable(id, element);
      if (found) return found;
    }
  }

  // Search header tables
  for (const header of doc.headers.values()) {
    for (const table of header.tables) {
      const found = searchTableInTable(id, table);
      if (found) return found;
    }
  }

  // Search footer tables
  for (const footer of doc.footers.values()) {
    for (const table of footer.tables) {
      const found = searchTableInTable(id, table);
      if (found) return found;
    }
  }

  return undefined;
}

function searchTableInTable(
  id: string,
  table: NormalizedTable,
): NormalizedTable | undefined {
  if (table.id === id) return table;

  // Search nested tables
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const nested of cell.nestedTables) {
        const found = searchTableInTable(id, nested);
        if (found) return found;
      }
    }
  }

  return undefined;
}

function searchParagraphInElement(
  predicate: (p: NormalizedParagraph) => boolean,
  element: BodyElement,
): NormalizedParagraph | undefined {
  if (element.kind === "paragraph") {
    return predicate(element) ? element : undefined;
  }

  if (element.kind === "table") {
    return searchParagraphInTable(predicate, element);
  }

  return undefined;
}

function searchParagraphInTable(
  predicate: (p: NormalizedParagraph) => boolean,
  table: NormalizedTable,
): NormalizedParagraph | undefined {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const para of cell.paragraphs) {
        if (predicate(para)) return para;
      }
      // Search nested tables
      for (const nested of cell.nestedTables) {
        const found = searchParagraphInTable(predicate, nested);
        if (found) return found;
      }
    }
  }
  return undefined;
}

/**
 * Find any top-level body element (paragraph or table) by its ID.
 * Accepts `para:` ids for body paragraphs.
 */
export function findElementById(
  id: string,
  doc: NormalizedDocument,
): BodyElement | undefined {
  if (id.startsWith(PARA_ID_PREFIX)) {
    const wanted = id.slice(PARA_ID_PREFIX.length).toUpperCase();
    return doc.body.find((el) => el.kind === "paragraph" && el.paraId === wanted);
  }
  for (const element of doc.body) {
    if (element.kind === "paragraph" && element.id === id) return element;
    if (element.kind === "table") {
      if (element.id === id) return element;
    }
  }
  return undefined;
}
