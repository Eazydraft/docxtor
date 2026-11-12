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
  checkParagraphEditable,
  checkParagraphDelete,
  checkRowInsert,
  checkRowDelete,
  findCheckbox,
} from "./control-ops.js";
import { resolveHeaderFooterContainer } from "./container-utils.js";
import { allNotes, findNoteOfParagraph, holdsNoteMark } from "../normalizer/notes.js";

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
      if (patch.type === "apply_style" && patch.find === undefined) checkParagraphEditable(para);
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
      // The new block lands next to the element, inside the control that wraps it
      const after = patch.afterId && !patch.container ? findElementById(patch.afterId, doc) : undefined;
      if (after) checkBlockInsert(after);
      return;
    }
    case "set_list_properties":
      checkParagraphEditable(findParagraphById(patch.target, doc)!);
      return;
    case "convert_to_list":
    case "remove_list":
      for (const target of patch.targets) checkParagraphEditable(findParagraphById(target, doc)!);
      return;
    case "add_table_row":
      checkRowInsert(findTableById(patch.target, doc)!, patch.afterRow);
      return;
    case "delete_table_row":
      checkRowDelete(findTableById(patch.target, doc)!, patch.row);
      return;
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
        return { valid: false, error: paragraphNotFound(patch.target, doc) };
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
        return { valid: false, error: paragraphNotFound(patch.target, doc) };
      }
      return validateNoteParagraphRemoval(para, doc);
    }
    case "add_comment": {
      const para = findParagraphById(patch.target, doc);
      if (!para) {
        return { valid: false, error: paragraphNotFound(patch.target, doc) };
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
        return { valid: false, error: paragraphNotFound(patch.target, doc) };
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
          error: paragraphNotFound(patch.target, doc),
        };
      }
      if (patch.type === "insert_text") {
        // Patches arrive as JSON from callers, so a field may be missing or
        // of the wrong type whatever the patch's TypeScript type says.
        const raw = patch as { anchor?: unknown; text?: unknown; position?: unknown };
        if (typeof raw.text !== "string") {
          return { valid: false, error: "insert_text needs text (the text to insert)" };
        }
        // A JSON null is how many serializers write an optional field that
        // is not set, so it reads as no anchor, like leaving the field out.
        if (raw.anchor !== undefined && raw.anchor !== null && typeof raw.anchor !== "string") {
          return {
            valid: false,
            error: "insert_text anchor must be text already in the paragraph; omit it to insert at the start or end of the paragraph",
          };
        }
        if (raw.position !== undefined && !["before", "after", "start", "end"].includes(raw.position as string)) {
          return {
            valid: false,
            error: `Invalid position: ${JSON.stringify(raw.position)} (must be "before" or "after" an anchor, or "start" or "end" with no anchor)`,
          };
        }
        if ((raw.position === "start" || raw.position === "end") && typeof raw.anchor === "string" && raw.anchor.length > 0) {
          return {
            valid: false,
            error: `position "${raw.position}" takes no anchor: drop the anchor to insert at the ${raw.position} of the paragraph, or use "before" or "after" with it`,
          };
        }
      } else if (typeof (patch as { find?: unknown }).find !== "string") {
        return { valid: false, error: `${patch.type} needs find (text already in the paragraph)` };
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
          error: paragraphNotFound(patch.target, doc),
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
          error: paragraphNotFound(patch.target, doc),
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
        return { valid: false, error: paragraphNotFound(patch.target, doc) };
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
        return { valid: false, error: paragraphNotFound(patch.target, doc) };
      }
      if (!para.listInfo) {
        return { valid: false, error: `Paragraph "${patch.target}" is not a list item` };
      }
      return validateNoteParagraphRemoval(para, doc);
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
          return { valid: false, error: paragraphNotFound(target, doc) };
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
          return { valid: false, error: paragraphNotFound(target, doc) };
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
        return { valid: false, error: paragraphNotFound(patch.target, doc) };
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

/**
 * A note keeps the paragraph that holds its number, and at least one
 * paragraph: removing either would leave a note Word cannot show. Deleting a
 * whole note (with its reference in the text) is not supported.
 */
function validateNoteParagraphRemoval(
  para: NormalizedParagraph,
  doc: NormalizedDocument,
): ValidationResult {
  const note = findNoteOfParagraph(doc, para);
  if (!note) return { valid: true };
  if (note.paragraphs.length === 1 || holdsNoteMark(para.xmlRef)) {
    return {
      valid: false,
      error: `Paragraph "${para.id}" is the ${note.kind}'s own paragraph (it holds the note's number); deleting a ${note.kind} is not supported. Use delete_text to remove its text`,
    };
  }
  return { valid: true };
}

function validateContainerTarget(
  container: { type: string; variant: string; section?: number },
  doc: NormalizedDocument,
): ValidationResult {
  if (!["header", "footer"].includes(container.type)) {
    return { valid: false, error: `Invalid container type: ${container.type}` };
  }
  if (!["default", "first", "even"].includes(container.variant)) {
    return { valid: false, error: `Invalid container variant: ${container.variant}` };
  }

  if (
    container.section !== undefined &&
    (!Number.isInteger(container.section) || container.section < 1)
  ) {
    return { valid: false, error: `Invalid container section: ${container.section} (sections are numbered from 1)` };
  }

  const resolved = resolveHeaderFooterContainer(
    doc,
    container as { type: "header" | "footer"; variant: string; section?: number },
  );
  if ("error" in resolved) return { valid: false, error: resolved.error };

  return { valid: true };
}

/**
 * The error for a paragraph id that names no paragraph.
 *
 * A table cell with no paragraph at all (a malformed file: Word requires
 * one in every cell) is shown by the projections under the id its first
 * paragraph would have, `tbl0.r1.c2.p0`. A patch that targets that id is
 * told what it found, since the id did come from the view.
 */
function paragraphNotFound(target: string, doc: NormalizedDocument): string {
  const m = /^(.+)\.r(\d+)\.c(\d+)\.p0$/.exec(target);
  const row = m ? findTableById(m[1]!, doc)?.rows.find((r) => r.index === Number(m[2])) : undefined;
  const cell = row?.cells.find((c) => c.colIndex === Number(m![3]));
  if (cell && cell.paragraphs.length === 0) {
    return (
      `Paragraph "${target}" not found: the cell ${m![1]}.r${m![2]}.c${m![3]} holds no paragraph, so it has no text to edit ` +
      `(the file is malformed there; Word puts a paragraph in every cell)`
    );
  }
  return `Paragraph "${target}" not found`;
}

/**
 * Find a paragraph by its ID anywhere in the document
 * (body, tables, headers, footers, footnotes, endnotes).
 *
 * Accepts either a positional id (`p12`, `tbl0.r1.c0.p0`, `hdr.default.p0`,
 * `fn1.p0`) or a native id in the form `para:4F2A1C7B`.
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
 * headers, footers, footnotes, endnotes. Case-insensitive.
 */
export function findParagraphByParaId(
  paraId: string,
  doc: NormalizedDocument,
): NormalizedParagraph | undefined {
  const wanted = paraId.toUpperCase();
  return findParagraph(doc, (p) => p.paraId === wanted);
}

/**
 * Ids of the paragraphs whose text contains `text`, in document order
 * (body with tables, headers, footers, notes), leaving out `except`.
 */
export function paragraphsContaining(
  doc: NormalizedDocument,
  text: string,
  except?: NormalizedParagraph,
): string[] {
  const ids: string[] = [];
  if (text.length === 0) return ids;
  findParagraph(doc, (p) => {
    if (p !== except && p.text.includes(text)) ids.push(p.id);
    return false;
  });
  return ids;
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

  // Search footnotes, then endnotes
  for (const note of allNotes(doc)) {
    for (const para of note.paragraphs) {
      if (predicate(para)) return para;
    }
    for (const table of note.tables) {
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

  // Search footnote and endnote tables
  for (const note of allNotes(doc)) {
    for (const table of note.tables) {
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
