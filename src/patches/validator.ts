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

export interface ValidationResult {
  valid: boolean;
  error?: string;
}

/**
 * Validate a single patch against the document.
 */
export function validatePatch(
  patch: Patch,
  doc: NormalizedDocument,
): ValidationResult {
  switch (patch.type) {
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
      if (patch.occurrence !== undefined) {
        if (patch.occurrence < 1 || !Number.isInteger(patch.occurrence)) {
          return {
            valid: false,
            error: `Invalid occurrence: ${patch.occurrence} (must be a positive integer)`,
          };
        }
      }
      // Must have at least one formatting property
      if (!patch.formatting || Object.keys(patch.formatting).length === 0) {
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
      if (patch.occurrence !== undefined) {
        if (patch.occurrence < 1 || !Number.isInteger(patch.occurrence)) {
          return {
            valid: false,
            error: `Invalid occurrence: ${patch.occurrence} (must be a positive integer)`,
          };
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
    default:
      return {
        valid: false,
        error: `Unknown patch type: ${(patch as Patch).type}`,
      };
  }
}

/**
 * Find a paragraph by its ID anywhere in the document
 * (body, tables, headers, footers).
 */
export function findParagraphById(
  id: string,
  doc: NormalizedDocument,
): NormalizedParagraph | undefined {
  // Search body elements
  for (const element of doc.body) {
    const found = searchParagraphInElement(id, element);
    if (found) return found;
  }

  // Search headers
  for (const header of doc.headers.values()) {
    for (const para of header.paragraphs) {
      if (para.id === id) return para;
    }
  }

  // Search footers
  for (const footer of doc.footers.values()) {
    for (const para of footer.paragraphs) {
      if (para.id === id) return para;
    }
  }

  return undefined;
}

/**
 * Find a table by its ID in the document body (including nested tables).
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
  id: string,
  element: BodyElement,
): NormalizedParagraph | undefined {
  if (element.kind === "paragraph") {
    return element.id === id ? element : undefined;
  }

  if (element.kind === "table") {
    return searchParagraphInTable(id, element);
  }

  return undefined;
}

function searchParagraphInTable(
  id: string,
  table: NormalizedTable,
): NormalizedParagraph | undefined {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const para of cell.paragraphs) {
        if (para.id === id) return para;
      }
      // Search nested tables
      for (const nested of cell.nestedTables) {
        const found = searchParagraphInTable(id, nested);
        if (found) return found;
      }
    }
  }
  return undefined;
}
