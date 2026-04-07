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
 * Find a table by its ID in the document body.
 */
export function findTableById(
  id: string,
  doc: NormalizedDocument,
): NormalizedTable | undefined {
  for (const element of doc.body) {
    if (element.kind === "table" && element.id === id) {
      return element;
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
    for (const row of element.rows) {
      for (const cell of row.cells) {
        for (const para of cell.paragraphs) {
          if (para.id === id) return para;
        }
      }
    }
  }

  return undefined;
}
