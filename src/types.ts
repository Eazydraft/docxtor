/**
 * Core type definitions for docxtor.
 *
 * The key architectural concept: every normalized structure holds `xmlRef`
 * pointers back to the original fast-xml-parser (FXP) nodes. Patches
 * execute by following these pointers and mutating the FXP object directly.
 * The FXP object IS the XML — serializing it produces the output document.
 */

// =============================================================================
// FXP Types (fast-xml-parser preserveOrder structure)
// =============================================================================

/**
 * A node in the FXP ordered array.
 * With `preserveOrder: true`, every element is a single-key object where:
 * - The key is the tag name (e.g., "w:p", "w:r", "w:t")
 * - The value is an array of child FXPNode[]
 * - Attributes are stored under the ":@" key
 * - Text nodes use "#text" as the key
 */
export interface FXPNode {
  [tagName: string]: FXPNode[] | string | FXPAttributes;
}

export interface FXPAttributes {
  [attrName: string]: string;
}

// =============================================================================
// Normalized AST Types
// =============================================================================

/** The full normalized document, holding references into FXP objects. */
export interface NormalizedDocument {
  /** All body-level elements in document order. */
  body: BodyElement[];
  /** Headers keyed by relationship ID. */
  headers: Map<string, HeaderFooter>;
  /** Footers keyed by relationship ID. */
  footers: Map<string, HeaderFooter>;
  /**
   * The raw FXP objects for each XML part.
   * These are the mutable structures that patches operate on.
   */
  xmlParts: DocxXmlParts;
}

export interface DocxXmlParts {
  document: FXPNode[];
  headers: Map<string, FXPNode[]>;
  footers: Map<string, FXPNode[]>;
  /** All other XML files in the zip (styles, numbering, etc.) — preserved as-is. */
  [key: string]: FXPNode[] | Map<string, FXPNode[]> | undefined;
}

/** A body element is either a paragraph or a table. */
export type BodyElement =
  | NormalizedParagraph
  | NormalizedTable;

/** Header or footer content. */
export interface HeaderFooter {
  type: "header" | "footer";
  /** e.g., "default", "first", "even" */
  variant: string;
  /** Relationship ID from sectPr. */
  rId: string;
  /** Paragraphs inside this header/footer. */
  paragraphs: NormalizedParagraph[];
  /** The FXP object for this header/footer XML file. */
  xmlRef: FXPNode[];
}

// =============================================================================
// Paragraph
// =============================================================================

export interface NormalizedParagraph {
  kind: "paragraph";
  /** Stable ID, e.g., "p0", "tbl0.r1.c0.p0", "hdr.default.p0" */
  id: string;
  /** Concatenated text from all runs. */
  text: string;
  /** Ordered run mappings with character offsets and XML refs. */
  runs: RunMapping[];
  /** Direct reference to the w:p FXP node. */
  xmlRef: FXPNode;
  /** Paragraph style name (if any). */
  style?: string;
  /** List/numbering info (if any). */
  listInfo?: ListInfo;
}

export interface RunMapping {
  /** This run's text content. */
  text: string;
  /** Start character offset in the paragraph's normalized text (inclusive). */
  startOffset: number;
  /** End character offset in the paragraph's normalized text (exclusive). */
  endOffset: number;
  /** Direct reference to the w:r FXP node. */
  xmlRef: FXPNode;
  /** Direct reference to the w:t FXP node within this run. */
  textXmlRef: FXPNode;
  /** Captured formatting properties. */
  formatting: RunFormatting;
}

export interface RunFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  highlight?: string;
}

export interface ListInfo {
  /** Numbering ID from w:numId. */
  numId: string;
  /** Indentation level from w:ilvl. */
  level: number;
}

// =============================================================================
// Table
// =============================================================================

export interface NormalizedTable {
  kind: "table";
  /** Stable ID, e.g., "tbl0", "tbl1" */
  id: string;
  /** Table rows. */
  rows: NormalizedTableRow[];
  /** Direct reference to the w:tbl FXP node. */
  xmlRef: FXPNode;
  /** Table style name (if any). */
  style?: string;
}

export interface NormalizedTableRow {
  /** Row index (0-based). */
  index: number;
  /** Cells in this row. */
  cells: NormalizedTableCell[];
  /** Direct reference to the w:tr FXP node. */
  xmlRef: FXPNode;
}

export interface NormalizedTableCell {
  /** Row index (0-based). */
  rowIndex: number;
  /** Column index (0-based). */
  colIndex: number;
  /** Paragraphs inside this cell. */
  paragraphs: NormalizedParagraph[];
  /** Direct reference to the w:tc FXP node. */
  xmlRef: FXPNode;
}

// =============================================================================
// Patch Types
// =============================================================================

export type Patch =
  | ReplaceTextPatch
  | InsertTextPatch
  | DeleteTextPatch
  | AddTableRowPatch
  | DeleteTableRowPatch
  | ReplaceCellTextPatch;

export interface ReplaceTextPatch {
  type: "replace_text";
  /** Target paragraph ID. */
  target: string;
  /** Text to find in normalized paragraph text. */
  find: string;
  /** Replacement text. */
  replace: string;
  /** Which occurrence to replace (1-based). Default: 1. */
  occurrence?: number;
}

export interface InsertTextPatch {
  type: "insert_text";
  /** Target paragraph ID. */
  target: string;
  /** Insert before or after the anchor text. */
  position: "before" | "after";
  /** Text to anchor against. */
  anchor: string;
  /** Text to insert. */
  text: string;
  /** Which occurrence of the anchor (1-based). Default: 1. */
  occurrence?: number;
}

export interface DeleteTextPatch {
  type: "delete_text";
  /** Target paragraph ID. */
  target: string;
  /** Text to delete. */
  find: string;
  /** Which occurrence to delete (1-based). Default: 1. */
  occurrence?: number;
}

export interface AddTableRowPatch {
  type: "add_table_row";
  /** Target table ID. */
  target: string;
  /** Row index to insert after. -1 inserts at the beginning. */
  afterRow: number;
  /** Cell text values for the new row. */
  cells: string[];
}

export interface DeleteTableRowPatch {
  type: "delete_table_row";
  /** Target table ID. */
  target: string;
  /** Row index to delete (0-based). */
  row: number;
}

export interface ReplaceCellTextPatch {
  type: "replace_cell_text";
  /** Target paragraph ID inside a table cell. */
  target: string;
  /** Text to find. */
  find: string;
  /** Replacement text. */
  replace: string;
  /** Which occurrence (1-based). Default: 1. */
  occurrence?: number;
}

// =============================================================================
// Patch Result
// =============================================================================

export interface PatchResult {
  /** Whether the patch was applied successfully. */
  success: boolean;
  /** The patch that was applied. */
  patch: Patch;
  /** Error message if the patch failed. */
  error?: string;
}

// =============================================================================
// LLM Projection
// =============================================================================

export type ProjectionFormat = "json" | "markdown";

export interface ProjectionOptions {
  /** Include formatting hints in the projection. Default: false. */
  includeFormatting?: boolean;
  /** Include headers and footers. Default: true. */
  includeHeadersFooters?: boolean;
  /** Include table structure. Default: true. */
  includeTables?: boolean;
}

// =============================================================================
// Public API Options
// =============================================================================

export interface LoadOptions {
  /** Whether to normalize headers and footers. Default: true. */
  includeHeadersFooters?: boolean;
}
