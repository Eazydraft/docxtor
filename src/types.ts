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
  /**
   * Style definitions extracted from word/styles.xml (read-only).
   * Empty array if styles.xml is not present.
   */
  styles: StyleInfo[];
  /**
   * The options the document was loaded with. Used when the document is
   * re-normalized (e.g., after accepting revisions or an atomic batch).
   */
  loadOptions?: LoadOptions;
}

export interface DocxXmlParts {
  document: FXPNode[];
  headers: Map<string, FXPNode[]>;
  footers: Map<string, FXPNode[]>;
  /** Parsed word/_rels/document.xml.rels for mutation by header/footer ops. */
  documentRels?: FXPNode[];
  /** Parsed [Content_Types].xml for mutation by header/footer ops. */
  contentTypes?: FXPNode[];
  /** Parsed word/settings.xml for mutation (e.g., w:evenAndOddHeaders). */
  settings?: FXPNode[];
  /** All other XML files in the zip (styles, numbering, etc.) — preserved as-is. */
  [key: string]: FXPNode[] | Map<string, FXPNode[]> | undefined;
  /** Parsed word/numbering.xml for list management. */
  numbering?: FXPNode[];
  /** Parsed word/footnotes.xml (if present). Walked by accept/reject revisions. */
  footnotes?: FXPNode[];
  /** Parsed word/endnotes.xml (if present). Walked by accept/reject revisions. */
  endnotes?: FXPNode[];
  /** Parsed word/comments.xml (if present, or created on demand by add_comment). */
  comments?: FXPNode[];
  /** Parsed docProps/core.xml (if present, or created on demand by setCoreProperties). */
  core?: FXPNode[];
  /** Parsed docProps/app.xml (if present). */
  app?: FXPNode[];
  /** Parsed _rels/.rels (package relationships). */
  rootRels?: FXPNode[];
}

/** A body element is a paragraph, table, or page break. */
export type BodyElement =
  | NormalizedParagraph
  | NormalizedTable
  | NormalizedPageBreak;

/**
 * Names a header or footer as the target of a patch. `section` (1-based, in
 * document order) is required when sections have different headers or
 * footers of this type and variant. Sections that share one part need no
 * `section`: either names the same part.
 */
export interface HeaderFooterContainer {
  type: "header" | "footer";
  variant: HeaderFooterVariant;
  section?: number;
}

/** One `w:headerReference` / `w:footerReference` to a header or footer part. */
export interface HeaderFooterReference {
  /** 1-based index of the referencing section, in document order. */
  sectionIndex: number;
  /** "default", "first" or "even". */
  variant: string;
  /** Relationship id used by the reference. */
  rId: string;
}

/** Header or footer content. */
export interface HeaderFooter {
  type: "header" | "footer";
  /** e.g., "default", "first", "even" */
  variant: string;
  /** Relationship ID from sectPr. */
  rId: string;
  /** 1-based index, in document order, of the first section that references this part. */
  sectionIndex: number;
  /**
   * Every reference to this part, in document order. Usually one. A part
   * shared by sections, or used as two variants, has one entry per
   * reference; the first is the one `variant`, `rId` and `sectionIndex` give.
   */
  references: HeaderFooterReference[];
  /**
   * What ids inside this part start with: `hdr` / `ftr`, or `hdr.s2` / `ftr.s2`
   * when more than one section of the document has headers or footers.
   * A paragraph id is `${idPrefix}.${variant}.p${n}`.
   */
  idPrefix: string;
  /** Paragraphs inside this header/footer. */
  paragraphs: NormalizedParagraph[];
  /** Tables inside this header/footer. */
  tables: NormalizedTable[];
  /** The FXP object for this header/footer XML file. */
  xmlRef: FXPNode[];
}

// =============================================================================
// Page Break
// =============================================================================

/**
 * A page break element in the document body.
 *
 * Surfaces paragraphs whose only purpose is a page break as a distinct
 * body element for clean projection output. Detects three OOXML mechanisms:
 * - `<w:br w:type="page"/>` — explicit page break inserted by the user
 * - `<w:lastRenderedPageBreak/>` — Word rendering hint
 * - `<w:pageBreakBefore/>` in `w:pPr` — paragraph-level "page break before"
 */
export interface NormalizedPageBreak {
  kind: "page_break";
  /** Stable ID, e.g., "pb0", "pb1" */
  id: string;
  /** Direct reference to the w:p FXP node containing the page break. */
  xmlRef: FXPNode;
}

// =============================================================================
// Paragraph
// =============================================================================

export interface NormalizedParagraph {
  kind: "paragraph";
  /** Stable ID, e.g., "p0", "tbl0.r1.c0.p0", "hdr.default.p0" */
  id: string;
  /**
   * Native Word paragraph identity from `w14:paraId` (8 uppercase hex chars).
   * Survives round trips through Word, unlike the positional `id`.
   * Present on Word 2010+ files, or on any paragraph when loaded with
   * `assignParaIds: true`. Addressable in patches as `para:XXXXXXXX`.
   */
  paraId?: string;
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
  /** Revision marks captured from w:ins / w:del elements (for projection use). */
  revisions?: RevisionMark[];
  /**
   * True if this paragraph contains a page break.
   * Detects `w:br w:type="page"`, `w:lastRenderedPageBreak`, and
   * `w:pageBreakBefore` in `w:pPr`.
   * Only set when the paragraph has text AND a page break — paragraphs
   * that consist solely of a page break are surfaced as NormalizedPageBreak instead.
   */
  hasPageBreak?: boolean;
  /**
   * Character offset within `.text` where the page break occurs.
   * Only set when `hasPageBreak` is true.
   *
   * - For `w:pageBreakBefore`: always `0` (the entire paragraph starts on a new page).
   * - For `w:br w:type="page"` or `w:lastRenderedPageBreak`: the character position
   *   in the concatenated paragraph text where the break element appears.
   *
   * In an ingestion pipeline, text before this offset is on the previous page,
   * and text from this offset onward is on the next page.
   */
  pageBreakOffset?: number;
}

/**
 * A revision mark captured during normalization, representing
 * a tracked insert or delete in the paragraph.
 */
export interface RevisionMark {
  type: "insert" | "delete";
  /** The text that was inserted or deleted. */
  text: string;
  /** For inserts: position in .text. For deletes: position where deletion occurred (zero-width). */
  startOffset: number;
  endOffset: number;
  /** The revision author. */
  author: string;
  /** The revision date (ISO 8601). */
  date: string;
  /**
   * The `w:id` of the `w:ins` / `w:del` element, when it has one. Pass it to
   * `acceptAllRevisions({ ids })` / `rejectAllRevisions({ ids })` to resolve
   * this change alone.
   */
  id?: string;
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
  /**
   * Direct reference to the w:t FXP node within this run.
   * Absent for runs that carry no `w:t` — a run holding only a line break
   * (`<w:r><w:br/></w:r>`) still maps to a `\n` in the paragraph text.
   */
  textXmlRef?: FXPNode;
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
  /** Raw `w:color w:val` (hex RGB such as "FF0000", or "auto"). */
  color?: string;
  /** `w:color w:themeColor` when the colour is theme-mapped (e.g., "accent2"). */
  themeColor?: string;
  highlight?: string;
}

/**
 * Filter for selecting runs in a projection by their formatting.
 * Every property given must match. String comparisons are case-insensitive
 * (so `color: "ff0000"` matches `w:color w:val="FF0000"`).
 */
export interface RunFilter {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strikethrough?: boolean;
  /** Hex RGB colour, e.g. "FF0000". Matches the raw `w:val` regardless of theme mapping. */
  color?: string;
  highlight?: string;
}

export interface ListInfo {
  /** Numbering ID from w:numId. */
  numId: string;
  /** Indentation level from w:ilvl (0-based). */
  level: number;
  /** Resolved list type. Undefined if numbering.xml is unavailable or unresolvable. */
  listType?: "bullet" | "ordered" | "multilevel";
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
  /** Nested tables inside this cell. */
  nestedTables: NormalizedTable[];
  /** Direct reference to the w:tc FXP node. */
  xmlRef: FXPNode;
}

// =============================================================================
// Style Info (extracted from styles.xml)
// =============================================================================

/**
 * A style definition extracted from word/styles.xml.
 * Read-only — we never create or modify style definitions.
 */
export interface StyleInfo {
  /** The style ID used in w:pStyle or w:rStyle (e.g., "Heading1"). */
  id: string;
  /** The human-readable name (e.g., "heading 1"). */
  name: string;
  /** OOXML style type. */
  type: "paragraph" | "character" | "table" | "numbering";
}

// =============================================================================
// List Patch Types
// =============================================================================

export interface AddListPatch {
  type: "add_list";
  /** Bullet, ordered (independent counters per level), or multilevel (compound numbering like 1.2.3). */
  listType: "bullet" | "ordered" | "multilevel";
  /** Text for each list item. Supports inline formatting markers and {{FIELD}} placeholders. */
  items: Array<string | { text: string; level?: number }>;
  /** Default indentation level for items without explicit level. Default: 0. */
  level?: number;
  /** Reuse an existing numbering definition instead of creating a new one.
   *  Pass the numId from an existing list to continue its numbering sequence. */
  numId?: string;
  /** Insert after this element ID. If omitted, appends to end of document body. */
  afterId?: string;
  /** Target a header/footer instead of the document body. */
  container?: HeaderFooterContainer;
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface AddListItemPatch {
  type: "add_list_item";
  /** Paragraph ID of an existing list item to insert relative to. */
  target: string;
  /** Insert before or after the target list item. */
  position: "before" | "after";
  /** Text for the new list item. Supports inline formatting markers and {{FIELD}} placeholders. */
  text: string;
  /** Indentation level. If omitted, inherits from target. */
  level?: number;
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface RemoveListItemPatch {
  type: "remove_list_item";
  /** Paragraph ID of the list item to remove. */
  target: string;
}

export interface ConvertToListPatch {
  type: "convert_to_list";
  /** Paragraph IDs to convert into list items. */
  targets: string[];
  /** List type to apply. */
  listType: "bullet" | "ordered" | "multilevel";
  /** Indentation level for all targets. Default: 0. */
  level?: number;
}

export interface RemoveListPatch {
  type: "remove_list";
  /** Paragraph IDs to strip list formatting from (convert back to plain paragraphs). */
  targets: string[];
}

export interface SetListPropertiesPatch {
  type: "set_list_properties";
  /** Paragraph ID of the list item to modify. */
  target: string;
  /** Change list type. Creates a new numbering definition if needed. */
  listType?: "bullet" | "ordered" | "multilevel";
  /** Change indentation level. */
  level?: number;
}

// =============================================================================
// Patch Types
// =============================================================================

export type Patch =
  | ReplaceTextPatch
  | InsertTextPatch
  | DeleteTextPatch
  | AppendTextPatch
  | InsertParagraphPatch
  | DeleteParagraphPatch
  | AddCommentPatch
  | AddTablePatch
  | AddTableRowPatch
  | DeleteTableRowPatch
  | ReplaceCellTextPatch
  | ApplyFormattingPatch
  | ApplyStylePatch
  | AddHeaderPatch
  | AddFooterPatch
  | RemoveHeaderPatch
  | RemoveFooterPatch
  | AddListPatch
  | AddListItemPatch
  | RemoveListItemPatch
  | ConvertToListPatch
  | RemoveListPatch
  | SetListPropertiesPatch;

export interface InsertParagraphPatch {
  type: "insert_paragraph";
  /** Target paragraph ID (positional `p12` or native `para:4F2A1C7B`). */
  target: string;
  /** Insert before or after the target paragraph. */
  position: "before" | "after";
  /**
   * Text for the new paragraph(s). Supports inline formatting markers and
   * {{FIELD}} placeholders. `\n` splits into multiple paragraphs.
   * Markdown tables and lists are NOT supported here — use `add_table` /
   * `add_list` with `afterId` instead.
   */
  text: string;
  /**
   * Copy the target's paragraph properties (`w:pPr`: style, numbering,
   * indentation, spacing) and the run formatting of its last run onto the
   * new paragraph(s). Default: true.
   */
  inheritStyle?: boolean;
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface DeleteParagraphPatch {
  type: "delete_paragraph";
  /** Target paragraph ID (positional or `para:` form). */
  target: string;
}

export interface AddCommentPatch {
  type: "add_comment";
  /** Target paragraph ID (positional or `para:` form). */
  target: string;
  /** Optional substring to anchor the comment to. If omitted, the whole paragraph is commented. */
  find?: string;
  /** Which occurrence of `find` (1-based). Default: 1. */
  occurrence?: number;
  /** Comment text. `\n` splits into multiple paragraphs inside the comment. */
  text: string;
  /** Comment author. Falls back to `revisionAuthor` from PatchOptions, then "docxtor". */
  author?: string;
  /** Author initials shown in Word's balloon. Derived from `author` if omitted. */
  initials?: string;
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

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
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
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
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface DeleteTextPatch {
  type: "delete_text";
  /** Target paragraph ID. */
  target: string;
  /** Text to delete. */
  find: string;
  /** Which occurrence to delete (1-based). Default: 1. */
  occurrence?: number;
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface AppendTextPatch {
  type: "append_text";
  /** Text for the new paragraph appended to the end of the document body.
   *  Supports markdown table syntax — pipe-delimited rows with a separator
   *  line are automatically converted to a real OOXML table.
   *  Supports {{FIELD}} placeholders (e.g., {{PAGE}}, {{NUMPAGES}}). */
  text: string;
  /** Target a header/footer instead of the document body. */
  container?: HeaderFooterContainer;
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface AddTablePatch {
  type: "add_table";
  /** Element ID to insert after. If omitted, appends to end of document. */
  afterId?: string;
  /** Optional header row cells. */
  headers?: string[];
  /** Data rows. Each sub-array is one row of cell text values. */
  rows: string[][];
  /** Target a header/footer instead of the document body. */
  container?: HeaderFooterContainer;
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
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface ApplyFormattingPatch {
  type: "apply_formatting";
  /** Target paragraph ID. */
  target: string;
  /**
   * Exact substring to apply formatting to.
   * If omitted, formatting is applied to all runs in the paragraph.
   */
  find?: string;
  /** Which occurrence of `find` (1-based). Default: 1. Ignored when `find` is omitted. */
  occurrence?: number;
  /** The formatting properties to set or remove. */
  formatting: FormattingUpdate;
}

/**
 * Inline formatting properties to set or remove on run(s).
 *
 * - `true` / string / number → set the property (add or update w:rPr child element)
 * - `false` / `null` → remove the property element from w:rPr
 * - Property not mentioned → leave unchanged
 *
 * Only touches inline w:rPr on targeted runs. Never modifies global styles.
 */
export interface FormattingUpdate {
  bold?: boolean | null;
  italic?: boolean | null;
  /** Underline style: "single", "double", "wave", "thick", "dash", "dotted", etc. */
  underline?: string | null;
  strikethrough?: boolean | null;
  /** Font size in half-points (e.g., 24 = 12pt). */
  fontSize?: number | null;
  /** Font family name (e.g., "Arial"). Sets both w:ascii and w:hAnsi on w:rFonts. */
  fontFamily?: string | null;
  /** Hex color (e.g., "FF0000" for red). */
  color?: string | null;
  /** Highlight color name (e.g., "yellow", "green", "cyan"). */
  highlight?: string | null;
}

export interface ApplyStylePatch {
  type: "apply_style";
  /** Target paragraph ID. */
  target: string;
  /**
   * Exact substring to apply a character style to.
   * If omitted, a paragraph style is applied (w:pStyle in w:pPr).
   * If provided, a character style is applied (w:rStyle in w:rPr) to matched runs.
   */
  find?: string;
  /** Which occurrence of `find` (1-based). Default: 1. Ignored when `find` is omitted. */
  occurrence?: number;
  /**
   * Style ID that must already exist in the document's styles.xml.
   * For paragraph styles (no find): must be a "paragraph" type style.
   * For character styles (with find): must be a "character" type style.
   */
  style: string;
}

// =============================================================================
// Header/Footer Patch Types
// =============================================================================

/** Variant for header/footer targeting. */
export type HeaderFooterVariant = "default" | "first" | "even";

export interface AddHeaderPatch {
  type: "add_header";
  /** Which header variant to create. */
  variant: HeaderFooterVariant;
  /** Text content for each paragraph. Supports inline formatting markers and {{FIELD}} placeholders. */
  paragraphs: string[];
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface AddFooterPatch {
  type: "add_footer";
  /** Which footer variant to create. */
  variant: HeaderFooterVariant;
  /** Text content for each paragraph. Supports inline formatting markers and {{FIELD}} placeholders. */
  paragraphs: string[];
  /**
   * Parse inline formatting markers in this patch's text fields. Overrides
   * `PatchOptions.inlineMarkers`. Set `false` to write the text verbatim.
   */
  markers?: boolean;
}

export interface RemoveHeaderPatch {
  type: "remove_header";
  /** Which header variant to remove. */
  variant: HeaderFooterVariant;
}

export interface RemoveFooterPatch {
  type: "remove_footer";
  /** Which footer variant to remove. */
  variant: HeaderFooterVariant;
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
  /** Include revision marks (tracked changes) in the projection. Default: false. */
  includeRevisions?: boolean;
  /** Include page break elements and hints in the projection. Default: false. */
  includePageBreaks?: boolean;
  /**
   * JSON projection only. When set, every paragraph's `formatting` array
   * contains exactly the runs matching the filter (with their offsets), and
   * nothing else. Implies `includeFormatting`. Use it to read, e.g., answers
   * a client typed in red: `{ runFilter: { color: "FF0000" } }`.
   */
  runFilter?: RunFilter;
}

// =============================================================================
// Patch Options (Revision Tracking)
// =============================================================================

/**
 * Options for controlling revision tracking when applying patches.
 */
export interface PatchOptions {
  /** Whether to generate OOXML tracked changes (w:ins, w:del, etc.). Default: false. */
  trackRevisions?: boolean;
  /** The author name for revision marks. Default: "docxtor". */
  revisionAuthor?: string;
  /**
   * The revision date (ISO 8601, e.g. "2026-10-13T00:00:00Z") stamped on every
   * revision mark. Default: current time. Pin it for byte-stable output.
   */
  revisionDate?: string;
  /**
   * All-or-nothing batch. Patches are applied against a deep clone of the
   * document; if every patch succeeds the clone replaces the document,
   * otherwise the document is left untouched. The returned results still
   * describe every patch (a `success: true` entry means the patch would have
   * applied, not that it did). Default: false.
   */
  atomic?: boolean;
  /**
   * Parse GFM inline formatting markers (`**bold**`, `_italic_`,
   * `~~strike~~`, `<u>underline</u>`) and backslash escapes in patch text
   * fields (`find`, `replace`, `anchor`, `text`, list items, header and
   * footer paragraphs, comment text). Default: true. Set `false` to write
   * and match text verbatim; a patch's own `markers` field overrides this.
   */
  inlineMarkers?: boolean;
}

// =============================================================================
// Revisions (accept / reject)
// =============================================================================

export interface RevisionFilterOptions {
  /** Only process revisions by these authors. Omit to process every revision. */
  authors?: string[];
  /**
   * Only process revisions whose `w:id` is in this list (combined with
   * `authors` when both are given). Ids come from `RevisionMark.id` in a
   * projection with `includeRevisions`. A replacement is usually two
   * revisions, a `w:del` and a `w:ins`, each with its own id.
   */
  ids?: string[];
}

// =============================================================================
// Document metadata (docProps)
// =============================================================================

/** Core document properties from docProps/core.xml. */
export interface CoreProperties {
  /** dc:creator */
  creator?: string;
  /** cp:lastModifiedBy */
  lastModifiedBy?: string;
  /** dc:title */
  title?: string;
  /** dc:subject */
  subject?: string;
  /** dc:description */
  description?: string;
  /** cp:keywords */
  keywords?: string;
  /** cp:category */
  category?: string;
  /** cp:contentStatus */
  contentStatus?: string;
  /** cp:revision */
  revision?: string;
  /** dcterms:created (W3CDTF) */
  created?: string;
  /** dcterms:modified (W3CDTF) */
  modified?: string;
}

// =============================================================================
// Public API Options
// =============================================================================

export interface LoadOptions {
  /** Whether to normalize headers and footers. Default: true. */
  includeHeadersFooters?: boolean;
  /**
   * Stamp a `w14:paraId` (and `w14:textId`) on every paragraph that lacks one,
   * in the body, headers, and footers. Declares the `w14` namespace and adds
   * it to `mc:Ignorable` on each part's root if missing. Default: false.
   */
  assignParaIds?: boolean;
}
