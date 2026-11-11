/**
 * docxtor — Surgical, LLM-friendly DOCX editing.
 *
 * @example
 * ```ts
 * import { Docxtor } from "@lucio-ai/docxtor";
 * import { readFile, writeFile } from "fs/promises";
 *
 * // Load a document
 * const buffer = await readFile("contract.docx");
 * const doc = await Docxtor.load(buffer);
 *
 * // Get an LLM-friendly projection
 * const projection = doc.project("markdown");
 * // → "[p0] (Heading1) Service Agreement\n[p1] This agreement is valid..."
 *
 * // Apply patches (e.g., from LLM output)
 * const results = doc.applyPatches([
 *   { type: "replace_text", target: "p1", find: "valid", replace: "enforceable" },
 * ]);
 *
 * // Export the modified document
 * const output = await doc.toBuffer();
 * await writeFile("contract-edited.docx", output);
 * ```
 *
 * @packageDocumentation
 */

import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import type {
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  Patch,
  PatchResult,
  PatchOptions,
  ProjectionFormat,
  ProjectionOptions,
  LoadOptions,
  BodyElement,
  HeaderFooter,
  StyleInfo,
  CoreProperties,
  RevisionFilterOptions,
  FXPNode,
} from "./types.js";
import { readDocx, type DocxZipEntries } from "./zip/reader.js";
import { writeDocx } from "./zip/writer.js";
import { serializeXml } from "./xml/serializer.js";
import { normalizeDocument, reNormalizeDocument } from "./normalizer/document.js";
import { applyPatches } from "./patches/engine.js";
import { project, type JsonProjection } from "./projection/llm.js";
import {
  findParagraphById as _findParagraphById,
  findParagraphByParaId as _findParagraphByParaId,
  findTableById as _findTableById,
  validatePatch as _validatePatch,
} from "./patches/validator.js";
import {
  acceptAllRevisions as _acceptAllRevisions,
  rejectAllRevisions as _rejectAllRevisions,
} from "./revisions/accept-reject.js";
import {
  getCoreProperties as _getCoreProperties,
  setCoreProperties as _setCoreProperties,
  scrubMetadata as _scrubMetadata,
} from "./metadata/properties.js";

/**
 * The main entry point for docxtor.
 *
 * A Docxtor instance represents a loaded DOCX document with:
 * - A normalized AST for inspecting document structure
 * - An LLM projection system for generating simplified views
 * - A patch engine for applying surgical edits
 * - Export capabilities to produce modified DOCX files
 */
export class Docxtor {
  private zipEntries: DocxZipEntries;
  private doc: NormalizedDocument;
  /** Zip paths to drop on the next export (e.g., docProps/custom.xml after a scrub). */
  private removedParts = new Set<string>();

  private constructor(
    zipEntries: DocxZipEntries,
    doc: NormalizedDocument,
  ) {
    this.zipEntries = zipEntries;
    this.doc = doc;
  }

  /**
   * Load a DOCX document from a buffer or file path.
   */
  static async load(
    input: Buffer | ArrayBuffer | Uint8Array | string,
    options?: LoadOptions,
  ): Promise<Docxtor> {
    let buffer: Buffer | ArrayBuffer | Uint8Array;

    if (typeof input === "string") {
      buffer = await readFile(input);
    } else {
      buffer = input;
    }

    const zipEntries = await readDocx(buffer);
    const doc = normalizeDocument(zipEntries, options);

    return new Docxtor(zipEntries, doc);
  }

  /**
   * Create a new Docxtor instance.
   *
   * If `input` is provided, delegates to {@link Docxtor.load}.
   * Otherwise, creates a valid empty DOCX document in memory.
   *
   * @param input - Optional buffer, ArrayBuffer, Uint8Array, or file path to load from
   * @param options - Optional load options
   *
   * @example
   * ```ts
   * // Create a blank document
   * const doc = await Docxtor.create();
   *
   * // Equivalent to Docxtor.load(buffer)
   * const doc2 = await Docxtor.create(buffer);
   * ```
   */
  static async create(
    input?: Buffer | ArrayBuffer | Uint8Array | string,
    options?: LoadOptions,
  ): Promise<Docxtor> {
    if (input !== undefined) {
      return Docxtor.load(input, options);
    }

    // Build a minimal valid DOCX in memory
    const zip = new JSZip();

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

    const topRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t></w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", topRels);
    zip.file("word/document.xml", documentXml);

    const buffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    });

    return Docxtor.load(buffer, options);
  }

  // ===========================================================================
  // Projection
  // ===========================================================================

  /**
   * Generate an LLM-friendly projection of the document.
   *
   * @param format - "json" for structured data, "markdown" for text
   * @param options - Control what's included in the projection
   */
  project(format?: "json", options?: ProjectionOptions): JsonProjection;
  project(format: "markdown", options?: ProjectionOptions): string;
  project(
    format: ProjectionFormat = "json",
    options: ProjectionOptions = {},
  ): JsonProjection | string {
    return project(this.doc, format, options);
  }

  // ===========================================================================
  // Patching
  // ===========================================================================

  /**
   * Apply patches to the document.
   *
   * Patches are applied sequentially. After each patch, affected elements
   * are re-normalized so subsequent patches see updated text/mappings.
   *
   * @param patches - Array of patches to apply
   * @param options - Optional options (e.g., revision tracking)
   * @returns Array of results, one per patch
   */
  applyPatches(patches: Patch[], options?: PatchOptions): PatchResult[] {
    return applyPatches(this.doc, patches, options);
  }

  // ===========================================================================
  // Revisions
  // ===========================================================================

  /**
   * Accept every tracked change (optionally only those by `authors`):
   * insertions become plain content, deletions are removed, property changes
   * are dropped, deleted rows and paragraphs disappear.
   *
   * Positional ids are reassigned afterwards; `paraId`s are unchanged.
   */
  acceptAllRevisions(options?: RevisionFilterOptions): { accepted: number } {
    const result = _acceptAllRevisions(this.doc, options);
    reNormalizeDocument(this.doc);
    return result;
  }

  /**
   * Reject every tracked change (optionally only those by `authors`):
   * insertions are removed, deletions restored, old properties put back,
   * inserted rows and paragraphs disappear.
   *
   * Positional ids are reassigned afterwards; `paraId`s are unchanged.
   */
  rejectAllRevisions(options?: RevisionFilterOptions): { rejected: number } {
    const result = _rejectAllRevisions(this.doc, options);
    reNormalizeDocument(this.doc);
    return result;
  }

  // ===========================================================================
  // Metadata
  // ===========================================================================

  /** Read docProps/core.xml properties. */
  getCoreProperties(): CoreProperties {
    return _getCoreProperties(this.doc);
  }

  /** Set docProps/core.xml properties (creating the part if needed). */
  setCoreProperties(props: Partial<CoreProperties>): void {
    _setCoreProperties(this.doc, props);
  }

  /**
   * Strip identifying metadata: clears creator, lastModifiedBy, description,
   * keywords, category and contentStatus, sets revision to "1", clears
   * Company / Manager / HyperlinkBase in app.xml, and removes
   * docProps/custom.xml. Title, created and modified are left untouched.
   */
  scrubMetadata(): void {
    const { removedParts } = _scrubMetadata(this.doc);
    for (const path of removedParts) this.removedParts.add(path);
  }

  // ===========================================================================
  // Export
  // ===========================================================================

  /**
   * Export the current document as a raw OOXML string (word/document.xml content).
   *
   * Returns the serialized XML for the document body, reflecting any patches
   * that have been applied. This is the same XML that gets written into the
   * DOCX ZIP as word/document.xml.
   */
  toString(): string {
    return serializeXml(this.doc.xmlParts.document);
  }

  /**
   * Export the modified document as a Buffer.
   */
  async toBuffer(): Promise<Buffer> {
    const documentXml = serializeXml(this.doc.xmlParts.document);

    const headers = new Map<string, string>();
    for (const [path, fxp] of this.doc.xmlParts.headers) {
      headers.set(path, serializeXml(fxp));
    }

    const footers = new Map<string, string>();
    for (const [path, fxp] of this.doc.xmlParts.footers) {
      footers.set(path, serializeXml(fxp));
    }

    // Determine which header/footer zip entries should be removed from the
    // current in-memory zip state. This handles multi-export flows where a
    // part may be added in one export and removed in a later export.
    const removedParts: string[] = [];
    for (const path of Object.keys(this.zipEntries.zip.files)) {
      if (/^word\/header\d+\.xml$/i.test(path)) {
        if (!this.doc.xmlParts.headers.has(path)) {
          removedParts.push(path);
        }
      } else if (/^word\/footer\d+\.xml$/i.test(path)) {
        if (!this.doc.xmlParts.footers.has(path)) {
          removedParts.push(path);
        }
      }
    }

    for (const path of this.removedParts) {
      if (!removedParts.includes(path)) removedParts.push(path);
    }

    // Serialize every optional part that was parsed (or created on demand)
    const serialize = (part: FXPNode[] | undefined): string | undefined =>
      part ? serializeXml(part) : undefined;
    const parts = this.doc.xmlParts;

    return writeDocx({
      zip: this.zipEntries.zip,
      documentXml,
      headers: headers.size > 0 ? headers : undefined,
      footers: footers.size > 0 ? footers : undefined,
      documentRelsXml: serialize(parts.documentRels),
      contentTypesXml: serialize(parts.contentTypes),
      settingsXml: serialize(parts.settings),
      numberingXml: serialize(parts.numbering),
      footnotesXml: serialize(parts.footnotes),
      endnotesXml: serialize(parts.endnotes),
      commentsXml: serialize(parts.comments),
      coreXml: serialize(parts.core),
      appXml: serialize(parts.app),
      rootRelsXml: serialize(parts.rootRels),
      removedParts: removedParts.length > 0 ? removedParts : undefined,
    });
  }

  /**
   * Export the modified document to a file.
   */
  async toFile(filePath: string): Promise<void> {
    const { writeFile: fsWriteFile } = await import("node:fs/promises");
    const buffer = await this.toBuffer();
    await fsWriteFile(filePath, buffer);
  }

  // ===========================================================================
  // Inspection
  // ===========================================================================

  /**
   * Get the full normalized AST.
   */
  getAST(): NormalizedDocument {
    return this.doc;
  }

  /**
   * Get all body-level elements (paragraphs and tables).
   */
  getBody(): BodyElement[] {
    return this.doc.body;
  }

  /**
   * Get all paragraphs in the document body (not inside tables).
   */
  getParagraphs(): NormalizedParagraph[] {
    return this.doc.body.filter(
      (e): e is NormalizedParagraph => e.kind === "paragraph",
    );
  }

  /**
   * Get all tables in the document body.
   */
  getTables(): NormalizedTable[] {
    return this.doc.body.filter(
      (e): e is NormalizedTable => e.kind === "table",
    );
  }

  /**
   * Get all headers.
   */
  getHeaders(): Map<string, HeaderFooter> {
    return this.doc.headers;
  }

  /**
   * Get all footers.
   */
  getFooters(): Map<string, HeaderFooter> {
    return this.doc.footers;
  }

  /**
   * Find a paragraph by its ID anywhere in the document
   * (body, tables, headers, footers, nested tables).
   */
  findParagraphById(id: string): NormalizedParagraph | undefined {
    return _findParagraphById(id, this.doc);
  }

  /**
   * Find a paragraph by its native `w14:paraId` (8 hex chars, case-insensitive).
   * Resolution order: body (including table cells), headers, footers.
   */
  findParagraphByParaId(paraId: string): NormalizedParagraph | undefined {
    return _findParagraphByParaId(paraId, this.doc);
  }

  /**
   * Find a table by its ID anywhere in the document body
   * (including nested tables).
   */
  findTableById(id: string): NormalizedTable | undefined {
    return _findTableById(id, this.doc);
  }

  /**
   * Get available styles from the document's styles.xml.
   * Useful for knowing which styles can be applied via apply_style patches.
   */
  getStyles(): StyleInfo[] {
    return this.doc.styles;
  }

  /**
   * Validate patches without applying them.
   * Returns validation results for each patch.
   */
  validatePatches(patches: Patch[]): { valid: boolean; error?: string }[] {
    return patches.map((patch) => _validatePatch(patch, this.doc));
  }
}

// Re-export types
export type {
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  NormalizedPageBreak,
  NormalizedTableRow,
  NormalizedTableCell,
  RunMapping,
  RunFormatting,
  RunFilter,
  ListInfo,
  HeaderFooter,
  BodyElement,
  Patch,
  ReplaceTextPatch,
  InsertTextPatch,
  DeleteTextPatch,
  AppendTextPatch,
  InsertParagraphPatch,
  DeleteParagraphPatch,
  AddCommentPatch,
  AddTablePatch,
  AddTableRowPatch,
  DeleteTableRowPatch,
  ReplaceCellTextPatch,
  ApplyFormattingPatch,
  ApplyStylePatch,
  AddHeaderPatch,
  AddFooterPatch,
  RemoveHeaderPatch,
  RemoveFooterPatch,
  AddListPatch,
  AddListItemPatch,
  RemoveListItemPatch,
  ConvertToListPatch,
  RemoveListPatch,
  SetListPropertiesPatch,
  HeaderFooterVariant,
  HeaderFooterContainer,
  HeaderFooterReference,
  FormattingUpdate,
  StyleInfo,
  PatchResult,
  PatchOptions,
  RevisionMark,
  RevisionFilterOptions,
  CoreProperties,
  ProjectionFormat,
  ProjectionOptions,
  LoadOptions,
  FXPNode,
} from "./types.js";

// Re-export low-level utilities for advanced usage
export { readDocx } from "./zip/reader.js";
export { writeDocx } from "./zip/writer.js";
export { parseXml } from "./xml/parser.js";
export { serializeXml } from "./xml/serializer.js";
export { normalizeDocument, reNormalizeDocument } from "./normalizer/document.js";
export type { HeaderFooterRef } from "./normalizer/document.js";
export { applyPatches, cloneDocument } from "./patches/engine.js";
export { project, projectToJson, projectToMarkdown } from "./projection/llm.js";
export type { JsonProjection, JsonProjectionPageBreak, JsonProjectionRunFormatting } from "./projection/llm.js";
export { validatePatch, findParagraphById, findParagraphByParaId, findTableById } from "./patches/validator.js";
export type { ValidationResult } from "./patches/validator.js";
export { acceptAllRevisions, rejectAllRevisions } from "./revisions/accept-reject.js";
export { getCoreProperties, setCoreProperties, scrubMetadata } from "./metadata/properties.js";
