/**
 * docxtor — Surgical, LLM-friendly DOCX editing.
 *
 * @example
 * ```ts
 * import { Docxtor } from "docxtor";
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
import type {
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  Patch,
  PatchResult,
  ProjectionFormat,
  ProjectionOptions,
  LoadOptions,
  BodyElement,
  HeaderFooter,
} from "./types.js";
import { readDocx, type DocxZipEntries } from "./zip/reader.js";
import { writeDocx } from "./zip/writer.js";
import { serializeXml } from "./xml/serializer.js";
import { normalizeDocument } from "./normalizer/document.js";
import { applyPatches } from "./patches/engine.js";
import { project, type JsonProjection } from "./projection/llm.js";
import {
  findParagraphById as _findParagraphById,
  findTableById as _findTableById,
  validatePatch as _validatePatch,
} from "./patches/validator.js";

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
   * @returns Array of results, one per patch
   */
  applyPatches(patches: Patch[]): PatchResult[] {
    return applyPatches(this.doc, patches);
  }

  // ===========================================================================
  // Export
  // ===========================================================================

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

    return writeDocx({
      zip: this.zipEntries.zip,
      documentXml,
      headers: headers.size > 0 ? headers : undefined,
      footers: footers.size > 0 ? footers : undefined,
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
   * Find a table by its ID anywhere in the document body
   * (including nested tables).
   */
  findTableById(id: string): NormalizedTable | undefined {
    return _findTableById(id, this.doc);
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
  NormalizedTableRow,
  NormalizedTableCell,
  RunMapping,
  RunFormatting,
  ListInfo,
  HeaderFooter,
  BodyElement,
  Patch,
  ReplaceTextPatch,
  InsertTextPatch,
  DeleteTextPatch,
  AddTableRowPatch,
  DeleteTableRowPatch,
  ReplaceCellTextPatch,
  PatchResult,
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
export { normalizeDocument } from "./normalizer/document.js";
export { applyPatches } from "./patches/engine.js";
export { project, projectToJson, projectToMarkdown } from "./projection/llm.js";
export type { JsonProjection } from "./projection/llm.js";
export { validatePatch, findParagraphById, findTableById } from "./patches/validator.js";
export type { ValidationResult } from "./patches/validator.js";
