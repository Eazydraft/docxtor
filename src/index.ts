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
} from "./types.js";
import { readDocx, type DocxZipEntries } from "./zip/reader.js";
import { writeDocx } from "./zip/writer.js";
import { serializeXml } from "./xml/serializer.js";
import { parseFlatOpc } from "./flat-opc/parser.js";
import { serializeFlatOpc } from "./flat-opc/serializer.js";
import type { FlatOpcPart } from "./flat-opc/types.js";
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
  /** Original Flat OPC parts, set only when loaded via loadFlatOpc(). */
  private flatOpcParts: FlatOpcPart[] | null = null;

  private constructor(
    zipEntries: DocxZipEntries,
    doc: NormalizedDocument,
    flatOpcParts?: FlatOpcPart[],
  ) {
    this.zipEntries = zipEntries;
    this.doc = doc;
    this.flatOpcParts = flatOpcParts ?? null;
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

  /**
   * Load a DOCX document from a Flat OPC XML string.
   *
   * Flat OPC is the format returned by Word's JavaScript API `getOoxml()`
   * method. It wraps all document parts in a single XML envelope using
   * `<pkg:package>` / `<pkg:part>` elements.
   *
   * @param xml - The Flat OPC XML string
   * @param options - Optional load options
   */
  static loadFlatOpc(xml: string, options?: LoadOptions): Docxtor {
    const { entries, parts } = parseFlatOpc(xml);
    const doc = normalizeDocument(entries, options);
    return new Docxtor(entries, doc, parts);
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

    // Serialize rels and content types if they were parsed
    const documentRelsXml = this.doc.xmlParts.documentRels
      ? serializeXml(this.doc.xmlParts.documentRels)
      : undefined;
    const contentTypesXml = this.doc.xmlParts.contentTypes
      ? serializeXml(this.doc.xmlParts.contentTypes)
      : undefined;
    const settingsXml = this.doc.xmlParts.settings
      ? serializeXml(this.doc.xmlParts.settings)
      : undefined;

    return writeDocx({
      zip: this.zipEntries.zip,
      documentXml,
      headers: headers.size > 0 ? headers : undefined,
      footers: footers.size > 0 ? footers : undefined,
      documentRelsXml,
      contentTypesXml,
      settingsXml,
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

  /**
   * Export the modified document as a Flat OPC XML string.
   *
   * This produces a string suitable for Word's `insertOoxml()` API.
   * Only works if the document was loaded via `loadFlatOpc()`, because
   * the original Flat OPC parts (including binary data like images)
   * are needed to reconstruct the full package.
   *
   * For documents loaded from a ZIP buffer, use `toBuffer()` instead.
   *
   * @throws Error if the document was not loaded from Flat OPC
   */
  toFlatOpc(): string {
    if (!this.flatOpcParts) {
      throw new Error(
        "toFlatOpc() is only available for documents loaded via loadFlatOpc(). " +
          "Use toBuffer() for ZIP-based documents.",
      );
    }

    // Serialize current XML state for parts that may have been modified
    const documentXml = serializeXml(this.doc.xmlParts.document);

    const headerXmls = new Map<string, string>();
    for (const [path, fxp] of this.doc.xmlParts.headers) {
      headerXmls.set(path, serializeXml(fxp));
    }

    const footerXmls = new Map<string, string>();
    for (const [path, fxp] of this.doc.xmlParts.footers) {
      footerXmls.set(path, serializeXml(fxp));
    }

    const documentRelsXml = this.doc.xmlParts.documentRels
      ? serializeXml(this.doc.xmlParts.documentRels)
      : undefined;
    const contentTypesXml = this.doc.xmlParts.contentTypes
      ? serializeXml(this.doc.xmlParts.contentTypes)
      : undefined;
    const settingsXml = this.doc.xmlParts.settings
      ? serializeXml(this.doc.xmlParts.settings)
      : undefined;

    // Build updated parts array: replace XML content for modified parts,
    // pass through unmodified parts (including binary) as-is.
    const updatedParts: FlatOpcPart[] = this.flatOpcParts.map((part) => {
      const name = part.name.startsWith("/") ? part.name.slice(1) : part.name;

      if (name === "word/document.xml") {
        return { ...part, xmlContent: documentXml };
      }
      if (name === "word/_rels/document.xml.rels" && documentRelsXml) {
        return { ...part, xmlContent: documentRelsXml };
      }
      if (name === "[Content_Types].xml" && contentTypesXml) {
        return { ...part, xmlContent: contentTypesXml };
      }
      if (name === "word/settings.xml" && settingsXml) {
        return { ...part, xmlContent: settingsXml };
      }
      if (headerXmls.has(name)) {
        return { ...part, xmlContent: headerXmls.get(name)! };
      }
      if (footerXmls.has(name)) {
        return { ...part, xmlContent: footerXmls.get(name)! };
      }

      // Unmodified part (styles, numbering, theme, images, etc.)
      return part;
    });

    // Handle newly added header/footer parts that weren't in the original
    const existingNames = new Set(
      this.flatOpcParts.map((p) =>
        p.name.startsWith("/") ? p.name.slice(1) : p.name,
      ),
    );

    for (const [path, xml] of headerXmls) {
      if (!existingNames.has(path)) {
        updatedParts.push({
          name: "/" + path,
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
          xmlContent: xml,
          binaryData: undefined,
        });
      }
    }

    for (const [path, xml] of footerXmls) {
      if (!existingNames.has(path)) {
        updatedParts.push({
          name: "/" + path,
          contentType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
          xmlContent: xml,
          binaryData: undefined,
        });
      }
    }

    return serializeFlatOpc(updatedParts);
  }

  /**
   * Check whether this document was loaded from Flat OPC format.
   */
  isFromFlatOpc(): boolean {
    return this.flatOpcParts !== null;
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
  ListInfo,
  HeaderFooter,
  BodyElement,
  Patch,
  ReplaceTextPatch,
  InsertTextPatch,
  DeleteTextPatch,
  AppendTextPatch,
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
  HeaderFooterVariant,
  FormattingUpdate,
  StyleInfo,
  PatchResult,
  PatchOptions,
  RevisionMark,
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
export type { HeaderFooterRef } from "./normalizer/document.js";
export { applyPatches } from "./patches/engine.js";
export { project, projectToJson, projectToMarkdown } from "./projection/llm.js";
export type {
  JsonProjection,
  JsonProjectionPageBreak,
} from "./projection/llm.js";
export {
  validatePatch,
  findParagraphById,
  findTableById,
} from "./patches/validator.js";
export type { ValidationResult } from "./patches/validator.js";
export { parseFlatOpc } from "./flat-opc/parser.js";
export { serializeFlatOpc } from "./flat-opc/serializer.js";
export type { FlatOpcPart } from "./flat-opc/types.js";
