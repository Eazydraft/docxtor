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

import JSZip from "jszip";
import type {
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  FXPNode,
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
import { normalizeDocument } from "./normalizer/document.js";
import { applyPatches } from "./patches/engine.js";
import { project, type JsonProjection } from "./projection/llm.js";
import {
  findParagraphById as _findParagraphById,
  findTableById as _findTableById,
  validatePatch as _validatePatch,
} from "./patches/validator.js";
import {
  isFlatOpc,
  parseFlatOpc,
  buildFlatOpc,
  buildDocumentFragment,
  opcPathToZipPath,
  zipPathToOpcPath,
  opcContentType,
  type FlatOpcPart,
} from "./ooxml/flat-opc.js";

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
  /** Stored when loaded via {@link fromOoxml} for lossless Flat OPC round-trips. */
  private flatOpcParts?: FlatOpcPart[];

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
      const { readFile } = await import("node:fs/promises");
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

    // Use "arraybuffer" in browser environments where Node Buffer is unavailable.
    const zipType = typeof Buffer !== "undefined" ? "nodebuffer" : "arraybuffer";
    const buffer = await zip.generateAsync({
      type: zipType as "nodebuffer",
      compression: "DEFLATE",
    });

    return Docxtor.load(buffer, options);
  }

  /**
   * Create a Docxtor instance from OOXML already extracted by Office.js.
   *
   * Accepts two formats:
   * - **Flat OPC XML** — the format returned by `Word.Body.getOoxml()` and
   *   accepted by `body.insertOoxml()`.  The root element is `<pkg:package>`.
   * - **Raw document XML** — the raw `<w:document>…</w:document>` string.
   *
   * Unlike {@link load}, this method is **synchronous** — ideal for
   * Office.js add-in code that already has the XML in memory.
   *
   * Use {@link toFlatOpc} to export the modified document back to Flat OPC
   * so it can be inserted via `body.insertOoxml()`.
   *
   * @example
   * ```ts
   * // Inside a Word add-in (Office.js)
   * await Word.run(async (context) => {
   *   const body = context.document.body;
   *   body.load("ooxml");
   *   await context.sync();
   *
   *   const doc = Docxtor.fromOoxml(body.getOoxml().value);
   *   doc.applyPatches([{ type: "replace_text", target: "p0", find: "foo", replace: "bar" }]);
   *
   *   body.insertOoxml(doc.toFlatOpc(), Word.InsertLocation.replace);
   *   await context.sync();
   * });
   * ```
   */
  static fromOoxml(xml: string, options?: LoadOptions): Docxtor {
    if (isFlatOpc(xml)) {
      return Docxtor._fromFlatOpc(xml, options);
    }
    return Docxtor._fromDocumentXml(xml, options);
  }

  /** @internal Load from a Flat OPC XML string. */
  private static _fromFlatOpc(xml: string, options?: LoadOptions): Docxtor {
    const parts = parseFlatOpc(xml);
    const zipEntries = _buildZipEntriesFromOpcParts(parts);
    const doc = normalizeDocument(zipEntries, options);
    const instance = new Docxtor(zipEntries, doc);
    instance.flatOpcParts = parts;
    return instance;
  }

  /** @internal Load from a raw `<w:document>` XML string. */
  private static _fromDocumentXml(
    documentXml: string,
    options?: LoadOptions,
  ): Docxtor {
    const zip = new JSZip();
    zip.file("word/document.xml", documentXml);

    const zipEntries: DocxZipEntries = {
      documentXml,
      headers: new Map(),
      footers: new Map(),
      zip,
    };

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
   * Export the modified document as a **Flat OPC XML** string.
   *
   * This is the format that Office.js accepts for `body.insertOoxml()`.
   * When the document was loaded via {@link fromOoxml} from Flat OPC, all
   * original parts (images, fonts, styles, etc.) are preserved — only the
   * modified XML parts (document, headers, footers, settings, rels) are
   * replaced.  When loaded from raw XML, a minimal single-part package is
   * produced.
   *
   * @example
   * ```ts
   * // Apply patches, then push back to Word
   * doc.applyPatches(patches);
   * body.insertOoxml(doc.toFlatOpc(), Word.InsertLocation.replace);
   * ```
   */
  toFlatOpc(): string {
    // Serialize all modified XML parts
    const overrides = new Map<string, string>();

    overrides.set("/word/document.xml", serializeXml(this.doc.xmlParts.document));

    for (const [zipPath, fxp] of this.doc.xmlParts.headers) {
      overrides.set(zipPathToOpcPath(zipPath), serializeXml(fxp));
    }
    for (const [zipPath, fxp] of this.doc.xmlParts.footers) {
      overrides.set(zipPathToOpcPath(zipPath), serializeXml(fxp));
    }
    if (this.doc.xmlParts.documentRels) {
      overrides.set(
        "/word/_rels/document.xml.rels",
        serializeXml(this.doc.xmlParts.documentRels),
      );
    }
    if (this.doc.xmlParts.contentTypes) {
      overrides.set("/[Content_Types].xml", serializeXml(this.doc.xmlParts.contentTypes));
    }
    if (this.doc.xmlParts.settings) {
      overrides.set("/word/settings.xml", serializeXml(this.doc.xmlParts.settings));
    }

    if (this.flatOpcParts) {
      // Merge overrides into the original parts for a lossless round-trip
      return buildFlatOpc(this.flatOpcParts, overrides);
    }

    // Build a minimal Flat OPC from the available parts
    const parts: FlatOpcPart[] = Array.from(overrides.entries()).map(
      ([name, content]) => ({
        name,
        contentType: opcContentType(name),
        content,
        isBinary: false,
      }),
    );
    return buildFlatOpc(parts);
  }

  /**
   * Convert patch results to a **Flat OPC XML** fragment for targeted
   * Office.js insertion.
   *
   * Only the elements affected by **successful** patches that have a `target`
   * ID are included (paragraphs and tables).  The fragment is a minimal
   * `<w:document>` with all namespace declarations preserved from the source
   * document, wrapped in a Flat OPC envelope so it can be passed directly to
   * `range.insertOoxml()` or `body.insertOoxml()`.
   *
   * Patches without a `target` (e.g. `append_text`, header/footer ops) do
   * not produce entries in the fragment — use {@link toFlatOpc} to export
   * the full modified document in those cases.
   *
   * @param results - The array returned by {@link applyPatches}.  Works for
   *   both tracked-change and plain patches.
   *
   * @example
   * ```ts
   * const results = doc.applyPatches(patches, { trackRevisions: true });
   * // Insert only the changed paragraphs/tables at the cursor
   * const ooxml = doc.patchResultsToOoxml(results);
   * range.insertOoxml(ooxml, Word.InsertLocation.replace);
   * ```
   */
  patchResultsToOoxml(results: PatchResult[]): string {
    const nodes = _collectPatchedNodes(results, this.doc);

    const docPart = this.flatOpcParts?.find(
      (p) => p.name === "/word/document.xml",
    );

    const fragmentXml = buildDocumentFragment(
      nodes,
      this.doc.xmlParts.document,
      serializeXml,
    );

    const part: FlatOpcPart = {
      name: "/word/document.xml",
      contentType:
        docPart?.contentType ??
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      content: fragmentXml,
      isBinary: false,
    };

    return buildFlatOpc([part]);
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

// =============================================================================
// Module-level helpers (not exported — internal to this file)
// =============================================================================

/**
 * Build a DocxZipEntries from Flat OPC parts.
 *
 * Populates a JSZip so the existing write machinery (`toBuffer`, `toFile`)
 * works even when the document was loaded from Flat OPC.  Binary parts are
 * added as base64.
 */
function _buildZipEntriesFromOpcParts(parts: FlatOpcPart[]): DocxZipEntries {
  const zip = new JSZip();
  const headers = new Map<string, string>();
  const footers = new Map<string, string>();
  let documentXml = "";
  let documentRels: string | undefined;
  let contentTypes: string | undefined;
  let stylesXml: string | undefined;
  let settingsXml: string | undefined;

  for (const part of parts) {
    const zipPath = opcPathToZipPath(part.name);

    if (part.isBinary) {
      zip.file(zipPath, part.content, { base64: true });
      continue;
    }

    zip.file(zipPath, part.content);

    const n = part.name; // OPC path with leading "/"
    if (n === "/word/document.xml") {
      documentXml = part.content;
    } else if (/^\/word\/header\d+\.xml$/i.test(n)) {
      headers.set(zipPath, part.content);
    } else if (/^\/word\/footer\d+\.xml$/i.test(n)) {
      footers.set(zipPath, part.content);
    } else if (n === "/word/_rels/document.xml.rels") {
      documentRels = part.content;
    } else if (n === "/[Content_Types].xml") {
      contentTypes = part.content;
    } else if (n === "/word/styles.xml") {
      stylesXml = part.content;
    } else if (n === "/word/settings.xml") {
      settingsXml = part.content;
    }
  }

  if (!documentXml) {
    throw new Error("Invalid OOXML: missing /word/document.xml part");
  }

  return {
    documentXml,
    headers,
    footers,
    documentRels,
    contentTypes,
    stylesXml,
    settingsXml,
    zip,
  };
}

/**
 * Collect the FXP nodes for all elements targeted by successful patches.
 * Each element ID is included at most once, preserving document order.
 */
function _collectPatchedNodes(
  results: PatchResult[],
  doc: NormalizedDocument,
): FXPNode[] {
  const seen = new Set<string>();
  const nodes: FXPNode[] = [];

  for (const result of results) {
    if (!result.success) continue;

    const patch = result.patch;
    let targetId: string | undefined;

    if ("target" in patch) {
      targetId = patch.target as string;
    }

    if (!targetId || seen.has(targetId)) continue;
    seen.add(targetId);

    const para = _findParagraphById(targetId, doc);
    if (para) {
      nodes.push(para.xmlRef);
      continue;
    }
    const table = _findTableById(targetId, doc);
    if (table) {
      nodes.push(table.xmlRef);
    }
  }

  return nodes;
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
export type { JsonProjection, JsonProjectionPageBreak } from "./projection/llm.js";
export { validatePatch, findParagraphById, findTableById } from "./patches/validator.js";
export type { ValidationResult } from "./patches/validator.js";
export {
  isFlatOpc,
  parseFlatOpc,
  buildFlatOpc,
  buildDocumentFragment,
  opcContentType,
  opcPathToZipPath,
  zipPathToOpcPath,
} from "./ooxml/flat-opc.js";
export type { FlatOpcPart } from "./ooxml/flat-opc.js";
