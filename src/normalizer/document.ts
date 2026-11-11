/**
 * Document normalizer: orchestrates the full normalization pipeline.
 *
 * Takes the raw XML parts from the ZIP reader, parses them with
 * fast-xml-parser, and builds the full NormalizedDocument with
 * paragraphs, tables, headers, and footers — all holding xmlRef
 * pointers back to the mutable FXP objects.
 */

import type {
  FXPNode,
  NormalizedDocument,
  BodyElement,
  HeaderFooter,
  DocxXmlParts,
  LoadOptions,
  StyleInfo,
  NormalizedPageBreak,
  NormalizedParagraph,
} from "../types.js";
import type { DocxZipEntries } from "../zip/reader.js";
import { parseXml } from "../xml/parser.js";
import {
  findChild,
  findChildren,
  getChildren,
  getTagName,
  getAttr,
} from "../xml/helpers.js";
import { normalizeParagraph } from "./paragraph.js";
import { normalizeTable } from "./table.js";
import { normalizeHeaderFooter } from "./header-footer.js";
import { resolveListType } from "../patches/numbering.js";
import { assignParaIdsToDocument } from "./para-id.js";

/**
 * Normalize a full DOCX document from zip entries.
 *
 * Returns a NormalizedDocument where every element holds xmlRef pointers
 * into the FXP objects. The FXP objects are the mutable XML — patches
 * mutate them, and serializing them produces the output document.
 */
export function normalizeDocument(
  zipEntries: DocxZipEntries,
  options?: LoadOptions,
): NormalizedDocument {
  // Parse all XML parts
  const documentFxp = parseXml(zipEntries.documentXml);

  const headersFxp = new Map<string, FXPNode[]>();
  for (const [path, xml] of zipEntries.headers) {
    headersFxp.set(path, parseXml(xml));
  }

  const footersFxp = new Map<string, FXPNode[]>();
  for (const [path, xml] of zipEntries.footers) {
    footersFxp.set(path, parseXml(xml));
  }

  // Parse styles.xml if present
  const styles = zipEntries.stylesXml
    ? parseStyles(zipEntries.stylesXml)
    : [];

  const parseOptional = (xml: string | undefined): FXPNode[] | undefined =>
    xml ? parseXml(xml) : undefined;

  const xmlParts: DocxXmlParts = {
    document: documentFxp,
    headers: headersFxp,
    footers: footersFxp,
    documentRels: parseOptional(zipEntries.documentRels),
    contentTypes: parseOptional(zipEntries.contentTypes),
    settings: parseOptional(zipEntries.settingsXml),
    numbering: parseOptional(zipEntries.numberingXml),
    footnotes: parseOptional(zipEntries.footnotesXml),
    endnotes: parseOptional(zipEntries.endnotesXml),
    comments: parseOptional(zipEntries.commentsXml),
    core: parseOptional(zipEntries.coreXml),
    app: parseOptional(zipEntries.appXml),
    rootRels: parseOptional(zipEntries.rootRels),
  };

  // Stamp native paragraph ids before normalizing so they show up in the AST
  if (options?.assignParaIds) {
    assignParaIdsToDocument(xmlParts);
  }

  const doc: NormalizedDocument = {
    body: [],
    headers: new Map(),
    footers: new Map(),
    xmlParts,
    styles,
    loadOptions: options,
  };

  reNormalizeDocument(doc);

  return doc;
}

/**
 * Rebuild the normalized view (body, headers, footers, list types) of a
 * document from its current XML parts, in place.
 *
 * Call this after a structural change that invalidates positional ids
 * across the whole document — e.g., accepting/rejecting all revisions.
 * `paraId`s are read from the XML so they are unaffected.
 */
export function reNormalizeDocument(doc: NormalizedDocument): void {
  const includeHeadersFooters = doc.loadOptions?.includeHeadersFooters ?? true;

  doc.body = normalizeBody(doc.xmlParts.document);
  doc.headers = new Map<string, HeaderFooter>();
  doc.footers = new Map<string, HeaderFooter>();

  if (includeHeadersFooters) {
    normalizeHeadersFooters(doc);
  }

  // Post-pass: enrich listType on all paragraphs with listInfo
  enrichAllListTypes(doc);
}

/**
 * Normalize the document body (w:document > w:body children).
 */
function normalizeBody(documentFxp: FXPNode[]): BodyElement[] {
  const elements: BodyElement[] = [];

  // Find w:document > w:body
  const docNode = documentFxp.find((n) => getTagName(n) === "w:document");
  if (!docNode) return elements;

  const bodyNode = findChild(docNode, "w:body");
  if (!bodyNode) return elements;

  const bodyChildren = getChildren(bodyNode);
  let paraIdx = 0;
  let tableIdx = 0;
  let pageBreakIdx = 0;

  for (const child of bodyChildren) {
    const tag = getTagName(child);
    if (tag === "w:p") {
      // Normalize once and reuse the result for both the page-break-only
      // check and the final elements.push — avoids duplicating the full
      // run mapping + revision scanning work.
      const normalized = normalizeParagraph(child, `p${paraIdx}`);

      if (normalized.hasPageBreak && normalized.text.trim().length === 0) {
        // Page-break-only paragraph — emit as NormalizedPageBreak
        const pb: NormalizedPageBreak = {
          kind: "page_break",
          id: `pb${pageBreakIdx}`,
          xmlRef: child,
        };
        elements.push(pb);
        pageBreakIdx++;
        // Don't increment paraIdx — this isn't a paragraph in the projection
        continue;
      }

      elements.push(normalized);
      paraIdx++;
    } else if (tag === "w:tbl") {
      elements.push(normalizeTable(child, `tbl${tableIdx}`));
      tableIdx++;
    }
    // Other elements (w:sectPr, etc.) are ignored in the normalized view
    // but preserved in the FXP object
  }

  return elements;
}

/**
 * Build a relationship map (rId → target filename) from parsed
 * document.xml.rels nodes.
 */
function parseRelationshipsFxp(
  parsed: FXPNode[] | undefined,
): Map<string, string> {
  const map = new Map<string, string>();
  if (!parsed) return map;

  // Find the Relationships root element
  const relsNode = parsed.find(
    (n) => getTagName(n) === "Relationships",
  );
  if (!relsNode) return map;

  const rels = findChildren(relsNode, "Relationship");
  for (const rel of rels) {
    const id = getAttr(rel, "Id");
    const target = getAttr(rel, "Target");
    if (id && target) {
      map.set(id, target);
    }
  }

  return map;
}

/**
 * Rebuild `doc.headers` and `doc.footers` from the section properties and
 * the header/footer parts, in document order (section 1 first).
 *
 * In a document where one section has headers or footers the ids are
 * `hdr.default.p0`. When more than one section has them, every id carries
 * its section: `hdr.s1.first.p0`, `hdr.s2.first.p0`.
 */
export function normalizeHeadersFooters(doc: NormalizedDocument): void {
  doc.headers = new Map<string, HeaderFooter>();
  doc.footers = new Map<string, HeaderFooter>();

  const relMap = parseRelationshipsFxp(doc.xmlParts.documentRels);
  const resolved: { ref: HeaderFooterRef; fxp: FXPNode[] }[] = [];
  const seen = new Set<string>();
  for (const ref of findHeaderFooterRefs(doc.xmlParts.document)) {
    const target = relMap.get(ref.rId);
    if (!target) continue;
    const parts = ref.type === "header" ? doc.xmlParts.headers : doc.xmlParts.footers;
    const fxp = parts.get(`word/${target}`);
    // A part referenced again by a later section keeps its first section
    const key = `${ref.type}:${ref.rId}`;
    if (!fxp || seen.has(key)) continue;
    seen.add(key);
    resolved.push({ ref, fxp });
  }

  const withSection = new Set(resolved.map((r) => r.ref.sectionIndex)).size > 1;
  for (const { ref, fxp } of resolved) {
    const hf = normalizeHeaderFooter(fxp, ref.type, ref.variant, ref.rId, ref.sectionIndex, withSection);
    (ref.type === "header" ? doc.headers : doc.footers).set(ref.rId, hf);
  }
}

export interface HeaderFooterRef {
  type: "header" | "footer";
  variant: string; // "default", "first", "even"
  rId: string;
  /** 1-based index of the section, in document order. */
  sectionIndex: number;
}

/**
 * Find header/footer references in w:sectPr elements, in document order:
 * the sections that end at a section-break paragraph first, in body order,
 * then the last section (the w:sectPr that is a child of w:body).
 */
export function findHeaderFooterRefs(documentFxp: FXPNode[]): HeaderFooterRef[] {
  const refs: HeaderFooterRef[] = [];

  const docNode = documentFxp.find((n) => getTagName(n) === "w:document");
  if (!docNode) return refs;

  const bodyNode = findChild(docNode, "w:body");
  if (!bodyNode) return refs;

  // A section's properties sit in the w:pPr of its last paragraph (a section
  // break); the last section's are a child of w:body.
  const sectPrs: FXPNode[] = [];
  for (const child of getChildren(bodyNode)) {
    if (getTagName(child) === "w:p") {
      const pPr = findChild(child, "w:pPr");
      if (pPr) {
        const sectPr = findChild(pPr, "w:sectPr");
        if (sectPr) sectPrs.push(sectPr);
      }
    }
  }
  sectPrs.push(...findChildren(bodyNode, "w:sectPr"));

  for (const [i, sectPr] of sectPrs.entries()) {
    const sectionIndex = i + 1;
    const children = getChildren(sectPr);
    for (const child of children) {
      const tag = getTagName(child);
      if (tag === "w:headerReference") {
        const rId = getAttr(child, "r:id");
        const variant = getAttr(child, "w:type") ?? "default";
        if (rId) refs.push({ type: "header", variant, rId, sectionIndex });
      } else if (tag === "w:footerReference") {
        const rId = getAttr(child, "r:id");
        const variant = getAttr(child, "w:type") ?? "default";
        if (rId) refs.push({ type: "footer", variant, rId, sectionIndex });
      }
    }
  }

  return refs;
}

// =============================================================================
// Styles
// =============================================================================

/**
 * Parse word/styles.xml and extract style definitions.
 *
 * We only extract the style ID, human-readable name, and type.
 * We never modify styles.xml — this is read-only metadata used for
 * validation and projection.
 */
function parseStyles(stylesXml: string): StyleInfo[] {
  const parsed = parseXml(stylesXml);
  const styles: StyleInfo[] = [];

  // Find w:styles root
  const stylesNode = parsed.find((n) => getTagName(n) === "w:styles");
  if (!stylesNode) return styles;

  const styleNodes = findChildren(stylesNode, "w:style");
  for (const styleNode of styleNodes) {
    const id = getAttr(styleNode, "w:styleId");
    const type = getAttr(styleNode, "w:type") as StyleInfo["type"] | undefined;

    if (!id || !type) continue;

    // Valid types: paragraph, character, table, numbering
    if (!["paragraph", "character", "table", "numbering"].includes(type)) continue;

    // Get the human-readable name from w:name
    const nameNode = findChild(styleNode, "w:name");
    const name = nameNode ? (getAttr(nameNode, "w:val") ?? id) : id;

    styles.push({ id, name, type });
  }

  return styles;
}

// =============================================================================
// List type enrichment
// =============================================================================

/**
 * Enrich listType on all paragraphs with listInfo across the entire document.
 */
function enrichAllListTypes(doc: NormalizedDocument): void {
  const numberingFxp = doc.xmlParts.numbering;
  if (!numberingFxp) return;

  for (const el of doc.body) {
    if (el.kind === "paragraph") {
      enrichListType(el, numberingFxp);
    } else if (el.kind === "table") {
      for (const row of el.rows) {
        for (const cell of row.cells) {
          for (const para of cell.paragraphs) {
            enrichListType(para, numberingFxp);
          }
        }
      }
    }
  }

  for (const hdr of doc.headers.values()) {
    for (const para of hdr.paragraphs) {
      enrichListType(para, numberingFxp);
    }
  }

  for (const ftr of doc.footers.values()) {
    for (const para of ftr.paragraphs) {
      enrichListType(para, numberingFxp);
    }
  }
}

/**
 * Enrich a single paragraph's listInfo with the resolved listType.
 * Exported for use by the patch engine after re-normalization.
 */
export function enrichListType(para: NormalizedParagraph, numberingFxp: FXPNode[]): void {
  if (!para.listInfo) return;
  const lt = resolveListType(numberingFxp, para.listInfo.numId);
  if (lt) {
    para.listInfo.listType = lt;
  }
}
