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
  contentChildren,
} from "../xml/helpers.js";
import { ControlResolver } from "./control.js";
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
    const relMap = parseRelationshipsFxp(doc.xmlParts.documentRels);
    const sectPrRefs = findHeaderFooterRefs(doc.xmlParts.document);

    for (const ref of sectPrRefs) {
      const target = relMap.get(ref.rId);
      if (!target) continue;

      const fullPath = `word/${target}`;

      if (ref.type === "header") {
        const fxp = doc.xmlParts.headers.get(fullPath);
        if (fxp) {
          doc.headers.set(ref.rId, normalizeHeaderFooter(fxp, "header", ref.variant, ref.rId));
        }
      } else {
        const fxp = doc.xmlParts.footers.get(fullPath);
        if (fxp) {
          doc.footers.set(ref.rId, normalizeHeaderFooter(fxp, "footer", ref.variant, ref.rId));
        }
      }
    }
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

  const resolver = new ControlResolver();
  let paraIdx = 0;
  let tableIdx = 0;
  let pageBreakIdx = 0;

  // Read through block-level content controls: a paragraph or table inside
  // one is a body element like any other.
  for (const { node: child, sdts } of contentChildren(bodyNode)) {
    const tag = getTagName(child);
    const control = resolver.resolve(sdts);
    if (tag === "w:p") {
      // Normalize once and reuse the result for both the page-break-only
      // check and the final elements.push — avoids duplicating the full
      // run mapping + revision scanning work.
      const normalized = normalizeParagraph(child, `p${paraIdx}`, control);

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
      elements.push(normalizeTable(child, `tbl${tableIdx}`, control, resolver));
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

export interface HeaderFooterRef {
  type: "header" | "footer";
  variant: string; // "default", "first", "even"
  rId: string;
}

/**
 * Find header/footer references in w:sectPr elements.
 */
export function findHeaderFooterRefs(documentFxp: FXPNode[]): HeaderFooterRef[] {
  const refs: HeaderFooterRef[] = [];

  const docNode = documentFxp.find((n) => getTagName(n) === "w:document");
  if (!docNode) return refs;

  const bodyNode = findChild(docNode, "w:body");
  if (!bodyNode) return refs;

  // sectPr can be a child of w:body (last section) or w:pPr (section break)
  const sectPrs = findChildren(bodyNode, "w:sectPr");

  // Also check paragraphs for an embedded sectPr (a section break)
  for (const { node: child } of contentChildren(bodyNode)) {
    if (getTagName(child) === "w:p") {
      const pPr = findChild(child, "w:pPr");
      if (pPr) {
        const sectPr = findChild(pPr, "w:sectPr");
        if (sectPr) sectPrs.push(sectPr);
      }
    }
  }

  for (const sectPr of sectPrs) {
    const children = getChildren(sectPr);
    for (const child of children) {
      const tag = getTagName(child);
      if (tag === "w:headerReference") {
        const rId = getAttr(child, "r:id");
        const variant = getAttr(child, "w:type") ?? "default";
        if (rId) refs.push({ type: "header", variant, rId });
      } else if (tag === "w:footerReference") {
        const rId = getAttr(child, "r:id");
        const variant = getAttr(child, "w:type") ?? "default";
        if (rId) refs.push({ type: "footer", variant, rId });
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
