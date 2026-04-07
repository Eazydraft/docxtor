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
  const includeHeadersFooters = options?.includeHeadersFooters ?? true;

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

  // Parse document relationships to map rId → filename
  const relMap = parseRelationships(zipEntries.documentRels);

  // Normalize body elements
  const body = normalizeBody(documentFxp);

  // Normalize headers and footers (unless disabled)
  const headers = new Map<string, HeaderFooter>();
  const footers = new Map<string, HeaderFooter>();

  if (includeHeadersFooters) {
    // Find section properties to discover header/footer references
    const sectPrRefs = findHeaderFooterRefs(documentFxp);

    for (const ref of sectPrRefs) {
      const target = relMap.get(ref.rId);
      if (!target) continue;

      const fullPath = `word/${target}`;

      if (ref.type === "header") {
        const fxp = headersFxp.get(fullPath);
        if (fxp) {
          const hdr = normalizeHeaderFooter(fxp, "header", ref.variant, ref.rId);
          headers.set(ref.rId, hdr);
        }
      } else {
        const fxp = footersFxp.get(fullPath);
        if (fxp) {
          const ftr = normalizeHeaderFooter(fxp, "footer", ref.variant, ref.rId);
          footers.set(ref.rId, ftr);
        }
      }
    }
  }

  const xmlParts: DocxXmlParts = {
    document: documentFxp,
    headers: headersFxp,
    footers: footersFxp,
  };

  return { body, headers, footers, xmlParts };
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

  for (const child of bodyChildren) {
    const tag = getTagName(child);
    if (tag === "w:p") {
      elements.push(normalizeParagraph(child, `p${paraIdx}`));
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
 * Parse document.xml.rels to build a relationship map (rId → target filename).
 */
function parseRelationships(
  relsXml: string | undefined,
): Map<string, string> {
  const map = new Map<string, string>();
  if (!relsXml) return map;

  const parsed = parseXml(relsXml);

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

interface HeaderFooterRef {
  type: "header" | "footer";
  variant: string; // "default", "first", "even"
  rId: string;
}

/**
 * Find header/footer references in w:sectPr elements.
 */
function findHeaderFooterRefs(documentFxp: FXPNode[]): HeaderFooterRef[] {
  const refs: HeaderFooterRef[] = [];

  const docNode = documentFxp.find((n) => getTagName(n) === "w:document");
  if (!docNode) return refs;

  const bodyNode = findChild(docNode, "w:body");
  if (!bodyNode) return refs;

  // sectPr can be a child of w:body (last section) or w:pPr (section break)
  const sectPrs = findChildren(bodyNode, "w:sectPr");

  // Also check the last paragraph for embedded sectPr
  const bodyChildren = getChildren(bodyNode);
  for (const child of bodyChildren) {
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
