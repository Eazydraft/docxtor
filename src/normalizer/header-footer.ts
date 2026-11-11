/**
 * Header/footer normalizer: reads header/footer XML parts and
 * normalizes their paragraphs.
 *
 * Headers and footers are separate XML files (word/header1.xml, etc.)
 * with w:hdr or w:ftr as the root element. Inside, the structure is
 * the same as document.xml: paragraphs and tables.
 */

import type {
  FXPNode,
  HeaderFooter,
  NormalizedParagraph,
  NormalizedTable,
} from "../types.js";
import { contentChildren, getTagName } from "../xml/helpers.js";
import { ControlResolver } from "./control.js";
import { normalizeParagraph } from "./paragraph.js";
import { normalizeTable } from "./table.js";

/**
 * Normalize a header or footer FXP structure.
 *
 * @param fxpNodes - The parsed FXP nodes from the header/footer XML
 * @param type - "header" or "footer"
 * @param variant - "default", "first", or "even"
 * @param rId - The relationship ID
 * @param sectionIndex - 1-based index of the section that references the part
 * @param withSection - Put the section into the ids (`hdr.s2.first.p0`); set
 *   when more than one section of the document has headers or footers
 */
export function normalizeHeaderFooter(
  fxpNodes: FXPNode[],
  type: "header" | "footer",
  variant: string,
  rId: string,
  sectionIndex: number = 1,
  withSection: boolean = false,
): HeaderFooter {
  const idPrefix = headerFooterIdPrefix(type, sectionIndex, withSection);
  const prefix = idPrefix;

  // Find the root element (w:hdr or w:ftr)
  const rootTag = type === "header" ? "w:hdr" : "w:ftr";
  const rootNode = fxpNodes.find((node) => getTagName(node) === rootTag);

  const paragraphs: NormalizedParagraph[] = [];
  const tables: NormalizedTable[] = [];

  if (rootNode) {
    const resolver = new ControlResolver();
    let paraIdx = 0;
    let tableIdx = 0;

    // Read through block-level content controls (a page-number gallery, say)
    for (const { node: child, sdts } of contentChildren(rootNode)) {
      const tag = getTagName(child);
      const control = resolver.resolve(sdts);
      if (tag === "w:p") {
        paragraphs.push(
          normalizeParagraph(
            child,
            `${prefix}.${variant}.p${paraIdx}`,
            control,
          ),
        );
        paraIdx++;
      } else if (tag === "w:tbl") {
        tables.push(
          normalizeTable(
            child,
            `${prefix}.${variant}.tbl${tableIdx}`,
            control,
            resolver,
          ),
        );
        tableIdx++;
      }
    }
  }

  return {
    type,
    variant,
    rId,
    sectionIndex,
    references: [{ sectionIndex, variant, rId }],
    idPrefix,
    paragraphs,
    tables,
    xmlRef: fxpNodes,
  };
}

/**
 * The start of every id inside a header or footer: `hdr` / `ftr`, or
 * `hdr.s2` / `ftr.s2` when the ids carry the section.
 */
export function headerFooterIdPrefix(
  type: "header" | "footer",
  sectionIndex: number,
  withSection: boolean,
): string {
  const base = type === "header" ? "hdr" : "ftr";
  return withSection ? `${base}.s${sectionIndex}` : base;
}

/** Whether a header or footer's ids carry its section (`hdr.s2.first.p0`). */
export function idsCarrySection(hf: HeaderFooter): boolean {
  return hf.idPrefix !== (hf.type === "header" ? "hdr" : "ftr");
}
