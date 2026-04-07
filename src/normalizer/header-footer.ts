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
} from "../types.js";
import { findChildren, getTagName } from "../xml/helpers.js";
import { normalizeParagraph } from "./paragraph.js";

/**
 * Normalize a header or footer FXP structure.
 *
 * @param fxpNodes - The parsed FXP nodes from the header/footer XML
 * @param type - "header" or "footer"
 * @param variant - "default", "first", or "even"
 * @param rId - The relationship ID
 */
export function normalizeHeaderFooter(
  fxpNodes: FXPNode[],
  type: "header" | "footer",
  variant: string,
  rId: string,
): HeaderFooter {
  const prefix = type === "header" ? "hdr" : "ftr";

  // Find the root element (w:hdr or w:ftr)
  const rootTag = type === "header" ? "w:hdr" : "w:ftr";
  const rootNode = fxpNodes.find((node) => getTagName(node) === rootTag);

  const paragraphs: NormalizedParagraph[] = [];

  if (rootNode) {
    const pNodes = findChildren(rootNode, "w:p");
    for (let i = 0; i < pNodes.length; i++) {
      paragraphs.push(
        normalizeParagraph(
          pNodes[i]!,
          `${prefix}.${variant}.p${i}`,
        ),
      );
    }
  }

  return {
    type,
    variant,
    rId,
    paragraphs,
    xmlRef: fxpNodes,
  };
}
