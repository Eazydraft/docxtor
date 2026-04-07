/**
 * Paragraph normalizer: walks a w:p FXP node and produces a
 * NormalizedParagraph with text, run mappings, and xmlRefs.
 *
 * This is the core of the mapping layer. Each RunMapping holds:
 * - The run's text content
 * - Character offset range in the concatenated paragraph text
 * - Direct reference to the w:r and w:t FXP nodes
 * - Captured formatting
 *
 * Patches use these mappings to find the right XML nodes to mutate.
 */

import type {
  FXPNode,
  NormalizedParagraph,
  RunMapping,
  RunFormatting,
  ListInfo,
} from "../types.js";
import {
  findChild,
  findChildren,
  getAttr,
  getChildren,
  getRunText,
  findTextNode,
  isTag,
} from "../xml/helpers.js";

/**
 * Normalize a single w:p FXP node into a NormalizedParagraph.
 *
 * @param paragraphNode - The w:p FXP node (must have "w:p" key)
 * @param id - The stable ID to assign (e.g., "p0", "tbl0.r1.c0.p0")
 */
export function normalizeParagraph(
  paragraphNode: FXPNode,
  id: string,
): NormalizedParagraph {
  const runs = findChildren(paragraphNode, "w:r");
  const runMappings: RunMapping[] = [];
  let offset = 0;

  for (const runNode of runs) {
    const text = getRunText(runNode);
    if (text.length === 0) continue; // Skip runs with no text (e.g., break-only runs)

    const textNode = findTextNode(runNode);
    if (!textNode) continue;

    const formatting = extractFormatting(runNode);

    runMappings.push({
      text,
      startOffset: offset,
      endOffset: offset + text.length,
      xmlRef: runNode,
      textXmlRef: textNode,
      formatting,
    });

    offset += text.length;
  }

  const fullText = runMappings.map((r) => r.text).join("");

  return {
    kind: "paragraph",
    id,
    text: fullText,
    runs: runMappings,
    xmlRef: paragraphNode,
    style: extractStyle(paragraphNode),
    listInfo: extractListInfo(paragraphNode),
  };
}

/**
 * Extract formatting properties from a w:r run node's w:rPr element.
 */
function extractFormatting(runNode: FXPNode): RunFormatting {
  const rPr = findChild(runNode, "w:rPr");
  if (!rPr) return {};

  const formatting: RunFormatting = {};
  const children = getChildren(rPr);

  for (const child of children) {
    if (isTag(child, "w:b")) {
      // <w:b/> means bold on; <w:b w:val="0"/> means bold off
      const val = getAttr(child, "w:val");
      formatting.bold = val !== "0" && val !== "false";
    } else if (isTag(child, "w:i")) {
      const val = getAttr(child, "w:val");
      formatting.italic = val !== "0" && val !== "false";
    } else if (isTag(child, "w:u")) {
      formatting.underline = getAttr(child, "w:val") ?? "single";
    } else if (isTag(child, "w:strike")) {
      const val = getAttr(child, "w:val");
      formatting.strikethrough = val !== "0" && val !== "false";
    } else if (isTag(child, "w:sz")) {
      const val = getAttr(child, "w:val");
      if (val) formatting.fontSize = parseInt(val, 10);
    } else if (isTag(child, "w:rFonts")) {
      formatting.fontFamily = getAttr(child, "w:ascii") ?? getAttr(child, "w:hAnsi");
    } else if (isTag(child, "w:color")) {
      formatting.color = getAttr(child, "w:val");
    } else if (isTag(child, "w:highlight")) {
      formatting.highlight = getAttr(child, "w:val");
    }
  }

  return formatting;
}

/**
 * Extract the paragraph style name from w:pPr/w:pStyle.
 */
function extractStyle(paragraphNode: FXPNode): string | undefined {
  const pPr = findChild(paragraphNode, "w:pPr");
  if (!pPr) return undefined;

  const pStyle = findChild(pPr, "w:pStyle");
  if (!pStyle) return undefined;

  return getAttr(pStyle, "w:val");
}

/**
 * Extract list/numbering info from w:pPr/w:numPr.
 */
function extractListInfo(paragraphNode: FXPNode): ListInfo | undefined {
  const pPr = findChild(paragraphNode, "w:pPr");
  if (!pPr) return undefined;

  const numPr = findChild(pPr, "w:numPr");
  if (!numPr) return undefined;

  const numIdNode = findChild(numPr, "w:numId");
  const ilvlNode = findChild(numPr, "w:ilvl");

  const numId = numIdNode ? getAttr(numIdNode, "w:val") : undefined;
  const level = ilvlNode ? getAttr(ilvlNode, "w:val") : undefined;

  if (!numId) return undefined;

  return {
    numId,
    level: level ? parseInt(level, 10) : 0,
  };
}
