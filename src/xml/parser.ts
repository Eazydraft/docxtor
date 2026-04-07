/**
 * XML parser: wraps fast-xml-parser with the correct options for
 * lossless OOXML roundtripping.
 *
 * Uses `preserveOrder: true` which produces an ordered array structure.
 * This is critical for DOCX because sibling element order matters
 * (e.g., w:pPr must come before w:r inside w:p).
 */

import { XMLParser } from "fast-xml-parser";
import type { FXPNode } from "../types.js";

/** Parser options tuned for lossless OOXML roundtripping. */
const PARSER_OPTIONS = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  // Preserve everything
  commentPropName: "#comment",
  cdataPropName: "#cdata",
  // Don't mangle values
  processEntities: false,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  // Keep declarations
  ignoreDeclaration: false,
  ignorePiTags: false,
};

const parser = new XMLParser(PARSER_OPTIONS);

/**
 * Parse an XML string into a FXP ordered array.
 *
 * The returned structure is an array of FXPNode objects where
 * element order is preserved. This IS the mutable XML representation —
 * modifying these objects and serializing them back produces the output XML.
 */
export function parseXml(xml: string): FXPNode[] {
  return parser.parse(xml) as FXPNode[];
}
