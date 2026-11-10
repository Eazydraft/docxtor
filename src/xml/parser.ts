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
  const nodes = parser.parse(xml) as FXPNode[];
  // The parser drops the whitespace between the XML declaration and the root
  // element. Remember it so serializeXml can write the part back unchanged.
  const m = DECLARATION_RE.exec(xml);
  if (m) declarationWhitespace.set(nodes, m[1]!);
  return nodes;
}

/** The XML declaration at the start of a part, capturing the whitespace after it. */
const DECLARATION_RE = /^\uFEFF?<\?xml[^>]*\?>([ \t\r\n]*)/;

/**
 * Whitespace that followed the XML declaration in each parsed part, keyed by
 * the array parseXml returned. A part docxtor created has no entry.
 */
const declarationWhitespace = new WeakMap<FXPNode[], string>();

/** Whitespace Word writes after the XML declaration. */
export const DEFAULT_DECLARATION_WHITESPACE = "\r\n";

/**
 * The whitespace to write after a part's XML declaration: what the parsed
 * source had, or Word's `\r\n` for a part that was not parsed.
 */
export function getDeclarationWhitespace(nodes: FXPNode[]): string {
  return declarationWhitespace.get(nodes) ?? DEFAULT_DECLARATION_WHITESPACE;
}

/**
 * Deep-clone a parsed part, keeping the declaration whitespace recorded for it.
 */
export function cloneXmlPart(nodes: FXPNode[]): FXPNode[] {
  const copy = JSON.parse(JSON.stringify(nodes)) as FXPNode[];
  const ws = declarationWhitespace.get(nodes);
  if (ws !== undefined) declarationWhitespace.set(copy, ws);
  return copy;
}
