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
  // Word writes empty elements self-closed in word/*.xml and as an open-close
  // pair in docProps/*.xml. The form is a property of the part, so remember
  // which one this part uses.
  if (usesOnlyEmptyPairs(xml)) expandedEmptyElements.add(nodes);
  return nodes;
}

/** Attributes of a start tag. A quoted value may hold `>` or `/>`. */
const ATTRIBUTES = String.raw`(?:\s+[^\s=<>"'\/]+\s*=\s*(?:"[^"]*"|'[^']*'))*\s*`;

/** A self-closing tag: `<w:b/>`, `<w:sz w:val="24"/>`. */
const SELF_CLOSING_RE = new RegExp(String.raw`<[^\s<>\/!?"'=]+${ATTRIBUTES}\/>`);

/** An empty element written as an open-close pair: `<dc:title></dc:title>`. */
const EMPTY_PAIR_RE = new RegExp(String.raw`<([^\s<>\/!?"'=]+)${ATTRIBUTES}><\/\1\s*>`);

/** Comments, CDATA sections and processing instructions: text that can hold tag-like strings. */
const NON_ELEMENT_RE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/g;

/**
 * Whether a part writes its empty elements as open-close pairs: it has at
 * least one such pair and no self-closing tag. Tags are matched as tags, so
 * `/>` inside an attribute value, a comment or a CDATA section does not count.
 */
function usesOnlyEmptyPairs(xml: string): boolean {
  // Only strip when there is something to strip besides the XML declaration
  const markup = xml.includes("<!") || xml.lastIndexOf("<?") > 1 ? xml.replace(NON_ELEMENT_RE, "") : xml;
  // Nearly every word/*.xml part has a self-closing tag near its start
  if (markup.includes("/>") && SELF_CLOSING_RE.test(markup)) return false;
  return EMPTY_PAIR_RE.test(markup);
}

/**
 * Parsed parts that write empty elements as an open-close pair: parts with
 * at least one such pair and no self-closing tag. Every other part, and every
 * part docxtor creates, is written self-closed.
 */
const expandedEmptyElements = new WeakSet<FXPNode[]>();

/**
 * Whether a part writes empty elements as `<a></a>` (true) or `<a/>` (false).
 */
export function usesExpandedEmptyElements(nodes: FXPNode[]): boolean {
  return expandedEmptyElements.has(nodes);
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
 * Deep-clone a parsed part, keeping what was recorded for it at parse time:
 * the declaration whitespace and the form of its empty elements.
 */
export function cloneXmlPart(nodes: FXPNode[]): FXPNode[] {
  const copy = JSON.parse(JSON.stringify(nodes)) as FXPNode[];
  const ws = declarationWhitespace.get(nodes);
  if (ws !== undefined) declarationWhitespace.set(copy, ws);
  if (expandedEmptyElements.has(nodes)) expandedEmptyElements.add(copy);
  return copy;
}
