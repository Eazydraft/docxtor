/**
 * XML serializer: converts a FXP ordered array back to an XML string.
 *
 * Uses the same options as the parser to ensure lossless roundtripping.
 * The FXP object passed here may have been mutated by the patch engine —
 * that's the whole point. We serialize the mutated state.
 */

import { XMLBuilder } from "fast-xml-parser";
import type { FXPNode } from "../types.js";
import { getDeclarationWhitespace, usesExpandedEmptyElements } from "./parser.js";

/** Builder options must match parser options for correct roundtripping. */
const BUILDER_OPTIONS = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  commentPropName: "#comment",
  cdataPropName: "#cdata",
  processEntities: false,
  // Produce well-formed XML
  format: false,
  suppressBooleanAttributes: false,
  // The builder escapes ' to &apos; in every attribute value, whatever
  // processEntities says, while the parser keeps a literal ' as it is. This
  // runs before that escape, so hide the apostrophe from it; serializeXml
  // puts it back. An &apos; in the source was never decoded, so it stays.
  attributeValueProcessor: (_name: string, value: unknown): string =>
    String(value).replaceAll("'", APOSTROPHE_SENTINEL),
};

/**
 * Stands in for ' inside attribute values while the builder runs. U+FFFF is
 * not a legal XML character, so no well-formed part can already contain it.
 */
const APOSTROPHE_SENTINEL = "\uFFFF";

/** Writes empty elements self-closed (<w:b/>), as Word does in word/*.xml. */
const builder = new XMLBuilder({ ...BUILDER_OPTIONS, suppressEmptyNode: true });

/** Writes empty elements as a pair (<dc:title></dc:title>), as Word does in docProps/*.xml. */
const expandedBuilder = new XMLBuilder({ ...BUILDER_OPTIONS, suppressEmptyNode: false });

/**
 * Serialize a FXP ordered array back to an XML string.
 */
export function serializeXml(nodes: FXPNode[]): string {
  const built = (usesExpandedEmptyElements(nodes) ? expandedBuilder : builder).build(nodes) as string;
  const xml = built.includes(APOSTROPHE_SENTINEL) ? built.replaceAll(APOSTROPHE_SENTINEL, "'") : built;
  if (!xml.startsWith("<?xml")) return xml;
  // The builder writes the root directly after the declaration. Put back the
  // whitespace the source had there (or Word's \r\n for a new part).
  const end = xml.indexOf("?>") + 2;
  return xml.slice(0, end) + getDeclarationWhitespace(nodes) + xml.slice(end);
}
