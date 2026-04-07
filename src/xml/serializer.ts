/**
 * XML serializer: converts a FXP ordered array back to an XML string.
 *
 * Uses the same options as the parser to ensure lossless roundtripping.
 * The FXP object passed here may have been mutated by the patch engine —
 * that's the whole point. We serialize the mutated state.
 */

import { XMLBuilder } from "fast-xml-parser";
import type { FXPNode } from "../types.js";

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
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

const builder = new XMLBuilder(BUILDER_OPTIONS);

/**
 * Serialize a FXP ordered array back to an XML string.
 */
export function serializeXml(nodes: FXPNode[]): string {
  return builder.build(nodes) as string;
}
