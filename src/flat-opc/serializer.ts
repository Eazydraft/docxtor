/**
 * Flat OPC serializer: converts document XML parts back into a Flat OPC
 * XML string that can be consumed by Word's insertOoxml() API.
 *
 * This is the inverse of the parser: takes the modified XML parts and
 * wraps them in a <pkg:package> envelope with <pkg:part> entries.
 */

import type { FlatOpcPart } from "./types.js";

const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

/**
 * Serialize an array of FlatOpcPart back into a Flat OPC XML string.
 *
 * @param parts - The original parts array (from parseFlatOpc), potentially
 *   with updated xmlContent for parts that were modified by patches.
 * @returns A complete Flat OPC XML string.
 */
export function serializeFlatOpc(parts: FlatOpcPart[]): string {
  const lines: string[] = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<?mso-application progid="Word.Document"?>`,
    `<pkg:package xmlns:pkg="${PKG_NS}">`,
  ];

  for (const part of parts) {
    const attrs = [
      `pkg:name="${escapeXmlAttr(part.name)}"`,
      `pkg:contentType="${escapeXmlAttr(part.contentType)}"`,
    ];

    // Add padding attribute for binary parts (Word expects it)
    if (part.binaryData) {
      attrs.push(`pkg:padding="512"`);
    }

    lines.push(`  <pkg:part ${attrs.join(" ")}>`);

    if (part.xmlContent !== undefined) {
      lines.push(`    <pkg:xmlData>`);
      lines.push(`      ${part.xmlContent}`);
      lines.push(`    </pkg:xmlData>`);
    } else if (part.binaryData !== undefined) {
      lines.push(`    <pkg:binaryData>${part.binaryData}</pkg:binaryData>`);
    }

    lines.push(`  </pkg:part>`);
  }

  lines.push(`</pkg:package>`);
  return lines.join("\n");
}

/** Escape special characters in XML attribute values. */
function escapeXmlAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
