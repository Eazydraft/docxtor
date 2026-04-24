/**
 * Flat OPC parser: extracts XML parts from a Flat OPC XML string.
 *
 * Flat OPC (Office Open XML Flat Package) is an alternative packaging format
 * where instead of a ZIP archive, all parts are embedded in a single XML file
 * using the `<pkg:package>` / `<pkg:part>` structure. This is the format
 * returned by Word's JavaScript API `getOoxml()` method.
 *
 * Structure:
 * ```xml
 * <?xml version="1.0" standalone="yes"?>
 * <?mso-application progid="Word.Document"?>
 * <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
 *   <pkg:part pkg:name="/word/document.xml"
 *             pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
 *     <pkg:xmlData>
 *       <w:document xmlns:w="...">...</w:document>
 *     </pkg:xmlData>
 *   </pkg:part>
 *   <pkg:part pkg:name="/word/styles.xml" pkg:contentType="...">
 *     <pkg:xmlData>...</pkg:xmlData>
 *   </pkg:part>
 *   <!-- binary parts use pkg:binaryData with base64 content -->
 * </pkg:package>
 * ```
 */

import type { DocxZipEntries } from "../zip/reader.js";
import type { FlatOpcPart } from "./types.js";

/**
 * Parse a Flat OPC XML string and extract parts into a DocxZipEntries-
 * compatible structure that can be fed directly into normalizeDocument().
 *
 * We use regex-based extraction rather than a full XML parser to avoid
 * parsing the inner XML content (which will be parsed separately by
 * the normalizer using fast-xml-parser with specific options).
 */
export function parseFlatOpc(xml: string): {
  entries: DocxZipEntries;
  parts: FlatOpcPart[];
} {
  const parts = extractParts(xml);

  // Build a DocxZipEntries-compatible structure (without a JSZip instance)
  let documentXml: string | undefined;
  const headers = new Map<string, string>();
  const footers = new Map<string, string>();
  let documentRels: string | undefined;
  let contentTypes: string | undefined;
  let stylesXml: string | undefined;
  let settingsXml: string | undefined;

  for (const part of parts) {
    if (part.binaryData) continue; // Skip binary parts (images, etc.)

    const name = normalizePath(part.name);
    const content = part.xmlContent!;

    if (name === "word/document.xml") {
      documentXml = content;
    } else if (name === "word/_rels/document.xml.rels") {
      documentRels = content;
    } else if (name === "[Content_Types].xml") {
      contentTypes = content;
    } else if (name === "word/styles.xml") {
      stylesXml = content;
    } else if (name === "word/settings.xml") {
      settingsXml = content;
    } else if (/^word\/header\d+\.xml$/i.test(name)) {
      headers.set(name, content);
    } else if (/^word\/footer\d+\.xml$/i.test(name)) {
      footers.set(name, content);
    }
  }

  if (!documentXml) {
    throw new Error(
      "Invalid Flat OPC: missing word/document.xml part",
    );
  }

  const entries: DocxZipEntries = {
    documentXml,
    headers,
    footers,
    documentRels,
    contentTypes,
    stylesXml,
    settingsXml,
    // No JSZip instance for Flat OPC — the writer uses a different path
    zip: null as any,
  };

  return { entries, parts };
}

/**
 * Extract all <pkg:part> elements from a Flat OPC XML string.
 *
 * Uses a streaming approach to avoid parsing inner XML content.
 */
function extractParts(xml: string): FlatOpcPart[] {
  const parts: FlatOpcPart[] = [];

  // Match each <pkg:part ...>...</pkg:part> block
  const partRegex = /<pkg:part\s+([^>]*)>([\s\S]*?)<\/pkg:part>/gi;
  let partMatch: RegExpExecArray | null;

  while ((partMatch = partRegex.exec(xml)) !== null) {
    const attrs = partMatch[1]!;
    const body = partMatch[2]!;

    // Extract pkg:name and pkg:contentType attributes
    const nameMatch = /pkg:name="([^"]*)"/.exec(attrs);
    const contentTypeMatch = /pkg:contentType="([^"]*)"/.exec(attrs);

    if (!nameMatch?.[1]) continue;

    const name = nameMatch[1];
    const contentType = contentTypeMatch?.[1] ?? "";

    // Check for <pkg:xmlData> or <pkg:binaryData>
    const xmlDataMatch = /<pkg:xmlData\b[^>]*>([\s\S]*?)<\/pkg:xmlData>/i.exec(body);
    const binaryDataMatch = /<pkg:binaryData\b[^>]*>([\s\S]*?)<\/pkg:binaryData>/i.exec(body);

    if (xmlDataMatch?.[1]) {
      parts.push({
        name,
        contentType,
        xmlContent: xmlDataMatch[1].trim(),
        binaryData: undefined,
      });
    } else if (binaryDataMatch?.[1]) {
      parts.push({
        name,
        contentType,
        xmlContent: undefined,
        binaryData: binaryDataMatch[1].trim(),
      });
    }
  }

  return parts;
}

/**
 * Normalize a part name by stripping the leading "/" if present.
 * Flat OPC uses "/word/document.xml" but JSZip/DocxZipEntries uses "word/document.xml".
 */
function normalizePath(name: string): string {
  return name.startsWith("/") ? name.slice(1) : name;
}
