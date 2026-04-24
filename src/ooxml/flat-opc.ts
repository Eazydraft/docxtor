/**
 * Flat OPC (Open XML Package Convention) utilities.
 *
 * Flat OPC is the single-XML-file format that Office.js exposes via
 * `body.getOoxml()` and accepts via `body.insertOoxml()`.  It wraps
 * every ZIP entry from a DOCX as a `<pkg:part>` element in one document.
 *
 * Round-trip:
 *   Office.js body.getOoxml()  →  parseFlatOpc  →  DocxZipEntries
 *   Modified DocxZipEntries    →  buildFlatOpc  →  body.insertOoxml()
 */

import { XMLParser } from "fast-xml-parser";
import type { FXPNode } from "../types.js";

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** One part inside a Flat OPC package. */
export interface FlatOpcPart {
  /** Package path, e.g. "/word/document.xml". Always starts with "/". */
  name: string;
  /** MIME content type. */
  contentType: string;
  /**
   * Raw content string.
   * - For XML parts (`isBinary: false`): the inner XML of `<pkg:xmlData>`.
   * - For binary parts (`isBinary: true`): the base64 string inside `<pkg:binaryData>`.
   */
  content: string;
  /** True when the part uses `<pkg:binaryData>` (images, fonts, etc.). */
  isBinary: boolean;
  /** Optional `pkg:padding` attribute (preserved verbatim). */
  padding?: string;
}

// ---------------------------------------------------------------------------
// Parser (Office.js → Docxtor)
// ---------------------------------------------------------------------------

// Capture both <pkg:xmlData> and <pkg:binaryData> as raw strings so that
// inner XML is not re-parsed and base64 blobs are preserved intact.
const OPC_PARSER = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: false,
  stopNodes: ["*.pkg:xmlData", "*.pkg:binaryData"],
  // pkg:part elements must always be an array (handles single-part packages)
  isArray: (tagName) => tagName === "pkg:part",
});

/**
 * Detect whether a string is in Flat OPC format.
 * Checks for the standard Microsoft Office XML-package namespace.
 */
export function isFlatOpc(xml: string): boolean {
  // The package root or its namespace declaration appears within the first
  // 2 KB — no need to scan the whole string.
  const head = xml.slice(0, 2048);
  return (
    head.includes("pkg:package") ||
    head.includes("xmlPackage")
  );
}

/**
 * Parse a Flat OPC XML string into its constituent parts.
 *
 * Handles both XML parts (`<pkg:xmlData>`) and binary parts
 * (`<pkg:binaryData>`).  Unknown / malformed parts are silently skipped.
 */
export function parseFlatOpc(xml: string): FlatOpcPart[] {
  const parsed = OPC_PARSER.parse(xml) as Record<string, unknown>;

  // The root element may appear as "pkg:package" or just "package"
  const pkg =
    (parsed["pkg:package"] ?? parsed["package"]) as
      | Record<string, unknown>
      | undefined;

  if (!pkg) return [];

  const rawParts =
    (pkg["pkg:part"] ?? pkg["part"]) as unknown[] | undefined;

  if (!Array.isArray(rawParts)) return [];

  const result: FlatOpcPart[] = [];

  for (const raw of rawParts) {
    if (!raw || typeof raw !== "object") continue;
    const p = raw as Record<string, unknown>;

    const name = (
      p["@_pkg:name"] ?? p["@_name"]
    ) as string | undefined;

    const contentType = (
      p["@_pkg:contentType"] ?? p["@_contentType"]
    ) as string | undefined;

    const padding = (
      p["@_pkg:padding"] ?? p["@_padding"]
    ) as string | undefined;

    // One of xmlData or binaryData must be present
    const xmlData = p["pkg:xmlData"] ?? p["xmlData"];
    const binaryData = p["pkg:binaryData"] ?? p["binaryData"];

    if (!name || !contentType) continue;

    if (xmlData !== undefined) {
      result.push({
        name,
        contentType,
        content: String(xmlData).trim(),
        isBinary: false,
        padding,
      });
    } else if (binaryData !== undefined) {
      result.push({
        name,
        contentType,
        content: String(binaryData).trim(),
        isBinary: true,
        padding,
      });
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// Builder (Docxtor → Office.js)
// ---------------------------------------------------------------------------

const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

/**
 * Build a Flat OPC XML string from a list of parts.
 *
 * Pass `overrides` to replace specific parts (e.g., the modified
 * `word/document.xml`) while keeping all others from `originalParts`.
 */
export function buildFlatOpc(
  parts: FlatOpcPart[],
  overrides?: Map<string, string>,
): string {
  const resolved = overrides
    ? parts.map((p) => {
        const ov = overrides.get(p.name);
        return ov !== undefined ? { ...p, content: ov } : p;
      })
    : parts;

  const partsXml = resolved.map(serializePart).join("");

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<?mso-application progid="Word.Document"?>` +
    `<pkg:package xmlns:pkg="${PKG_NS}">${partsXml}</pkg:package>`
  );
}

function serializePart(part: FlatOpcPart): string {
  const paddingAttr = part.padding
    ? ` pkg:padding="${escapeAttr(part.padding)}"`
    : "";

  const dataTag = part.isBinary ? "pkg:binaryData" : "pkg:xmlData";

  return (
    `<pkg:part` +
    ` pkg:name="${escapeAttr(part.name)}"` +
    ` pkg:contentType="${escapeAttr(part.contentType)}"` +
    `${paddingAttr}>` +
    `<${dataTag}>${part.content}</${dataTag}>` +
    `</pkg:part>`
  );
}

// ---------------------------------------------------------------------------
// Fragment builder (for patchResultsToOoxml)
// ---------------------------------------------------------------------------

/**
 * Known content types for Flat OPC paths.
 * Returned by `opcContentType()` when building minimal packages.
 */
const KNOWN_CONTENT_TYPES: Record<string, string> = {
  "/word/document.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  "/word/_rels/document.xml.rels":
    "application/vnd.openxmlformats-package.relationships+xml",
  "/_rels/.rels":
    "application/vnd.openxmlformats-package.relationships+xml",
  "/[Content_Types].xml":
    "application/vnd.openxmlformats-package.content-types+xml",
  "/word/styles.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  "/word/settings.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  "/word/fontTable.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
  "/word/webSettings.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
  "/word/numbering.xml":
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
};

export function opcContentType(path: string): string {
  if (KNOWN_CONTENT_TYPES[path]) return KNOWN_CONTENT_TYPES[path];
  if (/\/word\/header\d+\.xml$/i.test(path))
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
  if (/\/word\/footer\d+\.xml$/i.test(path))
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
  return "application/xml";
}

/**
 * Build a minimal `<w:document>` XML string that wraps the provided
 * FXP body nodes and preserves all namespace declarations from the
 * original document root.
 *
 * This is used by `patchResultsToOoxml` to produce a well-formed
 * OOXML fragment that Office.js `insertOoxml()` can accept.
 */
export function buildDocumentFragment(
  nodes: FXPNode[],
  documentFxp: FXPNode[],
  serializeFxp: (nodes: FXPNode[]) => string,
): string {
  // Extract namespace declarations from the w:document root node
  const docRootNode = documentFxp.find((n) => "w:document" in n) as
    | Record<string, unknown>
    | undefined;

  const rawAttrs = (
    docRootNode?.[":@"] as Record<string, string> | undefined
  ) ?? {};

  const nsAttrStr = Object.entries(rawAttrs)
    .filter(([k]) => k.startsWith("@_xmlns") || k === "@_mc:Ignorable")
    .map(([k, v]) => `${k.slice(2)}="${escapeAttr(v)}"`)
    .join(" ");

  const ns = nsAttrStr
    ? ` ${nsAttrStr}`
    : ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;

  const bodyContent = nodes.map((n) => serializeFxp([n])).join("");

  return `<w:document${ns}><w:body>${bodyContent}</w:body></w:document>`;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Convert a Flat OPC path ("/word/foo.xml") to a ZIP path ("word/foo.xml"). */
export function opcPathToZipPath(path: string): string {
  return path.startsWith("/") ? path.slice(1) : path;
}

/** Convert a ZIP path ("word/foo.xml") to a Flat OPC path ("/word/foo.xml"). */
export function zipPathToOpcPath(path: string): string {
  return path.startsWith("/") ? path : `/${path}`;
}

function escapeAttr(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
