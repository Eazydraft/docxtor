/**
 * ZIP writer: takes modified XML strings and writes them back into
 * a DOCX buffer using the original JSZip instance.
 *
 * Unmodified entries (images, styles, fonts, etc.) pass through untouched
 * because JSZip lazily decompresses — only entries we replace are reprocessed.
 */

import type JSZip from "jszip";

export interface WriteDocxOptions {
  /** The original JSZip instance from reading. */
  zip: JSZip;
  /** The modified document.xml content. */
  documentXml: string;
  /** Modified header XML files keyed by zip path. */
  headers?: Map<string, string>;
  /** Modified footer XML files keyed by zip path. */
  footers?: Map<string, string>;
  /** Modified document.xml.rels content. */
  documentRelsXml?: string;
  /** Modified [Content_Types].xml content. */
  contentTypesXml?: string;
}

/**
 * Write modified XML back into a DOCX buffer.
 *
 * Only the entries we explicitly set are recompressed.
 * Everything else (images, styles, fonts, etc.) passes through from the
 * original zip with minimal overhead.
 */
export async function writeDocx(
  options: WriteDocxOptions,
): Promise<Buffer> {
  const { zip, documentXml, headers, footers, documentRelsXml, contentTypesXml } = options;

  // Replace document.xml
  zip.file("word/document.xml", documentXml);

  // Replace modified headers
  if (headers) {
    for (const [path, content] of headers) {
      zip.file(path, content);
    }
  }

  // Replace modified footers
  if (footers) {
    for (const [path, content] of footers) {
      zip.file(path, content);
    }
  }

  // Replace document.xml.rels if modified
  if (documentRelsXml) {
    zip.file("word/_rels/document.xml.rels", documentRelsXml);
  }

  // Replace [Content_Types].xml if modified
  if (contentTypesXml) {
    zip.file("[Content_Types].xml", contentTypesXml);
  }

  // Generate the output buffer
  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return buffer;
}
