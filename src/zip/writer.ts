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
  /** Modified word/settings.xml content. */
  settingsXml?: string;
  /** Modified word/numbering.xml content. */
  numberingXml?: string;
  /** Modified word/footnotes.xml content. */
  footnotesXml?: string;
  /** Modified word/endnotes.xml content. */
  endnotesXml?: string;
  /** Modified word/comments.xml content. */
  commentsXml?: string;
  /** Modified docProps/core.xml content. */
  coreXml?: string;
  /** Modified docProps/app.xml content. */
  appXml?: string;
  /** Modified _rels/.rels content. */
  rootRelsXml?: string;
  /** Zip paths to remove from the archive (e.g., deleted header/footer parts). */
  removedParts?: string[];
}

/** Fixed zip path for each optional XML part. */
const PART_PATHS: Array<[keyof WriteDocxOptions, string]> = [
  ["documentRelsXml", "word/_rels/document.xml.rels"],
  ["contentTypesXml", "[Content_Types].xml"],
  ["settingsXml", "word/settings.xml"],
  ["numberingXml", "word/numbering.xml"],
  ["footnotesXml", "word/footnotes.xml"],
  ["endnotesXml", "word/endnotes.xml"],
  ["commentsXml", "word/comments.xml"],
  ["coreXml", "docProps/core.xml"],
  ["appXml", "docProps/app.xml"],
  ["rootRelsXml", "_rels/.rels"],
];

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
  const { zip, documentXml, headers, footers, removedParts } = options;

  // Remove deleted parts (e.g., header/footer XML files after remove_header/remove_footer)
  if (removedParts) {
    for (const path of removedParts) {
      zip.remove(path);
    }
  }

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

  // Replace every other part that was supplied
  for (const [key, path] of PART_PATHS) {
    const content = options[key];
    if (typeof content === "string") {
      zip.file(path, content);
    }
  }

  // Generate the output buffer
  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return buffer;
}
