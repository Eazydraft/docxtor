/**
 * ZIP reader: extracts XML parts from a DOCX buffer using JSZip.
 *
 * DOCX files are ZIP archives. We extract the XML files we need to normalize
 * and keep the JSZip instance alive so we can write modifications back.
 */

import JSZip from "jszip";

/** All the XML parts we extract from a DOCX file. */
export interface DocxZipEntries {
  /** The raw document.xml content. */
  documentXml: string;
  /** Header XML files keyed by filename (e.g., "word/header1.xml"). */
  headers: Map<string, string>;
  /** Footer XML files keyed by filename (e.g., "word/footer1.xml"). */
  footers: Map<string, string>;
  /** The relationships file for the main document part. */
  documentRels?: string;
  /** The content types file. */
  contentTypes?: string;
  /** The styles.xml file (may not be present). */
  stylesXml?: string;
  /** The JSZip instance (kept for writing back). */
  zip: JSZip;
}

/**
 * Read a DOCX buffer and extract all relevant XML parts.
 */
export async function readDocx(
  input: Buffer | ArrayBuffer | Uint8Array,
): Promise<DocxZipEntries> {
  const zip = await JSZip.loadAsync(input);

  // document.xml is required
  const documentXml = await readEntry(zip, "word/document.xml");
  if (documentXml === undefined) {
    throw new Error(
      "Invalid DOCX: missing word/document.xml",
    );
  }

  // Read document relationships to discover headers/footers
  const documentRels = await readEntry(
    zip,
    "word/_rels/document.xml.rels",
  );
  const contentTypes = await readEntry(zip, "[Content_Types].xml");
  const stylesXml = await readEntry(zip, "word/styles.xml");

  // Extract header and footer files
  const headers = new Map<string, string>();
  const footers = new Map<string, string>();

  // Scan zip for header/footer files
  for (const path of Object.keys(zip.files)) {
    if (/^word\/header\d+\.xml$/i.test(path)) {
      const content = await readEntry(zip, path);
      if (content !== undefined) {
        headers.set(path, content);
      }
    } else if (/^word\/footer\d+\.xml$/i.test(path)) {
      const content = await readEntry(zip, path);
      if (content !== undefined) {
        footers.set(path, content);
      }
    }
  }

  return {
    documentXml,
    headers,
    footers,
    documentRels,
    contentTypes,
    stylesXml,
    zip,
  };
}

/** Read a single zip entry as a string. Returns undefined if not found. */
async function readEntry(
  zip: JSZip,
  path: string,
): Promise<string | undefined> {
  const file = zip.file(path);
  if (!file) return undefined;
  return file.async("string");
}
