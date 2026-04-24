/**
 * Types for Flat OPC (Office Open XML Flat Package) support.
 */

/**
 * A single part extracted from a Flat OPC XML envelope.
 */
export interface FlatOpcPart {
  /** Part name/path, e.g., "/word/document.xml" */
  name: string;
  /** Content type, e.g., "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" */
  contentType: string;
  /** Raw XML string content (for XML parts). */
  xmlContent: string | undefined;
  /** Base64-encoded binary content (for binary parts like images). */
  binaryData: string | undefined;
}
