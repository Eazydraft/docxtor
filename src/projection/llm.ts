/**
 * LLM projection: generates simplified document representations
 * for LLM consumption.
 *
 * The projection strips away XML complexity and presents the document
 * as a flat list of paragraphs with IDs and text content. The LLM
 * can then emit patches referencing these IDs.
 */

import type {
  NormalizedDocument,
  ProjectionFormat,
  ProjectionOptions,
} from "../types.js";
import { projectToJson, projectToMarkdown } from "./formats.js";
import type { JsonProjection, JsonProjectionPageBreak, JsonProjectionRunFormatting } from "./formats.js";

/**
 * Project a NormalizedDocument to an LLM-friendly format.
 *
 * @param doc - The normalized document
 * @param format - "json" or "markdown"
 * @param options - Projection options
 * @returns The projected representation (object for JSON, string for markdown)
 */
export function project(
  doc: NormalizedDocument,
  format: ProjectionFormat = "json",
  options: ProjectionOptions = {},
): JsonProjection | string {
  switch (format) {
    case "json":
      return projectToJson(doc, options);
    case "markdown":
      return projectToMarkdown(doc, options);
    default:
      throw new Error(`Unknown projection format: ${format}`);
  }
}

export { projectToJson, projectToMarkdown };
export type { JsonProjection, JsonProjectionPageBreak, JsonProjectionRunFormatting };
