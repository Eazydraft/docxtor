/**
 * LLM projection formats: convert a NormalizedDocument into
 * simplified representations that are easy for LLMs to understand
 * and generate patches against.
 */

import type {
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  ProjectionOptions,
} from "../types.js";

// =============================================================================
// JSON Projection
// =============================================================================

interface JsonProjectionParagraph {
  id: string;
  text: string;
  style?: string;
  formatting?: { bold?: boolean; italic?: boolean }[];
}

interface JsonProjectionTableCell {
  id: string;
  text: string;
}

interface JsonProjectionTable {
  id: string;
  style?: string;
  rows: JsonProjectionTableCell[][];
}

interface JsonProjectionHeaderFooter {
  type: "header" | "footer";
  variant: string;
  paragraphs: JsonProjectionParagraph[];
}

export interface JsonProjection {
  body: (JsonProjectionParagraph | JsonProjectionTable)[];
  headers?: JsonProjectionHeaderFooter[];
  footers?: JsonProjectionHeaderFooter[];
}

/**
 * Project a NormalizedDocument to a JSON structure suitable for LLMs.
 */
export function projectToJson(
  doc: NormalizedDocument,
  options: ProjectionOptions = {},
): JsonProjection {
  const includeFormatting = options.includeFormatting ?? false;
  const includeHeadersFooters = options.includeHeadersFooters ?? true;
  const includeTables = options.includeTables ?? true;

  const body: (JsonProjectionParagraph | JsonProjectionTable)[] = [];

  for (const element of doc.body) {
    if (element.kind === "paragraph") {
      body.push(projectParagraphToJson(element, includeFormatting));
    } else if (element.kind === "table" && includeTables) {
      body.push(projectTableToJson(element));
    }
  }

  const result: JsonProjection = { body };

  if (includeHeadersFooters) {
    if (doc.headers.size > 0) {
      result.headers = [];
      for (const header of doc.headers.values()) {
        result.headers.push({
          type: "header",
          variant: header.variant,
          paragraphs: header.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting),
          ),
        });
      }
    }
    if (doc.footers.size > 0) {
      result.footers = [];
      for (const footer of doc.footers.values()) {
        result.footers.push({
          type: "footer",
          variant: footer.variant,
          paragraphs: footer.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting),
          ),
        });
      }
    }
  }

  return result;
}

function projectParagraphToJson(
  para: NormalizedParagraph,
  includeFormatting: boolean,
): JsonProjectionParagraph {
  const result: JsonProjectionParagraph = {
    id: para.id,
    text: para.text,
  };

  if (para.style) {
    result.style = para.style;
  }

  if (includeFormatting && para.runs.length > 0) {
    result.formatting = para.runs.map((r) => ({
      ...(r.formatting.bold ? { bold: true } : {}),
      ...(r.formatting.italic ? { italic: true } : {}),
    }));
  }

  return result;
}

function projectTableToJson(
  table: NormalizedTable,
): JsonProjectionTable {
  return {
    id: table.id,
    style: table.style,
    rows: table.rows.map((row) =>
      row.cells.map((cell) => ({
        id: cell.paragraphs[0]?.id ?? `${table.id}.r${row.index}.c${cell.colIndex}.p0`,
        text: cell.paragraphs.map((p) => p.text).join("\n"),
      })),
    ),
  };
}

// =============================================================================
// Markdown Projection
// =============================================================================

/**
 * Project a NormalizedDocument to a markdown-like string for LLMs.
 *
 * Format:
 * ```
 * [p0] This is a paragraph
 * [p1] (Heading1) This is a heading
 *
 * [tbl0] Table
 * | [tbl0.r0.c0.p0] Header A | [tbl0.r0.c1.p0] Header B |
 * | [tbl0.r1.c0.p0] Cell 1   | [tbl0.r1.c1.p0] Cell 2   |
 *
 * --- Header (default) ---
 * [hdr.default.p0] Header text
 *
 * --- Footer (default) ---
 * [ftr.default.p0] Footer text
 * ```
 */
export function projectToMarkdown(
  doc: NormalizedDocument,
  options: ProjectionOptions = {},
): string {
  const includeHeadersFooters = options.includeHeadersFooters ?? true;
  const includeTables = options.includeTables ?? true;
  const lines: string[] = [];

  for (const element of doc.body) {
    if (element.kind === "paragraph") {
      lines.push(formatParagraphMarkdown(element));
    } else if (element.kind === "table" && includeTables) {
      lines.push(""); // blank line before table
      lines.push(...formatTableMarkdown(element));
      lines.push(""); // blank line after table
    }
  }

  if (includeHeadersFooters) {
    for (const header of doc.headers.values()) {
      lines.push("");
      lines.push(`--- Header (${header.variant}) ---`);
      for (const para of header.paragraphs) {
        lines.push(formatParagraphMarkdown(para));
      }
    }

    for (const footer of doc.footers.values()) {
      lines.push("");
      lines.push(`--- Footer (${footer.variant}) ---`);
      for (const para of footer.paragraphs) {
        lines.push(formatParagraphMarkdown(para));
      }
    }
  }

  return lines.join("\n").trim();
}

function formatParagraphMarkdown(para: NormalizedParagraph): string {
  const styleHint = para.style ? ` (${para.style})` : "";
  return `[${para.id}]${styleHint} ${para.text}`;
}

function formatTableMarkdown(table: NormalizedTable): string[] {
  const lines: string[] = [];

  for (const row of table.rows) {
    const cells = row.cells.map((cell) => {
      const text = cell.paragraphs.map((p) => p.text).join(" ");
      const id = cell.paragraphs[0]?.id ?? `${table.id}.r${row.index}.c${cell.colIndex}.p0`;
      return `[${id}] ${text}`;
    });
    lines.push(`| ${cells.join(" | ")} |`);
  }

  return lines;
}
