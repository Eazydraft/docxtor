/**
 * LLM projection formats: convert a NormalizedDocument into
 * simplified representations that are easy for LLMs to understand
 * and generate patches against.
 */

import type {
  NormalizedDocument,
  NormalizedParagraph,
  NormalizedTable,
  RevisionMark,
  RunFormatting,
  StyleInfo,
  ProjectionOptions,
} from "../types.js";

// =============================================================================
// JSON Projection
// =============================================================================

interface JsonProjectionParagraph {
  id: string;
  text: string;
  style?: string;
  listLevel?: number;
  formatting?: RunFormatting[];
  revisions?: RevisionMark[];
  hasPageBreak?: boolean;
  pageBreakOffset?: number;
}

interface JsonProjectionTableCell {
  id: string;
  text: string;
  paragraphs?: JsonProjectionParagraph[];
}

interface JsonProjectionTable {
  id: string;
  style?: string;
  rows: JsonProjectionTableCell[][];
  nestedTables?: JsonProjectionTable[];
}

interface JsonProjectionHeaderFooter {
  type: "header" | "footer";
  variant: string;
  paragraphs: JsonProjectionParagraph[];
}

export interface JsonProjectionPageBreak {
  id: string;
  type: "page_break";
}

export interface JsonProjection {
  body: (JsonProjectionParagraph | JsonProjectionTable | JsonProjectionPageBreak)[];
  headers?: JsonProjectionHeaderFooter[];
  footers?: JsonProjectionHeaderFooter[];
  styles?: StyleInfo[];
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
  const includeRevisions = options.includeRevisions ?? false;
  const includePageBreaks = options.includePageBreaks ?? false;

  const body: (JsonProjectionParagraph | JsonProjectionTable | JsonProjectionPageBreak)[] = [];

  for (const element of doc.body) {
    if (element.kind === "paragraph") {
      body.push(projectParagraphToJson(element, includeFormatting, includeRevisions, includePageBreaks));
    } else if (element.kind === "table" && includeTables) {
      body.push(projectTableToJson(element, includeFormatting, includeRevisions));
    } else if (element.kind === "page_break" && includePageBreaks) {
      body.push({ id: element.id, type: "page_break" });
    }
  }

  const result: JsonProjection = { body };

  // Include available styles if present
  if (doc.styles.length > 0) {
    result.styles = doc.styles;
  }

  if (includeHeadersFooters) {
    if (doc.headers.size > 0) {
      result.headers = [];
      for (const header of doc.headers.values()) {
        result.headers.push({
          type: "header",
          variant: header.variant,
          paragraphs: header.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting, includeRevisions),
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
            projectParagraphToJson(p, includeFormatting, includeRevisions),
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
  includeRevisions: boolean = false,
  includePageBreaks: boolean = false,
): JsonProjectionParagraph {
  const result: JsonProjectionParagraph = {
    id: para.id,
    text: para.text,
  };

  if (para.style) {
    result.style = para.style;
  }

  if (para.listInfo) {
    result.listLevel = para.listInfo.level;
  }

  if (includeFormatting && para.runs.length > 0) {
    result.formatting = para.runs.map((r) => {
      const fmt: RunFormatting = {};
      if (r.formatting.bold) fmt.bold = true;
      if (r.formatting.italic) fmt.italic = true;
      if (r.formatting.underline) fmt.underline = r.formatting.underline;
      if (r.formatting.strikethrough) fmt.strikethrough = true;
      if (r.formatting.fontSize) fmt.fontSize = r.formatting.fontSize;
      if (r.formatting.fontFamily) fmt.fontFamily = r.formatting.fontFamily;
      if (r.formatting.color) fmt.color = r.formatting.color;
      if (r.formatting.highlight) fmt.highlight = r.formatting.highlight;
      return fmt;
    });
  }

  if (includeRevisions && para.revisions && para.revisions.length > 0) {
    result.revisions = para.revisions;
  }

  if (includePageBreaks && para.hasPageBreak) {
    result.hasPageBreak = true;
    if (para.pageBreakOffset !== undefined) {
      result.pageBreakOffset = para.pageBreakOffset;
    }
  }

  return result;
}

function projectTableToJson(
  table: NormalizedTable,
  includeFormatting: boolean = false,
  includeRevisions: boolean = false,
): JsonProjectionTable {
  const result: JsonProjectionTable = {
    id: table.id,
    style: table.style,
    rows: table.rows.map((row) =>
      row.cells.map((cell) => {
        const cellResult: JsonProjectionTableCell = {
          id: cell.paragraphs[0]?.id ?? `${table.id}.r${row.index}.c${cell.colIndex}.p0`,
          text: cell.paragraphs.map((p) => p.text).join("\n"),
        };
        // Include individual paragraph details for multi-paragraph cells
        if (cell.paragraphs.length > 1) {
          cellResult.paragraphs = cell.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting, includeRevisions),
          );
        }
        return cellResult;
      }),
    ),
  };

  // Include nested tables if any exist
  const allNested: JsonProjectionTable[] = [];
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const nested of cell.nestedTables) {
        allNested.push(projectTableToJson(nested, includeFormatting, includeRevisions));
      }
    }
  }
  if (allNested.length > 0) {
    result.nestedTables = allNested;
  }

  return result;
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
  const includeRevisions = options.includeRevisions ?? false;
  const includePageBreaks = options.includePageBreaks ?? false;
  const lines: string[] = [];

  for (const element of doc.body) {
    if (element.kind === "paragraph") {
      lines.push(formatParagraphMarkdown(element, includeRevisions, includePageBreaks));
    } else if (element.kind === "table" && includeTables) {
      lines.push(""); // blank line before table
      lines.push(...formatTableMarkdown(element, includeRevisions));
      lines.push(""); // blank line after table
    } else if (element.kind === "page_break" && includePageBreaks) {
      lines.push("");
      lines.push(`[${element.id}] --- PAGE BREAK ---`);
      lines.push("");
    }
  }

  if (includeHeadersFooters) {
    for (const header of doc.headers.values()) {
      lines.push("");
      lines.push(`--- Header (${header.variant}) ---`);
      for (const para of header.paragraphs) {
        lines.push(formatParagraphMarkdown(para, includeRevisions));
      }
    }

    for (const footer of doc.footers.values()) {
      lines.push("");
      lines.push(`--- Footer (${footer.variant}) ---`);
      for (const para of footer.paragraphs) {
        lines.push(formatParagraphMarkdown(para, includeRevisions));
      }
    }
  }

  // Include available styles
  if (doc.styles.length > 0) {
    lines.push("");
    lines.push("--- Available Styles ---");
    for (const style of doc.styles) {
      lines.push(`[${style.type}] ${style.id} — ${style.name}`);
    }
  }

  return lines.join("\n").trim();
}

function formatParagraphMarkdown(
  para: NormalizedParagraph,
  includeRevisions: boolean = false,
  includePageBreaks: boolean = false,
): string {
  const styleHint = para.style ? ` (${para.style})` : "";
  const listHint = para.listInfo ? ` [list L${para.listInfo.level}]` : "";

  let text = para.text;
  const hasRevisions = includeRevisions && para.revisions && para.revisions.length > 0;

  if (hasRevisions) {
    text = annotateTextWithRevisions(para.text, para.revisions!);
  }

  // Insert [PAGE BREAK] marker at the correct position in the text.
  // When revisions are also active, pageBreakOffset is relative to
  // para.text (the original string). We must map it into the annotated
  // string whose length has changed due to {+...+} / {-...-} markers.
  if (includePageBreaks && para.hasPageBreak) {
    let offset = para.pageBreakOffset ?? 0;

    if (hasRevisions) {
      offset = mapOffsetToAnnotatedText(offset, para.revisions!);
    }

    const before = text.slice(0, offset);
    const after = text.slice(offset);
    if (before.length === 0) {
      text = `[PAGE BREAK] ${after}`;
    } else if (after.length === 0) {
      text = `${before} [PAGE BREAK]`;
    } else {
      text = `${before} [PAGE BREAK] ${after}`;
    }
  }

  return `[${para.id}]${styleHint}${listHint} ${text}`;
}

/**
 * Map an offset in the original paragraph text to the corresponding
 * position in the revision-annotated string.
 *
 * Revision markers ({+...+} for inserts, {-...-} for deletes) add
 * characters that shift later positions. This walks revisions in
 * document order (ascending startOffset) and accumulates the shift.
 */
function mapOffsetToAnnotatedText(
  originalOffset: number,
  revisions: RevisionMark[],
): number {
  // Sort by startOffset ascending — we accumulate shift left-to-right
  const sorted = [...revisions].sort((a, b) => a.startOffset - b.startOffset);
  let shift = 0;

  for (const rev of sorted) {
    if (rev.type === "insert") {
      // {+...+} wraps text at [startOffset, endOffset).
      // The 2-char prefix "{+" is inserted at startOffset.
      // If the break is past (or at) startOffset, it shifts by +2 for the prefix.
      // The 2-char suffix "+}" is inserted at endOffset.
      // If the break is past (or at) endOffset, it shifts by another +2.
      if (originalOffset >= rev.startOffset) shift += 2; // "{+"
      if (originalOffset >= rev.endOffset) shift += 2;   // "+}"
    } else if (rev.type === "delete") {
      // {-text-} is inserted at startOffset (zero-width in original).
      // Total insertion = text.length + 4 chars ("{-" + text + "-}").
      // If the break is at or past the delete position, shift by the full marker length.
      if (originalOffset >= rev.startOffset) {
        shift += rev.text.length + 4;
      }
    }
  }

  return originalOffset + shift;
}

/**
 * Annotate paragraph text with revision markers.
 *
 * Inserts are shown as `{+text+}` and deletes as `{-text-}`.
 * Revisions are applied from right-to-left to avoid offset shifting.
 */
function annotateTextWithRevisions(
  text: string,
  revisions: RevisionMark[],
): string {
  // Sort by startOffset descending so we can splice from right-to-left
  const sorted = [...revisions].sort((a, b) => b.startOffset - a.startOffset);

  let result = text;
  for (const rev of sorted) {
    if (rev.type === "insert") {
      // The inserted text is already in .text at [startOffset, endOffset)
      // Wrap it with {+...+}
      const before = result.slice(0, rev.startOffset);
      const inserted = result.slice(rev.startOffset, rev.endOffset);
      const after = result.slice(rev.endOffset);
      result = `${before}{+${inserted}+}${after}`;
    } else if (rev.type === "delete") {
      // Deleted text is NOT in .text (startOffset === endOffset is a zero-width position)
      // Insert the deleted text marker at that position
      const before = result.slice(0, rev.startOffset);
      const after = result.slice(rev.startOffset);
      result = `${before}{-${rev.text}-}${after}`;
    }
  }

  return result;
}

function formatTableMarkdown(
  table: NormalizedTable,
  includeRevisions: boolean = false,
): string[] {
  const lines: string[] = [];

  for (const row of table.rows) {
    const cells = row.cells.map((cell) => {
      let text = cell.paragraphs.map((p) => p.text).join(" ");
      if (includeRevisions) {
        // Annotate each paragraph's revisions in the concatenated text
        const annotated = cell.paragraphs.map((p) => {
          if (p.revisions && p.revisions.length > 0) {
            return annotateTextWithRevisions(p.text, p.revisions);
          }
          return p.text;
        });
        text = annotated.join(" ");
      }
      const id = cell.paragraphs[0]?.id ?? `${table.id}.r${row.index}.c${cell.colIndex}.p0`;
      return `[${id}] ${text}`;
    });
    lines.push(`| ${cells.join(" | ")} |`);
  }

  return lines;
}
