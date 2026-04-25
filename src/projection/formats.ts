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
  RunMapping,
  StyleInfo,
  ProjectionOptions,
} from "../types.js";
import { getLevelInfo } from "../patches/numbering.js";

// =============================================================================
// JSON Projection
// =============================================================================

export interface JsonProjectionRunFormatting extends RunFormatting {
  /** The run's text content. */
  text: string;
  /** Start character offset in paragraph text (inclusive). */
  startOffset: number;
  /** End character offset in paragraph text (exclusive). */
  endOffset: number;
}

interface JsonProjectionParagraph {
  id: string;
  text: string;
  style?: string;
  listLevel?: number;
  listType?: string;
  numId?: string;
  formatting?: JsonProjectionRunFormatting[];
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
  tables?: JsonProjectionTable[];
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
        const hfEntry: JsonProjectionHeaderFooter = {
          type: "header",
          variant: header.variant,
          paragraphs: header.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting, includeRevisions),
          ),
        };
        if (includeTables && header.tables.length > 0) {
          hfEntry.tables = header.tables.map((t) =>
            projectTableToJson(t, includeFormatting, includeRevisions),
          );
        }
        result.headers.push(hfEntry);
      }
    }
    if (doc.footers.size > 0) {
      result.footers = [];
      for (const footer of doc.footers.values()) {
        const hfEntry: JsonProjectionHeaderFooter = {
          type: "footer",
          variant: footer.variant,
          paragraphs: footer.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting, includeRevisions),
          ),
        };
        if (includeTables && footer.tables.length > 0) {
          hfEntry.tables = footer.tables.map((t) =>
            projectTableToJson(t, includeFormatting, includeRevisions),
          );
        }
        result.footers.push(hfEntry);
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
    if (para.listInfo.listType) {
      result.listType = para.listInfo.listType;
    }
    result.numId = para.listInfo.numId;
  }

  if (includeFormatting && para.runs.length > 0) {
    const fmtRuns: JsonProjectionRunFormatting[] = [];
    for (const r of para.runs) {
      // Only include runs that have at least one formatting property set
      const hasAny =
        r.formatting.bold ||
        r.formatting.italic ||
        r.formatting.underline ||
        r.formatting.strikethrough ||
        r.formatting.fontSize ||
        r.formatting.fontFamily ||
        r.formatting.color ||
        r.formatting.highlight;
      if (!hasAny) continue;

      const fmt: JsonProjectionRunFormatting = {
        text: r.text,
        startOffset: r.startOffset,
        endOffset: r.endOffset,
      };
      if (r.formatting.bold) fmt.bold = true;
      if (r.formatting.italic) fmt.italic = true;
      if (r.formatting.underline) fmt.underline = r.formatting.underline;
      if (r.formatting.strikethrough) fmt.strikethrough = true;
      if (r.formatting.fontSize) fmt.fontSize = r.formatting.fontSize;
      if (r.formatting.fontFamily) fmt.fontFamily = r.formatting.fontFamily;
      if (r.formatting.color) fmt.color = r.formatting.color;
      if (r.formatting.highlight) fmt.highlight = r.formatting.highlight;
      fmtRuns.push(fmt);
    }
    if (fmtRuns.length > 0) {
      result.formatting = fmtRuns;
    }
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
// List prefix rendering
// =============================================================================

/**
 * Format a number according to a w:numFmt value.
 */
function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case "decimal":
      return String(n);
    case "lowerLetter": {
      // 1→a, 2→b, ..., 26→z, 27→aa
      let result = "";
      let v = n;
      while (v > 0) {
        v--;
        result = String.fromCharCode(97 + (v % 26)) + result;
        v = Math.floor(v / 26);
      }
      return result;
    }
    case "upperLetter": {
      let result = "";
      let v = n;
      while (v > 0) {
        v--;
        result = String.fromCharCode(65 + (v % 26)) + result;
        v = Math.floor(v / 26);
      }
      return result;
    }
    case "lowerRoman":
      return toRoman(n).toLowerCase();
    case "upperRoman":
      return toRoman(n);
    default:
      return String(n);
  }
}

function toRoman(n: number): string {
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
  let result = "";
  let v = n;
  for (let i = 0; i < vals.length; i++) {
    while (v >= vals[i]!) {
      result += syms[i];
      v -= vals[i]!;
    }
  }
  return result;
}

/** Default ordered formats per level (cycling). */
const DEFAULT_ORDERED_FMTS = ["decimal", "lowerLetter", "lowerRoman"];

/**
 * Build list prefix string for a paragraph, updating counter state.
 * Returns empty string for non-list paragraphs.
 */
function buildListPrefix(
  para: NormalizedParagraph,
  counters: Map<string, number>,
  numberingFxp?: import("../types.js").FXPNode[],
): string {
  if (!para.listInfo) return "";

  const { numId, level, listType } = para.listInfo;
  const indent = "  ".repeat(level);

  // Clear all deeper-level counters for this numId
  for (const key of counters.keys()) {
    const [kid, klvl] = key.split(":");
    if (kid === numId && Number(klvl) > level) {
      counters.delete(key);
    }
  }

  // Increment this level's counter
  const counterKey = `${numId}:${level}`;
  const count = (counters.get(counterKey) ?? 0) + 1;
  counters.set(counterKey, count);

  if (listType === "bullet" || !listType) {
    return `${indent}- `;
  }

  if (listType === "multilevel" && numberingFxp) {
    // Resolve compound prefix from w:lvlText pattern
    const info = getLevelInfo(numberingFxp, numId, level);
    if (info) {
      // Replace %N placeholders with counter values
      let prefix = info.pattern;
      const placeholders = prefix.match(/%(\d+)/g);
      if (placeholders) {
        for (const ph of placeholders) {
          const lvlIdx = parseInt(ph.slice(1), 10) - 1;
          const lvlCount = counters.get(`${numId}:${lvlIdx}`) ?? 1;
          const fmt = info.ancestorFormats.get(lvlIdx) ?? "decimal";
          prefix = prefix.replaceAll(ph, formatNumber(lvlCount, fmt));
        }
      }
      return `${indent}${prefix} `;
    }
  }

  // Ordered (independent counters) — resolve format from numbering.xml or use defaults
  let fmt = DEFAULT_ORDERED_FMTS[level % DEFAULT_ORDERED_FMTS.length] ?? "decimal";
  if (numberingFxp) {
    const info = getLevelInfo(numberingFxp, numId, level);
    if (info) {
      fmt = info.format;
    }
  }
  return `${indent}${formatNumber(count, fmt)}. `;
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
  const includeFormatting = options.includeFormatting ?? false;
  const lines: string[] = [];

  // List counter state: "numId:level" → current count
  const listCounters = new Map<string, number>();
  const numberingFxp = doc.xmlParts.numbering;

  for (const element of doc.body) {
    if (element.kind === "paragraph") {
      const listPrefix = buildListPrefix(element, listCounters, numberingFxp);
      lines.push(formatParagraphMarkdown(element, includeRevisions, includePageBreaks, includeFormatting, listPrefix));
    } else if (element.kind === "table" && includeTables) {
      lines.push(""); // blank line before table
      lines.push(...formatTableMarkdown(element, includeRevisions, includeFormatting));
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
      const hdrCounters = new Map<string, number>();
      for (const para of header.paragraphs) {
        const listPrefix = buildListPrefix(para, hdrCounters, numberingFxp);
        lines.push(formatParagraphMarkdown(para, includeRevisions, false, includeFormatting, listPrefix));
      }
      if (includeTables) {
        for (const table of header.tables) {
          lines.push("");
          lines.push(...formatTableMarkdown(table, includeRevisions));
        }
      }
    }

    for (const footer of doc.footers.values()) {
      lines.push("");
      lines.push(`--- Footer (${footer.variant}) ---`);
      const ftrCounters = new Map<string, number>();
      for (const para of footer.paragraphs) {
        const listPrefix = buildListPrefix(para, ftrCounters, numberingFxp);
        lines.push(formatParagraphMarkdown(para, includeRevisions, false, includeFormatting, listPrefix));
      }
      if (includeTables) {
        for (const table of footer.tables) {
          lines.push("");
          lines.push(...formatTableMarkdown(table, includeRevisions));
        }
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
  includeFormatting: boolean = false,
  listPrefix: string = "",
): string {
  const styleHint = para.style ? ` (${para.style})` : "";

  let text = includeFormatting
    ? annotateTextWithFormatting(para.text, para.runs)
    : para.text;
  const hasRevisions = includeRevisions && para.revisions && para.revisions.length > 0;

  if (hasRevisions) {
    if (includeFormatting) {
      // Map revision offsets through formatting annotation shift
      const mappedRevisions = para.revisions!.map((rev) => ({
        ...rev,
        startOffset: mapOffsetToFormattedText(rev.startOffset, para.runs),
        endOffset: mapOffsetToFormattedText(rev.endOffset, para.runs),
      }));
      text = annotateTextWithRevisions(text, mappedRevisions);
    } else {
      text = annotateTextWithRevisions(para.text, para.revisions!);
    }
  }

  // Insert [PAGE BREAK] marker at the correct position in the text.
  if (includePageBreaks && para.hasPageBreak) {
    let offset = para.pageBreakOffset ?? 0;

    if (includeFormatting) {
      offset = mapOffsetToFormattedText(offset, para.runs);
    }

    if (hasRevisions) {
      if (includeFormatting) {
        // Revisions have already been mapped through formatting, so use mapped revisions
        const mappedRevisions = para.revisions!.map((rev) => ({
          ...rev,
          startOffset: mapOffsetToFormattedText(rev.startOffset, para.runs),
          endOffset: mapOffsetToFormattedText(rev.endOffset, para.runs),
        }));
        offset = mapOffsetToAnnotatedText(offset, mappedRevisions);
      } else {
        offset = mapOffsetToAnnotatedText(offset, para.revisions!);
      }
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

  const listHint = para.listInfo ? ` {list:${para.listInfo.numId} L${para.listInfo.level}}` : "";
  return `[${para.id}]${styleHint}${listHint} ${listPrefix}${text}`;
}

/**
 * Annotate paragraph text with GFM inline formatting markers.
 *
 * Walks the runs and wraps text spans with:
 * - **bold**
 * - _italic_
 * - ~~strikethrough~~
 * - <u>underline</u>
 * - Combined: **_bold+italic_**
 *
 * Each run is wrapped independently. Runs with no representable
 * formatting are emitted as plain text.
 */
function annotateTextWithFormatting(
  text: string,
  runs: RunMapping[],
): string {
  if (runs.length === 0) return text;

  const parts: string[] = [];
  let lastEnd = 0;

  for (const run of runs) {
    // Fill any gap between runs with plain text
    if (run.startOffset > lastEnd) {
      parts.push(text.slice(lastEnd, run.startOffset));
    }

    const runText = text.slice(run.startOffset, run.endOffset);
    const { bold, italic, strikethrough, underline } = run.formatting;

    if (!bold && !italic && !strikethrough && !underline) {
      parts.push(runText);
    } else {
      let wrapped = runText;
      // Inner wrapping first, then outer
      if (underline) wrapped = `<u>${wrapped}</u>`;
      if (italic) wrapped = `_${wrapped}_`;
      if (bold) wrapped = `**${wrapped}**`;
      if (strikethrough) wrapped = `~~${wrapped}~~`;
      parts.push(wrapped);
    }

    lastEnd = run.endOffset;
  }

  // Trailing text after last run
  if (lastEnd < text.length) {
    parts.push(text.slice(lastEnd));
  }

  return parts.join("");
}

/**
 * Map an offset in the original paragraph text to the corresponding
 * position in the formatting-annotated string.
 *
 * Formatting markers (**, _, ~~) add characters that shift positions.
 * This walks runs in document order and accumulates the shift.
 */
function mapOffsetToFormattedText(
  originalOffset: number,
  runs: RunMapping[],
): number {
  let shift = 0;

  for (const run of runs) {
    const { bold, italic, strikethrough, underline } = run.formatting;
    if (!bold && !italic && !strikethrough && !underline) continue;

    // Calculate marker lengths for this run
    let prefixLen = 0;
    let suffixLen = 0;
    if (strikethrough) { prefixLen += 2; suffixLen += 2; }
    if (bold) { prefixLen += 2; suffixLen += 2; }
    if (italic) { prefixLen += 1; suffixLen += 1; }
    if (underline) { prefixLen += 3; suffixLen += 4; } // <u> and </u>

    if (originalOffset >= run.startOffset) {
      shift += prefixLen;
    }
    if (originalOffset >= run.endOffset) {
      shift += suffixLen;
    }
  }

  return originalOffset + shift;
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
  includeFormatting: boolean = false,
): string[] {
  const lines: string[] = [];

  for (const row of table.rows) {
    const cells = row.cells.map((cell) => {
      let text: string;
      if (includeFormatting) {
        text = cell.paragraphs
          .map((p) => {
            let t = annotateTextWithFormatting(p.text, p.runs);
            if (includeRevisions && p.revisions && p.revisions.length > 0) {
              const mappedRevisions = p.revisions.map((rev) => ({
                ...rev,
                startOffset: mapOffsetToFormattedText(rev.startOffset, p.runs),
                endOffset: mapOffsetToFormattedText(rev.endOffset, p.runs),
              }));
              t = annotateTextWithRevisions(t, mappedRevisions);
            }
            return t;
          })
          .join(" ");
      } else if (includeRevisions) {
        const annotated = cell.paragraphs.map((p) => {
          if (p.revisions && p.revisions.length > 0) {
            return annotateTextWithRevisions(p.text, p.revisions);
          }
          return p.text;
        });
        text = annotated.join(" ");
      } else {
        text = cell.paragraphs.map((p) => p.text).join(" ");
      }
      const id = cell.paragraphs[0]?.id ?? `${table.id}.r${row.index}.c${cell.colIndex}.p0`;
      return `[${id}] ${text}`;
    });
    lines.push(`| ${cells.join(" | ")} |`);
  }

  return lines;
}
