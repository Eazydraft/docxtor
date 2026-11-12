/**
 * LLM projection formats: convert a NormalizedDocument into
 * simplified representations that are easy for LLMs to understand
 * and generate patches against.
 */

import type {
  HeaderFooter,
  NormalizedDocument,
  NormalizedNote,
  NormalizedParagraph,
  NormalizedTable,
  NoteReference,
  RevisionMark,
  RunFormatting,
  RunFilter,
  RunMapping,
  StyleInfo,
  ProjectionOptions,
  ControlInfo,
  ControlRange,
  FXPNode,
} from "../types.js";
import { idsCarrySection } from "../normalizer/header-footer.js";
import { getChildren, getTagName, runHasContent } from "../xml/helpers.js";
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

/** A content control in a paragraph: its properties and the text range it covers. */
export type JsonProjectionControl = Omit<ControlInfo, "xmlRef" | "parent"> & {
  /** Start character offset in paragraph text (inclusive). */
  startOffset: number;
  /** End character offset in paragraph text (exclusive). */
  endOffset: number;
};

interface JsonProjectionParagraph {
  id: string;
  /** Native `w14:paraId`, when the paragraph has one. */
  paraId?: string;
  text: string;
  style?: string;
  listLevel?: number;
  listType?: string;
  numId?: string;
  formatting?: JsonProjectionRunFormatting[];
  revisions?: RevisionMark[];
  /**
   * Footnote and endnote references in the paragraph: the note's id prefix
   * (`fn1`, `en2`) and the offset in `text` where the mark sits. The mark is
   * not part of `text`.
   */
  noteRefs?: { note: string; offset: number }[];
  hasPageBreak?: boolean;
  pageBreakOffset?: number;
  /**
   * Content controls: one wrapped around the paragraph (covering all of its
   * text) first, then those inside it in document order.
   */
  controls?: JsonProjectionControl[];
}

interface JsonProjectionTableCell {
  id: string;
  /** Native `w14:paraId` of the cell's first paragraph, when it has one. */
  paraId?: string;
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
  /** 1-based section, present when more than one section has headers or footers. */
  section?: number;
  paragraphs: JsonProjectionParagraph[];
  tables?: JsonProjectionTable[];
}

export interface JsonProjectionNote {
  /** `fn1`, `en2`: what the ids of the note's paragraphs start with. */
  id: string;
  /** The note's number (the one Word displays, for a note with an automatic mark). */
  number: number;
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
  footnotes?: JsonProjectionNote[];
  endnotes?: JsonProjectionNote[];
  styles?: StyleInfo[];
}

/**
 * Project a NormalizedDocument to a JSON structure suitable for LLMs.
 */
export function projectToJson(
  doc: NormalizedDocument,
  options: ProjectionOptions = {},
): JsonProjection {
  const runFilter = options.runFilter;
  const includeFormatting = (options.includeFormatting ?? false) || runFilter !== undefined;
  const includeHeadersFooters = options.includeHeadersFooters ?? true;
  const includeTables = options.includeTables ?? true;
  const includeRevisions = options.includeRevisions ?? false;
  const includePageBreaks = options.includePageBreaks ?? false;
  const includeNotes = options.includeNotes ?? true;
  activeNoteIds = noteIdLookup(doc);

  const body: (JsonProjectionParagraph | JsonProjectionTable | JsonProjectionPageBreak)[] = [];

  for (const element of doc.body) {
    if (element.kind === "paragraph") {
      body.push(projectParagraphToJson(element, includeFormatting, includeRevisions, includePageBreaks, runFilter));
    } else if (element.kind === "table" && includeTables) {
      body.push(projectTableToJson(element, includeFormatting, includeRevisions, runFilter));
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
          ...(idsCarrySection(header) ? { section: header.sectionIndex } : {}),
          paragraphs: header.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting, includeRevisions, false, runFilter),
          ),
        };
        if (includeTables && header.tables.length > 0) {
          hfEntry.tables = header.tables.map((t) =>
            projectTableToJson(t, includeFormatting, includeRevisions, runFilter),
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
          ...(idsCarrySection(footer) ? { section: footer.sectionIndex } : {}),
          paragraphs: footer.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting, includeRevisions, false, runFilter),
          ),
        };
        if (includeTables && footer.tables.length > 0) {
          hfEntry.tables = footer.tables.map((t) =>
            projectTableToJson(t, includeFormatting, includeRevisions, runFilter),
          );
        }
        result.footers.push(hfEntry);
      }
    }
  }

  if (includeNotes) {
    const projectNotes = (notes: Map<string, NormalizedNote>): JsonProjectionNote[] =>
      [...notes.values()].map((note) => ({
        id: note.idPrefix,
        number: note.number,
        paragraphs: note.paragraphs.map((p) =>
          projectParagraphToJson(p, includeFormatting, includeRevisions, false, runFilter),
        ),
        ...(includeTables && note.tables.length > 0
          ? { tables: note.tables.map((t) => projectTableToJson(t, includeFormatting, includeRevisions, runFilter)) }
          : {}),
      }));
    if (doc.footnotes.size > 0) result.footnotes = projectNotes(doc.footnotes);
    if (doc.endnotes.size > 0) result.endnotes = projectNotes(doc.endnotes);
  }

  return result;
}

// =============================================================================
// Footnote and endnote marks
// =============================================================================

/** Looks up the id prefix (`fn1`, `en2`) of the note a reference points at. */
type NoteIdLookup = (ref: NoteReference) => NormalizedNote | undefined;

function noteIdLookup(doc: NormalizedDocument): NoteIdLookup {
  return (ref) => (ref.kind === "footnote" ? doc.footnotes : doc.endnotes).get(ref.noteId);
}

/**
 * The lookup for the document being projected. Set at the start of each
 * projection call, which is synchronous, so the paragraph and table
 * formatters below do not each need the document passed down to them.
 */
let activeNoteIds: NoteIdLookup = () => undefined;

/**
 * The mark the markdown view prints for a note reference: `[^1]` for a
 * footnote, `[^e1]` for an endnote. A reference with a custom mark prints
 * nothing: the mark is the text that follows it. So does a reference to a
 * note the part does not hold.
 */
function noteMark(ref: NoteReference): string {
  if (ref.customMark) return "";
  const note = activeNoteIds(ref);
  if (!note) return "";
  return ref.kind === "footnote" ? `[^${note.number}]` : `[^e${note.number}]`;
}

/**
 * Whether a run's formatting satisfies every property of a RunFilter.
 * String comparisons are case-insensitive so colour hex casing doesn't matter.
 */
export function runMatchesFilter(formatting: RunFormatting, filter: RunFilter): boolean {
  const eqStr = (a: string | undefined, b: string | undefined): boolean =>
    a !== undefined && b !== undefined && a.toUpperCase() === b.toUpperCase();

  if (filter.bold !== undefined && (formatting.bold ?? false) !== filter.bold) return false;
  if (filter.italic !== undefined && (formatting.italic ?? false) !== filter.italic) return false;
  if (filter.strikethrough !== undefined && (formatting.strikethrough ?? false) !== filter.strikethrough) return false;
  if (filter.underline !== undefined && !eqStr(formatting.underline, filter.underline)) return false;
  if (filter.color !== undefined && !eqStr(formatting.color, filter.color)) return false;
  if (filter.highlight !== undefined && !eqStr(formatting.highlight, filter.highlight)) return false;
  return true;
}

function projectParagraphToJson(
  para: NormalizedParagraph,
  includeFormatting: boolean,
  includeRevisions: boolean = false,
  includePageBreaks: boolean = false,
  runFilter?: RunFilter,
): JsonProjectionParagraph {
  const result: JsonProjectionParagraph = {
    id: para.id,
    text: para.text,
  };

  if (para.paraId) {
    result.paraId = para.paraId;
  }

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
      if (runFilter) {
        if (!runMatchesFilter(r.formatting, runFilter)) continue;
      } else {
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
      }

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
      if (r.formatting.themeColor) fmt.themeColor = r.formatting.themeColor;
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

  if (para.noteRefs && para.noteRefs.length > 0) {
    const refs = para.noteRefs.flatMap((ref) => {
      const note = activeNoteIds(ref);
      return note ? [{ note: note.idPrefix, offset: ref.offset }] : [];
    });
    if (refs.length > 0) result.noteRefs = refs;
  }

  if (includePageBreaks && para.hasPageBreak) {
    result.hasPageBreak = true;
    if (para.pageBreakOffset !== undefined) {
      result.pageBreakOffset = para.pageBreakOffset;
    }
  }

  const controls = controlRangesOf(para);
  if (controls.length > 0) {
    result.controls = controls.map(({ control: c, startOffset, endOffset }) => {
      const out: JsonProjectionControl = { id: c.id, kind: c.kind, placeholder: c.placeholder, startOffset, endOffset };
      if (c.alias !== undefined) out.alias = c.alias;
      if (c.tag !== undefined) out.tag = c.tag;
      if (c.options) out.options = c.options;
      if (c.checked !== undefined) out.checked = c.checked;
      if (c.locked) out.locked = c.locked;
      if (c.bound) out.bound = true;
      return out;
    });
  }

  return result;
}

/**
 * Every content control of a paragraph with the text range it covers: the
 * control wrapped around the paragraph (or its cell) first, covering all of
 * the text, then the inline ones.
 */
function controlRangesOf(para: NormalizedParagraph): ControlRange[] {
  const ranges: ControlRange[] = [];
  if (para.control) {
    ranges.push({ control: para.control, startOffset: 0, endOffset: para.text.length });
  }
  if (para.controls) ranges.push(...para.controls);
  return ranges;
}

function projectTableToJson(
  table: NormalizedTable,
  includeFormatting: boolean = false,
  includeRevisions: boolean = false,
  runFilter?: RunFilter,
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
        if (cell.paragraphs[0]?.paraId) {
          cellResult.paraId = cell.paragraphs[0].paraId;
        }
        // Include individual paragraph details for multi-paragraph cells,
        // when filtering runs (so matching runs in cells are reported), when
        // the cell holds content controls, and when a paragraph holds a
        // footnote or endnote reference: the mark is not in the cell's text,
        // so its `noteRefs` are the only trace of it
        const hasNoteRefs = cell.paragraphs.some((p) => p.noteRefs !== undefined && p.noteRefs.length > 0);
        const hasControls = cell.paragraphs.some((p) => p.control !== undefined || p.controls !== undefined);
        if (cell.paragraphs.length > 1 || runFilter !== undefined || hasNoteRefs || hasControls) {
          cellResult.paragraphs = cell.paragraphs.map((p) =>
            projectParagraphToJson(p, includeFormatting, includeRevisions, false, runFilter),
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
        allNested.push(projectTableToJson(nested, includeFormatting, includeRevisions, runFilter));
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
 *
 * --- Footnotes ---
 * [fn1.p0] Footnote text
 * ```
 *
 * A tab in the text prints as a tab character. A footnote or endnote
 * reference prints as `[^1]` / `[^e1]`; the mark is not part of the
 * paragraph's text.
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
  const includeNotes = options.includeNotes ?? true;
  activeNoteIds = noteIdLookup(doc);
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
      lines.push(`--- Header (${headerFooterLabel(header)}) ---`);
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
      lines.push(`--- Footer (${headerFooterLabel(footer)}) ---`);
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

  if (includeNotes) {
    for (const [title, notes] of [["Footnotes", doc.footnotes], ["Endnotes", doc.endnotes]] as const) {
      if (notes.size === 0) continue;
      lines.push("");
      lines.push(`--- ${title} ---`);
      for (const note of notes.values()) {
        const noteCounters = new Map<string, number>();
        for (const para of note.paragraphs) {
          const listPrefix = buildListPrefix(para, noteCounters, numberingFxp);
          lines.push(formatParagraphMarkdown(para, includeRevisions, false, includeFormatting, listPrefix));
        }
        if (includeTables) {
          for (const table of note.tables) {
            lines.push("");
            lines.push(...formatTableMarkdown(table, includeRevisions, includeFormatting));
            lines.push("");
          }
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
  const text = renderParagraphText(para, includeRevisions, includePageBreaks, includeFormatting);

  const listHint = para.listInfo ? ` {list:${para.listInfo.numId} L${para.listInfo.level}}` : "";
  return `[${formatIdToken(para)}]${styleHint}${listHint} ${listPrefix}${text}`;
}

/**
 * A paragraph's text as the markdown view shows it: formatting markers and
 * revision marks when asked for, a footnote or endnote mark (`[^1]`) where
 * each reference sits, the `[PAGE BREAK]` marker when asked for, and
 * content-control markers (`{field: …}`) around the controls that get one.
 *
 * The marks and markers are not part of `para.text`. Each has an offset in
 * that text, which is mapped through whatever the formatting and revision
 * annotations added before it is inserted.
 */
function renderParagraphText(
  para: NormalizedParagraph,
  includeRevisions: boolean,
  includePageBreaks: boolean,
  includeFormatting: boolean,
): string {
  let text = includeFormatting
    ? annotateTextWithFormatting(para.text, para.runs)
    : para.text;

  // Revision offsets, mapped through the formatting markers when those are shown
  let revisions: RevisionMark[] | undefined;
  if (includeRevisions && para.revisions && para.revisions.length > 0) {
    revisions = includeFormatting
      ? para.revisions.map((rev) => ({
          ...rev,
          startOffset: mapOffsetToFormattedText(rev.startOffset, para.runs),
          endOffset: mapOffsetToFormattedText(rev.endOffset, para.runs),
        }))
      : para.revisions;
    text = annotateTextWithRevisions(text, revisions);
  }

  // Everything else is spliced into the annotated text in one pass, so no
  // marker shifts the position of another.
  const splices: Splice[] = [];

  /** An offset in `para.text` → the same place in the annotated text. */
  const mapOffset = (offset: number): number => {
    let mapped = includeFormatting ? mapOffsetToFormattedText(offset, para.runs) : offset;
    if (revisions) mapped = mapOffsetToAnnotatedText(mapped, revisions);
    return mapped;
  };

  // Footnote and endnote marks, in order where several share an offset
  (para.noteRefs ?? []).forEach((ref, i) => {
    const mark = noteMark(ref);
    if (mark.length > 0) splices.push({ at: mapOffset(ref.offset), rank: 1, seq: i, text: mark });
  });

  // The [PAGE BREAK] marker goes at the break's position in the text, after
  // any note mark there
  if (includePageBreaks && para.hasPageBreak) {
    const offset = mapOffset(para.pageBreakOffset ?? 0);
    const marker =
      offset === 0 ? "[PAGE BREAK] " : offset >= text.length ? " [PAGE BREAK]" : " [PAGE BREAK] ";
    splices.push({ at: offset, rank: 1, seq: STANDALONE_AFTER_NOTES, text: marker });
  }

  // {field …} markers around content controls.
  const ranges = controlRangesOf(para);
  ranges.forEach((range, seq) => {
    const whole = range.control === para.control;
    const empty = range.startOffset === range.endOffset;
    const marker = controlMarker(range.control, empty, whole);
    if (!marker) return;
    const place = (offset: number): number => {
      let at = includeFormatting ? mapBoundaryToFormattedText(offset, para.runs) : offset;
      if (revisions) at = mapBoundaryToAnnotatedText(at, revisions);
      return at;
    };
    if (empty) {
      splices.push({ at: place(range.startOffset), rank: 1, seq: STANDALONE_AFTER_NOTES + 1 + seq, text: marker.open + marker.close });
    } else {
      // At one position: closers first (innermost first), then openers (outermost first)
      splices.push({ at: place(range.startOffset), rank: 2, seq, text: marker.open });
      splices.push({ at: place(range.endOffset), rank: 0, seq: -seq, text: marker.close });
    }
  });

  if (splices.length === 0) return text;
  splices.sort((a, b) => a.at - b.at || a.rank - b.rank || a.seq - b.seq);
  let out = "";
  let last = 0;
  for (const splice of splices) {
    out += text.slice(last, splice.at) + splice.text;
    last = splice.at;
  }
  return out + text.slice(last);
}

/** Order of the page-break marker and empty-control markers after note marks at one position. */
const STANDALONE_AFTER_NOTES = 1_000_000;

/** Text to put into a rendered paragraph at a position. */
interface Splice {
  at: number;
  /** Order at one position: 0 closes a marker, 1 stands alone, 2 opens one. */
  rank: number;
  seq: number;
  text: string;
}

/** The word a control's kind adds to its marker. */
const CONTROL_WORD: Partial<Record<ControlInfo["kind"], string>> = {
  date: " date",
  dropDownList: " dropdown",
  comboBox: " combo",
  checkbox: " checkbox",
  picture: " picture",
};

/**
 * The marker shown around a content control's text, if it gets one.
 *
 * - An empty control showing its prompt: `{field: Insert name}`, with the
 *   kind and the list items for a typed one:
 *   `{field dropdown "One-off" OR "Ongoing": Choose an item.}`.
 * - A filled drop-down list, combo box or date: `{field date= 1 May 2026}`,
 *   so the kind and the options stay visible.
 * - A control with no text at all: `{field:}`, or `{field date:}` and the
 *   like for a typed one; a picture: `{field picture}`.
 * - A filled rich-text or plain-text control, a checkbox (its glyph reads
 *   by itself) and a building-block gallery: no marker; they are content.
 *
 * @param empty - the control covers no text
 * @param whole - the control wraps the whole paragraph rather than sitting in it
 */
function controlMarker(
  control: ControlInfo,
  empty: boolean,
  whole: boolean,
): { open: string; close: string } | undefined {
  const word = CONTROL_WORD[control.kind] ?? "";
  const options =
    control.options && control.options.length > 0
      ? ` ${control.options.map((o) => `"${o}"`).join(" OR ")}`
      : "";

  if (control.kind === "picture") {
    return { open: empty ? "{field picture" : "{field picture: ", close: "}" };
  }
  if (control.placeholder) {
    // Flagged as showing its prompt but holding no text: no separator space
    return { open: empty ? `{field${word}${options}:` : `{field${word}${options}: `, close: "}" };
  }
  if (control.kind === "dropDownList" || control.kind === "comboBox" || control.kind === "date") {
    // With no text there is no value to show, whatever the prompt flag says
    return { open: empty ? `{field${word}${options}:` : `{field${word}${options}= `, close: "}" };
  }
  // A text control with nothing in it is still a field to see. An empty
  // paragraph inside a block control that holds something elsewhere is just
  // an empty paragraph.
  if (
    empty &&
    (control.kind === "richText" || control.kind === "text") &&
    (!whole || !controlHasContent(control))
  ) {
    return { open: "{field:", close: "}" };
  }
  return undefined;
}

/** True if a control holds any current text or object (deleted text does not count). */
function controlHasContent(control: ControlInfo): boolean {
  const visit = (nodes: FXPNode[]): boolean =>
    nodes.some((node) => {
      const tag = getTagName(node);
      if (tag === undefined || tag === "w:del" || tag === "w:moveFrom" || tag === "w:sdtPr") return false;
      if (tag === "w:r") return runHasContent(node);
      return visit(getChildren(node));
    });
  return visit(getChildren(control.xmlRef));
}

/**
 * The id token shown in markdown brackets: `p12`, or `p12 ¶4F2A1C7B` when
 * the paragraph carries a native paraId.
 */
function formatIdToken(para: NormalizedParagraph): string {
  return para.paraId ? `${para.id} ¶${para.paraId}` : para.id;
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
 * Like {@link mapOffsetToFormattedText}, for a boundary between runs: the
 * position after the markers that close at `originalOffset` and before the
 * ones that open there.
 */
function mapBoundaryToFormattedText(
  originalOffset: number,
  runs: RunMapping[],
): number {
  let shift = 0;
  for (const run of runs) {
    const { bold, italic, strikethrough, underline } = run.formatting;
    if (!bold && !italic && !strikethrough && !underline) continue;
    let prefixLen = 0;
    let suffixLen = 0;
    if (strikethrough) { prefixLen += 2; suffixLen += 2; }
    if (bold) { prefixLen += 2; suffixLen += 2; }
    if (italic) { prefixLen += 1; suffixLen += 1; }
    if (underline) { prefixLen += 3; suffixLen += 4; }
    if (originalOffset > run.startOffset) shift += prefixLen;
    if (originalOffset >= run.endOffset) shift += suffixLen;
  }
  return originalOffset + shift;
}

/**
 * Like {@link mapOffsetToAnnotatedText}, for a boundary: the position after
 * the insertion marks that close at `offset` and before the insertion and
 * deletion marks that start there.
 */
function mapBoundaryToAnnotatedText(
  offset: number,
  revisions: RevisionMark[],
): number {
  let shift = 0;
  for (const rev of revisions) {
    if (rev.type === "insert") {
      if (offset > rev.startOffset) shift += 2;
      if (offset >= rev.endOffset) shift += 2;
    } else if (offset > rev.startOffset) {
      shift += rev.text.length + 4;
    }
  }
  return offset + shift;
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
      // Every paragraph carries its own id, so a find can name the paragraph
      // that holds the text and a paragraph break never reads as a space.
      // A cell with no paragraph (a malformed file) keeps its place in the
      // row under the id its first paragraph would have. No paragraph has
      // that id: a patch that targets it is refused with a message that
      // says so (paragraphNotFound in the validator).
      if (cell.paragraphs.length === 0) return `[${table.id}.r${row.index}.c${cell.colIndex}.p0] `;
      return cell.paragraphs
        .map((p) => `[${formatIdToken(p)}] ${renderParagraphText(p, includeRevisions, false, includeFormatting)}`)
        .join(" ");
    });
    lines.push(`| ${cells.join(" | ")} |`);
  }

  return lines;
}

/** `first`, or `first, section 2` when the document's ids carry the section. */
function headerFooterLabel(hf: HeaderFooter): string {
  return idsCarrySection(hf) ? `${hf.variant}, section ${hf.sectionIndex}` : hf.variant;
}
