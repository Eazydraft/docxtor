/**
 * Block-level markdown segmentation for multi-line patch text.
 *
 * Splits patch text (`append_text`, `insert_paragraph` validation) into
 * plain text lines, markdown tables, and markdown lists.
 */

/** A contiguous block of text lines, a markdown table, or a markdown list. */
export type TextSegment =
  | { kind: "text"; lines: string[] }
  | { kind: "table"; headers: string[]; rows: string[][] }
  | { kind: "list"; listType: "bullet" | "ordered"; items: { text: string; level: number }[] };

/**
 * Parse text into segments of plain lines, markdown tables, and markdown lists.
 *
 * A markdown table is detected as:
 * - A header row: `| col1 | col2 | ...`
 * - A separator row: `| --- | --- | ...` (dashes, colons, spaces)
 * - Zero or more data rows: `| val1 | val2 | ...`
 *
 * A markdown list is detected as consecutive lines starting with:
 * - Bullet: `- `, `* `, `+ ` (after optional whitespace)
 * - Ordered: `\d+\. ` (after optional whitespace)
 * Nesting uses 2-space increments. Blank lines within a list don't break it.
 */
export function parseTextSegments(text: string): TextSegment[] {
  const lines = text.split("\n");
  const segments: TextSegment[] = [];
  let i = 0;

  while (i < lines.length) {
    // Try to detect a markdown table starting at line i
    if (
      i + 1 < lines.length &&
      isMarkdownTableRow(lines[i]!) &&
      isMarkdownSeparator(lines[i + 1]!)
    ) {
      const headerCells = parseMarkdownRow(lines[i]!);
      // Skip header and separator
      i += 2;
      const dataRows: string[][] = [];
      while (i < lines.length && isMarkdownTableRow(lines[i]!)) {
        const cells = parseMarkdownRow(lines[i]!);
        // Pad or trim to match header count
        while (cells.length < headerCells.length) cells.push("");
        if (cells.length > headerCells.length)
          cells.length = headerCells.length;
        dataRows.push(cells);
        i++;
      }
      segments.push({ kind: "table", headers: headerCells, rows: dataRows });
      continue;
    }

    // Try to detect a markdown list
    const listMatch = parseListLine(lines[i]!);
    if (listMatch) {
      const items: { text: string; level: number }[] = [];
      const listType = listMatch.listType;
      items.push({ text: listMatch.text, level: listMatch.level });
      i++;

      while (i < lines.length) {
        const line = lines[i]!;
        // Blank lines don't break the list
        if (line.trim().length === 0) {
          i++;
          continue;
        }
        const nextMatch = parseListLine(line);
        if (!nextMatch) break;
        // Type change breaks the list
        if (nextMatch.listType !== listType) break;

        items.push({ text: nextMatch.text, level: nextMatch.level });
        i++;
      }

      segments.push({ kind: "list", listType, items });
      continue;
    }

    // Plain text line
    const last = segments[segments.length - 1];
    if (last && last.kind === "text") {
      last.lines.push(lines[i]!);
    } else {
      segments.push({ kind: "text", lines: [lines[i]!] });
    }
    i++;
  }

  return segments;
}

/** True if the text contains a markdown table or list block. */
export function hasBlockMarkdown(text: string): boolean {
  return parseTextSegments(text).some((s) => s.kind !== "text");
}

/** Bullet pattern: optional whitespace + `- `, `* `, or `+ ` */
const BULLET_RE = /^(\s*)[*+-]\s(.*)$/;
/** Ordered pattern: optional whitespace + `N. ` */
const ORDERED_RE = /^(\s*)\d+\.\s(.*)$/;

function parseListLine(
  line: string,
): { listType: "bullet" | "ordered"; text: string; level: number } | undefined {
  // Normalize tabs to 2 spaces
  const normalized = line.replace(/\t/g, "  ");
  let m = BULLET_RE.exec(normalized);
  if (m) {
    const indent = m[1]!.length;
    return { listType: "bullet", text: m[2]!, level: Math.floor(indent / 2) };
  }
  m = ORDERED_RE.exec(normalized);
  if (m) {
    const indent = m[1]!.length;
    return { listType: "ordered", text: m[2]!, level: Math.floor(indent / 2) };
  }
  return undefined;
}

function isMarkdownTableRow(line: string): boolean {
  const trimmed = line.trim();
  return trimmed.startsWith("|") && trimmed.endsWith("|") && trimmed.length > 1;
}

function isMarkdownSeparator(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed.startsWith("|") || !trimmed.endsWith("|")) return false;
  // Each cell should be dashes (with optional colons for alignment)
  const cells = trimmed.slice(1, -1).split("|");
  return cells.every((cell) => /^\s*:?-+:?\s*$/.test(cell));
}

function parseMarkdownRow(line: string): string[] {
  const trimmed = line.trim();
  // Remove leading/trailing pipes and split
  const inner = trimmed.startsWith("|") ? trimmed.slice(1) : trimmed;
  const content = inner.endsWith("|") ? inner.slice(0, -1) : inner;
  return content.split("|").map((cell) => cell.trim());
}
