/**
 * Paragraph enumeration and occurrence helpers shared by the CLI commands.
 */

import type { Docxtor, NormalizedParagraph, NormalizedTable, BodyElement } from "../src/index.js";

export interface Located {
  para: NormalizedParagraph;
  /** "body", "table", "header" or "footer". */
  where: "body" | "table" | "header" | "footer";
}

function tableParagraphs(table: NormalizedTable, out: Located[]): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const p of cell.paragraphs) out.push({ para: p, where: "table" });
      for (const t of cell.nestedTables) tableParagraphs(t, out);
    }
  }
}

/** Body paragraphs and table-cell paragraphs, in document order. */
export function bodyParagraphs(body: BodyElement[]): Located[] {
  const out: Located[] = [];
  for (const el of body) {
    if (el.kind === "paragraph") out.push({ para: el, where: "body" });
    else if (el.kind === "table") tableParagraphs(el, out);
  }
  return out;
}

/** Every paragraph: body (with tables), then headers, then footers. */
export function allParagraphs(doc: Docxtor): Located[] {
  const out = bodyParagraphs(doc.getBody());
  for (const [kind, map] of [["header", doc.getHeaders()], ["footer", doc.getFooters()]] as const) {
    for (const hf of map.values()) {
      for (const p of hf.paragraphs) out.push({ para: p, where: kind });
      for (const t of hf.tables) {
        const cells: Located[] = [];
        tableParagraphs(t, cells);
        for (const c of cells) out.push({ para: c.para, where: kind });
      }
    }
  }
  return out;
}

/**
 * True when the text at `[start, end)` is the prompt of an empty content
 * control. An edit there fills the control with the whole replacement, so
 * the find and replace must not be trimmed to the words that differ.
 */
export function inPlaceholder(para: NormalizedParagraph, start: number, end: number): boolean {
  const inline = (para.controls ?? []).filter((r) => start >= r.startOffset && end <= r.endOffset && r.endOffset > r.startOffset);
  // Nested controls follow the one that holds them, so the last match is the innermost.
  const innermost = inline[inline.length - 1]?.control ?? para.control;
  return innermost?.placeholder === true;
}

/**
 * The library's occurrence number for a match starting at `start`: docxtor
 * counts matches with `indexOf(find, idx + 1)`, so overlapping matches count.
 */
export function occurrenceAt(text: string, find: string, start: number): number {
  let n = 0;
  let idx = text.indexOf(find);
  while (idx !== -1 && idx <= start) {
    n++;
    idx = text.indexOf(find, idx + 1);
  }
  return n;
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Non-overlapping match starts of `find` in `text`. */
export function matchStarts(text: string, find: string, wholeWord: boolean): number[] {
  if (find.length === 0) return [];
  const src = wholeWord
    ? `(?<![\\p{L}\\p{N}_])${escapeRegExp(find)}(?![\\p{L}\\p{N}_])`
    : escapeRegExp(find);
  const re = new RegExp(src, "gu");
  return [...text.matchAll(re)].map((m) => m.index!);
}

/** Straight-quote and curly-quote variants of a search string, original first. */
export function quoteVariants(s: string): string[] {
  const out = [s];
  if (/['"]/.test(s)) {
    out.push(s.replace(/'/g, "’").replace(/"(?=\S)/g, "“").replace(/"/g, "”"));
  }
  if (/[‘’“”]/.test(s)) {
    out.push(s.replace(/[‘’]/g, "'").replace(/[“”]/g, '"'));
  }
  return [...new Set(out)];
}

export function snippet(text: string, start: number, end: number, width = 50): string {
  const a = Math.max(0, start - width);
  const b = Math.min(text.length, end + width);
  return `${a > 0 ? "…" : ""}${text.slice(a, b)}${b < text.length ? "…" : ""}`.replace(/\s+/g, " ");
}

/**
 * Trim the words `find` and `replace` share at either end, so a long find
 * ("within thirty (30) days of receipt" -> "within sixty (60) days of receipt")
 * tracks only what differs ("thirty (30)" -> "sixty (60)"), as Word does.
 * Returns the offset of the trimmed find inside the original find.
 */
export function minimalEdit(find: string, replace: string): { offset: number; find: string; replace: string } {
  // Whole whitespace-delimited words, so "thirty (30)" -> "sixty (60)" is marked as such, not "thirty (30" -> "sixty (60".
  const words = (s: string): string[] => s.match(/\s+|\S+/g) ?? [];
  const a = words(find);
  const b = words(replace);
  if (a.join("") !== find || b.join("") !== replace) return { offset: 0, find, replace };
  let pre = 0;
  while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre++;
  let suf = 0;
  while (suf < a.length - pre && suf < b.length - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
  // Keep at least one token of find: the edit needs something to anchor on.
  while (pre + suf >= a.length && (pre > 0 || suf > 0)) {
    if (suf > 0) suf--; else pre--;
  }
  const offset = a.slice(0, pre).join("").length;
  return {
    offset,
    find: a.slice(pre, a.length - suf).join(""),
    replace: b.slice(pre, b.length - suf).join(""),
  };
}

export type PlannedEdit =
  | { kind: "replace"; offset: number; find: string; replace: string }
  | { kind: "delete"; offset: number; find: string }
  | { kind: "insert"; offset: number; anchor: string; position: "before" | "after"; text: string };

/**
 * The tracked edit for one find/replace pair. Whole words are trimmed first (minimalEdit). If what is
 * left is a pure deletion or insertion at the character level ("including, without limitation," ->
 * "including"; "Seller" -> "Sellers"), it becomes exactly that, as Word would mark it. Otherwise the
 * whole differing words are replaced, so "thirty (30)" -> "sixty (60)" is not split mid-word.
 */
export function planEdit(find: string, replace: string): PlannedEdit {
  const m = minimalEdit(find, replace);
  const f = m.find, r = m.replace;
  let p = 0;
  while (p < f.length && p < r.length && f[p] === r[p]) p++;
  let q = 0;
  while (q < f.length - p && q < r.length - p && f[f.length - 1 - q] === r[r.length - 1 - q]) q++;
  if (r.length - p - q === 0 && f.length - p - q > 0) {
    return { kind: "delete", offset: m.offset + p, find: f.slice(p, f.length - q) };
  }
  if (f.length - p - q === 0 && r.length - p - q > 0) {
    const text = r.slice(p, r.length - q);
    return p > 0
      ? { kind: "insert", offset: m.offset, anchor: f.slice(0, p), position: "after", text }
      : { kind: "insert", offset: m.offset, anchor: f, position: "before", text };
  }
  return m.replace === "" ? { kind: "delete", offset: m.offset, find: f } : { kind: "replace", ...m };
}
