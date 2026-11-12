/**
 * Paragraph enumeration and occurrence helpers shared by the CLI commands.
 */

import type { Docxtor, NormalizedParagraph, NormalizedTable, BodyElement, ControlInfo } from "../src/index.js";

export interface Located {
  para: NormalizedParagraph;
  /** "body", "table", "header", "footer", "footnote" or "endnote". */
  where: "body" | "table" | "header" | "footer" | "footnote" | "endnote";
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

/** Every paragraph: body (with tables), then headers, footers, footnotes and endnotes. */
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
  for (const note of [...doc.getFootnotes().values(), ...doc.getEndnotes().values()]) {
    for (const p of note.paragraphs) out.push({ para: p, where: note.kind });
    for (const t of note.tables) {
      const cells: Located[] = [];
      tableParagraphs(t, cells);
      for (const c of cells) out.push({ para: c.para, where: note.kind });
    }
  }
  return out;
}

/**
 * The empty content control whose prompt holds the text at `[start, end)`,
 * if there is one, with the range of paragraph text the prompt covers. An
 * edit there fills the control with the whole replacement, so the find and
 * replace must not be trimmed to the words that differ, and one control
 * takes one edit however often the prompt contains the find.
 */
export function promptRange(
  para: NormalizedParagraph,
  start: number,
  end: number,
): { control: ControlInfo; startOffset: number; endOffset: number } | undefined {
  const inline = (para.controls ?? []).filter((r) => start >= r.startOffset && end <= r.endOffset && r.endOffset > r.startOffset);
  // Nested controls follow the one that holds them, so the last match is the innermost.
  const innermost = inline[inline.length - 1] ?? (para.control ? { control: para.control, startOffset: 0, endOffset: para.text.length } : undefined);
  return innermost?.control.placeholder === true ? innermost : undefined;
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

/**
 * Non-overlapping matches of `find` in `text`. With `looseSpace`, a space or a tab in `find` matches
 * either, so a clause typed as "17.2 For goods" finds the "17.2<tab>For goods" a Word file holds.
 */
export function matchRanges(text: string, find: string, wholeWord: boolean, looseSpace = false): { start: number; matched: string }[] {
  if (find.length === 0) return [];
  const body = looseSpace
    ? [...find].map((ch) => (ch === " " || ch === "\t" ? "[ \\t]" : escapeRegExp(ch))).join("")
    : escapeRegExp(find);
  const src = wholeWord ? `(?<![\\p{L}\\p{N}_])${body}(?![\\p{L}\\p{N}_])` : body;
  const re = new RegExp(src, "gu");
  return [...text.matchAll(re)].map((m) => ({ start: m.index!, matched: m[0] }));
}

/** Non-overlapping match starts of `find` in `text`. */
export function matchStarts(text: string, find: string, wholeWord: boolean): number[] {
  return matchRanges(text, find, wholeWord).map((m) => m.start);
}

export interface Hit {
  para: NormalizedParagraph;
  start: number;
  /** The text in the paragraph that matched; differs from what was typed in quotes, or in a space typed for a tab. */
  matched: string;
}

/**
 * Every match of `find` in the given paragraphs. The text as typed is tried first, then its curly and
 * straight quote variants, then the same with a space and a tab treated as one character. The first
 * form that matches anywhere wins; each hit reports the text that actually matched.
 */
export function findHits(paras: Located[], find: string, wholeWord: boolean): Hit[] {
  const variants = quoteVariants(find);
  for (const loose of /[ \t]/.test(find) ? [false, true] : [false]) {
    for (const variant of variants) {
      const hits = paras.flatMap(({ para }) =>
        matchRanges(para.text, variant, wholeWord, loose).map((m) => ({ para, start: m.start, matched: m.matched })));
      if (hits.length) return hits;
    }
  }
  return [];
}

/**
 * The replacement for text that matched as `matched` when `find` was typed. Where the replacement
 * repeats the start or the end of what was typed, it takes the document's own characters there, so a
 * tab typed as a space, or a curly quote typed straight, is kept and not rewritten by the edit.
 */
export function alignReplacement(find: string, matched: string, replace: string): string {
  if (find === matched || find.length !== matched.length) return replace;
  const max = Math.min(find.length, replace.length);
  let pre = 0;
  while (pre < max && find[pre] === replace[pre]) pre++;
  let suf = 0;
  while (suf < max - pre && find[find.length - 1 - suf] === replace[replace.length - 1 - suf]) suf++;
  return matched.slice(0, pre) + replace.slice(pre, replace.length - suf) + matched.slice(matched.length - suf);
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

/**
 * How a patch should name each paragraph: by its native id (`para:4F2A1C7B`) when the paragraph has one
 * that no other paragraph shares, otherwise by its positional id. A native id names one paragraph whatever
 * part it is in, so the patch cannot resolve to a paragraph in another header or footer.
 */
export function patchTargets(paras: Located[]): (para: NormalizedParagraph) => string {
  const count = new Map<string, number>();
  for (const { para } of paras) {
    if (para.paraId) count.set(para.paraId.toUpperCase(), (count.get(para.paraId.toUpperCase()) ?? 0) + 1);
  }
  return (para) => (para.paraId && count.get(para.paraId.toUpperCase()) === 1 ? `para:${para.paraId}` : para.id);
}

/** The container a paragraph id belongs to: `tbl0.r3.c1` for `tbl0.r3.c1.p0`, `fn1` for `fn1.p0`, `""` for a body paragraph. */
function containerOf(id: string): string {
  return id.replace(/\.?p\d+$/, "");
}

/**
 * Two neighbouring paragraphs of one container (body, table cell, header, footer, note) that `find`
 * runs across: its first part ends one paragraph and the rest starts the next. Any spaces typed
 * between the two parts are ignored, since `view` shows a table cell's paragraphs on one line with a
 * space between them.
 */
export function findAcrossBreak(paras: Located[], find: string): [NormalizedParagraph, NormalizedParagraph] | undefined {
  for (const variant of quoteVariants(find)) {
    for (let i = 0; i + 1 < paras.length; i++) {
      const a = paras[i]!.para;
      const b = paras[i + 1]!.para;
      if (containerOf(a.id) !== containerOf(b.id)) continue;
      const head = a.text.trimEnd();
      const tail = b.text.trimStart();
      if (head.length === 0 || tail.length === 0) continue;
      const lastChar = head[head.length - 1]!;
      for (let k = 1; k < variant.length; k++) {
        if (variant[k - 1] !== lastChar) continue;
        const rest = variant.slice(k).trimStart();
        if (rest.length > 0 && head.endsWith(variant.slice(0, k)) && tail.startsWith(rest)) return [a, b];
      }
    }
  }
  return undefined;
}

/**
 * Why a find that matched nothing failed, when docxtor can tell: the text is in a paragraph other than
 * the `--in` one, or it runs across a paragraph break. Returns undefined when the text is nowhere, so
 * the caller keeps its own "not found" message. `every` is all paragraphs of the document; `splitAdvice`
 * says what to do about a find that runs across a break.
 */
export function explainNoMatch(every: Located[], find: string, wholeWord: boolean, inId: string | undefined, splitAdvice: string): string | undefined {
  if (inId) {
    const hits = findHits(every, find, wholeWord);
    if (hits.length === 1) {
      const id = hits[0]!.para.id;
      return `"${find}" is not in ${inId}; it is in ${id}. Use --in ${id}.`;
    }
    if (hits.length > 1) {
      const list = hits.slice(0, 20).map((h, i) => `  ${i + 1}. [${h.para.id}] ${snippet(h.para.text, h.start, h.start + h.matched.length, 40)}`);
      return `"${find}" is not in ${inId}; it matches ${hits.length} times elsewhere. Use --in with one of these ids:\n${list.join("\n")}${hits.length > 20 ? "\n  …" : ""}`;
    }
  }
  const across = findAcrossBreak(every, find);
  if (across) {
    return `"${find}" runs across two paragraphs (${across[0].id} and ${across[1].id}); a find must stay inside one paragraph. ${splitAdvice}`;
  }
  return undefined;
}
