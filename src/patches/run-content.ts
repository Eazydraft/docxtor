/**
 * Run content: a run's text plus the zero-width children that sit between
 * its characters.
 *
 * A `w:r` holds `w:rPr`, text (`w:t`, tabs that read as `\t`, and `w:br`
 * soft breaks that read as `\n`) and children that contribute nothing to the
 * paragraph text: page and column breaks, `w:lastRenderedPageBreak`,
 * drawings, symbols, field characters, footnote references. Every text edit
 * that rebuilds a run must carry those children across at the character
 * position they occupied, or a page break vanishes because a word next to it
 * was replaced.
 */

import type { FXPNode } from "../types.js";
import {
  getChildren,
  cloneNode,
  createWt,
  runChildChar,
  takeCharNode,
  textElementContent,
} from "../xml/helpers.js";

/** A non-text child of a run, anchored before the character at `offset`. */
export interface AnchoredChild {
  node: FXPNode;
  /** Characters of run text that precede this child (0 = before the first). */
  offset: number;
}

/** A run taken apart: its properties, its text and its anchored children. */
export interface RunContent {
  rPr?: FXPNode;
  text: string;
  anchored: AnchoredChild[];
  /**
   * The node behind each `\t` and `\n` of `text`, by character offset. A run
   * rebuilt from this content writes that node back, so a `w:ptab` stays a
   * `w:ptab` and a `w:cr` a `w:cr`. A character with no entry is written as
   * a new `w:tab` or `w:br`.
   */
  charNodes?: Map<number, FXPNode>;
}

/** True for children that are part of the run's text: `w:t`, tabs and soft breaks. */
export function isTextChild(child: FXPNode): boolean {
  return "w:t" in child || runChildChar(child) !== undefined;
}

/**
 * Read a run as text plus anchored children. Mirrors `getRunText`: `w:t`
 * contributes its characters, a tab one `\t`, a soft break one `\n`,
 * everything else none.
 */
export function readRunContent(runNode: FXPNode): RunContent {
  const content: RunContent = { text: "", anchored: [], charNodes: new Map() };
  for (const child of getChildren(runNode)) {
    if ("w:rPr" in child) {
      content.rPr = child;
    } else if ("w:t" in child) {
      content.text += textElementContent(child);
    } else {
      const ch = runChildChar(child);
      if (ch !== undefined) {
        content.charNodes!.set(content.text.length, child);
        content.text += ch;
      } else {
        content.anchored.push({ node: child, offset: content.text.length });
      }
    }
  }
  return content;
}

/**
 * Rebuild a run's children from content: `w:rPr` first, then the text as
 * `w:t`, `w:tab` and `w:br` nodes with each anchored child placed before the
 * character at its offset (offsets past the end go last). Anchored nodes and
 * the nodes in `charNodes` are moved, not cloned.
 */
export function writeRunContent(runNode: FXPNode, content: RunContent): void {
  const children = getChildren(runNode);
  children.length = 0;
  if (content.rPr) children.push(content.rPr);

  const anchored = [...content.anchored].sort((a, b) => a.offset - b.offset);
  const text = content.text;
  let next = 0;
  let buf = "";
  const flush = (): void => {
    if (buf.length > 0) children.push(createWt(buf));
    buf = "";
  };
  for (let i = 0; i < text.length; i++) {
    while (next < anchored.length && anchored[next]!.offset <= i) {
      flush();
      children.push(anchored[next]!.node);
      next++;
    }
    const ch = text[i]!;
    if (ch === "\n" || ch === "\t") {
      flush();
      const own = content.charNodes?.get(i);
      children.push(own && runChildChar(own) === ch ? own : takeCharNode(ch));
    } else {
      buf += ch;
    }
  }
  flush();
  while (next < anchored.length) {
    children.push(anchored[next]!.node);
    next++;
  }
  // A run that ends up with neither text nor anchored children keeps an
  // empty w:t, as a cleared run always has.
  if (text.length === 0 && anchored.length === 0) {
    children.push(createWt(""));
  }
}

/**
 * Where the tab and soft-break nodes of `oldText` belong in `newText`.
 *
 * Characters in the part the two texts share at the start and at the end
 * keep their own node. In the part that changed, each `\t` or `\n` of the
 * new text takes the next old node that reads as that character, first from
 * the changed part of the old text and then from `extra` (the nodes of other
 * runs the edit removed), so a tab that a replacement keeps is the original
 * element.
 */
export function carryCharNodes(
  oldText: string,
  oldNodes: Map<number, FXPNode> | undefined,
  newText: string,
  extra: FXPNode[] = [],
): Map<number, FXPNode> {
  const result = new Map<number, FXPNode>();
  const old = oldNodes ?? new Map<number, FXPNode>();
  const max = Math.min(oldText.length, newText.length);
  let pre = 0;
  while (pre < max && oldText[pre] === newText[pre]) pre++;
  let suf = 0;
  while (suf < max - pre && oldText[oldText.length - 1 - suf] === newText[newText.length - 1 - suf]) suf++;

  const pool: FXPNode[] = [];
  for (const [offset, node] of [...old].sort((a, b) => a[0] - b[0])) {
    if (offset < pre) result.set(offset, node);
    else if (offset >= oldText.length - suf) result.set(offset - oldText.length + newText.length, node);
    else pool.push(node);
  }
  pool.push(...extra);
  for (let i = pre; i < newText.length - suf && pool.length > 0; i++) {
    const ch = newText[i]!;
    if (ch !== "\t" && ch !== "\n") continue;
    const idx = pool.findIndex((n) => runChildChar(n) === ch);
    if (idx !== -1) result.set(i, pool.splice(idx, 1)[0]!);
  }
  return result;
}

/** Build a fresh `w:r` from content, cloning the run properties. */
export function createRunFromContent(content: RunContent): FXPNode {
  const run: FXPNode = { "w:r": [] as FXPNode[] };
  writeRunContent(run, {
    rPr: content.rPr ? cloneNode(content.rPr) : undefined,
    text: content.text,
    anchored: content.anchored,
    charNodes: content.charNodes,
  });
  return run;
}

/**
 * Cut content at `[start, end)` into three parts. Anchored children at or
 * before `start` go to the first part and those at or after `end` to the
 * last, so a break that sits right next to a match stays outside it; only
 * children strictly inside the range travel with the middle part.
 */
export function partitionContent(
  content: RunContent,
  start: number,
  end: number,
  atStart: "before" | "after" = "before",
): { before: RunContent; matched: RunContent; after: RunContent } {
  const before: RunContent = { rPr: content.rPr, text: content.text.slice(0, start), anchored: [] };
  const matched: RunContent = { rPr: content.rPr, text: content.text.slice(start, end), anchored: [] };
  const after: RunContent = { rPr: content.rPr, text: content.text.slice(end), anchored: [] };
  for (const [offset, node] of content.charNodes ?? []) {
    if (offset < start) (before.charNodes ??= new Map()).set(offset, node);
    else if (offset >= end) (after.charNodes ??= new Map()).set(offset - end, node);
    else (matched.charNodes ??= new Map()).set(offset - start, node);
  }
  for (const a of content.anchored) {
    const leading = atStart === "before" ? a.offset <= start : a.offset < start;
    if (leading) before.anchored.push(a);
    else if (a.offset >= end) after.anchored.push({ node: a.node, offset: a.offset - end });
    else matched.anchored.push({ node: a.node, offset: a.offset - start });
  }
  return { before, matched, after };
}

/** True if the content has anything worth keeping as a run. */
export function hasContent(content: RunContent): boolean {
  return content.text.length > 0 || content.anchored.length > 0;
}
