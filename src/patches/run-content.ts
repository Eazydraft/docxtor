/**
 * Run content: a run's text plus the zero-width children that sit between
 * its characters.
 *
 * A `w:r` holds `w:rPr`, text (`w:t`, and `w:br` soft breaks that read as
 * `\n`) and children that contribute nothing to the paragraph text: page and
 * column breaks, `w:lastRenderedPageBreak`, tabs, drawings, symbols, field
 * characters. Every text edit that rebuilds a run must carry those children
 * across at the character position they occupied, or a page break vanishes
 * because a word next to it was replaced.
 */

import type { FXPNode } from "../types.js";
import {
  getChildren,
  getAttr,
  cloneNode,
  createWt,
  unescapeXmlText,
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
}

/** True for children that are part of the run's text: `w:t` and soft breaks. */
export function isTextChild(child: FXPNode): boolean {
  if ("w:t" in child) return true;
  if ("w:br" in child) {
    const brType = getAttr(child, "w:type");
    return !brType || brType === "textWrapping";
  }
  return false;
}

/**
 * Read a run as text plus anchored children. Mirrors `getRunText`: `w:t`
 * contributes its characters, a soft `w:br` one `\n`, everything else none.
 */
export function readRunContent(runNode: FXPNode): RunContent {
  const content: RunContent = { text: "", anchored: [] };
  for (const child of getChildren(runNode)) {
    if ("w:rPr" in child) {
      content.rPr = child;
    } else if ("w:t" in child) {
      const tChildren = child["w:t"];
      if (Array.isArray(tChildren)) {
        for (const tChild of tChildren) {
          if ("#text" in tChild) {
            content.text += unescapeXmlText(String(tChild["#text"]));
          }
        }
      }
    } else if (isTextChild(child)) {
      content.text += "\n";
    } else {
      content.anchored.push({ node: child, offset: content.text.length });
    }
  }
  return content;
}

/**
 * Rebuild a run's children from content: `w:rPr` first, then the text as
 * `w:t` / `w:br` pairs with each anchored child placed before the character
 * at its offset (offsets past the end go last). Anchored nodes are moved,
 * not cloned.
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
    if (ch === "\n") {
      flush();
      children.push({ "w:br": [] as FXPNode[] });
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

/** Build a fresh `w:r` from content, cloning the run properties. */
export function createRunFromContent(content: RunContent): FXPNode {
  const run: FXPNode = { "w:r": [] as FXPNode[] };
  writeRunContent(run, {
    rPr: content.rPr ? cloneNode(content.rPr) : undefined,
    text: content.text,
    anchored: content.anchored,
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
