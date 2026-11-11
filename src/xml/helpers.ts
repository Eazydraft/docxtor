/**
 * Utility functions for navigating fast-xml-parser's preserveOrder structure.
 *
 * With `preserveOrder: true`, FXP produces arrays of single-key objects:
 *
 *   [{ "w:p": [...children...], ":@": { "@_w:rsidR": "00A7" } }]
 *
 * These helpers abstract over that structure so the rest of the codebase
 * doesn't need to deal with the verbosity directly.
 */

import type { FXPNode } from "../types.js";

// =============================================================================
// Tag name extraction
// =============================================================================

/**
 * Get the tag name of an FXP node.
 * The tag name is the first key that isn't ":@" (attributes) or "#text".
 */
export function getTagName(node: FXPNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key !== ":@" && key !== "#text") {
      return key;
    }
  }
  return undefined;
}

/**
 * Check if a node has a specific tag name.
 */
export function isTag(node: FXPNode, tagName: string): boolean {
  return tagName in node;
}

/**
 * Check if a node is a text node.
 */
export function isTextNode(node: FXPNode): boolean {
  return "#text" in node;
}

// =============================================================================
// Child navigation
// =============================================================================

/**
 * Get the children array of a node (the value under the tag name key).
 */
export function getChildren(node: FXPNode): FXPNode[] {
  const tagName = getTagName(node);
  if (!tagName) return [];
  const children = node[tagName];
  if (Array.isArray(children)) return children;
  return [];
}

/**
 * Find the first child with a specific tag name.
 */
export function findChild(
  node: FXPNode,
  tagName: string,
): FXPNode | undefined {
  const children = getChildren(node);
  return children.find((child) => tagName in child);
}

/**
 * Find all children with a specific tag name.
 */
export function findChildren(
  node: FXPNode,
  tagName: string,
): FXPNode[] {
  const children = getChildren(node);
  return children.filter((child) => tagName in child);
}

/**
 * Find the first descendant matching a path of tag names.
 * E.g., findByPath(node, ["w:body", "w:p"]) finds the first w:p inside w:body.
 */
export function findByPath(
  node: FXPNode,
  path: string[],
): FXPNode | undefined {
  let current: FXPNode | undefined = node;
  for (const tag of path) {
    if (!current) return undefined;
    current = findChild(current, tag);
  }
  return current;
}

// =============================================================================
// Content controls (w:sdt)
// =============================================================================

/**
 * The children array of a content control's `w:sdtContent`: the real array,
 * so splicing it edits the control. A control without `w:sdtContent` gives a
 * detached empty array.
 */
export function sdtContentChildren(sdtNode: FXPNode): FXPNode[] {
  const content = findChild(sdtNode, "w:sdtContent");
  return content ? getChildren(content) : [];
}

/** One child seen through content-control wrappers. */
export interface ContentChild {
  /** The child node (never a `w:sdt`). */
  node: FXPNode;
  /** The array that really holds `node`: the parent's children, or a `w:sdtContent`'s. */
  siblings: FXPNode[];
  /** The `w:sdt` wrappers crossed to reach `node`, outermost first. */
  sdts: FXPNode[];
}

/**
 * The children of `children` in document order, reading through content
 * controls: a `w:sdt` child is replaced by the children of its
 * `w:sdtContent`, recursively. Returns a snapshot, so the caller may splice
 * `siblings` while iterating.
 */
export function contentChildrenOf(children: FXPNode[], sdts: FXPNode[] = []): ContentChild[] {
  const out: ContentChild[] = [];
  for (const child of children) {
    if ("w:sdt" in child) {
      out.push(...contentChildrenOf(sdtContentChildren(child), [...sdts, child]));
    } else {
      out.push({ node: child, siblings: children, sdts });
    }
  }
  return out;
}

/** {@link contentChildrenOf} for the children of a node. */
export function contentChildren(node: FXPNode): ContentChild[] {
  return contentChildrenOf(getChildren(node));
}

// =============================================================================
// Attribute access
// =============================================================================

/**
 * Get an attribute value from a node.
 * Attribute names should NOT include the "@_" prefix — it's added automatically.
 */
export function getAttr(
  node: FXPNode,
  attrName: string,
): string | undefined {
  const attrs = node[":@"] as Record<string, string> | undefined;
  if (!attrs) return undefined;
  return attrs[`@_${attrName}`];
}

/**
 * Set an attribute value on a node.
 */
export function setAttr(
  node: FXPNode,
  attrName: string,
  value: string,
): void {
  if (!node[":@"]) {
    node[":@"] = {};
  }
  (node[":@"] as Record<string, string>)[`@_${attrName}`] = value;
}

/**
 * Remove an attribute from a node.
 */
export function removeAttr(node: FXPNode, attrName: string): void {
  const attrs = node[":@"] as Record<string, string> | undefined;
  if (attrs) {
    delete attrs[`@_${attrName}`];
  }
}

/**
 * Get all attributes as a plain object (without the "@_" prefix).
 */
export function getAttrs(
  node: FXPNode,
): Record<string, string> {
  const attrs = node[":@"] as Record<string, string> | undefined;
  if (!attrs) return {};
  const result: Record<string, string> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (key.startsWith("@_")) {
      result[key.slice(2)] = value;
    }
  }
  return result;
}

// =============================================================================
// Text access
// =============================================================================

/**
 * Escape special XML characters in text content.
 *
 * Required because fast-xml-parser with processEntities: false does NOT
 * auto-escape on build. Entities from the original XML survive as-is,
 * but any NEW text we inject must be manually escaped.
 */
export function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Unescape XML entities in text content for display/comparison.
 *
 * Since processEntities: false keeps entities as-is during parse,
 * we unescape when extracting text for the normalized view so that
 * the paragraph text reads naturally (e.g., "A & B" not "A &amp; B").
 */
export function unescapeXmlText(text: string): string {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

/**
 * Get the text content of a text node.
 */
export function getTextContent(node: FXPNode): string {
  if ("#text" in node) {
    return String(node["#text"]);
  }
  return "";
}

/**
 * Set the text content of a text node.
 * Escapes XML special characters automatically.
 */
export function setTextContent(node: FXPNode, text: string): void {
  node["#text"] = escapeXmlText(text);
}

/**
 * The one character a run child other than `w:t` adds to the text:
 * - `w:tab` and `w:ptab` (the positional tab of headers and footers): `\t`
 * - `w:br` with no type or `textWrapping`, and `w:cr`: `\n`
 *
 * Everything else adds none: page and column breaks, drawings, field
 * characters, footnote references. The `w:tab` elements that define tab
 * stops live in `w:pPr/w:tabs`, not in a run, so they are never seen here.
 */
export function runChildChar(child: FXPNode): string | undefined {
  if ("w:tab" in child || "w:ptab" in child) return "\t";
  if ("w:cr" in child) return "\n";
  if ("w:br" in child) {
    const brType = getAttr(child, "w:type");
    return !brType || brType === "textWrapping" ? "\n" : undefined;
  }
  return undefined;
}

/** The unescaped text of a `w:t` (or `w:delText`) element. */
export function textElementContent(node: FXPNode, tag: "w:t" | "w:delText" = "w:t"): string {
  const children = node[tag];
  let text = "";
  if (Array.isArray(children)) {
    for (const child of children) {
      if ("#text" in child) text += unescapeXmlText(String(child["#text"]));
    }
  }
  return text;
}

/**
 * Get the text of a w:r run node: its w:t elements, unescaped, plus one
 * character for each tab (`\t`) and soft line break (`\n`), in order.
 * Structural breaks, page breaks (w:br w:type="page") and column breaks
 * (w:br w:type="column"), add nothing.
 */
export function getRunText(runNode: FXPNode): string {
  let text = "";
  for (const child of getChildren(runNode)) {
    if ("w:t" in child) text += textElementContent(child);
    else text += runChildChar(child) ?? "";
  }
  return text;
}

/**
 * Find the w:t node inside a w:r run node.
 * Returns the first w:t element found.
 */
export function findTextNode(runNode: FXPNode): FXPNode | undefined {
  const children = getChildren(runNode);
  return children.find((child) => "w:t" in child);
}

/**
 * Check if a run contains a page break.
 *
 * Detects two OOXML mechanisms:
 * - `<w:br w:type="page"/>` — explicit page break inserted by the user
 * - `<w:lastRenderedPageBreak/>` — rendering hint marking where Word
 *   paginated the document on last save
 */
export function runHasPageBreak(runNode: FXPNode): boolean {
  const children = getChildren(runNode);
  for (const child of children) {
    if ("w:br" in child) {
      const brType = getAttr(child, "w:type");
      if (brType === "page") return true;
    }
    if ("w:lastRenderedPageBreak" in child) {
      return true;
    }
  }
  return false;
}

/**
 * Find the character offset within a run's text where a page break occurs.
 *
 * Walks the run's children in document order, accumulating text from w:t
 * elements. When a page break element is found (w:br w:type="page" or
 * w:lastRenderedPageBreak), returns the number of text characters seen so
 * far — i.e., the break position relative to the run's own text.
 *
 * Returns -1 if no page break is found in this run.
 */
export function getRunPageBreakOffset(runNode: FXPNode): number {
  let charsSeen = 0;

  for (const child of getChildren(runNode)) {
    if ("w:t" in child) {
      charsSeen += textElementContent(child).length;
    } else if ("w:lastRenderedPageBreak" in child) {
      return charsSeen;
    } else if ("w:br" in child && getAttr(child, "w:type") === "page") {
      return charsSeen;
    } else {
      // A tab or a soft line break counts as one character
      charsSeen += (runChildChar(child) ?? "").length;
    }
  }

  return -1;
}

/**
 * Check if a paragraph contains a page break in any of its runs,
 * or has the `w:pageBreakBefore` property in its `w:pPr`.
 *
 * Covers three OOXML mechanisms:
 * 1. `<w:br w:type="page"/>` in a run — explicit page break
 * 2. `<w:lastRenderedPageBreak/>` in a run — Word rendering hint
 * 3. `<w:pageBreakBefore/>` in `w:pPr` — paragraph-level "page break before"
 */
export function paragraphHasPageBreak(paragraphNode: FXPNode): boolean {
  // Check paragraph property: w:pPr > w:pageBreakBefore
  const pPr = findChild(paragraphNode, "w:pPr");
  if (pPr) {
    const pbb = findChild(pPr, "w:pageBreakBefore");
    if (pbb) {
      // <w:pageBreakBefore/> means on; <w:pageBreakBefore w:val="0"/> means off
      const val = getAttr(pbb, "w:val");
      if (val !== "0" && val !== "false") {
        return true;
      }
    }
  }

  // Check runs for w:br type="page" or w:lastRenderedPageBreak
  return _scanForPageBreak(getChildren(paragraphNode));
}

function _scanForPageBreak(children: FXPNode[]): boolean {
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:r") {
      if (runHasPageBreak(child)) return true;
    } else if (tag === "w:hyperlink" || tag === "w:ins") {
      if (_scanForPageBreak(getChildren(child))) return true;
    } else if (tag === "w:sdt") {
      if (_scanForPageBreak(sdtContentChildren(child))) return true;
    }
  }
  return false;
}

// =============================================================================
// Node manipulation
// =============================================================================

/**
 * Deep clone an FXP node.
 * Creates a completely independent copy with no shared references.
 */
export function cloneNode(node: FXPNode): FXPNode {
  return JSON.parse(JSON.stringify(node)) as FXPNode;
}

/**
 * Insert a child node after a reference child in a parent's children array.
 * Returns the index where the new child was inserted.
 */
export function insertAfter(
  parent: FXPNode,
  refChild: FXPNode,
  newChild: FXPNode,
): number {
  const children = getChildren(parent);
  const index = children.indexOf(refChild);
  if (index === -1) {
    throw new Error("Reference child not found in parent");
  }
  children.splice(index + 1, 0, newChild);
  return index + 1;
}

/**
 * Insert a child node before a reference child in a parent's children array.
 * Returns the index where the new child was inserted.
 */
export function insertBefore(
  parent: FXPNode,
  refChild: FXPNode,
  newChild: FXPNode,
): number {
  const children = getChildren(parent);
  const index = children.indexOf(refChild);
  if (index === -1) {
    throw new Error("Reference child not found in parent");
  }
  children.splice(index, 0, newChild);
  return index;
}

/**
 * Remove a child node from a parent's children array.
 * Returns true if the child was found and removed.
 */
export function removeChild(
  parent: FXPNode,
  child: FXPNode,
): boolean {
  const children = getChildren(parent);
  const index = children.indexOf(child);
  if (index === -1) return false;
  children.splice(index, 1);
  return true;
}

/**
 * Insert a child node into a parent's children array at a specific index.
 */
export function insertChildAt(
  parent: FXPNode,
  index: number,
  newChild: FXPNode,
): void {
  const children = getChildren(parent);
  children.splice(index, 0, newChild);
}

/**
 * Create a new FXP element node.
 */
export function createElement(
  tagName: string,
  children: FXPNode[] = [],
  attrs?: Record<string, string>,
): FXPNode {
  const node: FXPNode = { [tagName]: children };
  if (attrs && Object.keys(attrs).length > 0) {
    const fxpAttrs: Record<string, string> = {};
    for (const [key, value] of Object.entries(attrs)) {
      fxpAttrs[`@_${key}`] = value;
    }
    node[":@"] = fxpAttrs;
  }
  return node;
}

/**
 * Create a text node.
 * Escapes XML special characters automatically.
 */
export function createTextNode(text: string): FXPNode {
  return { "#text": escapeXmlText(text) } as unknown as FXPNode;
}

/**
 * Create a w:t element with text content, preserving spaces.
 */
export function createWt(text: string): FXPNode {
  const node: FXPNode = {
    "w:t": [createTextNode(text)],
  };
  // Add xml:space="preserve" if text has leading/trailing whitespace
  if (needsSpacePreserve(text)) {
    setAttr(node, "xml:space", "preserve");
  }
  return node;
}

/**
 * Check if text needs xml:space="preserve" on a w:t element.
 * Returns true if text has leading/trailing whitespace or double spaces.
 */
export function needsSpacePreserve(text: string): boolean {
  return text.length > 0 && (text !== text.trim() || text.includes("  "));
}

/**
 * Turn run text into run children: `w:t` for the characters, `<w:tab/>` for
 * each `\t` and `<w:br/>` for each `\n`, the reverse of `getRunText`.
 * Empty text gives one empty `w:t`.
 *
 * `reuse` holds nodes the text had before an edit (`w:tab`, `w:ptab`,
 * `w:br`, `w:cr`), in order. A `\t` or `\n` takes the first of them that
 * reads as that character, so a positional tab that survives an edit is
 * still a `w:ptab`; one that is new in the text is a plain `w:tab`. Nodes
 * taken are removed from `reuse`.
 */
export function textToRunChildren(text: string, reuse?: FXPNode[]): FXPNode[] {
  const children: FXPNode[] = [];
  let buf = "";
  const flush = (): void => {
    if (buf.length > 0) children.push(createWt(buf));
    buf = "";
  };
  for (const ch of text) {
    if (ch === "\t" || ch === "\n") {
      flush();
      children.push(takeCharNode(ch, reuse));
    } else {
      buf += ch;
    }
  }
  flush();
  if (children.length === 0) children.push(createWt(""));
  return children;
}

/** The node for a `\t` or `\n`: the first matching one in `reuse`, or a new one. */
export function takeCharNode(ch: string, reuse?: FXPNode[]): FXPNode {
  if (reuse) {
    const idx = reuse.findIndex((n) => runChildChar(n) === ch);
    if (idx !== -1) return reuse.splice(idx, 1)[0]!;
  }
  return ch === "\t" ? { "w:tab": [] as FXPNode[] } : { "w:br": [] as FXPNode[] };
}

/** The tab and soft-break nodes of a run, in order (see `textToRunChildren`). */
export function runCharNodes(runNode: FXPNode): FXPNode[] {
  return getChildren(runNode).filter((c) => runChildChar(c) !== undefined);
}

/**
 * Create a w:r run element with text and optional formatting.
 * If rPrNode is provided, it's cloned and used as the run properties.
 * `\t` and `\n` in the text become `w:tab` and `w:br` (see `textToRunChildren`).
 */
export function createRun(
  text: string,
  rPrNode?: FXPNode,
  reuse?: FXPNode[],
): FXPNode {
  const children: FXPNode[] = [];
  if (rPrNode) {
    children.push(cloneNode(rPrNode));
  }
  children.push(...textToRunChildren(text, reuse));
  return { "w:r": children };
}
