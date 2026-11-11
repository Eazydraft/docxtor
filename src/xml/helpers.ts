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
 * Get all text from w:t elements inside a w:r run node.
 * Unescapes XML entities so the returned text is natural.
 * Converts w:br (line break) elements to \n characters.
 * Structural breaks — page breaks (w:br w:type="page") and column
 * breaks (w:br w:type="column") — are skipped (not included in text).
 */
export function getRunText(runNode: FXPNode): string {
  const children = getChildren(runNode);
  let text = "";
  for (const child of children) {
    if ("w:t" in child) {
      const tChildren = child["w:t"];
      if (Array.isArray(tChildren)) {
        for (const tChild of tChildren) {
          if ("#text" in tChild) {
            text += unescapeXmlText(String(tChild["#text"]));
          }
        }
      }
    } else if ("w:br" in child) {
      // Page breaks (w:type="page") and column breaks (w:type="column")
      // are structural breaks, not text — skip them
      const brType = getAttr(child, "w:type");
      if (!brType || brType === "textWrapping") {
        // Regular line break
        text += "\n";
      }
      // page/column breaks produce no text
    }
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
  const children = getChildren(runNode);
  let charsSeen = 0;

  for (const child of children) {
    if ("w:t" in child) {
      const tChildren = child["w:t"];
      if (Array.isArray(tChildren)) {
        for (const tChild of tChildren) {
          if ("#text" in tChild) {
            charsSeen += unescapeXmlText(String(tChild["#text"])).length;
          }
        }
      }
    } else if ("w:br" in child) {
      const brType = getAttr(child, "w:type");
      if (!brType || brType === "textWrapping") {
        // Regular line break counts as one char (\n)
        charsSeen += 1;
      } else if (brType === "page") {
        return charsSeen;
      }
    } else if ("w:lastRenderedPageBreak" in child) {
      return charsSeen;
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
 * Create a w:r run element with text and optional formatting.
 * If rPrNode is provided, it's cloned and used as the run properties.
 *
 * Handles newlines (\n) by interleaving w:t and w:br elements,
 * matching the behavior of getRunText() which converts w:br → \n.
 */
export function createRun(
  text: string,
  rPrNode?: FXPNode,
): FXPNode {
  const children: FXPNode[] = [];
  if (rPrNode) {
    children.push(cloneNode(rPrNode));
  }

  if (text.includes("\n")) {
    // Multi-line: interleave w:t and w:br
    const segments = text.split("\n");
    for (let i = 0; i < segments.length; i++) {
      const seg = segments[i]!;
      if (seg.length > 0 || i === 0) {
        children.push(createWt(seg));
      }
      if (i < segments.length - 1) {
        children.push({ "w:br": [] as FXPNode[] });
      }
    }
  } else {
    children.push(createWt(text));
  }

  return { "w:r": children };
}
