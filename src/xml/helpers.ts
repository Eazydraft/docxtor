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
      text += "\n";
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
 */
export function createRun(
  text: string,
  rPrNode?: FXPNode,
): FXPNode {
  const children: FXPNode[] = [];
  if (rPrNode) {
    children.push(cloneNode(rPrNode));
  }
  children.push(createWt(text));
  return { "w:r": children };
}
