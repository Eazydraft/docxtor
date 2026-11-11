/**
 * Native paragraph identity: `w14:paraId`.
 *
 * Word 2010+ stamps every `w:p` with a `w14:paraId` attribute — 8 uppercase
 * hex characters in the range (0, 0x80000000) — that survives round trips
 * through Word. Positional ids (`p12`) shift whenever a paragraph is inserted
 * or deleted; paraIds do not. This module:
 *
 * - generates fresh, document-unique paraIds
 * - stamps missing ids on load (`assignParaIds: true`)
 * - stamps new paragraphs created by patches, provided the part they live in
 *   declares the `w14` namespace (otherwise Word would reject the attribute)
 */

import type { FXPNode, NormalizedDocument, DocxXmlParts } from "../types.js";
import { getAttr, setAttr, getChildren, getTagName } from "../xml/helpers.js";

export const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** Root element tags whose paragraphs may carry paraIds. */
const PART_ROOT_TAGS = new Set(["w:document", "w:hdr", "w:ftr", "w:footnotes", "w:endnotes", "w:comments"]);

// =============================================================================
// Registry of used ids (per xmlParts object, so a cloned document gets its own)
// =============================================================================

const registry = new WeakMap<DocxXmlParts, Set<string>>();

/**
 * Get (building lazily) the set of every paraId currently used in the document.
 */
function getUsedParaIds(doc: NormalizedDocument): Set<string> {
  let used = registry.get(doc.xmlParts);
  if (!used) {
    used = new Set<string>();
    for (const nodes of iterateParts(doc.xmlParts)) {
      collectParaIds(nodes, used);
    }
    registry.set(doc.xmlParts, used);
  }
  return used;
}

function* iterateParts(xmlParts: DocxXmlParts): Generator<FXPNode[]> {
  for (const value of Object.values(xmlParts)) {
    if (Array.isArray(value)) {
      yield value;
    } else if (value instanceof Map) {
      for (const nodes of value.values()) yield nodes;
    }
  }
}

function collectParaIds(nodes: FXPNode[], into: Set<string>): void {
  for (const node of nodes) {
    const tag = getTagName(node);
    if (!tag) continue;
    if (tag === "w:p") {
      const id = getAttr(node, "w14:paraId");
      if (id) into.add(id.toUpperCase());
    }
    const children = getChildren(node);
    if (children.length > 0) collectParaIds(children, into);
  }
}

// =============================================================================
// Generation
// =============================================================================

/**
 * Generate a fresh paraId not present in `used`, and record it.
 * Word accepts values in (0, 0x80000000); we emit 8 uppercase hex chars.
 */
export function generateParaId(used: Set<string>): string {
  for (;;) {
    const value = Math.floor(Math.random() * 0x7ffffffe) + 1;
    const id = value.toString(16).toUpperCase().padStart(8, "0");
    if (!used.has(id)) {
      used.add(id);
      return id;
    }
  }
}

/** Validate the shape of a paraId (8 hex chars, < 0x80000000, non-zero). */
export function isValidParaId(id: string): boolean {
  if (!/^[0-9A-Fa-f]{8}$/.test(id)) return false;
  const n = parseInt(id, 16);
  return n > 0 && n < 0x80000000;
}

// =============================================================================
// Namespace handling
// =============================================================================

/** Whether a part root declares the `w14` namespace prefix. */
export function partDeclaresW14(rootNode: FXPNode | undefined): boolean {
  if (!rootNode) return false;
  return getAttr(rootNode, "xmlns:w14") !== undefined;
}

/**
 * Declare `xmlns:w14` and add `w14` to `mc:Ignorable` on a part root
 * (`w:document`, `w:hdr`, `w:ftr`). Idempotent.
 */
export function ensureW14Namespace(rootNode: FXPNode): void {
  if (getAttr(rootNode, "xmlns:w14") === undefined) {
    setAttr(rootNode, "xmlns:w14", W14_NAMESPACE);
  }
  if (getAttr(rootNode, "xmlns:mc") === undefined) {
    setAttr(rootNode, "xmlns:mc", MC_NAMESPACE);
  }
  const ignorable = getAttr(rootNode, "mc:Ignorable");
  if (ignorable === undefined) {
    setAttr(rootNode, "mc:Ignorable", "w14");
  } else {
    const prefixes = ignorable.split(/\s+/).filter((p) => p.length > 0);
    if (!prefixes.includes("w14")) {
      prefixes.push("w14");
      setAttr(rootNode, "mc:Ignorable", prefixes.join(" "));
    }
  }
}

/** Find the root element node of a part (the first node with a known root tag). */
export function findPartRoot(nodes: FXPNode[]): FXPNode | undefined {
  return nodes.find((n) => {
    const tag = getTagName(n);
    return tag !== undefined && PART_ROOT_TAGS.has(tag);
  });
}

// =============================================================================
// Stamping
// =============================================================================

/**
 * Stamp every `w:p` without a `w14:paraId` in the given part, declaring the
 * namespace on the root if needed. Used by `Docxtor.load(..., { assignParaIds: true })`.
 *
 * @returns number of paragraphs stamped
 */
export function assignMissingParaIds(
  partNodes: FXPNode[],
  used: Set<string>,
): number {
  const root = findPartRoot(partNodes);
  if (!root) return 0;

  let stamped = 0;
  const visit = (nodes: FXPNode[]): void => {
    for (const node of nodes) {
      const tag = getTagName(node);
      if (!tag) continue;
      if (tag === "w:p") {
        const existing = getAttr(node, "w14:paraId");
        if (!existing || !isValidParaId(existing)) {
          const id = generateParaId(used);
          setAttr(node, "w14:paraId", id);
          stamped++;
          if (getAttr(node, "w14:textId") === undefined) {
            setAttr(node, "w14:textId", id);
          }
        } else if (getAttr(node, "w14:textId") === undefined) {
          setAttr(node, "w14:textId", existing.toUpperCase());
        }
      }
      const children = getChildren(node);
      if (children.length > 0) visit(children);
    }
  };

  // Count first so we only touch the namespace if there's something to stamp
  visit(getChildren(root));
  if (stamped > 0) ensureW14Namespace(root);
  return stamped;
}

/**
 * Assign paraIds across every part of a freshly parsed document.
 * Builds the registry so later patches generate collision-free ids.
 */
export function assignParaIdsToDocument(xmlParts: DocxXmlParts): void {
  const used = new Set<string>();
  for (const nodes of iterateParts(xmlParts)) collectParaIds(nodes, used);

  // Only the content parts get stamped (document, headers, footers).
  // Footnotes/endnotes/comments are left alone — Word stamps those itself.
  assignMissingParaIds(xmlParts.document, used);
  for (const nodes of xmlParts.headers.values()) assignMissingParaIds(nodes, used);
  for (const nodes of xmlParts.footers.values()) assignMissingParaIds(nodes, used);

  registry.set(xmlParts, used);
}

/**
 * Stamp fresh paraIds on new paragraph nodes created by a patch.
 *
 * Walks `nodes` recursively (so table cells inside a new `w:tbl` are covered).
 * Does nothing unless the part root declares `w14` — Word would otherwise
 * refuse to open the file. With `overwrite`, existing ids are replaced too
 * (used after cloning a table row, which would otherwise duplicate ids).
 */
export function stampNewParagraphs(
  doc: NormalizedDocument,
  nodes: FXPNode[],
  partRoot: FXPNode | undefined,
  overwrite: boolean = false,
): void {
  if (!partDeclaresW14(partRoot)) return;
  const used = getUsedParaIds(doc);

  const visit = (list: FXPNode[]): void => {
    for (const node of list) {
      const tag = getTagName(node);
      if (!tag) continue;
      if (tag === "w:p" && (overwrite || getAttr(node, "w14:paraId") === undefined)) {
        const id = generateParaId(used);
        setAttr(node, "w14:paraId", id);
        setAttr(node, "w14:textId", id);
      }
      const children = getChildren(node);
      if (children.length > 0) visit(children);
    }
  };
  visit(nodes);
}

/**
 * Resolve the part root node that contains a normalized element, judged by
 * its positional id prefix (`hdr.` / `ftr.` → that header/footer, else body).
 */
export function findPartRootForId(doc: NormalizedDocument, id: string): FXPNode | undefined {
  const m = id.match(/^(hdr|ftr)\./);
  if (m) {
    const map = m[1] === "hdr" ? doc.headers : doc.footers;
    const rootTag = m[1] === "hdr" ? "w:hdr" : "w:ftr";
    for (const hf of map.values()) {
      // `hdr.default.` in a single-section document, `hdr.s2.default.` otherwise
      if (id.startsWith(`${hf.idPrefix}.${hf.variant}.`)) {
        return hf.xmlRef.find((n) => getTagName(n) === rootTag);
      }
    }
    return undefined;
  }
  return doc.xmlParts.document.find((n) => getTagName(n) === "w:document");
}
