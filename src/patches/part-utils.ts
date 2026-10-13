/**
 * Helpers for OPC package plumbing: relationships and content types.
 *
 * Used when a part is created on demand (comments.xml, core.xml) or removed
 * (custom.xml). Works on either relationship part — the package-level
 * `_rels/.rels` or the main document's `word/_rels/document.xml.rels`.
 */

import type { FXPNode } from "../types.js";
import { createElement, findChildren, getAttr, getChildren, getTagName } from "../xml/helpers.js";

const RELS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

/** Create an empty Relationships part. */
export function createRelsPart(): FXPNode[] {
  return [
    { "?xml": [{ "#text": "" } as unknown as FXPNode], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } } as FXPNode,
    createElement("Relationships", [], { xmlns: RELS_NAMESPACE }),
  ];
}

/** Create an empty [Content_Types] part with the standard defaults. */
export function createContentTypesPart(): FXPNode[] {
  return [
    createElement("Types", [
      createElement("Default", [], { Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" }),
      createElement("Default", [], { Extension: "xml", ContentType: "application/xml" }),
    ], { xmlns: CONTENT_TYPES_NAMESPACE }),
  ];
}

function getRelsRoot(rels: FXPNode[]): FXPNode | undefined {
  return rels.find((n) => getTagName(n) === "Relationships");
}

function getTypesRoot(contentTypes: FXPNode[]): FXPNode | undefined {
  return contentTypes.find((n) => getTagName(n) === "Types");
}

/** Next unused `rIdN` in a relationships part. */
export function getNextRelationshipId(rels: FXPNode[]): string {
  let maxId = 0;
  const root = getRelsRoot(rels);
  if (root) {
    for (const rel of findChildren(root, "Relationship")) {
      const id = getAttr(rel, "Id");
      if (id) {
        const num = parseInt(id.replace(/^rId/, ""), 10);
        if (!isNaN(num) && num > maxId) maxId = num;
      }
    }
  }
  return `rId${maxId + 1}`;
}

/** Find a relationship by type, returning its target. */
export function findRelationshipByType(
  rels: FXPNode[] | undefined,
  type: string,
): { id: string; target: string } | undefined {
  if (!rels) return undefined;
  const root = getRelsRoot(rels);
  if (!root) return undefined;
  for (const rel of findChildren(root, "Relationship")) {
    if (getAttr(rel, "Type") === type) {
      return { id: getAttr(rel, "Id") ?? "", target: getAttr(rel, "Target") ?? "" };
    }
  }
  return undefined;
}

/**
 * Ensure a relationship of the given type and target exists. Returns its Id.
 */
export function ensureRelationship(
  rels: FXPNode[],
  type: string,
  target: string,
): string {
  const root = getRelsRoot(rels);
  if (!root) throw new Error("Relationships root not found");
  for (const rel of findChildren(root, "Relationship")) {
    if (getAttr(rel, "Type") === type && getAttr(rel, "Target") === target) {
      return getAttr(rel, "Id") ?? "";
    }
  }
  const id = getNextRelationshipId(rels);
  getChildren(root).push(createElement("Relationship", [], { Id: id, Type: type, Target: target }));
  return id;
}

/** Remove every relationship of the given type. Returns the removed targets. */
export function removeRelationshipsByType(
  rels: FXPNode[] | undefined,
  type: string,
): string[] {
  if (!rels) return [];
  const root = getRelsRoot(rels);
  if (!root) return [];
  const children = getChildren(root);
  const removed: string[] = [];
  for (let i = children.length - 1; i >= 0; i--) {
    const child = children[i]!;
    if (getTagName(child) === "Relationship" && getAttr(child, "Type") === type) {
      removed.push(getAttr(child, "Target") ?? "");
      children.splice(i, 1);
    }
  }
  return removed;
}

/** Ensure an Override for `partName` (leading slash) exists in [Content_Types].xml. */
export function ensureContentTypeOverride(
  contentTypes: FXPNode[],
  partName: string,
  contentType: string,
): void {
  const root = getTypesRoot(contentTypes);
  if (!root) return;
  const normalized = partName.startsWith("/") ? partName : `/${partName}`;
  const exists = findChildren(root, "Override").some(
    (o) => getAttr(o, "PartName") === normalized,
  );
  if (!exists) {
    getChildren(root).push(createElement("Override", [], { PartName: normalized, ContentType: contentType }));
  }
}

/** Remove the Override for `partName` (with or without leading slash). */
export function removeContentTypeOverride(
  contentTypes: FXPNode[] | undefined,
  partName: string,
): void {
  if (!contentTypes) return;
  const root = getTypesRoot(contentTypes);
  if (!root) return;
  const children = getChildren(root);
  const normalized = partName.startsWith("/") ? partName : `/${partName}`;
  const idx = children.findIndex(
    (c) => getTagName(c) === "Override" && getAttr(c, "PartName") === normalized,
  );
  if (idx !== -1) children.splice(idx, 1);
}
