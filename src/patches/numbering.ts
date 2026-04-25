/**
 * Numbering definition management for list support.
 *
 * Manages word/numbering.xml — creating numbering definitions for
 * bullet, ordered, and multilevel lists, and resolving list types
 * from existing definitions.
 */

import type { FXPNode, NormalizedDocument } from "../types.js";
import {
  createElement,
  findChild,
  findChildren,
  getChildren,
  getAttr,
  getTagName,
} from "../xml/helpers.js";

// =============================================================================
// Ensure numbering part exists
// =============================================================================

/**
 * Ensure numbering.xml part exists in the document. Creates it if missing.
 * Also ensures content types and relationships are set up.
 * Returns the numbering FXP array.
 */
export function ensureNumberingPart(doc: NormalizedDocument): FXPNode[] {
  if (doc.xmlParts.numbering) {
    return doc.xmlParts.numbering;
  }

  // Create a minimal numbering.xml structure
  const numberingRoot = createElement(
    "w:numbering",
    [],
    {
      "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    },
  );

  doc.xmlParts.numbering = [numberingRoot];

  // Add content type override
  if (doc.xmlParts.contentTypes) {
    const typesNode = doc.xmlParts.contentTypes.find(
      (n) => getTagName(n) === "Types",
    );
    if (typesNode) {
      const children = getChildren(typesNode);
      // Check if override already exists
      const exists = children.some(
        (c) =>
          getTagName(c) === "Override" &&
          getAttr(c, "PartName") === "/word/numbering.xml",
      );
      if (!exists) {
        children.push(
          createElement("Override", [], {
            PartName: "/word/numbering.xml",
            ContentType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
          }),
        );
      }
    }
  }

  // Add relationship
  if (doc.xmlParts.documentRels) {
    const relsNode = doc.xmlParts.documentRels.find(
      (n) => getTagName(n) === "Relationships",
    );
    if (relsNode) {
      const rels = getChildren(relsNode);
      // Check if relationship already exists
      const exists = rels.some(
        (r) =>
          getAttr(r, "Type") ===
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
      );
      if (!exists) {
        // Find next rId
        let maxId = 0;
        for (const r of rels) {
          const id = getAttr(r, "Id");
          if (id) {
            const num = parseInt(id.replace("rId", ""), 10);
            if (!isNaN(num) && num > maxId) maxId = num;
          }
        }
        rels.push(
          createElement("Relationship", [], {
            Id: `rId${maxId + 1}`,
            Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
            Target: "numbering.xml",
          }),
        );
      }
    }
  }

  return doc.xmlParts.numbering;
}

// =============================================================================
// ID generation
// =============================================================================

/** Scan existing w:num entries and return the next available numId. */
export function getNextNumId(numberingFxp: FXPNode[]): number {
  const root = numberingFxp.find((n) => getTagName(n) === "w:numbering");
  if (!root) return 1;

  let max = 0;
  for (const child of findChildren(root, "w:num")) {
    const id = getAttr(child, "w:numId");
    if (id) {
      const num = parseInt(id, 10);
      if (!isNaN(num) && num > max) max = num;
    }
  }
  return max + 1;
}

/** Scan existing w:abstractNum entries and return the next available abstractNumId. */
export function getNextAbstractNumId(numberingFxp: FXPNode[]): number {
  const root = numberingFxp.find((n) => getTagName(n) === "w:numbering");
  if (!root) return 0;

  let max = -1;
  for (const child of findChildren(root, "w:abstractNum")) {
    const id = getAttr(child, "w:abstractNumId");
    if (id) {
      const num = parseInt(id, 10);
      if (!isNaN(num) && num > max) max = num;
    }
  }
  return max + 1;
}

// =============================================================================
// Create numbering definitions
// =============================================================================

/** Bullet symbols per level (cycling). */
const BULLET_CHARS = ["", "o", "", "", "o", "", "", "o", ""];

/** Ordered formats per level (cycling). */
const ORDERED_FORMATS: Array<{ fmt: string; textFn: (lvl: number) => string }> = [
  { fmt: "decimal", textFn: (l) => `%${l + 1}.` },
  { fmt: "lowerLetter", textFn: (l) => `%${l + 1}.` },
  { fmt: "lowerRoman", textFn: (l) => `%${l + 1}.` },
];

/**
 * Create a full numbering definition (w:abstractNum + w:num) for a list type.
 * Returns the numId string to use in w:numPr.
 */
export function createNumberingDefinition(
  doc: NormalizedDocument,
  listType: "bullet" | "ordered" | "multilevel",
  levels: number = 9,
): string {
  const numberingFxp = ensureNumberingPart(doc);
  const root = numberingFxp.find((n) => getTagName(n) === "w:numbering")!;
  const rootChildren = getChildren(root);

  const abstractNumId = getNextAbstractNumId(numberingFxp);
  const numId = getNextNumId(numberingFxp);

  // Build level definitions
  const lvlNodes: FXPNode[] = [];
  for (let i = 0; i < levels; i++) {
    let fmt: string;
    let lvlText: string;

    if (listType === "bullet") {
      fmt = "bullet";
      lvlText = BULLET_CHARS[i % BULLET_CHARS.length] ?? "";
    } else if (listType === "multilevel") {
      fmt = "decimal";
      // Compound: "%1.%2.%3." etc.
      lvlText = Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join(".") + ".";
    } else {
      // ordered — independent counters
      const style = ORDERED_FORMATS[i % ORDERED_FORMATS.length]!;
      fmt = style.fmt;
      lvlText = style.textFn(i);
    }

    lvlNodes.push(
      createElement("w:lvl", [
        createElement("w:start", [], { "w:val": "1" }),
        createElement("w:numFmt", [], { "w:val": fmt }),
        createElement("w:lvlText", [], { "w:val": lvlText }),
        createElement("w:lvlJc", [], { "w:val": "left" }),
      ], { "w:ilvl": String(i) }),
    );
  }

  const abstractNum = createElement(
    "w:abstractNum",
    lvlNodes,
    { "w:abstractNumId": String(abstractNumId) },
  );

  const num = createElement(
    "w:num",
    [createElement("w:abstractNumId", [], { "w:val": String(abstractNumId) })],
    { "w:numId": String(numId) },
  );

  // Insert abstractNum before any w:num elements (OOXML ordering requirement)
  const firstNumIdx = rootChildren.findIndex((c) => getTagName(c) === "w:num");
  if (firstNumIdx >= 0) {
    rootChildren.splice(firstNumIdx, 0, abstractNum);
  } else {
    rootChildren.push(abstractNum);
  }
  // Now append the w:num
  rootChildren.push(num);

  return String(numId);
}

// =============================================================================
// List type resolution
// =============================================================================

/**
 * Resolve list type from numId by inspecting numbering.xml definitions.
 */
export function resolveListType(
  numberingFxp: FXPNode[],
  numId: string,
): "bullet" | "ordered" | "multilevel" | undefined {
  const root = numberingFxp.find((n) => getTagName(n) === "w:numbering");
  if (!root) return undefined;

  // Find the w:num with this numId
  const numNode = findChildren(root, "w:num").find(
    (n) => getAttr(n, "w:numId") === numId,
  );
  if (!numNode) return undefined;

  // Get the abstractNumId
  const abstractNumIdNode = findChild(numNode, "w:abstractNumId");
  if (!abstractNumIdNode) return undefined;
  const abstractNumId = getAttr(abstractNumIdNode, "w:val");
  if (!abstractNumId) return undefined;

  // Find the w:abstractNum
  const abstractNum = findChildren(root, "w:abstractNum").find(
    (n) => getAttr(n, "w:abstractNumId") === abstractNumId,
  );
  if (!abstractNum) return undefined;

  const lvls = findChildren(abstractNum, "w:lvl");

  // Check for bullet
  let hasBullet = false;
  let hasMultiPlaceholder = false;

  for (const lvl of lvls) {
    const numFmt = findChild(lvl, "w:numFmt");
    if (numFmt && getAttr(numFmt, "w:val") === "bullet") {
      hasBullet = true;
    }

    const lvlText = findChild(lvl, "w:lvlText");
    if (lvlText) {
      const pattern = getAttr(lvlText, "w:val") ?? "";
      // Count %N placeholders
      const matches = pattern.match(/%\d+/g);
      if (matches && matches.length > 1) {
        hasMultiPlaceholder = true;
      }
    }
  }

  if (hasBullet) return "bullet";
  if (hasMultiPlaceholder) return "multilevel";
  return "ordered";
}

// =============================================================================
// Level info for projection
// =============================================================================

/**
 * Get the w:lvlText pattern and w:numFmt for a given numId + level.
 */
export function getLevelInfo(
  numberingFxp: FXPNode[],
  numId: string,
  level: number,
): {
  pattern: string;
  format: string;
  ancestorFormats: Map<number, string>;
} | null {
  const root = numberingFxp.find((n) => getTagName(n) === "w:numbering");
  if (!root) return null;

  // Find w:num → abstractNumId
  const numNode = findChildren(root, "w:num").find(
    (n) => getAttr(n, "w:numId") === numId,
  );
  if (!numNode) return null;

  const abstractNumIdNode = findChild(numNode, "w:abstractNumId");
  if (!abstractNumIdNode) return null;
  const abstractNumId = getAttr(abstractNumIdNode, "w:val");
  if (!abstractNumId) return null;

  const abstractNum = findChildren(root, "w:abstractNum").find(
    (n) => getAttr(n, "w:abstractNumId") === abstractNumId,
  );
  if (!abstractNum) return null;

  const lvls = findChildren(abstractNum, "w:lvl");

  // Find the target level
  const targetLvl = lvls.find(
    (l) => getAttr(l, "w:ilvl") === String(level),
  );
  if (!targetLvl) return null;

  const fmtNode = findChild(targetLvl, "w:numFmt");
  const textNode = findChild(targetLvl, "w:lvlText");

  const format = fmtNode ? (getAttr(fmtNode, "w:val") ?? "decimal") : "decimal";
  const pattern = textNode ? (getAttr(textNode, "w:val") ?? "") : "";

  // Build ancestor formats map
  const ancestorFormats = new Map<number, string>();
  const placeholders = pattern.match(/%(\d+)/g);
  if (placeholders) {
    for (const ph of placeholders) {
      const lvlIdx = parseInt(ph.slice(1), 10) - 1; // %1 → level 0
      const ancestorLvl = lvls.find(
        (l) => getAttr(l, "w:ilvl") === String(lvlIdx),
      );
      if (ancestorLvl) {
        const ancestorFmt = findChild(ancestorLvl, "w:numFmt");
        ancestorFormats.set(
          lvlIdx,
          ancestorFmt ? (getAttr(ancestorFmt, "w:val") ?? "decimal") : "decimal",
        );
      }
    }
  }

  return { pattern, format, ancestorFormats };
}
