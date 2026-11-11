/**
 * Header/footer structural operations: add and remove headers/footers.
 *
 * These ops create/remove XML parts, manage relationships in
 * document.xml.rels and [Content_Types].xml, and mutate sectPr
 * references in the document body.
 */

import type {
  NormalizedDocument,
  AddHeaderPatch,
  AddFooterPatch,
  RemoveHeaderPatch,
  RemoveFooterPatch,
  FXPNode,
  HeaderFooterVariant,
} from "../types.js";
import {
  createElement,
  findChild,
  findChildren,
  getChildren,
  getTagName,
  getAttr,
} from "../xml/helpers.js";
import { normalizeHeadersFooters } from "../normalizer/document.js";
import { buildFieldRuns } from "./field-codes.js";
import type { RevisionContext } from "./revision.js";
import { createInsElement } from "./revision.js";

// =============================================================================
// OOXML variant → w:type mapping
// =============================================================================

/** Map variant names to OOXML w:type attribute values. */
const VARIANT_TYPE_MAP: Record<HeaderFooterVariant, string> = {
  default: "default",
  first: "first",
  even: "even",
};

// =============================================================================
// Add Header
// =============================================================================

/**
 * Execute an add_header patch: creates a new header XML part and wires
 * it into the document's last section.
 */
export function executeAddHeader(
  doc: NormalizedDocument,
  patch: AddHeaderPatch,
  revCtx?: RevisionContext,
  markers: boolean = true,
): void {
  executeAddHeaderFooter(doc, "header", patch.variant, patch.paragraphs, revCtx, markers);
}

// =============================================================================
// Add Footer
// =============================================================================

/**
 * Execute an add_footer patch: creates a new footer XML part and wires
 * it into the document's last section.
 */
export function executeAddFooter(
  doc: NormalizedDocument,
  patch: AddFooterPatch,
  revCtx?: RevisionContext,
  markers: boolean = true,
): void {
  executeAddHeaderFooter(doc, "footer", patch.variant, patch.paragraphs, revCtx, markers);
}

// =============================================================================
// Remove Header
// =============================================================================

/**
 * Execute a remove_header patch: removes a header from the document's
 * last section.
 */
export function executeRemoveHeader(
  doc: NormalizedDocument,
  patch: RemoveHeaderPatch,
): void {
  executeRemoveHeaderFooter(doc, "header", patch.variant);
}

// =============================================================================
// Remove Footer
// =============================================================================

/**
 * Execute a remove_footer patch: removes a footer from the document's
 * last section.
 */
export function executeRemoveFooter(
  doc: NormalizedDocument,
  patch: RemoveFooterPatch,
): void {
  executeRemoveHeaderFooter(doc, "footer", patch.variant);
}

// =============================================================================
// Shared implementation
// =============================================================================

/**
 * Add a header or footer to the document.
 */
function executeAddHeaderFooter(
  doc: NormalizedDocument,
  type: "header" | "footer",
  variant: HeaderFooterVariant,
  paragraphs: string[],
  revCtx?: RevisionContext,
  markers: boolean = true,
): void {
  // 1. Determine next available filename number
  const existingParts = type === "header" ? doc.xmlParts.headers : doc.xmlParts.footers;
  const nextNum = getNextPartNumber(existingParts, type);
  const filename = `${type}${nextNum}.xml`;
  const fullPath = `word/${filename}`;

  // 2. Build the header/footer XML FXP
  const rootTag = type === "header" ? "w:hdr" : "w:ftr";
  const pNodes: FXPNode[] = [];

  for (const text of paragraphs) {
    const runs = buildFieldRuns(text, undefined, markers);

    if (revCtx) {
      // Wrap runs in w:ins for revision tracking
      const revRuns = runs.map((run) => {
        // For field code runs, wrap as-is; for plain runs, use as-is
        return run;
      });
      const insElement = createInsElement(revCtx, revRuns);
      pNodes.push({ "w:p": [insElement] });
    } else {
      pNodes.push({ "w:p": runs });
    }
  }

  const rootNode = createElement(rootTag, pNodes, {
    "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  });
  // Wrap attributes onto the tag using FXP convention (attrs on root via :@)
  const fxpPart: FXPNode[] = [rootNode];

  // 3. Generate a new rId
  const rId = getNextRelationshipId(doc);

  // 4. Add relationship to document.xml.rels
  addRelationship(doc, rId, filename, type);

  // 5. Add content type override
  addContentTypeOverride(doc, fullPath, type);

  // 6. Add header/footer reference to last sectPr
  addSectPrReference(doc, type, variant, rId);

  // 7. Store the new FXP part
  existingParts.set(fullPath, fxpPart);

  // 8. Rebuild doc.headers/doc.footers. A new part can turn a document with
  // headers in one section into one with headers in two, which changes the
  // ids of the parts that were already there.
  normalizeHeadersFooters(doc);
}

/**
 * Remove a header or footer from the document.
 */
function executeRemoveHeaderFooter(
  doc: NormalizedDocument,
  type: "header" | "footer",
  variant: HeaderFooterVariant,
): void {
  // 1. Find the reference in sectPr
  const refTag = type === "header" ? "w:headerReference" : "w:footerReference";
  const sectPr = getLastSectPr(doc);
  if (!sectPr) throw new Error("No sectPr found in document");

  const sectPrChildren = getChildren(sectPr);
  const refNode = sectPrChildren.find((child) => {
    const tag = getTagName(child);
    if (tag !== refTag) return false;
    const refType = getAttr(child, "w:type") ?? "default";
    return refType === VARIANT_TYPE_MAP[variant];
  });

  if (!refNode) {
    throw new Error(`No ${type} with variant "${variant}" found in sectPr`);
  }

  const rId = getAttr(refNode, "r:id");
  if (!rId) throw new Error(`${type} reference has no r:id`);

  // 2. Remove from sectPr
  const idx = sectPrChildren.indexOf(refNode);
  if (idx !== -1) sectPrChildren.splice(idx, 1);

  // 3. Find the target filename from rels
  const target = findRelationshipTarget(doc, rId);
  const fullPath = target ? `word/${target}` : undefined;

  // 4. Remove the relationship
  removeRelationship(doc, rId);

  // 5. Remove content type override
  if (fullPath) {
    removeContentTypeOverride(doc, fullPath);
  }

  // 6. Remove from xmlParts
  const parts = type === "header" ? doc.xmlParts.headers : doc.xmlParts.footers;
  if (fullPath) {
    parts.delete(fullPath);
  }

  // 7. Rebuild the normalized maps (see executeAddHeaderFooter)
  normalizeHeadersFooters(doc);
}

// =============================================================================
// Helpers: file numbering
// =============================================================================

/**
 * Find the next available number for a header/footer filename.
 */
function getNextPartNumber(
  parts: Map<string, FXPNode[]>,
  type: "header" | "footer",
): number {
  let maxNum = 0;
  const regex = new RegExp(`${type}(\\d+)\\.xml$`, "i");
  for (const path of parts.keys()) {
    const match = path.match(regex);
    if (match) {
      const num = parseInt(match[1]!, 10);
      if (num > maxNum) maxNum = num;
    }
  }
  return maxNum + 1;
}

// =============================================================================
// Helpers: relationship management
// =============================================================================

/**
 * Get the next available relationship ID.
 */
function getNextRelationshipId(doc: NormalizedDocument): string {
  const rels = doc.xmlParts.documentRels;
  if (!rels) return "rId1";

  let maxId = 0;
  const relsNode = rels.find((n) => getTagName(n) === "Relationships");
  if (relsNode) {
    const relNodes = findChildren(relsNode, "Relationship");
    for (const rel of relNodes) {
      const id = getAttr(rel, "Id");
      if (id) {
        const num = parseInt(id.replace(/^rId/, ""), 10);
        if (!isNaN(num) && num > maxId) maxId = num;
      }
    }
  }
  return `rId${maxId + 1}`;
}

/**
 * Add a Relationship node to document.xml.rels.
 */
function addRelationship(
  doc: NormalizedDocument,
  rId: string,
  target: string,
  type: "header" | "footer",
): void {
  if (!doc.xmlParts.documentRels) {
    // Initialize a minimal Relationships root on-demand so that
    // add_header / add_footer works on documents created via Docxtor.create()
    // (which doesn't include word/_rels/document.xml.rels).
    doc.xmlParts.documentRels = [
      createElement("Relationships", [], {
        "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships",
      }),
    ];
  }

  const relType = type === "header"
    ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

  const relNode = createElement("Relationship", [], {
    "Id": rId,
    "Type": relType,
    "Target": target,
  });

  const relsNode = doc.xmlParts.documentRels.find(
    (n) => getTagName(n) === "Relationships",
  );
  if (relsNode) {
    getChildren(relsNode).push(relNode);
  }
}

/**
 * Find the target filename for a relationship by rId.
 */
function findRelationshipTarget(
  doc: NormalizedDocument,
  rId: string,
): string | undefined {
  const rels = doc.xmlParts.documentRels;
  if (!rels) return undefined;

  const relsNode = rels.find((n) => getTagName(n) === "Relationships");
  if (!relsNode) return undefined;

  const relNodes = findChildren(relsNode, "Relationship");
  for (const rel of relNodes) {
    if (getAttr(rel, "Id") === rId) {
      return getAttr(rel, "Target");
    }
  }
  return undefined;
}

/**
 * Remove a Relationship node from document.xml.rels by rId.
 */
function removeRelationship(doc: NormalizedDocument, rId: string): void {
  const rels = doc.xmlParts.documentRels;
  if (!rels) return;

  const relsNode = rels.find((n) => getTagName(n) === "Relationships");
  if (!relsNode) return;

  const children = getChildren(relsNode);
  const idx = children.findIndex(
    (child) => getTagName(child) === "Relationship" && getAttr(child, "Id") === rId,
  );
  if (idx !== -1) children.splice(idx, 1);
}

// =============================================================================
// Helpers: content types management
// =============================================================================

/**
 * Add an Override entry to [Content_Types].xml.
 */
function addContentTypeOverride(
  doc: NormalizedDocument,
  partName: string,
  type: "header" | "footer",
): void {
  if (!doc.xmlParts.contentTypes) return;

  const contentType = type === "header"
    ? "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    : "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

  const overrideNode = createElement("Override", [], {
    "PartName": `/${partName}`,
    "ContentType": contentType,
  });

  const typesNode = doc.xmlParts.contentTypes.find(
    (n) => getTagName(n) === "Types",
  );
  if (typesNode) {
    getChildren(typesNode).push(overrideNode);
  }
}

/**
 * Remove an Override entry from [Content_Types].xml by part name.
 */
function removeContentTypeOverride(
  doc: NormalizedDocument,
  partName: string,
): void {
  if (!doc.xmlParts.contentTypes) return;

  const typesNode = doc.xmlParts.contentTypes.find(
    (n) => getTagName(n) === "Types",
  );
  if (!typesNode) return;

  const children = getChildren(typesNode);
  // PartName may be stored with or without leading /
  const idx = children.findIndex((child) => {
    if (getTagName(child) !== "Override") return false;
    const pn = getAttr(child, "PartName");
    return pn === `/${partName}` || pn === partName;
  });
  if (idx !== -1) children.splice(idx, 1);
}

// =============================================================================
// Helpers: sectPr management
// =============================================================================

/**
 * Get the last w:sectPr in the document body.
 */
function getLastSectPr(doc: NormalizedDocument): FXPNode | undefined {
  const docNode = doc.xmlParts.document.find(
    (n) => getTagName(n) === "w:document",
  );
  if (!docNode) return undefined;

  const bodyNode = findChild(docNode, "w:body");
  if (!bodyNode) return undefined;

  const bodyChildren = getChildren(bodyNode);
  // sectPr is typically the last child of w:body
  for (let i = bodyChildren.length - 1; i >= 0; i--) {
    if (getTagName(bodyChildren[i]!) === "w:sectPr") {
      return bodyChildren[i]!;
    }
  }
  return undefined;
}

/**
 * Add a header/footer reference to the last sectPr.
 * Creates sectPr if it doesn't exist.
 */
function addSectPrReference(
  doc: NormalizedDocument,
  type: "header" | "footer",
  variant: HeaderFooterVariant,
  rId: string,
): void {
  let sectPr = getLastSectPr(doc);

  if (!sectPr) {
    // Create a sectPr and append to body
    sectPr = createElement("w:sectPr", []);
    const docNode = doc.xmlParts.document.find(
      (n) => getTagName(n) === "w:document",
    );
    if (!docNode) throw new Error("Cannot find w:document");
    const bodyNode = findChild(docNode, "w:body");
    if (!bodyNode) throw new Error("Cannot find w:body");
    getChildren(bodyNode).push(sectPr);
  }

  const refTag = type === "header" ? "w:headerReference" : "w:footerReference";
  const refNode = createElement(refTag, [], {
    "w:type": VARIANT_TYPE_MAP[variant],
    "r:id": rId,
  });

  // Insert header/footer references at the beginning of sectPr children
  // (before pgSz, pgMar, etc.)
  getChildren(sectPr).unshift(refNode);

  // If variant is "first", ensure w:titlePg is present in sectPr
  if (variant === "first") {
    const sectPrChildren = getChildren(sectPr);
    const hasTitlePg = sectPrChildren.some(
      (c) => getTagName(c) === "w:titlePg",
    );
    if (!hasTitlePg) {
      sectPrChildren.push(createElement("w:titlePg", []));
    }
  }

  // If variant is "even", ensure w:evenAndOddHeaders is present in
  // word/settings.xml so Word honours the even-page header/footer reference.
  if (variant === "even") {
    ensureEvenAndOddHeaders(doc);
  }
}

// =============================================================================
// Helpers: settings.xml management
// =============================================================================

/**
 * Ensure `<w:evenAndOddHeaders/>` is present in word/settings.xml.
 *
 * Without this element, Word ignores even-page header/footer references.
 * If settings.xml was not loaded (e.g., created via Docxtor.create()),
 * a minimal settings part is initialized on-demand.
 */
function ensureEvenAndOddHeaders(doc: NormalizedDocument): void {
  if (!doc.xmlParts.settings) {
    // Create a minimal settings.xml FXP structure
    doc.xmlParts.settings = [
      createElement("w:settings", [
        createElement("w:evenAndOddHeaders", []),
      ], {
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      }),
    ];
  }

  const settingsNode = doc.xmlParts.settings.find(
    (n) => getTagName(n) === "w:settings",
  );
  if (!settingsNode) return;

  const children = getChildren(settingsNode);
  const hasFlag = children.some((c) => getTagName(c) === "w:evenAndOddHeaders");
  if (!hasFlag) {
    children.push(createElement("w:evenAndOddHeaders", []));
  }

  ensureSettingsRelationship(doc);
  ensureSettingsContentTypeOverride(doc);
}

function ensureSettingsRelationship(doc: NormalizedDocument): void {
  if (!doc.xmlParts.documentRels) {
    doc.xmlParts.documentRels = [
      createElement("Relationships", [], {
        "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships",
      }),
    ];
  }

  const relsNode = doc.xmlParts.documentRels.find(
    (n) => getTagName(n) === "Relationships",
  );
  if (!relsNode) return;

  const relType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
  const children = getChildren(relsNode);
  const hasSettingsRel = children.some((child) => {
    if (getTagName(child) !== "Relationship") return false;
    return getAttr(child, "Type") === relType && getAttr(child, "Target") === "settings.xml";
  });
  if (hasSettingsRel) return;

  children.push(createElement("Relationship", [], {
    "Id": getNextRelationshipId(doc),
    "Type": relType,
    "Target": "settings.xml",
  }));
}

function ensureSettingsContentTypeOverride(doc: NormalizedDocument): void {
  if (!doc.xmlParts.contentTypes) {
    doc.xmlParts.contentTypes = [
      createElement("Types", [], {
        "xmlns": "http://schemas.openxmlformats.org/package/2006/content-types",
      }),
    ];
  }

  const typesNode = doc.xmlParts.contentTypes.find(
    (n) => getTagName(n) === "Types",
  );
  if (!typesNode) return;

  const partName = "/word/settings.xml";
  const contentType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
  const children = getChildren(typesNode);
  const hasSettingsOverride = children.some((child) => {
    if (getTagName(child) !== "Override") return false;
    return getAttr(child, "PartName") === partName;
  });
  if (hasSettingsOverride) return;

  children.push(createElement("Override", [], {
    "PartName": partName,
    "ContentType": contentType,
  }));
}
