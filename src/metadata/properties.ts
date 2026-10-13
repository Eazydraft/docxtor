/**
 * Document metadata: docProps/core.xml, docProps/app.xml, docProps/custom.xml.
 *
 * - getCoreProperties / setCoreProperties read and write core.xml fields.
 * - scrubMetadata removes identifying information before a file leaves the
 *   building: author fields in core.xml, Company/Manager/HyperlinkBase in
 *   app.xml, and the whole custom.xml part.
 */

import type { NormalizedDocument, CoreProperties, FXPNode } from "../types.js";
import {
  createElement,
  escapeXmlText,
  getChildren,
  getTagName,
  setAttr,
  unescapeXmlText,
} from "../xml/helpers.js";
import {
  createRelsPart,
  createContentTypesPart,
  ensureRelationship,
  ensureContentTypeOverride,
  removeRelationshipsByType,
  removeContentTypeOverride,
} from "../patches/part-utils.js";

const CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS = "http://purl.org/dc/terms/";
const DCMITYPE_NS = "http://purl.org/dc/dcmitype/";
const XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";

const CORE_REL_TYPE =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const CORE_CONTENT_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";
const CUSTOM_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

/** Core property → qualified element name as Word writes it. */
const CORE_TAGS: Record<keyof CoreProperties, string> = {
  creator: "dc:creator",
  lastModifiedBy: "cp:lastModifiedBy",
  title: "dc:title",
  subject: "dc:subject",
  description: "dc:description",
  keywords: "cp:keywords",
  category: "cp:category",
  contentStatus: "cp:contentStatus",
  revision: "cp:revision",
  created: "dcterms:created",
  modified: "dcterms:modified",
};

/** Fields cleared by scrubMetadata. */
const SCRUB_CORE_FIELDS: (keyof CoreProperties)[] = [
  "creator",
  "lastModifiedBy",
  "description",
  "keywords",
  "category",
  "contentStatus",
];

/** app.xml elements removed by scrubMetadata (matched by local name). */
const SCRUB_APP_FIELDS = new Set(["Company", "Manager", "HyperlinkBase"]);

// =============================================================================
// Helpers
// =============================================================================

function localName(tag: string): string {
  const idx = tag.indexOf(":");
  return idx === -1 ? tag : tag.slice(idx + 1);
}

function findRoot(part: FXPNode[] | undefined, local: string): FXPNode | undefined {
  if (!part) return undefined;
  return part.find((n) => {
    const tag = getTagName(n);
    return tag !== undefined && localName(tag) === local;
  });
}

function findChildByLocalName(parent: FXPNode, local: string): FXPNode | undefined {
  return getChildren(parent).find((c) => {
    const tag = getTagName(c);
    return tag !== undefined && localName(tag) === local;
  });
}

function getElementText(node: FXPNode): string {
  let text = "";
  for (const c of getChildren(node)) {
    if ("#text" in c) text += unescapeXmlText(String(c["#text"]));
  }
  return text;
}

function setElementText(node: FXPNode, text: string): void {
  const tag = getTagName(node)!;
  node[tag] = text.length > 0 ? [{ "#text": escapeXmlText(text) } as unknown as FXPNode] : [];
}

function removeChildByLocalName(parent: FXPNode, local: string): boolean {
  const children = getChildren(parent);
  const idx = children.findIndex((c) => {
    const tag = getTagName(c);
    return tag !== undefined && localName(tag) === local;
  });
  if (idx === -1) return false;
  children.splice(idx, 1);
  return true;
}

// =============================================================================
// core.xml
// =============================================================================

function createCorePart(): FXPNode[] {
  return [
    { "?xml": [{ "#text": "" } as unknown as FXPNode], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } } as FXPNode,
    createElement("cp:coreProperties", [], {
      "xmlns:cp": CP_NS,
      "xmlns:dc": DC_NS,
      "xmlns:dcterms": DCTERMS_NS,
      "xmlns:dcmitype": DCMITYPE_NS,
      "xmlns:xsi": XSI_NS,
    }),
  ];
}

/**
 * Ensure docProps/core.xml exists as a parsed part, wired into the package.
 */
function ensureCorePart(doc: NormalizedDocument): FXPNode {
  if (!doc.xmlParts.core) {
    doc.xmlParts.core = createCorePart();
    if (!doc.xmlParts.rootRels) doc.xmlParts.rootRels = createRelsPart();
    ensureRelationship(doc.xmlParts.rootRels, CORE_REL_TYPE, "docProps/core.xml");
    if (!doc.xmlParts.contentTypes) doc.xmlParts.contentTypes = createContentTypesPart();
    ensureContentTypeOverride(doc.xmlParts.contentTypes, "/docProps/core.xml", CORE_CONTENT_TYPE);
  }
  const root = findRoot(doc.xmlParts.core, "coreProperties");
  if (!root) throw new Error("docProps/core.xml has no coreProperties root");
  return root;
}

/**
 * Read the core document properties. Missing fields are undefined.
 */
export function getCoreProperties(doc: NormalizedDocument): CoreProperties {
  const result: CoreProperties = {};
  const root = findRoot(doc.xmlParts.core, "coreProperties");
  if (!root) return result;

  for (const key of Object.keys(CORE_TAGS) as (keyof CoreProperties)[]) {
    const node = findChildByLocalName(root, localName(CORE_TAGS[key]));
    if (node) result[key] = getElementText(node);
  }
  return result;
}

/**
 * Set core document properties. Only the fields given are touched; an empty
 * string empties the element. Creates docProps/core.xml if the document has none.
 */
export function setCoreProperties(doc: NormalizedDocument, props: Partial<CoreProperties>): void {
  const root = ensureCorePart(doc);

  for (const key of Object.keys(props) as (keyof CoreProperties)[]) {
    const value = props[key];
    if (value === undefined || !(key in CORE_TAGS)) continue;

    const tag = CORE_TAGS[key];
    let node = findChildByLocalName(root, localName(tag));
    if (!node) {
      node = createElement(tag, []);
      getChildren(root).push(node);
    }
    setElementText(node, value);
    if (key === "created" || key === "modified") {
      setAttr(node, "xsi:type", "dcterms:W3CDTF");
    }
  }
}

// =============================================================================
// Scrub
// =============================================================================

/**
 * Strip identifying metadata.
 *
 * - core.xml: removes creator, lastModifiedBy, description, keywords,
 *   category, contentStatus; sets revision to "1". Title, created and
 *   modified are left untouched.
 * - app.xml: removes Company, Manager, HyperlinkBase.
 * - custom.xml: removes the part, its package relationship and content type.
 *
 * @returns zip paths the caller must drop from the archive on save
 */
export function scrubMetadata(doc: NormalizedDocument): { removedParts: string[] } {
  const removedParts: string[] = [];

  const coreRoot = findRoot(doc.xmlParts.core, "coreProperties");
  if (coreRoot) {
    for (const key of SCRUB_CORE_FIELDS) {
      removeChildByLocalName(coreRoot, localName(CORE_TAGS[key]));
    }
    setCoreProperties(doc, { revision: "1" });
  }

  const appRoot = findRoot(doc.xmlParts.app, "Properties");
  if (appRoot) {
    const children = getChildren(appRoot);
    for (let i = children.length - 1; i >= 0; i--) {
      const tag = getTagName(children[i]!);
      if (tag !== undefined && SCRUB_APP_FIELDS.has(localName(tag))) {
        children.splice(i, 1);
      }
    }
  }

  const customTargets = removeRelationshipsByType(doc.xmlParts.rootRels, CUSTOM_REL_TYPE);
  removeContentTypeOverride(doc.xmlParts.contentTypes, "/docProps/custom.xml");
  removedParts.push("docProps/custom.xml");
  for (const target of customTargets) {
    const path = target.replace(/^\//, "");
    if (path && !removedParts.includes(path)) removedParts.push(path);
  }

  return { removedParts };
}
