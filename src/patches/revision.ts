/**
 * Revision tracking helpers for OOXML tracked changes.
 *
 * Provides:
 * - RevisionContext: holds the auto-incrementing ID counter, author, date
 * - scanMaxRevisionId: finds the highest w:id in existing revisions
 * - createElement helpers for w:ins, w:del, w:rPrChange, w:pPrChange
 * - convertRunToDeleted: renames w:t → w:delText inside a run
 */

import type { FXPNode, NormalizedDocument, PatchOptions } from "../types.js";
import {
  getChildren,
  getTagName,
  getAttr,
  setAttr,
  createElement,
  cloneNode,
  escapeXmlText,
  needsSpacePreserve,
} from "../xml/helpers.js";

// =============================================================================
// RevisionContext
// =============================================================================

/**
 * Context object threaded through patch execution when trackRevisions is true.
 * Holds the auto-incrementing revision ID counter, author, and date.
 */
export interface RevisionContext {
  /** Next revision ID to assign. */
  nextId: number;
  /** Author for revision marks. */
  author: string;
  /** ISO 8601 date string for revision marks. */
  date: string;
}

/**
 * Create a RevisionContext from PatchOptions.
 * Scans the document for existing revision IDs to avoid collisions.
 */
export function createRevisionContext(
  doc: NormalizedDocument,
  options: PatchOptions,
): RevisionContext {
  const maxId = scanMaxRevisionId(doc);
  return {
    nextId: maxId + 1,
    author: options.revisionAuthor ?? "docxtor",
    date: options.revisionDate ?? new Date().toISOString(),
  };
}

/**
 * Get the next revision ID and increment the counter.
 */
export function nextRevisionId(ctx: RevisionContext): string {
  const id = String(ctx.nextId);
  ctx.nextId++;
  return id;
}

// =============================================================================
// Scan for existing revision IDs
// =============================================================================

/** Tag names that carry w:id for revision tracking. */
const REVISION_TAGS = new Set([
  "w:ins",
  "w:del",
  "w:rPrChange",
  "w:pPrChange",
  "w:trPrChange",
  "w:tblPrChange",
  "w:sectPrChange",
  "w:cellIns",
  "w:cellDel",
  "w:cellMerge",
]);

/**
 * Scan all XML parts in the document for the highest w:id value
 * on revision-related elements.
 *
 * Returns 0 if no existing revisions are found.
 */
export function scanMaxRevisionId(doc: NormalizedDocument): number {
  let maxId = 0;

  // Scan document.xml
  maxId = Math.max(maxId, scanNodesForMaxId(doc.xmlParts.document));

  // Scan headers
  for (const fxp of doc.xmlParts.headers.values()) {
    maxId = Math.max(maxId, scanNodesForMaxId(fxp));
  }

  // Scan footers
  for (const fxp of doc.xmlParts.footers.values()) {
    maxId = Math.max(maxId, scanNodesForMaxId(fxp));
  }

  // Scan the other revision-bearing parts so new ids never collide
  for (const key of ["footnotes", "endnotes", "comments"] as const) {
    const fxp = doc.xmlParts[key];
    if (fxp) maxId = Math.max(maxId, scanNodesForMaxId(fxp));
  }

  return maxId;
}

/**
 * Recursively scan an FXP node array for revision w:id values.
 */
function scanNodesForMaxId(nodes: FXPNode[]): number {
  let maxId = 0;

  for (const node of nodes) {
    const tag = getTagName(node);
    if (tag && REVISION_TAGS.has(tag)) {
      const idStr = getAttr(node, "w:id");
      if (idStr) {
        const id = parseInt(idStr, 10);
        if (!isNaN(id) && id > maxId) {
          maxId = id;
        }
      }
    }

    // Recurse into children
    if (tag) {
      const children = getChildren(node);
      if (children.length > 0) {
        maxId = Math.max(maxId, scanNodesForMaxId(children));
      }
    }
  }

  return maxId;
}

// =============================================================================
// Revision element creation helpers
// =============================================================================

/**
 * Create a w:ins element wrapping child nodes.
 */
export function createInsElement(
  ctx: RevisionContext,
  children: FXPNode[],
): FXPNode {
  const id = nextRevisionId(ctx);
  return createElement("w:ins", children, {
    "w:id": id,
    "w:author": ctx.author,
    "w:date": ctx.date,
  });
}

/**
 * Create a w:del element wrapping child nodes.
 */
export function createDelElement(
  ctx: RevisionContext,
  children: FXPNode[],
): FXPNode {
  const id = nextRevisionId(ctx);
  return createElement("w:del", children, {
    "w:id": id,
    "w:author": ctx.author,
    "w:date": ctx.date,
  });
}

/**
 * Create a w:rPrChange element containing the old w:rPr snapshot.
 */
export function createRPrChange(
  ctx: RevisionContext,
  oldRPr: FXPNode,
): FXPNode {
  const id = nextRevisionId(ctx);
  return createElement("w:rPrChange", [cloneNode(oldRPr)], {
    "w:id": id,
    "w:author": ctx.author,
    "w:date": ctx.date,
  });
}

/**
 * Create a w:pPrChange element containing the old w:pPr snapshot.
 */
export function createPPrChange(
  ctx: RevisionContext,
  oldPPr: FXPNode,
): FXPNode {
  const id = nextRevisionId(ctx);
  return createElement("w:pPrChange", [cloneNode(oldPPr)], {
    "w:id": id,
    "w:author": ctx.author,
    "w:date": ctx.date,
  });
}

/**
 * Create a w:ins element for table row properties (goes inside w:trPr).
 * This is an empty self-closing element with revision attributes.
 */
export function createTrPrIns(ctx: RevisionContext): FXPNode {
  const id = nextRevisionId(ctx);
  return createElement("w:ins", [], {
    "w:id": id,
    "w:author": ctx.author,
    "w:date": ctx.date,
  });
}

/**
 * Create a w:del element for table row properties (goes inside w:trPr).
 */
export function createTrPrDel(ctx: RevisionContext): FXPNode {
  const id = nextRevisionId(ctx);
  return createElement("w:del", [], {
    "w:id": id,
    "w:author": ctx.author,
    "w:date": ctx.date,
  });
}

// =============================================================================
// Run conversion helpers
// =============================================================================

/**
 * Convert a run's w:t elements to w:delText for use inside w:del.
 *
 * OOXML requires deleted text to use w:delText instead of w:t.
 * This mutates the run node in-place.
 */
export function convertRunToDeleted(runNode: FXPNode): void {
  const children = getChildren(runNode);
  for (let i = 0; i < children.length; i++) {
    const child = children[i]!;
    if ("w:t" in child) {
      // Replace the w:t with a w:delText that has the same children and attributes
      const tChildren = child["w:t"] as FXPNode[];
      const attrs = child[":@"];
      const newNode: FXPNode = { "w:delText": tChildren };
      if (attrs) {
        newNode[":@"] = attrs;
      }
      children[i] = newNode;
    } else if ("w:instrText" in child) {
      // Field instructions inside a deleted run use w:delInstrText
      const newNode: FXPNode = { "w:delInstrText": child["w:instrText"] as FXPNode[] };
      if (child[":@"]) newNode[":@"] = child[":@"];
      children[i] = newNode;
    }
  }
}

/**
 * Create a new w:r run node with the given text and optional formatting,
 * suitable for use inside w:ins. Handles newlines via w:br.
 */
export function createRevisionRun(
  text: string,
  rPrNode: FXPNode | undefined,
): FXPNode {
  const children: FXPNode[] = [];

  // Clone formatting if present
  if (rPrNode) {
    children.push(cloneNode(rPrNode));
  }

  if (text.includes("\n")) {
    // Multi-line: interleave w:t and w:br
    const segments = text.split("\n");
    for (let i = 0; i < segments.length; i++) {
      const seg = segments[i]!;
      if (seg.length > 0 || i === 0) {
        const wt: FXPNode = {
          "w:t": [{ "#text": escapeXmlText(seg) } as unknown as FXPNode],
        };
        if (needsSpacePreserve(seg)) {
          setAttr(wt, "xml:space", "preserve");
        }
        children.push(wt);
      }
      if (i < segments.length - 1) {
        children.push({ "w:br": [] as FXPNode[] });
      }
    }
  } else {
    // Single line
    const wt: FXPNode = {
      "w:t": [{ "#text": escapeXmlText(text) } as unknown as FXPNode],
    };
    if (needsSpacePreserve(text)) {
      setAttr(wt, "xml:space", "preserve");
    }
    children.push(wt);
  }

  return { "w:r": children };
}
