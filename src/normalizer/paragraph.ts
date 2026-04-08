/**
 * Paragraph normalizer: walks a w:p FXP node and produces a
 * NormalizedParagraph with text, run mappings, and xmlRefs.
 *
 * This is the core of the mapping layer. Each RunMapping holds:
 * - The run's text content
 * - Character offset range in the concatenated paragraph text
 * - Direct reference to the w:r and w:t FXP nodes
 * - Captured formatting
 *
 * Patches use these mappings to find the right XML nodes to mutate.
 *
 * Hyperlink support: w:hyperlink elements wrap w:r runs. We collect
 * runs from both direct children and from inside w:hyperlink wrappers,
 * so hyperlink text participates in the paragraph's normalized text
 * and can be targeted by patches.
 */

import type {
  FXPNode,
  NormalizedParagraph,
  RunMapping,
  RunFormatting,
  ListInfo,
  RevisionMark,
} from "../types.js";
import {
  findChild,
  getAttr,
  getChildren,
  getTagName,
  getRunText,
  findTextNode,
  isTag,
  unescapeXmlText,
} from "../xml/helpers.js";

/**
 * Collect all w:r run nodes from a paragraph, including those nested
 * inside w:hyperlink and w:ins elements. Skips w:del entirely.
 * Preserves document order.
 *
 * This ensures:
 * - Inserted text (w:ins) is included in .text — subsequent patches see it
 * - Deleted text (w:del) is excluded from .text — subsequent patches don't see it
 */
function collectRuns(paragraphNode: FXPNode): FXPNode[] {
  const children = getChildren(paragraphNode);
  const runs: FXPNode[] = [];

  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:r") {
      runs.push(child);
    } else if (tag === "w:hyperlink") {
      // Collect runs inside the hyperlink (recurse for nested w:ins within hyperlink)
      collectRunsFromContainer(child, runs);
    } else if (tag === "w:ins") {
      // Inserted runs are "current content" — recurse into them
      collectRunsFromContainer(child, runs);
    }
    // w:del is intentionally skipped — deleted text is excluded from normalized text
  }

  return runs;
}

/**
 * Recursively collect runs from a container element (w:hyperlink, w:ins).
 * Handles nested containers (e.g., w:ins inside w:hyperlink or vice versa).
 */
function collectRunsFromContainer(containerNode: FXPNode, runs: FXPNode[]): void {
  const children = getChildren(containerNode);
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:r") {
      runs.push(child);
    } else if (tag === "w:hyperlink" || tag === "w:ins") {
      collectRunsFromContainer(child, runs);
    }
    // w:del inside containers is also skipped
  }
}

/**
 * Normalize a single w:p FXP node into a NormalizedParagraph.
 *
 * @param paragraphNode - The w:p FXP node (must have "w:p" key)
 * @param id - The stable ID to assign (e.g., "p0", "tbl0.r1.c0.p0")
 */
export function normalizeParagraph(
  paragraphNode: FXPNode,
  id: string,
): NormalizedParagraph {
  const runs = collectRuns(paragraphNode);
  const runMappings: RunMapping[] = [];
  let offset = 0;

  for (const runNode of runs) {
    const text = getRunText(runNode);
    if (text.length === 0) continue; // Skip runs with no text (e.g., break-only runs)

    const textNode = findTextNode(runNode);
    if (!textNode) continue;

    const formatting = extractFormatting(runNode);

    runMappings.push({
      text,
      startOffset: offset,
      endOffset: offset + text.length,
      xmlRef: runNode,
      textXmlRef: textNode,
      formatting,
    });

    offset += text.length;
  }

  const fullText = runMappings.map((r) => r.text).join("");

  // Capture revision marks from w:ins / w:del elements
  const revisions = collectRevisionMarks(paragraphNode, runMappings);

  const result: NormalizedParagraph = {
    kind: "paragraph",
    id,
    text: fullText,
    runs: runMappings,
    xmlRef: paragraphNode,
    style: extractStyle(paragraphNode),
    listInfo: extractListInfo(paragraphNode),
  };

  if (revisions.length > 0) {
    result.revisions = revisions;
  }

  return result;
}

/**
 * Extract formatting properties from a w:r run node's w:rPr element.
 */
function extractFormatting(runNode: FXPNode): RunFormatting {
  const rPr = findChild(runNode, "w:rPr");
  if (!rPr) return {};

  const formatting: RunFormatting = {};
  const children = getChildren(rPr);

  for (const child of children) {
    if (isTag(child, "w:b")) {
      // <w:b/> means bold on; <w:b w:val="0"/> means bold off
      const val = getAttr(child, "w:val");
      formatting.bold = val !== "0" && val !== "false";
    } else if (isTag(child, "w:i")) {
      const val = getAttr(child, "w:val");
      formatting.italic = val !== "0" && val !== "false";
    } else if (isTag(child, "w:u")) {
      formatting.underline = getAttr(child, "w:val") ?? "single";
    } else if (isTag(child, "w:strike")) {
      const val = getAttr(child, "w:val");
      formatting.strikethrough = val !== "0" && val !== "false";
    } else if (isTag(child, "w:sz")) {
      const val = getAttr(child, "w:val");
      if (val) formatting.fontSize = parseInt(val, 10);
    } else if (isTag(child, "w:rFonts")) {
      formatting.fontFamily = getAttr(child, "w:ascii") ?? getAttr(child, "w:hAnsi");
    } else if (isTag(child, "w:color")) {
      formatting.color = getAttr(child, "w:val");
    } else if (isTag(child, "w:highlight")) {
      formatting.highlight = getAttr(child, "w:val");
    }
  }

  return formatting;
}

/**
 * Extract the paragraph style name from w:pPr/w:pStyle.
 */
function extractStyle(paragraphNode: FXPNode): string | undefined {
  const pPr = findChild(paragraphNode, "w:pPr");
  if (!pPr) return undefined;

  const pStyle = findChild(pPr, "w:pStyle");
  if (!pStyle) return undefined;

  return getAttr(pStyle, "w:val");
}

/**
 * Extract list/numbering info from w:pPr/w:numPr.
 */
function extractListInfo(paragraphNode: FXPNode): ListInfo | undefined {
  const pPr = findChild(paragraphNode, "w:pPr");
  if (!pPr) return undefined;

  const numPr = findChild(pPr, "w:numPr");
  if (!numPr) return undefined;

  const numIdNode = findChild(numPr, "w:numId");
  const ilvlNode = findChild(numPr, "w:ilvl");

  const numId = numIdNode ? getAttr(numIdNode, "w:val") : undefined;
  const level = ilvlNode ? getAttr(ilvlNode, "w:val") : undefined;

  if (!numId) return undefined;

  return {
    numId,
    level: level ? parseInt(level, 10) : 0,
  };
}

// =============================================================================
// Revision mark extraction
// =============================================================================

/**
 * Collect revision marks (w:ins, w:del) from a paragraph's XML.
 *
 * - For w:ins runs: startOffset/endOffset point into the .text (runs are included in normalized text)
 * - For w:del runs: startOffset === endOffset (zero-width position) since deleted text is excluded
 *   from .text; the deleted content is stored in the RevisionMark.text field
 */
function collectRevisionMarks(
  paragraphNode: FXPNode,
  runMappings: RunMapping[],
): RevisionMark[] {
  const revisions: RevisionMark[] = [];

  // Build a set of xmlRefs for quick lookup of which runs are "ins" runs
  // We need to track which runs came from w:ins to annotate their offsets
  const insRunRefs = new Set<FXPNode>();
  collectInsRunRefs(paragraphNode, insRunRefs);

  // Track revision marks for w:ins runs
  for (const mapping of runMappings) {
    if (insRunRefs.has(mapping.xmlRef)) {
      // Find the parent w:ins element to get author/date
      const insInfo = findParentInsInfo(paragraphNode, mapping.xmlRef);
      revisions.push({
        type: "insert",
        text: mapping.text,
        startOffset: mapping.startOffset,
        endOffset: mapping.endOffset,
        author: insInfo?.author ?? "",
        date: insInfo?.date ?? "",
      });
    }
  }

  // Track revision marks for w:del (zero-width in normalized text)
  collectDelRevisionMarks(paragraphNode, runMappings, revisions);

  return revisions;
}

/**
 * Recursively collect xmlRef pointers for runs inside w:ins elements.
 */
function collectInsRunRefs(node: FXPNode, refs: Set<FXPNode>): void {
  const children = getChildren(node);
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:ins") {
      // All runs inside this w:ins are "inserted"
      markAllRunsInContainer(child, refs);
    } else if (tag === "w:hyperlink") {
      collectInsRunRefs(child, refs);
    }
    // Don't recurse into w:del — those are excluded
  }
}

function markAllRunsInContainer(node: FXPNode, refs: Set<FXPNode>): void {
  const children = getChildren(node);
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:r") {
      refs.add(child);
    } else if (tag === "w:hyperlink" || tag === "w:ins") {
      markAllRunsInContainer(child, refs);
    }
  }
}

/**
 * Find the w:ins parent element's author and date for a given run node.
 */
function findParentInsInfo(
  paragraphNode: FXPNode,
  runNode: FXPNode,
): { author: string; date: string } | undefined {
  return findInsInfoInChildren(getChildren(paragraphNode), runNode);
}

function findInsInfoInChildren(
  children: FXPNode[],
  runNode: FXPNode,
): { author: string; date: string } | undefined {
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:ins") {
      // Check if this w:ins contains the run
      if (containsRun(child, runNode)) {
        return {
          author: getAttr(child, "w:author") ?? "",
          date: getAttr(child, "w:date") ?? "",
        };
      }
    } else if (tag === "w:hyperlink") {
      const result = findInsInfoInChildren(getChildren(child), runNode);
      if (result) return result;
    }
  }
  return undefined;
}

function containsRun(node: FXPNode, runNode: FXPNode): boolean {
  const children = getChildren(node);
  for (const child of children) {
    if (child === runNode) return true;
    const tag = getTagName(child);
    if (tag && tag !== "w:del") {
      if (containsRun(child, runNode)) return true;
    }
  }
  return false;
}

/**
 * Collect deletion revision marks from w:del elements.
 *
 * For each w:del, we need to figure out its position in the normalized text.
 * Since deleted runs are excluded from .text, the position is zero-width
 * and falls at the boundary where the deletion occurred.
 */
function collectDelRevisionMarks(
  paragraphNode: FXPNode,
  runMappings: RunMapping[],
  revisions: RevisionMark[],
): void {
  const children = getChildren(paragraphNode);
  let currentTextOffset = 0;

  // Walk children in order, tracking position
  for (let i = 0; i < children.length; i++) {
    const child = children[i]!;
    const tag = getTagName(child);

    if (tag === "w:r") {
      // Find this run in mappings to advance the offset
      const mapping = runMappings.find((m) => m.xmlRef === child);
      if (mapping) {
        currentTextOffset = mapping.endOffset;
      }
    } else if (tag === "w:hyperlink" || tag === "w:ins") {
      // Advance through contained runs
      currentTextOffset = advanceThroughContainer(child, runMappings, currentTextOffset);
    } else if (tag === "w:del") {
      // Collect deleted text from this w:del
      const delText = getDeletedText(child);
      if (delText.length > 0) {
        revisions.push({
          type: "delete",
          text: delText,
          startOffset: currentTextOffset,
          endOffset: currentTextOffset,
          author: getAttr(child, "w:author") ?? "",
          date: getAttr(child, "w:date") ?? "",
        });
      }
    }
  }
}

function advanceThroughContainer(
  node: FXPNode,
  runMappings: RunMapping[],
  offset: number,
): number {
  const children = getChildren(node);
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:r") {
      const mapping = runMappings.find((m) => m.xmlRef === child);
      if (mapping) {
        offset = mapping.endOffset;
      }
    } else if (tag === "w:hyperlink" || tag === "w:ins") {
      offset = advanceThroughContainer(child, runMappings, offset);
    }
  }
  return offset;
}

/**
 * Extract text content from runs inside a w:del element.
 * Reads from both w:delText and w:t elements (some docs use w:t inside w:del).
 */
function getDeletedText(delNode: FXPNode): string {
  const children = getChildren(delNode);
  let text = "";
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:r") {
      const runChildren = getChildren(child);
      for (const rc of runChildren) {
        if ("w:delText" in rc) {
          const dtChildren = rc["w:delText"];
          if (Array.isArray(dtChildren)) {
            for (const dtChild of dtChildren as FXPNode[]) {
              if ("#text" in dtChild) {
                text += unescapeXmlText(String(dtChild["#text"]));
              }
            }
          }
        } else if ("w:t" in rc) {
          const tChildren = rc["w:t"];
          if (Array.isArray(tChildren)) {
            for (const tChild of tChildren as FXPNode[]) {
              if ("#text" in tChild) {
                text += unescapeXmlText(String(tChild["#text"]));
              }
            }
          }
        }
      }
    }
  }
  return text;
}
