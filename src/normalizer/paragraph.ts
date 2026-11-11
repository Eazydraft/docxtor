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
  ControlInfo,
  ControlRange,
  NoteReference,
} from "../types.js";
import { runNoteReferences } from "./notes.js";
import {
  findChild,
  getAttr,
  getChildren,
  getTagName,
  getRunText,
  findTextNode,
  isTag,
  runChildChar,
  textElementContent,
  paragraphHasPageBreak,
  runHasPageBreak,
  getRunPageBreakOffset,
  sdtContentChildren,
} from "../xml/helpers.js";
import { controlInfoFor } from "./control.js";

/** One step of a paragraph's content, in document order. */
type ContentStep =
  | { kind: "run"; node: FXPNode; control?: ControlInfo }
  | { kind: "open" | "close"; control: ControlInfo };

/**
 * Collect all w:r run nodes from a paragraph, including those nested
 * inside w:hyperlink, w:ins and inline content controls (w:sdt). Skips
 * w:del entirely. Preserves document order, and marks where each inline
 * control opens and closes.
 *
 * This ensures:
 * - Inserted text (w:ins) is included in .text — subsequent patches see it
 * - Deleted text (w:del) is excluded from .text — subsequent patches don't see it
 * - Text in a content control, a prompt included, is part of .text
 *
 * @param outer - the control wrapped around the whole paragraph, if any
 * @param control - the inline control being walked (none at the top level)
 */
function collectContent(
  children: FXPNode[],
  steps: ContentStep[],
  outer: ControlInfo | undefined,
  control?: ControlInfo,
): void {
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:r") {
      steps.push(control ? { kind: "run", node: child, control } : { kind: "run", node: child });
    } else if (tag === "w:hyperlink" || tag === "w:ins") {
      // Hyperlink runs and inserted runs are "current content" — recurse
      collectContent(getChildren(child), steps, outer, control);
    } else if (tag === "w:sdt") {
      // An inline control sits in the one around it, or else in the control
      // wrapped around the paragraph.
      const inner = controlInfoFor(child, control ?? outer);
      steps.push({ kind: "open", control: inner });
      collectContent(sdtContentChildren(child), steps, outer, inner);
      steps.push({ kind: "close", control: inner });
    }
    // w:del is intentionally skipped — deleted text is excluded from normalized text
  }
}

/**
 * Normalize a single w:p FXP node into a NormalizedParagraph.
 *
 * @param paragraphNode - The w:p FXP node (must have "w:p" key)
 * @param id - The stable ID to assign (e.g., "p0", "tbl0.r1.c0.p0")
 * @param control - The innermost content control wrapped around the whole
 *                  paragraph (a block-level control, or one around its cell)
 */
export function normalizeParagraph(
  paragraphNode: FXPNode,
  id: string,
  control?: ControlInfo,
): NormalizedParagraph {
  const steps: ContentStep[] = [];
  collectContent(getChildren(paragraphNode), steps, control);
  const runMappings: RunMapping[] = [];
  const controls: ControlRange[] = [];
  const open = new Map<ControlInfo, ControlRange>();
  let offset = 0;
  let pageBreakOffset: number | undefined;
  const noteRefs: NoteReference[] = [];

  for (const step of steps) {
    if (step.kind !== "run") {
      if (step.kind === "open") {
        const range: ControlRange = { control: step.control, startOffset: offset, endOffset: offset };
        open.set(step.control, range);
        controls.push(range);
      } else {
        open.get(step.control)!.endOffset = offset;
      }
      continue;
    }

    const runNode = step.node;
    const text = getRunText(runNode);

    // A footnote or endnote reference adds no text; note where it sits. The
    // run that holds one usually has no text at all, so look before skipping.
    for (const ref of runNoteReferences(runNode)) {
      noteRefs.push({ ...ref, offset: offset + ref.offset });
    }

    // Check for page break in this run before skipping empty runs.
    // A page-break-only run has text.length === 0 but still contains a break.
    if (pageBreakOffset === undefined && runHasPageBreak(runNode)) {
      const withinRun = getRunPageBreakOffset(runNode);
      if (withinRun >= 0) {
        pageBreakOffset = offset + withinRun;
      }
    }

    // Skip runs that contribute nothing to the text (empty w:t, page-break-only,
    // field chars, drawings). A run holding only a tab or a line break still
    // yields "\t" or "\n", so every w:tab and w:br is one character in .text — the same text the projection
    // shows, so offsets and occurrences computed from it resolve correctly.
    if (text.length === 0) continue;

    const textNode = findTextNode(runNode);
    const formatting = extractFormatting(runNode);

    const mapping: RunMapping = {
      text,
      startOffset: offset,
      endOffset: offset + text.length,
      xmlRef: runNode,
      formatting,
    };
    if (textNode) mapping.textXmlRef = textNode;
    if (step.control) mapping.control = step.control;
    runMappings.push(mapping);

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

  const paraId = getAttr(paragraphNode, "w14:paraId");
  if (paraId) {
    result.paraId = paraId.toUpperCase();
  }

  if (control) result.control = control;
  if (controls.length > 0) result.controls = controls;

  if (revisions.length > 0) {
    result.revisions = revisions;
  }

  if (noteRefs.length > 0) {
    result.noteRefs = noteRefs;
  }

  // Detect page breaks — check w:pageBreakBefore (always offset 0) and run-level breaks
  if (paragraphHasPageBreak(paragraphNode)) {
    result.hasPageBreak = true;

    // w:pageBreakBefore is a paragraph-level property — always offset 0
    const pPr = findChild(paragraphNode, "w:pPr");
    if (pPr) {
      const pbb = findChild(pPr, "w:pageBreakBefore");
      if (pbb) {
        const val = getAttr(pbb, "w:val");
        if (val !== "0" && val !== "false") {
          result.pageBreakOffset = 0;
          return result;
        }
      }
    }

    // Run-level break — use the offset computed during the run loop
    if (pageBreakOffset !== undefined) {
      result.pageBreakOffset = pageBreakOffset;
    }
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
      const themeColor = getAttr(child, "w:themeColor");
      if (themeColor) formatting.themeColor = themeColor;
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

  // numId "0" means "no numbering" in OOXML (removes inherited list formatting)
  if (numId === "0") return undefined;

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
        ...(insInfo?.id !== undefined ? { id: insInfo.id } : {}),
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
    } else if (tag === "w:hyperlink" || tag === "w:sdt" || tag === "w:sdtContent") {
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
    } else if (tag === "w:hyperlink" || tag === "w:ins" || tag === "w:sdt" || tag === "w:sdtContent") {
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
): { author: string; date: string; id?: string } | undefined {
  return findInsInfoInChildren(getChildren(paragraphNode), runNode);
}

function findInsInfoInChildren(
  children: FXPNode[],
  runNode: FXPNode,
): { author: string; date: string; id?: string } | undefined {
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:ins") {
      // Check if this w:ins contains the run
      if (containsRun(child, runNode)) {
        return {
          author: getAttr(child, "w:author") ?? "",
          date: getAttr(child, "w:date") ?? "",
          id: getAttr(child, "w:id"),
        };
      }
    } else if (tag === "w:hyperlink" || tag === "w:sdt" || tag === "w:sdtContent") {
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
 *
 * This walker recurses into containers (w:hyperlink, w:ins) to find
 * w:del elements nested inside them — e.g., when a deletion occurs inside
 * a hyperlink or inside a previous insertion.
 */
function collectDelRevisionMarks(
  paragraphNode: FXPNode,
  runMappings: RunMapping[],
  revisions: RevisionMark[],
): void {
  // Build a constant-time lookup map: xmlRef → RunMapping
  // Avoids O(n²) linear scans when walking many runs
  const runMap = new Map<FXPNode, RunMapping>();
  for (const mapping of runMappings) {
    runMap.set(mapping.xmlRef, mapping);
  }

  walkForDeletions(getChildren(paragraphNode), runMap, revisions, 0);
}

/**
 * Recursively walk children, advancing the text offset through runs,
 * recording w:del revision marks, and recursing into containers.
 *
 * Returns the text offset after processing all children.
 */
function walkForDeletions(
  children: FXPNode[],
  runMap: Map<FXPNode, RunMapping>,
  revisions: RevisionMark[],
  offset: number,
): number {
  for (const child of children) {
    const tag = getTagName(child);

    if (tag === "w:r") {
      const mapping = runMap.get(child);
      if (mapping) {
        offset = mapping.endOffset;
      }
    } else if (tag === "w:del") {
      // Collect deleted text from this w:del
      const delText = getDeletedText(child);
      if (delText.length > 0) {
        revisions.push({
          type: "delete",
          text: delText,
          startOffset: offset,
          endOffset: offset,
          author: getAttr(child, "w:author") ?? "",
          date: getAttr(child, "w:date") ?? "",
          ...(getAttr(child, "w:id") !== undefined ? { id: getAttr(child, "w:id")! } : {}),
        });
      }
    } else if (tag === "w:hyperlink" || tag === "w:ins") {
      // Recurse into containers — they may contain w:r, w:del, or further nesting
      offset = walkForDeletions(getChildren(child), runMap, revisions, offset);
    } else if (tag === "w:sdt") {
      offset = walkForDeletions(sdtContentChildren(child), runMap, revisions, offset);
    }
  }
  return offset;
}

/**
 * Extract text content from runs inside a w:del element.
 * Reads from both w:delText and w:t elements (some docs use w:t inside w:del);
 * a deleted tab reads as "\t" and a deleted soft line break as "\n", as in
 * live text.
 */
function getDeletedText(delNode: FXPNode): string {
  const children = getChildren(delNode);
  let text = "";
  for (const child of children) {
    const tag = getTagName(child);
    if (tag === "w:r") {
      const runChildren = getChildren(child);
      for (const rc of runChildren) {
        if ("w:delText" in rc) text += textElementContent(rc, "w:delText");
        else if ("w:t" in rc) text += textElementContent(rc);
        else text += runChildChar(rc) ?? "";
      }
    }
  }
  return text;
}
