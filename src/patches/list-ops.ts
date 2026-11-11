/**
 * List operation executors for the 6 list patch types.
 *
 * Each executor mutates the FXP tree directly and updates the
 * NormalizedDocument accordingly.
 */

import type {
  NormalizedDocument,
  FXPNode,
  AddListPatch,
  AddListItemPatch,
  RemoveListItemPatch,
  ConvertToListPatch,
  RemoveListPatch,
  SetListPropertiesPatch,
  ControlInfo,
} from "../types.js";
import { createNumberingDefinition } from "./numbering.js";
import { enrichListType } from "../normalizer/document.js";
import { normalizeParagraph } from "../normalizer/paragraph.js";
import {
  createElement,
  findChild,
  getChildren,
  getTagName,
  setAttr,
  cloneNode,
} from "../xml/helpers.js";
import { buildFieldRuns } from "./field-codes.js";
import { findParagraphById, findElementById } from "./validator.js";
import { createInsElement, createPPrChange, convertRunToDeleted, createDelElement, type RevisionContext } from "./revision.js";
import {
  getBodyChildren,
  insertIntoBody,
  getNextPIdx,
  findParentChildren,
  findNormalizedContainer,
  deriveNewParagraphId,
  reNormalizeParagraph as reNormalize,
} from "./container-utils.js";
import { stampNewParagraphs, findPartRootForId } from "../normalizer/para-id.js";

// =============================================================================
// Helpers
// =============================================================================

/**
 * Build a w:p node with w:numPr for a list item.
 */
function buildListParagraphNode(
  text: string,
  numId: string,
  level: number,
  revCtx?: RevisionContext,
  markers: boolean = true,
): FXPNode {
  const numPr = createElement("w:numPr", [
    createElement("w:ilvl", [], { "w:val": String(level) }),
    createElement("w:numId", [], { "w:val": numId }),
  ]);
  const pPr = createElement("w:pPr", [numPr]);

  const runs = buildFieldRuns(text, undefined, markers);
  let children: FXPNode[];
  if (revCtx) {
    const insElement = createInsElement(revCtx, runs);
    children = [pPr, insElement];
  } else {
    children = [pPr, ...runs];
  }

  return { "w:p": children };
}

// =============================================================================
// Executors
// =============================================================================

export function executeAddList(
  doc: NormalizedDocument,
  patch: AddListPatch,
  revCtx?: RevisionContext,
  markers: boolean = true,
): void {
  const numId = patch.numId ?? createNumberingDefinition(doc, patch.listType);
  const defaultLevel = patch.level ?? 0;
  const nodes: FXPNode[] = [];

  for (const item of patch.items) {
    const text = typeof item === "string" ? item : item.text;
    const level = typeof item === "string" ? defaultLevel : (item.level ?? defaultLevel);
    nodes.push(buildListParagraphNode(text, numId, level, revCtx, markers));
  }

  if (patch.container) {
    // Insert into header/footer
    const map = patch.container.type === "header" ? doc.headers : doc.footers;
    for (const hf of map.values()) {
      if (hf.variant === patch.container.variant) {
        const rootTag = hf.type === "header" ? "w:hdr" : "w:ftr";
        const rootNode = hf.xmlRef.find((n) => getTagName(n) === rootTag);
        if (rootNode) {
          const rootChildren = getChildren(rootNode);
          stampNewParagraphs(doc, nodes, rootNode);
          rootChildren.push(...nodes);
          const prefix = patch.container.type === "header" ? "hdr" : "ftr";
          let nextPIdx = hf.paragraphs.length;
          for (const node of nodes) {
            const normalized = normalizeParagraph(
              node,
              `${prefix}.${patch.container.variant}.p${nextPIdx++}`,
            );
            if (doc.xmlParts.numbering) enrichListType(normalized, doc.xmlParts.numbering);
            hf.paragraphs.push(normalized);
          }
        }
        return;
      }
    }
    return;
  }

  const bodyChildren = getBodyChildren(doc);
  let afterXmlRef: FXPNode | undefined;
  let afterElementId: string | undefined;
  let afterControl: ControlInfo | undefined;
  if (patch.afterId) {
    const target = findElementById(patch.afterId, doc);
    if (target) {
      afterXmlRef = target.xmlRef;
      afterElementId = target.id;
      // Next to an element inside a block-level control, the list is inside it too
      if (target.kind !== "page_break") afterControl = target.control;
    }
  }

  stampNewParagraphs(doc, nodes, findPartRootForId(doc, "body"));
  insertIntoBody(bodyChildren, nodes, afterXmlRef);

  // Normalize and add to doc.body
  let nextPIdx = getNextPIdx(doc);
  const insertIdx = afterElementId !== undefined
    ? doc.body.findIndex((el) => el.id === afterElementId)
    : -1;

  for (let i = 0; i < nodes.length; i++) {
    const id = `p${nextPIdx++}`;
    const normalized = normalizeParagraph(nodes[i]!, id, afterControl);
    if (doc.xmlParts.numbering) enrichListType(normalized, doc.xmlParts.numbering);
    if (insertIdx !== -1) {
      doc.body.splice(insertIdx + 1 + i, 0, normalized);
    } else {
      doc.body.push(normalized);
    }
  }
}

export function executeAddListItem(
  doc: NormalizedDocument,
  patch: AddListItemPatch,
  revCtx?: RevisionContext,
  markers: boolean = true,
): void {
  const targetPara = findParagraphById(patch.target, doc)!;
  const numId = targetPara.listInfo!.numId;
  const level = patch.level ?? targetPara.listInfo!.level;

  const newNode = buildListParagraphNode(patch.text, numId, level, revCtx, markers);
  stampNewParagraphs(doc, [newNode], findPartRootForId(doc, targetPara.id));

  // Find the parent FXP container
  const parentChildren = findParentChildren(doc, targetPara.xmlRef);
  if (!parentChildren) throw new Error("Cannot find parent container for target paragraph");

  const targetIdx = parentChildren.indexOf(targetPara.xmlRef);
  if (targetIdx === -1) throw new Error("Target xmlRef not found in parent");

  if (patch.position === "after") {
    parentChildren.splice(targetIdx + 1, 0, newNode);
  } else {
    parentChildren.splice(targetIdx, 0, newNode);
  }

  // Normalize and insert into the correct normalized container
  const container = findNormalizedContainer(doc, targetPara.id);
  const containerArray = container ? container.array : doc.body;
  const id = deriveNewParagraphId(targetPara.id, containerArray);
  const normalized = normalizeParagraph(newNode, id, targetPara.control);
  if (doc.xmlParts.numbering) enrichListType(normalized, doc.xmlParts.numbering);

  if (container) {
    if (patch.position === "after") {
      container.array.splice(container.index + 1, 0, normalized);
    } else {
      container.array.splice(container.index, 0, normalized);
    }
  } else {
    doc.body.push(normalized);
  }
}

export function executeRemoveListItem(
  doc: NormalizedDocument,
  patch: RemoveListItemPatch,
  revCtx?: RevisionContext,
): void {
  const targetPara = findParagraphById(patch.target, doc)!;

  if (revCtx) {
    // Mark runs as deleted instead of removing
    const children = getChildren(targetPara.xmlRef);
    // Store old pPr for pPrChange
    const oldPPr = findChild(targetPara.xmlRef, "w:pPr");
    if (oldPPr) {
      const pPrChange = createPPrChange(revCtx, oldPPr);
      const pPrChildren = getChildren(oldPPr);
      pPrChildren.push(pPrChange);
    }
    // Wrap runs in w:del — handle both direct w:r and w:r nested inside w:ins
    for (let i = children.length - 1; i >= 0; i--) {
      const child = children[i]!;
      const tag = getTagName(child);
      if (tag === "w:r") {
        // Clone the run and convert w:t → w:delText for valid OOXML
        const cloned = cloneNode(child);
        convertRunToDeleted(cloned);
        const delElement = createDelElement(revCtx, [cloned]);
        children[i] = delElement;
      } else if (tag === "w:ins") {
        // Inserted but not accepted text: removing the w:ins entirely is the
        // correct OOXML semantics (the content was never part of the accepted
        // document, so wrapping it in w:del would corrupt the file).
        children.splice(i, 1);
      }
    }
    reNormalize(doc, targetPara);
    return;
  }

  // Physical removal
  const parentChildren = findParentChildren(doc, targetPara.xmlRef);
  if (parentChildren) {
    const idx = parentChildren.indexOf(targetPara.xmlRef);
    if (idx !== -1) parentChildren.splice(idx, 1);
  }

  // Remove from the correct normalized container (body, hf.paragraphs, or cell.paragraphs)
  const container = findNormalizedContainer(doc, targetPara.id);
  if (container) {
    container.array.splice(container.index, 1);
  }
}

export function executeConvertToList(
  doc: NormalizedDocument,
  patch: ConvertToListPatch,
  revCtx?: RevisionContext,
): void {
  const numId = createNumberingDefinition(doc, patch.listType);
  const level = patch.level ?? 0;

  for (const target of patch.targets) {
    const para = findParagraphById(target, doc)!;
    const pNode = para.xmlRef;

    // Find or create w:pPr
    let pPr = findChild(pNode, "w:pPr");
    const pChildren = getChildren(pNode);

    if (revCtx && pPr) {
      const pPrChange = createPPrChange(revCtx, pPr);
      getChildren(pPr).push(pPrChange);
    }

    if (!pPr) {
      pPr = createElement("w:pPr", []);
      pChildren.unshift(pPr);
    }

    const pPrChildren = getChildren(pPr);

    // Remove existing w:numPr if present
    const existingIdx = pPrChildren.findIndex((c) => getTagName(c) === "w:numPr");
    if (existingIdx !== -1) pPrChildren.splice(existingIdx, 1);

    // Add new w:numPr
    const numPr = createElement("w:numPr", [
      createElement("w:ilvl", [], { "w:val": String(level) }),
      createElement("w:numId", [], { "w:val": numId }),
    ]);
    pPrChildren.unshift(numPr);

    reNormalize(doc, para);
  }
}

export function executeRemoveList(
  doc: NormalizedDocument,
  patch: RemoveListPatch,
  revCtx?: RevisionContext,
): void {
  for (const target of patch.targets) {
    const para = findParagraphById(target, doc)!;
    const pPr = findChild(para.xmlRef, "w:pPr");
    if (!pPr) continue;

    if (revCtx) {
      const pPrChange = createPPrChange(revCtx, pPr);
      getChildren(pPr).push(pPrChange);
    }

    const pPrChildren = getChildren(pPr);
    const numPrIdx = pPrChildren.findIndex((c) => getTagName(c) === "w:numPr");
    if (numPrIdx !== -1) pPrChildren.splice(numPrIdx, 1);

    reNormalize(doc, para);
  }
}

export function executeSetListProperties(
  doc: NormalizedDocument,
  patch: SetListPropertiesPatch,
  revCtx?: RevisionContext,
): void {
  const para = findParagraphById(patch.target, doc)!;
  const pPr = findChild(para.xmlRef, "w:pPr");
  if (!pPr) throw new Error("No w:pPr found on target paragraph");

  if (revCtx) {
    const pPrChange = createPPrChange(revCtx, pPr);
    getChildren(pPr).push(pPrChange);
  }

  const numPr = findChild(pPr, "w:numPr");
  if (!numPr) throw new Error("No w:numPr found on target paragraph");

  if (patch.level !== undefined) {
    const ilvl = findChild(numPr, "w:ilvl");
    if (ilvl) {
      setAttr(ilvl, "w:val", String(patch.level));
    } else {
      getChildren(numPr).unshift(
        createElement("w:ilvl", [], { "w:val": String(patch.level) }),
      );
    }
  }

  if (patch.listType !== undefined) {
    const newNumId = createNumberingDefinition(doc, patch.listType);
    const numIdNode = findChild(numPr, "w:numId");
    if (numIdNode) {
      setAttr(numIdNode, "w:val", newNumId);
    } else {
      getChildren(numPr).push(
        createElement("w:numId", [], { "w:val": newNumId }),
      );
    }
  }

  reNormalize(doc, para);
}
