/**
 * add_comment: attach a Word comment to a paragraph or a substring of it.
 *
 * Creates `word/comments.xml` on demand (with its relationship and content
 * type), appends a `w:comment`, and anchors it in the paragraph with
 * `w:commentRangeStart` / `w:commentRangeEnd` plus a `w:commentReference` run.
 */

import type {
  NormalizedDocument,
  NormalizedParagraph,
  AddCommentPatch,
  FXPNode,
} from "../types.js";
import {
  createElement,
  findChild,
  findChildren,
  getAttr,
  getChildren,
  getTagName,
} from "../xml/helpers.js";
import { findOccurrence, resolveRunRanges } from "../mapping/text-to-runs.js";
import { splitRunsForMatch, findRunParentChildren } from "./run-splitter.js";
import { buildFieldRuns } from "./field-codes.js";
import { stripFormattingMarkers } from "./markdown-parser.js";
import {
  createRelsPart,
  createContentTypesPart,
  ensureRelationship,
  ensureContentTypeOverride,
} from "./part-utils.js";

const COMMENTS_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

export interface CommentOptions {
  /** Author to use when the patch does not name one. */
  defaultAuthor: string;
  /** ISO 8601 date for the comment. */
  date: string;
  /** Parse inline formatting markers in `find` and `text`. Default: true. */
  markers?: boolean;
}

/**
 * Ensure word/comments.xml exists (as a parsed part) and is wired into the
 * package. Returns the `w:comments` root node.
 */
export function ensureCommentsPart(doc: NormalizedDocument): FXPNode {
  if (!doc.xmlParts.comments) {
    doc.xmlParts.comments = [
      { "?xml": [{ "#text": "" } as unknown as FXPNode], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } } as FXPNode,
      createElement("w:comments", [], {
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      }),
    ];

    if (!doc.xmlParts.documentRels) doc.xmlParts.documentRels = createRelsPart();
    ensureRelationship(doc.xmlParts.documentRels, COMMENTS_REL_TYPE, "comments.xml");

    if (!doc.xmlParts.contentTypes) doc.xmlParts.contentTypes = createContentTypesPart();
    ensureContentTypeOverride(doc.xmlParts.contentTypes, "/word/comments.xml", COMMENTS_CONTENT_TYPE);
  }

  const root = doc.xmlParts.comments.find((n) => getTagName(n) === "w:comments");
  if (!root) throw new Error("word/comments.xml has no w:comments root");
  return root;
}

/** Next unused comment id (max existing + 1, starting at 0). */
function nextCommentId(commentsRoot: FXPNode): string {
  let max = -1;
  for (const c of findChildren(commentsRoot, "w:comment")) {
    const id = parseInt(getAttr(c, "w:id") ?? "", 10);
    if (!isNaN(id) && id > max) max = id;
  }
  return String(max + 1);
}

/** Initials from an author name: first letter of each word, up to three. */
export function deriveInitials(author: string): string {
  const words = author.split(/\s+/).filter((w) => w.length > 0);
  const initials = words.map((w) => w[0]!.toUpperCase()).join("");
  return initials.slice(0, 3) || "A";
}

/**
 * Build the `w:comment` element with its body paragraphs.
 */
function buildCommentElement(
  id: string,
  author: string,
  initials: string,
  date: string,
  text: string,
  doc: NormalizedDocument,
  markers: boolean,
): FXPNode {
  const hasCommentTextStyle = doc.styles.some((s) => s.id === "CommentText");
  const hasCommentRefStyle = doc.styles.some((s) => s.id === "CommentReference");

  const paragraphs = text.split("\n").map((line, idx) => {
    const children: FXPNode[] = [];
    if (hasCommentTextStyle) {
      children.push(createElement("w:pPr", [createElement("w:pStyle", [], { "w:val": "CommentText" })]));
    }
    if (idx === 0) {
      const refRun: FXPNode[] = [];
      if (hasCommentRefStyle) {
        refRun.push(createElement("w:rPr", [createElement("w:rStyle", [], { "w:val": "CommentReference" })]));
      }
      refRun.push(createElement("w:annotationRef", []));
      children.push({ "w:r": refRun });
    }
    children.push(...buildFieldRuns(line, undefined, markers));
    return { "w:p": children } as FXPNode;
  });

  return createElement("w:comment", paragraphs, {
    "w:id": id,
    "w:author": author,
    "w:date": date,
    "w:initials": initials,
  });
}

/** The `w:r` holding the comment reference mark. */
function buildReferenceRun(id: string, doc: NormalizedDocument): FXPNode {
  const children: FXPNode[] = [];
  if (doc.styles.some((s) => s.id === "CommentReference")) {
    children.push(createElement("w:rPr", [createElement("w:rStyle", [], { "w:val": "CommentReference" })]));
  }
  children.push(createElement("w:commentReference", [], { "w:id": id }));
  return { "w:r": children };
}

/**
 * Execute an add_comment patch on a paragraph.
 */
export function executeAddComment(
  doc: NormalizedDocument,
  paragraph: NormalizedParagraph,
  patch: AddCommentPatch,
  options: CommentOptions,
): void {
  const author = patch.author && patch.author.length > 0 ? patch.author : options.defaultAuthor;
  const initials = patch.initials ?? deriveInitials(author);

  // Resolve the anchor before touching the package, so a failed patch has
  // no side effects (no empty comments part, relationship or content type).
  let matchedRuns: FXPNode[] | undefined;
  if (patch.find !== undefined && patch.find.length > 0) {
    const occurrence = patch.occurrence ?? 1;
    const plainFind = stripFormattingMarkers(patch.find, options.markers ?? true);
    const matchStart = findOccurrence(paragraph.text, plainFind, occurrence);
    if (matchStart === -1) {
      throw new Error(
        `Text "${plainFind}" not found in paragraph "${paragraph.id}" (occurrence ${occurrence})`,
      );
    }
    const ranges = resolveRunRanges(paragraph.runs, matchStart, matchStart + plainFind.length);
    if (ranges.length === 0) {
      throw new Error(`No runs found for "${plainFind}" in paragraph "${paragraph.id}"`);
    }
    matchedRuns = splitRunsForMatch(paragraph.xmlRef, ranges);
    if (matchedRuns.length === 0) {
      throw new Error(`Failed to isolate matched runs in paragraph "${paragraph.id}"`);
    }
  }

  const commentsRoot = ensureCommentsPart(doc);
  const id = nextCommentId(commentsRoot);

  const rangeStart = createElement("w:commentRangeStart", [], { "w:id": id });
  const rangeEnd = createElement("w:commentRangeEnd", [], { "w:id": id });
  const refRun = buildReferenceRun(id, doc);

  if (matchedRuns) {
    const firstRun = matchedRuns[0]!;
    const lastRun = matchedRuns[matchedRuns.length - 1]!;

    const firstParent = findRunParentChildren(paragraph.xmlRef, firstRun);
    if (!firstParent) throw new Error("Matched run not found in paragraph");
    firstParent.splice(firstParent.indexOf(firstRun), 0, rangeStart);

    const lastParent = findRunParentChildren(paragraph.xmlRef, lastRun);
    if (!lastParent) throw new Error("Matched run not found in paragraph");
    lastParent.splice(lastParent.indexOf(lastRun) + 1, 0, rangeEnd, refRun);
  } else {
    const pChildren = getChildren(paragraph.xmlRef);
    const pPr = findChild(paragraph.xmlRef, "w:pPr");
    const startIdx = pPr ? pChildren.indexOf(pPr) + 1 : 0;
    pChildren.splice(startIdx, 0, rangeStart);
    pChildren.push(rangeEnd, refRun);
  }

  getChildren(commentsRoot).push(
    buildCommentElement(id, author, initials, options.date, patch.text, doc, options.markers ?? true),
  );
}
