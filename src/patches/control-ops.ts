/**
 * Edits that touch a content control (`w:sdt`).
 *
 * A text patch whose match lies inside a control is planned here before it
 * runs:
 *
 * - on the prompt of an empty control it fills the control: every prompt run
 *   goes, the new text arrives as one run, and `w:showingPlcHdr` is removed;
 * - on a filled rich-text or plain-text control it is an ordinary text edit;
 * - on a drop-down list the resulting text must be one of the list items;
 * - on a date the text is written as given and `w:fullDate` follows when it
 *   parses as a date;
 * - on a checkbox the glyph and `w14:checked` flip together;
 * - a locked control, a data-bound control, a picture control and a match
 *   that crosses the edge of a control are refused with a reason.
 */

import type {
  NormalizedParagraph,
  NormalizedTable,
  BodyElement,
  ControlInfo,
  RunMapping,
  ReplaceTextPatch,
  InsertTextPatch,
  DeleteTextPatch,
  ReplaceCellTextPatch,
  SetCheckboxPatch,
  FXPNode,
} from "../types.js";
import { findOccurrence, resolveRunRanges } from "../mapping/text-to-runs.js";
import { stripFormattingMarkers, parseFormattedText } from "./markdown-parser.js";
import { createFormattedRuns, findInsertTargetRun } from "./text-ops.js";
import { findRunParentChildren } from "./run-splitter.js";
import { markRunsDeleted } from "./paragraph-ops.js";
import { createInsElement, type RevisionContext } from "./revision.js";
import { controlChain, controlInfoFor, readControlInfo } from "../normalizer/control.js";
import {
  cloneNode,
  createElement,
  findChild,
  getAttr,
  getChildren,
  getTagName,
  removeAttr,
  setAttr,
  isAnchorRun,
  RUN_CONTAINER_TAGS,
} from "../xml/helpers.js";

export type TextPatch = ReplaceTextPatch | InsertTextPatch | DeleteTextPatch | ReplaceCellTextPatch;

/** Why an edit inside a content control cannot be made. */
export class ControlEditError extends Error {}

/** What a text patch does to the control its match lies in. */
export type ControlEditPlan =
  /** Replace the prompt of an empty control with `text`. */
  | { mode: "fill"; control: ControlInfo; scope: ControlScope; text: string; raw: string }
  /** Put a checkbox control in the given state. */
  | { mode: "checkbox"; control: ControlInfo; scope: ControlScope; checked: boolean }
  /** An ordinary text edit; afterwards the control's stored value follows `text`. */
  | { mode: "edit"; control: ControlInfo; text: string };

/** Whether the control sits inside the paragraph or wraps it. */
type ControlScope = "inline" | "block";

// =============================================================================
// Describing controls
// =============================================================================

const KIND_WORDS: Record<ControlInfo["kind"], string> = {
  richText: "",
  text: "",
  date: "date ",
  dropDownList: "dropdown ",
  comboBox: "combo box ",
  checkbox: "checkbox ",
  picture: "picture ",
  docPart: "",
  other: "",
};

/** `dropdown content control "Type of contract"`, for error messages. */
function describe(control: ControlInfo): string {
  const name = control.alias ?? control.tag;
  return `${KIND_WORDS[control.kind]}content control${name ? ` "${name}"` : ""}`;
}

function quote(items: string[]): string {
  return items.map((o) => `"${o}"`).join(", ");
}

// =============================================================================
// Locks and bindings
// =============================================================================
//
// Every patch that changes the document goes through one of two questions,
// so no patch type needs its own idea of where a control might sit:
//
// 1. Is the place being changed inside a control that is closed to edits
//    (content-locked or data-bound), at any level of nesting? Answered from
//    the chain of controls around the paragraph, row or table: assertOpen.
// 2. Does the thing being removed hold a locked control, at any depth?
//    Answered from the XML of the node itself: assertNoLockedInside. It
//    reads the XML rather than the normalized view, so a control around a
//    nested row or in a nested table cannot be overlooked.

const BOUND_REASON =
  "bound to a custom XML part; Word overwrites its content from that part when the file is opened, so an edit here would be lost";

/** True when the content of `c` itself may not be edited. */
function contentLocked(c: ControlInfo): boolean {
  return c.locked === "content" || c.locked === "both";
}

/**
 * Refuse a change to something inside `control` when it, or a control
 * around it, is content-locked or data-bound.
 *
 * @param where - what is being changed, for the message ("p7", "tbl0 row 1")
 * @param subject - how to name it ("the text", "the paragraph", "the table")
 * @param stopAt - controls that are outside the change and need no check
 */
function assertOpen(
  where: string,
  subject: string,
  control: ControlInfo | undefined,
  stopAt?: Set<ControlInfo>,
): void {
  for (const c of controlChain(control)) {
    if (stopAt?.has(c)) break;
    if (contentLocked(c)) {
      throw new ControlEditError(
        `${where}: ${subject} is in a locked ${describe(c)}; its content cannot be edited`,
      );
    }
    if (c.bound) {
      throw new ControlEditError(`${where}: ${subject} is in a ${describe(c)} ${BOUND_REASON}`);
    }
  }
}

/** The first locked content control at any depth under `node`, if any. */
function findLockedInside(node: FXPNode): ControlInfo | undefined {
  for (const child of getChildren(node)) {
    const tag = getTagName(child);
    if (!tag) continue;
    if (tag === "w:sdt") {
      const info = readControlInfo(child);
      if (info.locked) return info;
    }
    const found = findLockedInside(child);
    if (found) return found;
  }
  return undefined;
}

/**
 * Refuse to remove `node` when it holds a locked control at any depth. A
 * control locked against deletion would go with it; one locked against
 * editing would lose its content, or have it marked deleted when tracked.
 */
function assertNoLockedInside(where: string, subject: string, node: FXPNode): void {
  const locked = findLockedInside(node);
  if (!locked) return;
  if (locked.locked === "content") {
    throw new ControlEditError(`${where}: ${subject} holds a locked ${describe(locked)}; it cannot be deleted`);
  }
  throw new ControlEditError(
    `${where}: ${subject} holds a ${describe(locked)} that is locked against deletion`,
  );
}

/**
 * Check a patch that changes the runs of a paragraph without changing its
 * text (formatting, a character style): refused inside a closed control.
 * `runs` are the runs it touches; omit for the whole paragraph.
 */
export function checkRunsEditable(para: NormalizedParagraph, runs?: RunMapping[]): void {
  assertOpen(para.id, "the paragraph", para.control);
  const outer = new Set(controlChain(para.control));
  const seen = new Set<ControlInfo>();
  for (const run of runs ?? para.runs) {
    if (!run.control || seen.has(run.control)) continue;
    seen.add(run.control);
    assertOpen(para.id, `"${run.text}"`, run.control, outer);
  }
}

/**
 * Check a patch that changes a paragraph's own properties (its style, its
 * list membership or level): refused when the paragraph is inside a closed
 * control.
 */
export function checkParagraphEditable(para: NormalizedParagraph): void {
  assertOpen(para.id, "the paragraph", para.control);
}

/**
 * Check that new block content (a paragraph, a list item, a table) may be
 * put next to `target`. It lands in the control that wraps `target`, so
 * that control must be an open, filled rich-text one.
 */
export function checkBlockInsert(target: BodyElement): void {
  const control = target.control;
  if (!control) return;
  const what = target.kind === "table" ? "the table" : target.kind === "page_break" ? "the page break" : "the paragraph";
  assertOpen(target.id, what, control);
  if (control.placeholder) {
    throw new ControlEditError(
      `${target.id}: ${what} is the placeholder of an empty ${describe(control)}; ` +
        `fill the control with replace_text before adding content to it`,
    );
  }
  if (control.kind !== "richText" && control.kind !== "docPart" && control.kind !== "other") {
    throw new ControlEditError(
      `${target.id}: ${what} is in a ${describe(control)}, which holds a single value; ` +
        `content cannot be added to it`,
    );
  }
}

/**
 * Check that a paragraph may be deleted: not when it is inside a closed
 * control, and not when it holds a locked control at any depth.
 */
export function checkParagraphDelete(para: NormalizedParagraph): void {
  assertOpen(para.id, "the paragraph", para.control);
  assertNoLockedInside(para.id, "the paragraph", para.xmlRef);
}

/**
 * Check that the rows of a table may change (a row added or deleted): not
 * when the table is inside a closed control.
 */
export function checkTableEditable(table: NormalizedTable): void {
  assertOpen(table.id, "the table", table.control);
}

/**
 * Check that a row may be added after `afterRow` (-1: before the first row).
 * The new row is put next to the row it copies, so when a control wraps
 * that row the new one lands inside it: that control must be open too.
 */
export function checkRowInsert(table: NormalizedTable, afterRow: number): void {
  checkTableEditable(table);
  const source = table.rows[afterRow >= 0 ? afterRow : 0];
  if (!source) return;
  assertOpen(`${table.id} row ${source.index}`, "the row", source.control, new Set(controlChain(table.control)));
}

/**
 * Check that a table row may be deleted: not when the table or the row is
 * inside a closed control, not when a locked control wraps the row, and not
 * when the row holds a locked control at any depth (around a cell, in a
 * cell, around a nested row, in a nested table).
 */
export function checkRowDelete(table: NormalizedTable, rowIndex: number): void {
  checkTableEditable(table);
  const row = table.rows[rowIndex];
  if (!row) return;
  const where = `${table.id} row ${rowIndex}`;

  // Controls around the table stay. One wrapped around the row itself would
  // be left empty, so it counts as held by the row.
  const outer = new Set(controlChain(table.control));
  for (const c of controlChain(row.control)) {
    if (outer.has(c)) break;
    if (c.locked === "control") {
      throw new ControlEditError(`${where}: the row holds a ${describe(c)} that is locked against deletion`);
    }
  }
  assertOpen(where, "the row", row.control, outer);
  assertNoLockedInside(where, "the row", row.xmlRef);
}

// =============================================================================
// Planning a text patch
// =============================================================================

/**
 * Work out what a text patch does to a content control.
 *
 * Returns undefined when the match is not in a control (or the text is not
 * found: the patch itself reports that). Throws ControlEditError when the
 * edit cannot be made.
 */
export function planControlEdit(
  para: NormalizedParagraph,
  patch: TextPatch,
  markers: boolean = true,
): ControlEditPlan | undefined {
  if (!para.control && !para.controls) return undefined;

  const occurrence = patch.occurrence ?? 1;
  const needle = stripFormattingMarkers(patch.type === "insert_text" ? patch.anchor : patch.find, markers);
  const matchStart = findOccurrence(para.text, needle, occurrence);
  if (matchStart === -1) return undefined;
  const matchEnd = matchStart + needle.length;
  const matchRuns = resolveRunRanges(para.runs, matchStart, matchEnd).map((r) => r.run);

  // What the patch writes, and the range of paragraph text it replaces.
  const raw = patch.type === "insert_text" ? patch.text : patch.type === "delete_text" ? "" : patch.replace;
  const written = stripFormattingMarkers(raw, markers);
  let editStart = matchStart;
  let editEnd = matchEnd;
  let runs = matchRuns;

  if (patch.type === "insert_text") {
    // An anchor on a prompt fills the control. Otherwise the text goes next
    // to the anchor, in whichever run the insertion picks.
    const anchorControl = soleControl(matchRuns) ?? para.control;
    if (!(matchRuns.length > 0 && sameControl(matchRuns) && anchorControl?.placeholder)) {
      const offset = patch.position === "before" ? matchStart : matchEnd;
      const target = findInsertTargetRun(para, offset, patch.position);
      if (!target) return undefined;
      runs = [target];
      editStart = offset;
      editEnd = offset;
    }
  }
  if (runs.length === 0) return undefined;

  if (!sameControl(runs)) {
    const edge = crossedControl(runs);
    throw new ControlEditError(
      `${para.id}: "${needle}" crosses the edge of a ${describe(edge)}; ` +
        `edit the text inside the control and the text around it separately`,
    );
  }

  const inline = runs[0]!.control;
  const control = inline ?? para.control;
  if (!control) return undefined;
  const scope: ControlScope = inline ? "inline" : "block";

  assertOpen(para.id, "the text", control);
  if (control.kind === "picture") {
    throw new ControlEditError(`${para.id}: a ${describe(control)} cannot be edited`);
  }

  // A checkbox wrapped around a paragraph or cell is its glyph's run, not the
  // paragraph's other text
  const glyph = !inline && control.kind === "checkbox" ? blockGlyphRun(para, control) : undefined;
  if (glyph && (editStart < glyph.startOffset || editEnd > glyph.endOffset)) {
    throw new ControlEditError(
      `${para.id}: the text is in a ${describe(control)} beside its glyph; only the glyph can be changed ` +
        `(replace it, or use set_checkbox)`,
    );
  }
  const range = inline
    ? para.controls!.find((r) => r.control === inline)!
    : glyph ?? { startOffset: 0, endOffset: para.text.length };
  const current = para.text.slice(range.startOffset, range.endOffset);
  const single = scope === "inline" || countParagraphs(control) === 1;

  let text: string;
  if (control.placeholder) {
    if (patch.type === "delete_text" || written.length === 0) {
      throw new ControlEditError(
        `${para.id}: "${current}" is the placeholder of an empty ${describe(control)}; there is nothing to delete`,
      );
    }
    if (!single) {
      throw new ControlEditError(
        `${para.id}: the placeholder of this ${describe(control)} spans several paragraphs; filling it is not supported`,
      );
    }
    // A prompt is not document text: the control gets the new text, whole.
    // Everything it shows is replaced, so nothing in there may be locked.
    const shown = scope === "block" ? para.xmlRef : findChild(control.xmlRef, "w:sdtContent");
    if (shown) assertNoLockedInside(para.id, "the placeholder", shown);
    text = written;
  } else {
    text =
      current.slice(0, editStart - range.startOffset) +
      written +
      current.slice(editEnd - range.startOffset);
    if (text.length === 0 && single && control.kind !== "docPart" && control.kind !== "other") {
      throw new ControlEditError(
        `${para.id}: clearing a content control is not supported (the edit would leave the ${describe(control)} empty)`,
      );
    }
  }

  if (control.kind === "checkbox") {
    const states = checkboxStates(control);
    if (text !== states.checked.glyph && text !== states.unchecked.glyph) {
      throw new ControlEditError(
        `${para.id}: a checkbox control holds "${states.unchecked.glyph}" (unchecked) or ` +
          `"${states.checked.glyph}" (checked); replace one with the other, or use set_checkbox`,
      );
    }
    return { mode: "checkbox", control, scope, checked: text === states.checked.glyph };
  }

  if (control.kind === "dropDownList" && !(control.options ?? []).includes(text)) {
    throw new ControlEditError(
      `${para.id}: "${text}" is not an option of the ${describe(control)}; choose one of: ${quote(control.options ?? [])}`,
    );
  }

  return control.placeholder
    ? { mode: "fill", control, scope, text, raw }
    : { mode: "edit", control, text };
}

/**
 * The innermost control whose edge a match crosses: one that holds some of
 * the matched runs but not all of them.
 */
function crossedControl(runs: RunMapping[]): ControlInfo {
  for (const run of runs) {
    for (const c of controlChain(run.control)) {
      if (!runs.every((r) => controlChain(r.control).includes(c))) return c;
    }
  }
  return runs.find((r) => r.control)!.control!;
}

/** True when every run sits in the same inline control (or all in none). */
function sameControl(runs: RunMapping[]): boolean {
  return runs.every((r) => r.control === runs[0]!.control);
}

/** The inline control shared by all runs, if there is one. */
function soleControl(runs: RunMapping[]): ControlInfo | undefined {
  return runs.length > 0 && sameControl(runs) ? runs[0]!.control : undefined;
}

/** Number of paragraphs a control holds (one, for an inline control's host). */
function countParagraphs(control: ControlInfo): number {
  let n = 0;
  const visit = (nodes: FXPNode[]): void => {
    for (const node of nodes) {
      const tag = getTagName(node);
      if (!tag) continue;
      if (tag === "w:p") n++;
      else visit(getChildren(node));
    }
  };
  const content = findChild(control.xmlRef, "w:sdtContent");
  if (content) visit(getChildren(content));
  return n;
}

// =============================================================================
// Carrying out a plan
// =============================================================================

/** Children of a control's content that a fill replaces: runs and what holds them. */
function isRunBearing(node: FXPNode): boolean {
  const tag = getTagName(node);
  return tag !== undefined && (RUN_CONTAINER_TAGS.has(tag) || tag === "w:sdt");
}

/**
 * The array holding the control's runs in this paragraph: the control's
 * `w:sdtContent` for an inline control, the paragraph itself for a control
 * wrapped around it.
 */
function contentArray(para: NormalizedParagraph, control: ControlInfo, scope: ControlScope): FXPNode[] {
  if (scope === "block") return getChildren(para.xmlRef);
  let content = findChild(control.xmlRef, "w:sdtContent");
  if (!content) {
    content = createElement("w:sdtContent", []);
    getChildren(control.xmlRef).push(content);
  }
  return getChildren(content);
}

/**
 * Replace everything the control shows in this paragraph by `runs`.
 *
 * The new runs go right after where the last old run was. With `trackOld`
 * the old runs are marked deleted in place and the new ones follow them, so
 * accepting gives exactly the untracked result, and markers between the old
 * runs (a bookmark, a comment's range) keep their order around the new text. A run that only anchors a
 * comment, a footnote or an endnote is not content and is kept: one that
 * stood before the old text stays before the new, any other follows it (a
 * comment's anchor after the end of its range). The new runs are wrapped in
 * `w:ins` whenever a revision context is given.
 */
function replaceContent(
  content: FXPNode[],
  runs: FXPNode[],
  revCtx: RevisionContext | undefined,
  trackOld: boolean,
): void {
  const firstText = content.findIndex((c) => isRunBearing(c) && !isAnchorRun(c));
  const anchors: { node: FXPNode; leading: boolean }[] = [];
  for (let i = content.length - 1; i >= 0; i--) {
    if (isAnchorRun(content[i]!)) {
      anchors.unshift({ node: content.splice(i, 1)[0]!, leading: firstText !== -1 && i < firstText });
    }
  }

  // Where the new content goes: right after the last old run, tracked or
  // not; with nothing shown, inside a comment's range.
  const marker: FXPNode = { "docxtor:here": [] };
  let lastRun = -1;
  content.forEach((c, i) => {
    if (isRunBearing(c)) lastRun = i;
  });
  let at: number;
  if (lastRun === -1) {
    const rangeEnd = content.findIndex((c) => getTagName(c) === "w:commentRangeEnd");
    at = rangeEnd === -1 ? content.length : rangeEnd;
  } else {
    at = lastRun + 1;
  }
  content.splice(at, 0, marker);

  if (revCtx && trackOld) {
    markRunsDeleted(content, revCtx);
  } else {
    for (let i = content.length - 1; i >= 0; i--) {
      if (isRunBearing(content[i]!)) content.splice(i, 1);
    }
  }
  const added = revCtx ? [createInsElement(revCtx, runs)] : runs;
  content.splice(content.indexOf(marker), 1, ...added);

  // Put the anchors back
  const start = content.indexOf(added[0]!);
  const leading = anchors.filter((a) => a.leading).map((a) => a.node);
  content.splice(start, 0, ...leading);
  let after = start + leading.length + added.length;
  for (const { node, leading: isLeading } of anchors) {
    if (isLeading) continue;
    const id = commentIdOf(node);
    const end = id === undefined ? -1 : content.findIndex((c) => getTagName(c) === "w:commentRangeEnd" && getAttr(c, "w:id") === id);
    if (end !== -1) {
      content.splice(end + 1, 0, node);
      if (end < after) after++;
    } else {
      content.splice(after, 0, node);
      after++;
    }
  }
}

function commentIdOf(anchor: FXPNode): string | undefined {
  const ref = getChildren(anchor).find((c) => getTagName(c) === "w:commentReference");
  return ref ? getAttr(ref, "w:id") : undefined;
}

/** The run properties a control gives to text typed into it (`w:sdtPr/w:rPr`). */
function controlRunProperties(control: ControlInfo): FXPNode | undefined {
  const sdtPr = findChild(control.xmlRef, "w:sdtPr");
  return sdtPr ? findChild(sdtPr, "w:rPr") : undefined;
}

/** The control stops showing its prompt. `w:placeholder` stays: Word restores the prompt from it. */
function clearPlaceholderFlag(control: ControlInfo): void {
  const sdtPr = findChild(control.xmlRef, "w:sdtPr");
  if (!sdtPr) return;
  const children = getChildren(sdtPr);
  const idx = children.findIndex((c) => getTagName(c) === "w:showingPlcHdr");
  if (idx !== -1) children.splice(idx, 1);
}

/** Re-read a control's properties into the shared ControlInfo after an edit. */
function refresh(control: ControlInfo): void {
  controlInfoFor(control.xmlRef, control.parent);
}

/**
 * Carry out a "fill" or "checkbox" plan. An "edit" plan is an ordinary text
 * patch followed by {@link syncControlValue}.
 */
export function applyControlPlan(
  para: NormalizedParagraph,
  plan: Exclude<ControlEditPlan, { mode: "edit" }>,
  revCtx: RevisionContext | undefined,
  markers: boolean,
): void {
  if (plan.mode === "checkbox") {
    writeCheckbox(para, plan.control, plan.scope, plan.checked, revCtx);
    return;
  }

  // The new run takes the control's own run properties when it has any,
  // otherwise the paragraph's defaults; never the prompt's grey style.
  const segments = parseFormattedText(plan.raw, markers);
  const runs = createFormattedRuns(segments, controlRunProperties(plan.control));
  // With tracking the fill is an insertion and nothing else: a prompt was
  // never document text, so there is no deletion to record.
  replaceContent(contentArray(para, plan.control, plan.scope), runs, revCtx, false);
  clearPlaceholderFlag(plan.control);
  syncControlValue(plan.control, plan.text);
}

/**
 * Bring the value a typed control stores in `w:sdtPr` in line with the text
 * it now shows: `w:lastValue` of a drop-down list or combo box, `w:fullDate`
 * of a date picker.
 */
export function syncControlValue(control: ControlInfo, text: string): void {
  const sdtPr = findChild(control.xmlRef, "w:sdtPr");
  if (sdtPr) {
    if (control.kind === "dropDownList" || control.kind === "comboBox") {
      const list = findChild(sdtPr, control.kind === "dropDownList" ? "w:dropDownList" : "w:comboBox");
      if (list) {
        const item = getChildren(list).find(
          (c) =>
            getTagName(c) === "w:listItem" &&
            (getAttr(c, "w:displayText") ?? getAttr(c, "w:value")) === text,
        );
        // A combo box takes free text; its last value is then the text itself.
        setAttr(list, "w:lastValue", item ? (getAttr(item, "w:value") ?? text) : text);
      }
    } else if (control.kind === "date") {
      const date = findChild(sdtPr, "w:date");
      if (date) {
        const full = parseDate(text);
        if (full) setAttr(date, "w:fullDate", full);
        else removeAttr(date, "w:fullDate");
      }
    }
  }
  refresh(control);
}

// =============================================================================
// Dates
// =============================================================================

const MONTHS = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];

function monthNumber(name: string): number {
  const n = name.toLowerCase().replace(/\.$/, "");
  if (n.length < 3) return 0;
  const idx = MONTHS.findIndex((m) => m === n || (n.length === 3 && m.startsWith(n)) || (n === "sept" && m === "september"));
  return idx + 1;
}

/**
 * The `w:fullDate` value for a date written as text, or undefined when the
 * text is not an unambiguous date. Reads `2026-11-11`, `11 November 2026`
 * and `November 11, 2026` (month names in English, in full or abbreviated).
 * A numeric date with slashes is ambiguous between day-first and
 * month-first and is left alone.
 */
export function parseDate(text: string): string | undefined {
  const t = text.trim();
  let y = 0;
  let m = 0;
  let d = 0;
  let match: RegExpExecArray | null;
  if ((match = /^(\d{4})-(\d{2})-(\d{2})(?:T[\d:.]+(?:Z|[+-]\d{2}:?\d{2})?)?$/.exec(t))) {
    [y, m, d] = [Number(match[1]), Number(match[2]), Number(match[3])];
  } else if ((match = /^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+\.?),?\s+(\d{4})$/.exec(t))) {
    [d, m, y] = [Number(match[1]), monthNumber(match[2]!), Number(match[3])];
  } else if ((match = /^([A-Za-z]+\.?)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})$/.exec(t))) {
    [m, d, y] = [monthNumber(match[1]!), Number(match[2]), Number(match[3])];
  } else {
    return undefined;
  }
  if (m < 1 || m > 12 || d < 1) return undefined;
  const daysInMonth = new Date(Date.UTC(y, m, 0)).getUTCDate();
  if (d > daysInMonth) return undefined;
  const pad = (n: number, w: number): string => String(n).padStart(w, "0");
  return `${pad(y, 4)}-${pad(m, 2)}-${pad(d, 2)}T00:00:00Z`;
}

// =============================================================================
// Checkboxes
// =============================================================================

interface CheckboxState {
  glyph: string;
  font?: string;
}

/**
 * The glyph and font a checkbox control shows in each state, from
 * `w14:checkedState` / `w14:uncheckedState` (Word's defaults when absent:
 * ☒ and ☐ in MS Gothic).
 */
export function checkboxStates(control: ControlInfo): { checked: CheckboxState; unchecked: CheckboxState } {
  const states = {
    checked: { glyph: "☒", font: "MS Gothic" } as CheckboxState,
    unchecked: { glyph: "☐", font: "MS Gothic" } as CheckboxState,
  };
  const sdtPr = findChild(control.xmlRef, "w:sdtPr");
  const box = sdtPr ? findChild(sdtPr, "w14:checkbox") : undefined;
  if (!box) return states;
  for (const [key, tag] of [["checked", "w14:checkedState"], ["unchecked", "w14:uncheckedState"]] as const) {
    const el = findChild(box, tag);
    if (!el) continue;
    const code = parseInt(getAttr(el, "w14:val") ?? "", 16);
    if (!Number.isNaN(code) && code > 0) states[key].glyph = String.fromCodePoint(code);
    const font = getAttr(el, "w14:font");
    if (font !== undefined) states[key].font = font;
  }
  return states;
}

/** Put a checkbox control in a state: glyph, glyph font and `w14:checked` together. */
function writeCheckbox(
  para: NormalizedParagraph,
  control: ControlInfo,
  scope: ControlScope,
  checked: boolean,
  revCtx: RevisionContext | undefined,
): void {
  const state = checkboxStates(control)[checked ? "checked" : "unchecked"];

  // A checkbox wrapped around a paragraph or cell changes only the run of
  // its glyph: other text there is not the checkbox's.
  const glyphRun = scope === "block" ? blockGlyphRun(para, control) : undefined;

  // The glyph keeps the formatting of the one it replaces, in the state's font.
  const oldRun = scope === "inline" ? para.runs.find((r) => r.control === control) : glyphRun;
  const oldRPr = oldRun ? findChild(oldRun.xmlRef, "w:rPr") : controlRunProperties(control);
  const rPrChildren = oldRPr
    ? getChildren(oldRPr).filter((c) => getTagName(c) !== "w:rPrChange").map(cloneNode)
    : [];
  if (state.font) {
    let fonts = rPrChildren.find((c) => getTagName(c) === "w:rFonts");
    if (!fonts) {
      fonts = createElement("w:rFonts", []);
      // w:rFonts follows w:rStyle in CT_RPr
      const styleIdx = rPrChildren.findIndex((c) => getTagName(c) === "w:rStyle");
      rPrChildren.splice(styleIdx + 1, 0, fonts);
    }
    for (const attr of ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"]) setAttr(fonts, attr, state.font);
  }
  const rPr = rPrChildren.length > 0 ? createElement("w:rPr", rPrChildren) : undefined;
  const run = createFormattedRuns([{ text: state.glyph }], rPr);

  if (glyphRun) {
    replaceRun(para, glyphRun.xmlRef, run, revCtx, !control.placeholder);
  } else {
    replaceContent(contentArray(para, control, scope), run, revCtx, !control.placeholder);
  }
  clearPlaceholderFlag(control);

  const sdtPr = findChild(control.xmlRef, "w:sdtPr");
  const box = sdtPr ? findChild(sdtPr, "w14:checkbox") : undefined;
  if (box) {
    let flag = findChild(box, "w14:checked");
    if (!flag) {
      flag = createElement("w14:checked", []);
      getChildren(box).unshift(flag);
    }
    setAttr(flag, "w14:val", checked ? "1" : "0");
  }
  refresh(control);
}

/**
 * The run that shows the glyph of a checkbox wrapped around a paragraph or
 * cell: the first run outside any inline control whose text is the checked
 * or the unchecked glyph.
 */
function blockGlyphRun(para: NormalizedParagraph, control: ControlInfo): RunMapping | undefined {
  const { checked, unchecked } = checkboxStates(control);
  return para.runs.find((r) => !r.control && (r.text === checked.glyph || r.text === unchecked.glyph));
}

/**
 * Replace one run by `runs`, in whichever parent holds it. With `trackOld`
 * the old run is marked deleted and the new ones follow it; the new runs are
 * wrapped in `w:ins` whenever a revision context is given.
 */
function replaceRun(
  para: NormalizedParagraph,
  runNode: FXPNode,
  runs: FXPNode[],
  revCtx: RevisionContext | undefined,
  trackOld: boolean,
): void {
  const parent = findRunParentChildren(para.xmlRef, runNode);
  if (!parent) throw new Error(`The checkbox glyph was not found in paragraph "${para.id}"`);
  const old = [runNode];
  if (revCtx && trackOld) markRunsDeleted(old, revCtx);
  else old.length = 0;
  parent.splice(parent.indexOf(runNode), 1, ...old, ...(revCtx ? [createInsElement(revCtx, runs)] : runs));
}

/**
 * The checkbox controls of a paragraph, in order: one wrapped around the
 * paragraph (or its cell) first, then those inside it.
 */
function checkboxesOf(para: NormalizedParagraph): { control: ControlInfo; scope: ControlScope }[] {
  const out: { control: ControlInfo; scope: ControlScope }[] = [];
  if (para.control?.kind === "checkbox") out.push({ control: para.control, scope: "block" });
  for (const range of para.controls ?? []) {
    if (range.control.kind === "checkbox") out.push({ control: range.control, scope: "inline" });
  }
  return out;
}

/** Find the checkbox a set_checkbox patch names; throws when there is none or it is locked. */
export function findCheckbox(
  para: NormalizedParagraph,
  patch: SetCheckboxPatch,
): { control: ControlInfo; scope: ControlScope } {
  const boxes = checkboxesOf(para);
  const n = patch.occurrence ?? 1;
  if (boxes.length === 0) {
    throw new ControlEditError(`${para.id}: the paragraph has no checkbox control`);
  }
  if (n > boxes.length) {
    throw new ControlEditError(
      `${para.id}: occurrence ${n} asked for, but the paragraph has ${boxes.length} checkbox control${boxes.length === 1 ? "" : "s"}`,
    );
  }
  const box = boxes[n - 1]!;
  assertOpen(para.id, "the checkbox", box.control);
  // Wrapped around a paragraph or cell, the checkbox must show its glyph in
  // a run of its own, or hold nothing else, so that no other text is touched.
  if (box.scope === "block" && !blockGlyphRun(para, box.control) && para.text.trim().length > 0) {
    throw new ControlEditError(
      `${para.id}: the checkbox content control wrapped around this paragraph shows no checkbox glyph in a run of its own; ` +
        `it cannot be set without changing the paragraph's other text`,
    );
  }
  return box;
}

/** Execute a set_checkbox patch. Setting the state it already has changes nothing. */
export function executeSetCheckbox(
  para: NormalizedParagraph,
  patch: SetCheckboxPatch,
  revCtx?: RevisionContext,
): void {
  const { control, scope } = findCheckbox(para, patch);
  if (control.checked === patch.checked && !control.placeholder) return;
  writeCheckbox(para, control, scope, patch.checked, revCtx);
}
