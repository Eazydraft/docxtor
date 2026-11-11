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
import { markRunsDeleted } from "./paragraph-ops.js";
import { createInsElement, type RevisionContext } from "./revision.js";
import { controlChain, readControlInfo } from "../normalizer/control.js";
import {
  cloneNode,
  createElement,
  findChild,
  getAttr,
  getChildren,
  getTagName,
  removeAttr,
  setAttr,
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

/**
 * Refuse an edit to the content of `control` when it, or a control around
 * it, is content-locked or data-bound.
 */
function assertEditable(paraId: string, control: ControlInfo | undefined): void {
  for (const c of controlChain(control)) {
    if (c.locked === "content" || c.locked === "both") {
      throw new ControlEditError(
        `${paraId}: the text is in a locked ${describe(c)}; its content cannot be edited`,
      );
    }
    if (c.bound) {
      throw new ControlEditError(
        `${paraId}: the ${describe(c)} is bound to a custom XML part; Word overwrites its text ` +
          `from that part when the file is opened, so an edit here would be lost`,
      );
    }
  }
}

/**
 * Check a patch that changes the runs of a paragraph without changing its
 * text (formatting, a character style): refused inside a locked control.
 * `runs` are the runs it touches; omit for the whole paragraph.
 */
export function checkRunsEditable(para: NormalizedParagraph, runs?: RunMapping[]): void {
  for (const c of controlChain(para.control)) {
    if (c.locked === "content" || c.locked === "both") {
      throw new ControlEditError(
        `${para.id}: the paragraph is in a locked ${describe(c)}; its content cannot be edited`,
      );
    }
  }
  const seen = new Set<ControlInfo>();
  for (const run of runs ?? para.runs) {
    if (!run.control || seen.has(run.control)) continue;
    seen.add(run.control);
    for (const c of controlChain(run.control)) {
      if (c.locked === "content" || c.locked === "both") {
        throw new ControlEditError(
          `${para.id}: "${run.text}" is in a locked ${describe(c)}; its content cannot be edited`,
        );
      }
    }
  }
}

/**
 * Check that new block content (a paragraph, a list item, a table) may be
 * put next to `para`. It lands in the control that wraps `para`, so that
 * control must be an open, filled rich-text one.
 */
export function checkBlockInsert(para: NormalizedParagraph): void {
  const control = para.control;
  if (!control) return;
  assertEditable(para.id, control);
  if (control.placeholder) {
    throw new ControlEditError(
      `${para.id}: the paragraph is the placeholder of an empty ${describe(control)}; ` +
        `fill the control with replace_text before adding content to it`,
    );
  }
  if (control.kind !== "richText" && control.kind !== "docPart" && control.kind !== "other") {
    throw new ControlEditError(
      `${para.id}: the paragraph is in a ${describe(control)}, which holds a single value; ` +
        `content cannot be added to it`,
    );
  }
}

/**
 * Check that a paragraph may be deleted: not when a content-locked control
 * wraps it, and not when it holds a control that is locked against deletion.
 */
export function checkParagraphDelete(para: NormalizedParagraph): void {
  for (const c of controlChain(para.control)) {
    if (c.locked === "content" || c.locked === "both") {
      throw new ControlEditError(
        `${para.id}: the paragraph is in a locked ${describe(c)}; its content cannot be deleted`,
      );
    }
  }
  for (const range of para.controls ?? []) {
    if (range.control.locked === "control" || range.control.locked === "both") {
      throw new ControlEditError(
        `${para.id}: the paragraph holds a ${describe(range.control)} that is locked against deletion`,
      );
    }
  }
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
      const target = findInsertTargetRun(para, offset);
      if (!target) return undefined;
      runs = [target];
      editStart = offset;
      editEnd = offset;
    }
  }
  if (runs.length === 0) return undefined;

  if (!sameControl(runs)) {
    const edge = runs.find((r) => r.control)!.control!;
    throw new ControlEditError(
      `${para.id}: "${needle}" crosses the edge of a ${describe(edge)}; ` +
        `edit the text inside the control and the text around it separately`,
    );
  }

  const inline = runs[0]!.control;
  const control = inline ?? para.control;
  if (!control) return undefined;
  const scope: ControlScope = inline ? "inline" : "block";

  assertEditable(para.id, control);
  if (control.kind === "picture") {
    throw new ControlEditError(`${para.id}: a ${describe(control)} cannot be edited`);
  }

  const range = inline
    ? para.controls!.find((r) => r.control === inline)!
    : { startOffset: 0, endOffset: para.text.length };
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

/** Children of a control's content that hold runs: what a fill replaces. */
const RUN_BEARING = new Set([
  "w:r",
  "w:ins",
  "w:del",
  "w:moveTo",
  "w:moveFrom",
  "w:hyperlink",
  "w:smartTag",
  "w:fldSimple",
  "w:sdt",
]);

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
 * `track` marks the old content deleted instead of removing it; the new
 * runs are wrapped in `w:ins` whenever a revision context is given.
 */
function replaceContent(
  content: FXPNode[],
  runs: FXPNode[],
  revCtx: RevisionContext | undefined,
  trackOld: boolean,
): void {
  if (revCtx && trackOld) {
    markRunsDeleted(content, revCtx);
    content.push(createInsElement(revCtx, runs));
    return;
  }
  let at = content.findIndex((c) => RUN_BEARING.has(getTagName(c) ?? ""));
  if (at === -1) at = content.length;
  for (let i = content.length - 1; i >= 0; i--) {
    if (RUN_BEARING.has(getTagName(content[i]!) ?? "")) content.splice(i, 1);
  }
  content.splice(at, 0, ...(revCtx ? [createInsElement(revCtx, runs)] : runs));
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
  const fresh = readControlInfo(control.xmlRef, control.parent);
  for (const key of ["options", "checked", "locked", "bound", "alias", "tag"] as const) delete control[key];
  Object.assign(control, fresh);
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

  // The glyph keeps the formatting of the one it replaces, in the state's font.
  const oldRun = scope === "inline" ? para.runs.find((r) => r.control === control) : para.runs[0];
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

  replaceContent(contentArray(para, control, scope), run, revCtx, !control.placeholder);
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
  assertEditable(para.id, box.control);
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
