/**
 * Footnote and endnote normalizer.
 *
 * `word/footnotes.xml` and `word/endnotes.xml` hold one `w:footnote` /
 * `w:endnote` per note, with the same block content as the body: paragraphs
 * and tables. A note is tied to the text by a `w:footnoteReference` /
 * `w:endnoteReference` in a run, which names the note's `w:id`.
 *
 * The `w:id` is not the number Word displays. Word numbers notes by the
 * order of their references in the text, so that is the number used here,
 * in the mark the view prints (`[^1]`, `[^e1]`) and in the ids of the note's
 * paragraphs (`fn1.p0`, `en1.p0`).
 */

import type {
  FXPNode,
  NormalizedDocument,
  NormalizedNote,
  NormalizedParagraph,
  NormalizedTable,
  NoteKind,
  NoteReference,
} from "../types.js";
import { getAttr, getChildren, getTagName, textElementContent, runChildChar, contentChildren } from "../xml/helpers.js";
import { ControlResolver } from "./control.js";
import { normalizeParagraph } from "./paragraph.js";
import { normalizeTable } from "./table.js";

const PARTS: Record<NoteKind, { root: string; note: string; prefix: string }> = {
  footnote: { root: "w:footnotes", note: "w:footnote", prefix: "fn" },
  endnote: { root: "w:endnotes", note: "w:endnote", prefix: "en" },
};

/**
 * The note references in a run, each with the number of characters of the
 * run's text that precede it.
 */
export function runNoteReferences(runNode: FXPNode): NoteReference[] {
  const refs: NoteReference[] = [];
  let chars = 0;
  for (const child of getChildren(runNode)) {
    if ("w:t" in child) {
      chars += textElementContent(child).length;
    } else if ("w:footnoteReference" in child || "w:endnoteReference" in child) {
      const noteId = getAttr(child, "w:id");
      if (noteId === undefined) continue;
      const custom = getAttr(child, "w:customMarkFollows");
      refs.push({
        kind: "w:footnoteReference" in child ? "footnote" : "endnote",
        noteId,
        offset: chars,
        ...(custom !== undefined && custom !== "0" && custom !== "false" ? { customMark: true } : {}),
      });
    } else {
      chars += (runChildChar(child) ?? "").length;
    }
  }
  return refs;
}

/**
 * Every note reference in the body, in document order. Inside a table cell
 * the paragraphs and nested tables are taken in the order of the XML, so a
 * reference in a nested table is counted where it sits, between the
 * paragraphs around that table.
 */
function bodyNoteReferences(doc: NormalizedDocument): NoteReference[] {
  const refs: NoteReference[] = [];
  const fromTable = (table: NormalizedTable): void => {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        const paragraphs = new Map(cell.paragraphs.map((p) => [p.xmlRef, p]));
        const nested = new Map(cell.nestedTables.map((t) => [t.xmlRef, t]));
        // Read through content controls, as the table normalizer does
        for (const { node: child } of contentChildren(cell.xmlRef)) {
          const para = paragraphs.get(child);
          if (para) refs.push(...(para.noteRefs ?? []));
          const inner = nested.get(child);
          if (inner) fromTable(inner);
        }
      }
    }
  };
  for (const el of doc.body) {
    if (el.kind === "paragraph") refs.push(...(el.noteRefs ?? []));
    else if (el.kind === "table") fromTable(el);
  }
  return refs;
}

/**
 * Rebuild `doc.footnotes` and `doc.endnotes` from the note parts. Call after
 * the body is normalized: the numbers come from the references in it.
 *
 * Numbering, per kind:
 * 1. notes with an automatic mark, in order of their first reference: these
 *    are the numbers Word shows;
 * 2. notes whose reference carries its own mark (`w:customMarkFollows`),
 *    which Word leaves out of its count, in order of first reference;
 * 3. notes nothing refers to, in the order of the part.
 *
 * The separator entries (`w:type="separator"` and the like) are not notes
 * and are left out.
 */
export function normalizeNotes(doc: NormalizedDocument): void {
  const refs = bodyNoteReferences(doc);
  doc.footnotes = normalizeNotePart(doc.xmlParts.footnotes, "footnote", refs);
  doc.endnotes = normalizeNotePart(doc.xmlParts.endnotes, "endnote", refs);
}

function normalizeNotePart(
  part: FXPNode[] | undefined,
  kind: NoteKind,
  allRefs: NoteReference[],
): Map<string, NormalizedNote> {
  const notes = new Map<string, NormalizedNote>();
  if (!part) return notes;
  const { root: rootTag, note: noteTag, prefix } = PARTS[kind];
  const root = part.find((n) => getTagName(n) === rootTag);
  if (!root) return notes;

  const nodes = new Map<string, FXPNode>();
  for (const node of getChildren(root)) {
    if (getTagName(node) !== noteTag) continue;
    const type = getAttr(node, "w:type");
    if (type !== undefined && type !== "normal") continue;
    const id = getAttr(node, "w:id");
    if (id !== undefined && !nodes.has(id)) nodes.set(id, node);
  }

  // First reference to each note, in document order
  const first = new Map<string, NoteReference>();
  for (const ref of allRefs) {
    if (ref.kind === kind && nodes.has(ref.noteId) && !first.has(ref.noteId)) first.set(ref.noteId, ref);
  }
  const order = [
    ...[...first.values()].filter((r) => !r.customMark).map((r) => r.noteId),
    ...[...first.values()].filter((r) => r.customMark).map((r) => r.noteId),
    ...[...nodes.keys()].filter((id) => !first.has(id)),
  ];

  order.forEach((noteId, i) => {
    const node = nodes.get(noteId)!;
    const number = i + 1;
    const idPrefix = `${prefix}${number}`;
    const paragraphs: NormalizedParagraph[] = [];
    const tables: NormalizedTable[] = [];
    // Read through block-level content controls, as in the body
    const resolver = new ControlResolver();
    for (const { node: child, sdts } of contentChildren(node)) {
      const tag = getTagName(child);
      const control = resolver.resolve(sdts);
      if (tag === "w:p") paragraphs.push(normalizeParagraph(child, `${idPrefix}.p${paragraphs.length}`, control));
      else if (tag === "w:tbl") tables.push(normalizeTable(child, `${idPrefix}.tbl${tables.length}`, control, resolver));
    }
    const ref = first.get(noteId);
    notes.set(noteId, {
      kind,
      noteId,
      number,
      idPrefix,
      referenced: ref !== undefined,
      ...(ref?.customMark ? { customMark: true } : {}),
      paragraphs,
      tables,
      xmlRef: node,
    });
  });
  return notes;
}

/** Every footnote, then every endnote, in number order. */
export function allNotes(doc: NormalizedDocument): NormalizedNote[] {
  return [...doc.footnotes.values(), ...doc.endnotes.values()];
}

/** The note that holds a paragraph directly (not inside one of its tables). */
export function findNoteOfParagraph(doc: NormalizedDocument, para: NormalizedParagraph): NormalizedNote | undefined {
  return allNotes(doc).find((note) => note.paragraphs.includes(para));
}

/** True if the paragraph holds the note's own mark (`w:footnoteRef` / `w:endnoteRef`). */
export function holdsNoteMark(paragraphNode: FXPNode): boolean {
  const scan = (nodes: FXPNode[]): boolean =>
    nodes.some((n) => {
      const tag = getTagName(n);
      if (tag === "w:footnoteRef" || tag === "w:endnoteRef") return true;
      return tag !== undefined && tag !== "w:pPr" && scan(getChildren(n));
    });
  return scan(getChildren(paragraphNode));
}
