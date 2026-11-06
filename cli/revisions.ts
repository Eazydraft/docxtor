/**
 * List tracked changes grouped the way a reviewer reads them: a deletion
 * directly followed by an insertion is one replacement, with both w:ids.
 */

import type { Docxtor, FXPNode } from "../src/index.js";
import { getAttr, getChildren, getTagName, unescapeXmlText } from "../src/xml/helpers.js";
import { allParagraphs } from "./paragraphs.js";

export interface Change {
  /** 1-based, in document order. */
  n: number;
  paragraph: string;
  kind: "insertion" | "deletion" | "replacement" | "formatting" | "paragraph" | "other";
  ids: string[];
  authors: string[];
  date: string;
  deleted: string;
  inserted: string;
  /** The paragraph with this change marked as [-old-][+new+] and other changes applied as current text. */
  context: string;
}

const TEXT_REV = new Set(["w:ins", "w:del", "w:moveFrom", "w:moveTo"]);
const PROP_CHANGE = /^w:\w+PrChange$/;
/** Elements that sit between runs without breaking a change apart. */
const TRANSPARENT = new Set([
  "w:bookmarkStart", "w:bookmarkEnd", "w:proofErr", "w:commentRangeStart", "w:commentRangeEnd",
  "w:moveFromRangeStart", "w:moveFromRangeEnd", "w:moveToRangeStart", "w:moveToRangeEnd", "w:pPr",
]);

function runText(run: FXPNode): string {
  let s = "";
  for (const c of getChildren(run)) {
    const tag = getTagName(c);
    if (tag === "w:t" || tag === "w:delText") {
      for (const t of getChildren(c)) if (t["#text"] !== undefined) s += unescapeXmlText(String(t["#text"]));
    } else if (tag === "w:tab") s += "\t";
    else if (tag === "w:br" || tag === "w:cr") s += "\n";
    else if (tag === "w:noBreakHyphen") s += "-";
  }
  return s;
}

function textOf(node: FXPNode): string {
  const tag = getTagName(node);
  if (tag === "w:r") return runText(node);
  return getChildren(node).map((c) => textOf(c)).join("");
}

/** A run's property change, if any. */
function runPropChange(run: FXPNode): FXPNode | undefined {
  const rPr = getChildren(run).find((c) => getTagName(c) === "w:rPr");
  return rPr ? getChildren(rPr).find((c) => getTagName(c) === "w:rPrChange") : undefined;
}

interface Piece {
  kind: "text" | "ins" | "del" | "fmt";
  text: string;
  node?: FXPNode;
}

function pieces(children: FXPNode[], out: Piece[]): void {
  for (const c of children) {
    const tag = getTagName(c);
    if (!tag || TRANSPARENT.has(tag)) continue;
    if (tag === "w:ins" || tag === "w:moveTo") out.push({ kind: "ins", text: textOf(c), node: c });
    else if (tag === "w:del" || tag === "w:moveFrom") out.push({ kind: "del", text: textOf(c), node: c });
    else if (tag === "w:r") {
      const chg = runPropChange(c);
      out.push(chg ? { kind: "fmt", text: runText(c), node: chg } : { kind: "text", text: runText(c) });
    } else if (tag === "w:hyperlink" || tag === "w:smartTag" || tag === "w:sdt" || tag === "w:sdtContent" || tag === "w:fldSimple") {
      pieces(getChildren(c), out);
    }
  }
}

function meta(node: FXPNode): { id?: string; author: string; date: string } {
  return { id: getAttr(node, "w:id"), author: getAttr(node, "w:author") ?? "", date: getAttr(node, "w:date") ?? "" };
}

export function listChanges(doc: Docxtor): Change[] {
  const changes: Change[] = [];
  const seen = new Set<FXPNode>();

  for (const { para } of allParagraphs(doc)) {
    const all: Piece[] = [];
    pieces(getChildren(para.xmlRef), all);
    // Runs with no text (formatting-only runs, field chars) don't separate a deletion from its insertion.
    const ps = all.filter((p) => !(p.kind === "text" && p.text === ""));

    // Group consecutive ins/del pieces with no current text between them.
    const groups: { from: number; to: number }[] = [];
    for (let i = 0; i < ps.length; i++) {
      const p = ps[i]!;
      if (p.kind === "ins" || p.kind === "del") {
        const last = groups[groups.length - 1];
        const lastPiece = last ? ps[last.to]! : undefined;
        // Adjacent, or a deletion and an insertion separated only by whitespace ("[-Baa3-] [+Ba2+]").
        const gap = last ? ps.slice(last.to + 1, i) : [];
        const joins = last && lastPiece && (lastPiece.kind === "ins" || lastPiece.kind === "del")
          && (gap.length === 0 || (lastPiece.kind !== p.kind && gap.every((g) => g.kind === "text" && g.text.trim() === "")));
        if (joins) last!.to = i;
        else groups.push({ from: i, to: i });
      } else if (p.kind === "fmt") {
        groups.push({ from: i, to: i });
      }
    }

    for (const g of groups) {
      const members = ps.slice(g.from, g.to + 1);
      for (const m of members) if (m.node) seen.add(m.node);
      const deleted = members.filter((m) => m.kind === "del").map((m) => m.text).join("");
      const inserted = members.filter((m) => m.kind === "ins").map((m) => m.text).join("");
      const metas = members.filter((m) => m.node).map((m) => meta(m.node!));
      const fmt = members[0]!.kind === "fmt";
      const before = ps.slice(0, g.from).filter((p) => p.kind !== "del").map((p) => p.text).join("");
      const after = ps.slice(g.to + 1).filter((p) => p.kind !== "del").map((p) => p.text).join("");
      const mid = fmt ? `[fmt:${members[0]!.text}]` : members.map((m) => (m.kind === "del" ? `[-${m.text}-]` : m.kind === "ins" ? `[+${m.text}+]` : m.text)).join("");
      changes.push({
        n: 0,
        paragraph: para.id,
        kind: fmt ? "formatting" : deleted && inserted ? "replacement" : deleted ? "deletion" : "insertion",
        ids: metas.map((m) => m.id).filter((x): x is string => x !== undefined),
        authors: [...new Set(metas.map((m) => m.author))],
        date: metas[0]?.date ?? "",
        deleted: fmt ? "" : deleted,
        inserted: fmt ? members[0]!.text : inserted,
        context: clip(before, true) + mid + clip(after, false),
      });
    }

    // Paragraph-level marks: the paragraph mark (inserted/deleted paragraph) and pPrChange.
    const pPr = getChildren(para.xmlRef).find((c) => getTagName(c) === "w:pPr");
    if (pPr) {
      for (const c of getChildren(pPr)) {
        const tag = getTagName(c)!;
        if (tag === "w:rPr") {
          for (const m of getChildren(c)) {
            const mt = getTagName(m);
            if (mt === "w:ins" || mt === "w:del") {
              seen.add(m);
              const md = meta(m);
              // A wholly inserted or deleted paragraph: fold the mark into its text change.
              const kind = mt === "w:ins" ? "insertion" : "deletion";
              const own = changes.filter((c) => c.paragraph === para.id && c.kind !== "paragraph");
              const hasCurrentText = ps.some((p) => p.kind === "text" && p.text.length > 0) || (kind === "deletion" && ps.some((p) => p.kind === "ins"));
              if (own.length === 1 && own[0]!.kind === kind && !hasCurrentText && own[0]!.authors.includes(md.author)) {
                if (md.id) own[0]!.ids.push(md.id);
                own[0]!.kind = kind;
                own[0]!.context = `whole paragraph ${kind === "insertion" ? "inserted" : "deleted"}: ${own[0]!.context}`;
                continue;
              }
              changes.push({
                n: 0, paragraph: para.id, kind: "paragraph", ids: md.id ? [md.id] : [], authors: [md.author], date: md.date,
                deleted: "", inserted: "",
                context: `paragraph mark ${mt === "w:ins" ? "inserted" : "deleted"} (${mt === "w:ins" ? "new paragraph break" : "paragraph merged with the next"}): ${clip(para.text, false)}`,
              });
            } else if (mt === "w:rPrChange") seen.add(m);
          }
        } else if (PROP_CHANGE.test(tag)) {
          seen.add(c);
          const md = meta(c);
          changes.push({
            n: 0, paragraph: para.id, kind: "formatting", ids: md.id ? [md.id] : [], authors: [md.author], date: md.date,
            deleted: "", inserted: "", context: `paragraph properties changed: ${clip(para.text, false)}`,
          });
        }
      }
    }
  }

  // Anything else that carries a w:id and w:author (row/cell markers, section changes, notes).
  const parts: FXPNode[][] = [doc.getAST().xmlParts.document];
  const walk = (nodes: FXPNode[]): void => {
    for (const n of nodes) {
      const tag = getTagName(n);
      if (!tag || tag === "#text") continue;
      if (!seen.has(n) && (TEXT_REV.has(tag) || PROP_CHANGE.test(tag) || tag === "w:cellIns" || tag === "w:cellDel" || tag === "w:cellMerge")
        && getAttr(n, "w:author") !== undefined && getAttr(n, "w:id") !== undefined) {
        seen.add(n);
        const md = meta(n);
        changes.push({
          n: 0, paragraph: "-", kind: "other", ids: [md.id!], authors: [md.author], date: md.date,
          deleted: "", inserted: "", context: `${tag}${textOf(n) ? `: ${clip(textOf(n), false)}` : ""}`,
        });
        continue;
      }
      walk(getChildren(n));
    }
  };
  for (const p of parts) walk(p);

  changes.forEach((c, i) => (c.n = i + 1));
  return changes;
}

function clip(s: string, tail: boolean, width = 60): string {
  s = s.replace(/\s+/g, " ");
  if (s.length <= width) return s;
  return tail ? `…${s.slice(-width)}` : `${s.slice(0, width)}…`;
}

export function formatChanges(changes: Change[]): string {
  if (changes.length === 0) return "No tracked changes.";
  const lines = changes.map((c) => {
    const what =
      c.kind === "replacement" ? `replace "${c.deleted}" -> "${c.inserted}"`
      : c.kind === "deletion" ? `delete "${c.deleted}"`
      : c.kind === "insertion" ? `insert "${c.inserted}"`
      : c.kind === "formatting" ? `formatting${c.inserted ? ` of "${c.inserted}"` : ""}`
      : c.kind;
    return `#${c.n} [${c.paragraph}] ids=${c.ids.join(",")} by ${c.authors.join(", ")}: ${what}\n    ${c.context}`;
  });
  const authors = [...new Set(changes.flatMap((c) => c.authors))];
  return `${changes.length} change${changes.length === 1 ? "" : "s"} by ${authors.map((a) => `"${a}"`).join(", ")}\n${lines.join("\n")}`;
}
