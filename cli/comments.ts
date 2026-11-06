/**
 * Read existing Word comments: author, date, text, and the passage each is anchored to.
 * The library writes comments but does not read them, so this walks the XML directly.
 */

import type { Docxtor, FXPNode } from "../src/index.js";
import { getAttr, getChildren, getTagName, unescapeXmlText } from "../src/xml/helpers.js";
import { allParagraphs } from "./paragraphs.js";

export interface CommentInfo {
  id: string;
  author: string;
  initials: string;
  date: string;
  text: string;
  /** Id of the paragraph where the anchored passage starts ("" when the anchor isn't in the body, headers or footers). */
  paragraph: string;
  /** The commented passage (empty for a point comment). */
  anchor: string;
}

function textOf(nodes: FXPNode[]): string {
  let s = "";
  for (const n of nodes) {
    const tag = getTagName(n);
    if (tag === "w:t") for (const t of getChildren(n)) { if (t["#text"] !== undefined) s += unescapeXmlText(String(t["#text"])); }
    else if (tag === "w:tab") s += "\t";
    else if (tag === "w:br" || tag === "w:cr") s += "\n";
    else if (tag && tag !== "w:delText" && tag !== "w:instrText") s += textOf(getChildren(n));
  }
  return s;
}

export function listComments(doc: Docxtor): CommentInfo[] {
  const part = doc.getAST().xmlParts.comments;
  if (!part) return [];
  const root = part.find((n) => getTagName(n) === "w:comments");
  const comments: CommentInfo[] = [];
  for (const c of root ? getChildren(root) : []) {
    if (getTagName(c) !== "w:comment") continue;
    const paras = getChildren(c).filter((p) => getTagName(p) === "w:p").map((p) => textOf(getChildren(p)));
    comments.push({
      id: getAttr(c, "w:id") ?? "",
      author: getAttr(c, "w:author") ?? "",
      initials: getAttr(c, "w:initials") ?? "",
      date: getAttr(c, "w:date") ?? "",
      text: paras.join("\n"),
      paragraph: "",
      anchor: "",
    });
  }
  const byId = new Map(comments.map((c) => [c.id, c]));

  // Walk every paragraph in document order; text between commentRangeStart and commentRangeEnd is the anchor.
  const open = new Set<string>();
  const walk = (nodes: FXPNode[], paraId: string): void => {
    for (const n of nodes) {
      const tag = getTagName(n);
      if (tag === "w:commentRangeStart" || tag === "w:commentReference") {
        const c = byId.get(getAttr(n, "w:id") ?? "");
        if (c && !c.paragraph) c.paragraph = paraId;
        if (c && tag === "w:commentRangeStart") open.add(c.id);
      } else if (tag === "w:commentRangeEnd") {
        open.delete(getAttr(n, "w:id") ?? "");
      } else if (tag === "w:r") {
        const t = textOf(getChildren(n));
        for (const id of open) byId.get(id)!.anchor += t;
      } else if (tag === "w:del" || tag === "w:moveFrom") {
        continue; // deleted text is not part of the current passage
      } else if (tag) {
        walk(getChildren(n), paraId);
      }
    }
  };
  for (const { para } of allParagraphs(doc)) {
    walk(getChildren(para.xmlRef), para.id);
    for (const id of open) byId.get(id)!.anchor += "\n";
  }
  for (const c of comments) c.anchor = c.anchor.replace(/\n+$/, "");
  return comments;
}

export function formatComments(comments: CommentInfo[]): string {
  if (comments.length === 0) return "No comments.";
  const clip = (s: string, n: number): string => {
    const t = s.replace(/\s+/g, " ").trim();
    return t.length > n ? `${t.slice(0, n)}…` : t;
  };
  const lines = comments.map((c, i) =>
    `#${i + 1} [${c.paragraph || "-"}] by ${c.author || "(no author)"}${c.date ? ` ${c.date.slice(0, 10)}` : ""}: ${c.text.replace(/\n/g, " / ")}` +
    (c.anchor ? `\n    on: "${clip(c.anchor, 160)}"` : ""));
  return `${comments.length} comment${comments.length === 1 ? "" : "s"}\n${lines.join("\n")}`;
}
