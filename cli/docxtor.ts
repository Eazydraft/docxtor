/**
 * docxtor CLI: the command-line front end the agent skill drives.
 *
 * Bundled with the library into plugin/skills/docxtor/scripts/docxtor.mjs
 * (`pnpm build:skill`), so the skill needs only Node 20+.
 */

import { readFile } from "node:fs/promises";
import { parseArgs, type ParseArgsOptionsConfig } from "node:util";
import { Docxtor, type NormalizedTable, type Patch, type PatchOptions, type PatchResult } from "../src/index.js";
import { allParagraphs, matchStarts, occurrenceAt, patchTargets, planEdit, quoteVariants, snippet } from "./paragraphs.js";
import { formatChanges, listChanges } from "./revisions.js";
import { formatComments, listComments } from "./comments.js";
import { getAttr, getChildren, getTagName, unescapeXmlText } from "../src/xml/helpers.js";

const USAGE = `docxtor <command> [options]

Read
  view FILE [--grep REGEX [--context N]] [--formatting] [--no-revisions] [--styles] [--json]
      Paragraphs with their ids ([p12], [tbl0.r1.c0.p0], [hdr.default.p0]).
      Tracked changes show inline as {+inserted+} and {-deleted-}.
  revisions FILE [--json]
      Every tracked change, numbered, with its w:ids, author and context.
  comments FILE [--json]
      Existing comments: author, date, text, and the passage each is on.
  metadata FILE [--json]
      Document properties: author, last modified by, title, company, dates, ...

Write (tracked changes by default; -o may equal FILE to update in place)
  replace FILE -o OUT --find TEXT --replace TEXT [--find .. --replace ..]
          [--all] [--in ID] [--occurrence N] [--whole-word]
      Find text anywhere (body, tables, headers, footers) and replace it.
      Refuses a find that matches more than once unless --all, --in or
      --occurrence says which.
  comment FILE -o OUT --text TEXT (--find TEXT | --in ID) [--in ID]
      Add a Word comment anchored to TEXT (or to the whole paragraph ID).
  edit FILE -o OUT [PATCHES.json | -] [--dry-run]
      Apply a JSON array of docxtor patches (stdin when omitted or "-").
      All-or-nothing: if any patch fails nothing is written. --dry-run
      checks every patch and reports each one without writing.
  accept FILE -o OUT [--ids 3,4] [--by AUTHOR] [--others reject|keep]
  reject FILE -o OUT [--ids 3,4] [--by AUTHOR] [--others accept|keep]
      Resolve tracked changes: all of them, or only the listed w:ids
      (from \`revisions\`) / authors. --others decides the rest (default keep).
  metadata FILE -o OUT [--set key=value ...] [--scrub]
      Set document properties, or --scrub identifying ones (author,
      last modified by, company, manager, custom properties).

Common write options
  --author NAME     revision and comment author (default "docxtor")
  --untracked       edit directly instead of as tracked changes
`;

const STRING = { type: "string" } as const;
const BOOL = { type: "boolean" } as const;

function fail(msg: string): never {
  process.stderr.write(`error: ${msg}\n`);
  process.exit(1);
}

function parse<const T extends ParseArgsOptionsConfig>(argv: string[], options: T) {
  try {
    return parseArgs({ args: argv, options, allowPositionals: true, strict: true });
  } catch (e) {
    fail(`${(e as Error).message}\n\n${USAGE}`);
  }
}

const WRITE_OPTS = {
  out: { ...STRING, short: "o" },
  author: STRING,
  untracked: BOOL,
} as const;

function patchOptions(values: { author?: string; untracked?: boolean }): PatchOptions {
  return {
    trackRevisions: !values.untracked,
    revisionAuthor: values.author ?? "docxtor",
    inlineMarkers: false,
    atomic: true,
  };
}

async function load(file: string | undefined): Promise<Docxtor> {
  if (!file) fail(`missing input file\n\n${USAGE}`);
  try {
    return await Docxtor.load(file);
  } catch (e) {
    fail(`cannot open ${file}: ${(e as Error).message}`);
  }
}

function needOut(out: string | undefined): string {
  if (!out) fail("missing -o OUT (the file to write; it may be the input file to update it in place)");
  return out;
}

/** Apply atomically; on success write OUT and summarise, on failure explain and write nothing. */
async function applyAndSave(doc: Docxtor, patches: Patch[], opts: PatchOptions, out: string): Promise<PatchResult[]> {
  const before = await doc.toBuffer();
  const results = doc.applyPatches(patches, opts);
  const failed = results.filter((r) => !r.success);
  if (failed.length > 0) {
    const notes = await explainFailures(before, patches, results, opts);
    const lines = results.flatMap((r, i) => (r.success ? [] : [`  ${JSON.stringify(r.patch)}\n    -> ${notes.get(i)}`]));
    fail(`${failed.length} of ${results.length} patches failed; nothing was written.\n${lines.join("\n")}`);
  }
  await doc.toFile(out);
  return results;
}

/**
 * Explain each failed patch against the document as it stood when that patch ran: a batch is applied
 * in order, so an earlier patch (a new table row, an inserted paragraph) can change what an id points at.
 */
async function explainFailures(before: Buffer, patches: Patch[], results: PatchResult[], opts: PatchOptions): Promise<Map<number, string>> {
  const notes = new Map<number, string>();
  for (let i = 0; i < results.length; i++) {
    if (results[i]!.success) continue;
    const state = await Docxtor.load(before);
    state.applyPatches(patches.slice(0, i), { ...opts, atomic: false });
    notes.set(i, explain(state, results[i]!.error ?? ""));
  }
  return notes;
}

/** Add what the agent needs to fix a failed patch: the paragraph's actual text, or how to find ids. */
function explain(doc: Docxtor, error: string): string {
  const inPara = /in paragraph "([^"]+)"/.exec(error);
  if (inPara) {
    const p = doc.findParagraphById(inPara[1]!) ?? doc.findParagraphByParaId(inPara[1]!.replace(/^para:/, ""));
    if (p) return `${error}\n       ${inPara[1]} currently reads: "${p.text.length > 300 ? `${p.text.slice(0, 300)}…` : p.text}"`;
  }
  if (/not found/i.test(error) && /["'](p\d|tbl\d|hdr\.|ftr\.|para:)/.test(error)) {
    return `${error}\n       (ids come from \`view FILE\`; ids are positional, so re-run view after edits that add or remove paragraphs)`;
  }
  return error;
}

// ---------------------------------------------------------------------------

async function cmdView(argv: string[]): Promise<void> {
  const { values, positionals } = parse(argv, {
    grep: STRING, context: STRING, formatting: BOOL, "no-revisions": BOOL, styles: BOOL, json: BOOL,
  });
  const doc = await load(positionals[0]);
  if (values.json) {
    process.stdout.write(`${JSON.stringify(doc.project("json", { includeRevisions: !values["no-revisions"], includeFormatting: !!values.formatting }), null, 1)}\n`);
    return;
  }
  let md = doc.project("markdown", { includeRevisions: !values["no-revisions"], includeFormatting: !!values.formatting });
  if (!values.styles) md = md.replace(/\n*--- Available Styles ---[\s\S]*$/, "");
  if (!values["no-revisions"]) md = markTrackedRows(doc, md);
  if (values.grep) {
    let re: RegExp;
    try { re = new RegExp(values.grep, "iu"); } catch (e) { fail(`bad --grep: ${(e as Error).message}`); }
    const all = md.split("\n");
    const n = values.context === undefined ? 0 : Number(values.context);
    if (!Number.isInteger(n) || n < 0) fail(`--context must be a whole number of paragraphs, got "${values.context}"`);
    // Keep each match plus n lines either side; separate non-adjacent groups with "--", like grep -C.
    const keep = new Set<number>();
    all.forEach((l, i) => { if (re.test(l)) for (let k = Math.max(0, i - n); k <= Math.min(all.length - 1, i + n); k++) keep.add(k); });
    const idx = [...keep].sort((a, b) => a - b);
    const out: string[] = [];
    idx.forEach((i, k) => { if (n > 0 && k > 0 && i !== idx[k - 1]! + 1) out.push("--"); out.push(all[i]!); });
    md = out.length ? out.join("\n") : `no paragraph matches /${values.grep}/i (try a shorter or different pattern; view FILE without --grep shows everything)`;
  }
  process.stdout.write(`${md}\n`);
}

/**
 * A tracked row insertion or deletion lives on the row (w:trPr), not on its text, so the projection shows the
 * row as ordinary. Label such rows: "| [tbl0.r1.c0.p0] … | {+row inserted by Acme+}".
 */
function markTrackedRows(doc: Docxtor, md: string): string {
  const labels = new Map<string, string>();
  const walk = (tables: NormalizedTable[]): void => {
    for (const t of tables) {
      for (const row of t.rows) {
        const trPr = getChildren(row.xmlRef).find((c) => getTagName(c) === "w:trPr");
        const mark = trPr ? getChildren(trPr).find((c) => getTagName(c) === "w:ins" || getTagName(c) === "w:del") : undefined;
        if (mark) {
          const who = getAttr(mark, "w:author");
          labels.set(`${t.id}.r${row.index}.`, getTagName(mark) === "w:ins" ? `{+row inserted${who ? ` by ${who}` : ""}+}` : `{-row deleted${who ? ` by ${who}` : ""}-}`);
        }
        for (const cell of row.cells) walk(cell.nestedTables);
      }
    }
  };
  walk(doc.getTables());
  for (const part of [...doc.getHeaders().values(), ...doc.getFooters().values()]) walk(part.tables);
  if (labels.size === 0) return md;
  return md.split("\n").map((line) => {
    // The first cell's paragraph id; the greedy prefix picks the innermost row of a nested table.
    const m = /^\| \[(\S*\.r\d+\.)c\d+\.p\d+\]/.exec(line);
    const label = m ? labels.get(m[1]!) : undefined;
    return label ? `${line} ${label}` : line;
  }).join("\n");
}

async function cmdComments(argv: string[]): Promise<void> {
  const { values, positionals } = parse(argv, { json: BOOL });
  const doc = await load(positionals[0]);
  const comments = listComments(doc);
  process.stdout.write(values.json ? `${JSON.stringify(comments, null, 1)}\n` : `${formatComments(comments)}\n`);
}

async function cmdRevisions(argv: string[]): Promise<void> {
  const { values, positionals } = parse(argv, { json: BOOL });
  const doc = await load(positionals[0]);
  const changes = listChanges(doc);
  process.stdout.write(values.json ? `${JSON.stringify(changes, null, 1)}\n` : `${formatChanges(changes)}\n`);
}

async function cmdReplace(argv: string[]): Promise<void> {
  const { values, positionals } = parse(argv, {
    ...WRITE_OPTS,
    find: { ...STRING, multiple: true }, replace: { ...STRING, multiple: true },
    all: BOOL, in: STRING, occurrence: STRING, "whole-word": BOOL,
  });
  const finds = values.find ?? [];
  const repls = values.replace ?? [];
  if (finds.length === 0 || finds.length !== repls.length) fail("give --find and --replace in pairs");
  const out = needOut(values.out);
  const doc = await load(positionals[0]);
  const inScope = (id: string, paraId?: string): boolean => !values.in || id === values.in || paraId === values.in.replace(/^para:/, "");
  if (values.in && !allParagraphs(doc).some((l) => inScope(l.para.id, l.para.paraId))) fail(`no paragraph ${values.in}`);

  const opts = patchOptions(values);
  const report: string[] = [];
  // Pair by pair, each against the current text, so a later --find sees earlier replacements.
  // Nothing is written unless every pair applies.
  for (let k = 0; k < finds.length; k++) {
    let find = finds[k]!;
    const replace = repls[k]!;
    const every = allParagraphs(doc);
    const targetOf = patchTargets(every);
    const paras = every.filter((l) => inScope(l.para.id, l.para.paraId));
    // Every match. Try curly/straight quote variants if the literal text is absent.
    // `id` is the positional id shown to the user; `target` is what the patch names.
    let hits: { id: string; target: string; text: string; start: number }[] = [];
    for (const variant of quoteVariants(find)) {
      hits = paras.flatMap(({ para }) => matchStarts(para.text, variant, !!values["whole-word"]).map((start) => ({ id: para.id, target: targetOf(para), text: para.text, start })));
      if (hits.length) { find = variant; break; }
    }
    if (hits.length === 0) {
      fail(`"${finds[k]}" not found${values.in ? ` in ${values.in}` : ""}; nothing was written. Text in a docx may differ from what you expect (quotes, dashes, non-breaking spaces, a tracked deletion); check with: view FILE --grep '<a distinctive word>'`);
    }
    if (values.occurrence) {
      const n = Number(values.occurrence);
      if (!Number.isInteger(n) || n < 1 || n > hits.length) fail(`--occurrence ${values.occurrence}: there are ${hits.length} matches`);
      hits = [hits[n - 1]!];
    } else if (hits.length > 1 && !values.all) {
      const list = hits.slice(0, 20).map((h, i) => `  ${i + 1}. [${h.id}] ${snippet(h.text, h.start, h.start + find.length, 40)}`);
      fail(`"${find}" matches ${hits.length} times; nothing was written. Use --all, --in ID or --occurrence N (counted in this list):\n${list.join("\n")}${hits.length > 20 ? "\n  …" : ""}`);
    }
    // Right to left within each paragraph so earlier occurrence numbers stay valid.
    const m = planEdit(find, replace);
    const patches: Patch[] = [...hits].reverse().map((h): Patch => {
      const start = h.start + m.offset;
      if (m.kind === "insert") {
        return { type: "insert_text", target: h.target, position: m.position, anchor: m.anchor, text: m.text, occurrence: occurrenceAt(h.text, m.anchor, start) };
      }
      const occurrence = occurrenceAt(h.text, m.find, start);
      return m.kind === "delete"
        ? { type: "delete_text", target: h.target, find: m.find, occurrence }
        : { type: "replace_text", target: h.target, find: m.find, replace: m.replace, occurrence };
    });
    const failed = doc.applyPatches(patches, opts).filter((r) => !r.success);
    if (failed.length) fail(`${failed.length} replacements of "${find}" failed; nothing was written.\n${failed.map((r) => `  ${JSON.stringify(r.patch)} -> ${r.error}`).join("\n")}`);
    const ids = [...new Set(hits.map((h) => h.id))];
    const marked = m.kind === "insert" ? `inserted "${m.text}"` : m.kind === "delete" ? `deleted "${m.find}"` : `"${m.find}" -> "${m.replace}"`;
    const tracked = m.kind !== "replace" || m.find !== find || m.replace !== replace ? ` (marked: ${marked})` : "";
    report.push(`"${find}" -> "${replace}"${tracked}: ${hits.length} replacement${hits.length === 1 ? "" : "s"} in ${ids.length} paragraph${ids.length === 1 ? "" : "s"} (${ids.slice(0, 12).join(", ")}${ids.length > 12 ? ", …" : ""})`);
  }
  await doc.toFile(out);
  process.stdout.write(`${report.join("\n")}\nwrote ${out}${values.untracked ? "" : ` (tracked, author "${values.author ?? "docxtor"}")`}\n`);
}

async function cmdComment(argv: string[]): Promise<void> {
  const { values, positionals } = parse(argv, { ...WRITE_OPTS, find: STRING, text: STRING, in: STRING, occurrence: STRING });
  if (!values.text) fail("missing --text");
  if (!values.find && !values.in) fail("give --find TEXT (the passage to anchor on) or --in ID");
  const out = needOut(values.out);
  const doc = await load(positionals[0]);
  let target = values.in;
  let shown = values.in;
  let find = values.find;
  let occurrence: number | undefined;
  if (find) {
    const every = allParagraphs(doc);
    const targetOf = patchTargets(every);
    const paras = every.filter((l) => !values.in || l.para.id === values.in || l.para.paraId === values.in.replace(/^para:/, ""));
    let hits: { id: string; target: string; text: string; start: number }[] = [];
    for (const variant of quoteVariants(find)) {
      hits = paras.flatMap(({ para }) => matchStarts(para.text, variant, false).map((start) => ({ id: para.id, target: targetOf(para), text: para.text, start })));
      if (hits.length) { find = variant; break; }
    }
    if (hits.length === 0) fail(`"${values.find}" not found; check with: view FILE --grep '<a distinctive word>'`);
    if (values.occurrence) {
      const n = Number(values.occurrence);
      if (!Number.isInteger(n) || n < 1 || n > hits.length) fail(`--occurrence ${values.occurrence}: there are ${hits.length} matches`);
      hits = [hits[n - 1]!];
    } else if (hits.length > 1) {
      const list = hits.slice(0, 20).map((h, i) => `  ${i + 1}. [${h.id}] ${snippet(h.text, h.start, h.start + find!.length, 40)}`);
      fail(`"${find}" matches ${hits.length} times; give a longer --find, --in ID or --occurrence N:\n${list.join("\n")}`);
    }
    target = hits[0]!.target;
    shown = hits[0]!.id;
    occurrence = occurrenceAt(hits[0]!.text, find, hits[0]!.start);
  }
  const author = values.author ?? "docxtor";
  const patch: Patch = { type: "add_comment", target: target!, text: values.text, author, markers: false, ...(find ? { find, occurrence } : {}) };
  await applyAndSave(doc, [patch], { ...patchOptions(values), trackRevisions: false }, out);
  process.stdout.write(`comment on [${shown}]${find ? ` anchored to "${find}"` : ""} by "${author}"\nwrote ${out}\n`);
}

async function cmdEdit(argv: string[]): Promise<void> {
  const { values, positionals } = parse(argv, { ...WRITE_OPTS, "dry-run": BOOL });
  const out = values["dry-run"] ? "" : needOut(values.out);
  const doc = await load(positionals[0]);
  const src = positionals[1];
  let raw: string;
  if (!src || src === "-") {
    const chunks: Buffer[] = [];
    for await (const c of process.stdin) chunks.push(c as Buffer);
    raw = Buffer.concat(chunks).toString("utf8");
  } else raw = await readFile(src, "utf8");
  let patches: Patch[];
  try {
    const parsed = JSON.parse(raw) as unknown;
    patches = (Array.isArray(parsed) ? parsed : [parsed]) as Patch[];
  } catch (e) {
    fail(`patches are not valid JSON: ${(e as Error).message}`);
  }
  if (values["dry-run"]) {
    // Apply to the in-memory copy only (atomic, so a failure leaves it untouched) and report per patch.
    const before = await doc.toBuffer();
    const results = doc.applyPatches(patches, patchOptions(values));
    const bad = results.filter((r) => !r.success).length;
    const notes = await explainFailures(before, patches, results, patchOptions(values));
    results.forEach((r, i) => process.stdout.write(`${i + 1}. ${r.success ? "ok" : "FAILS"} ${r.patch.type}${"target" in r.patch && r.patch.target ? ` ${r.patch.target}` : ""}${r.success ? "" : `\n   -> ${notes.get(i)}`}\n`));
    process.stdout.write(bad ? `${bad} of ${results.length} would fail; fix them and run again.\n` : `all ${results.length} would apply; run again without --dry-run and with -o OUT to write.\n`);
    if (bad) process.exit(1);
    return;
  }
  const results = await applyAndSave(doc, patches, patchOptions(values), out);
  process.stdout.write(`applied ${results.length} patch${results.length === 1 ? "" : "es"}\nwrote ${out}${values.untracked ? "" : ` (tracked, author "${values.author ?? "docxtor"}")`}\n`);
}

/** Company and Manager from docProps/app.xml (read-only here; --scrub removes them). */
function appProperties(doc: Docxtor): Record<string, string> {
  const root = doc.getAST().xmlParts.app?.find((n) => getTagName(n)?.endsWith("Properties"));
  const out: Record<string, string> = {};
  for (const c of root ? getChildren(root) : []) {
    const local = (getTagName(c) ?? "").replace(/^.*:/, "");
    if (local === "Company" || local === "Manager") {
      const text = getChildren(c).map((t) => (t["#text"] !== undefined ? unescapeXmlText(String(t["#text"])) : "")).join("");
      if (text) out[local.toLowerCase()] = text;
    }
  }
  return out;
}

const META_KEYS = ["creator", "lastModifiedBy", "title", "subject", "description", "keywords", "category", "contentStatus", "revision", "created", "modified"] as const;

async function cmdMetadata(argv: string[]): Promise<void> {
  const { values, positionals } = parse(argv, { out: { ...STRING, short: "o" }, set: { ...STRING, multiple: true }, scrub: BOOL, json: BOOL });
  const doc = await load(positionals[0]);
  const show = (): void => {
    const props = { ...(doc.getCoreProperties() as Record<string, string | undefined>), ...appProperties(doc) };
    if (values.json) { process.stdout.write(`${JSON.stringify(props, null, 1)}\n`); return; }
    for (const k of [...META_KEYS, "company", "manager"]) if (props[k] !== undefined && props[k] !== "") process.stdout.write(`${k}: ${props[k]}\n`);
  };
  if (!values.set && !values.scrub) { show(); return; }
  const out = needOut(values.out);
  if (values.scrub) doc.scrubMetadata();
  if (values.set) {
    const props: Record<string, string> = {};
    for (const kv of values.set) {
      const eq = kv.indexOf("=");
      const key = kv.slice(0, eq);
      if (eq < 1 || !(META_KEYS as readonly string[]).includes(key)) fail(`--set takes key=value with key one of: ${META_KEYS.join(", ")}`);
      props[key] = kv.slice(eq + 1);
    }
    doc.setCoreProperties(props);
  }
  await doc.toFile(out);
  show();
  process.stdout.write(`wrote ${out}\n`);
}

async function cmdResolve(mode: "accept" | "reject", argv: string[]): Promise<void> {
  const { values, positionals } = parse(argv, { out: { ...STRING, short: "o" }, ids: STRING, by: { ...STRING, multiple: true }, others: STRING });
  const out = needOut(values.out);
  const doc = await load(positionals[0]);
  const ids = values.ids?.split(/[\s,]+/).map((s) => s.replace(/^#/, "")).filter(Boolean);
  const other = mode === "accept" ? "reject" : "accept";
  if (values.others && values.others !== other && values.others !== "keep") fail(`--others must be ${other} or keep`);
  if (ids) {
    const known = new Set(listChanges(doc).flatMap((c) => c.ids));
    const unknown = ids.filter((i) => !known.has(i));
    if (unknown.length) fail(`no tracked change has w:id ${unknown.join(", ")} (run: revisions FILE)`);
  }
  const filter = { ...(ids ? { ids } : {}), ...(values.by ? { authors: values.by } : {}) };
  const n = mode === "accept" ? doc.acceptAllRevisions(filter).accepted : doc.rejectAllRevisions(filter).rejected;
  let rest = "";
  if (values.others === other) {
    const m = other === "accept" ? doc.acceptAllRevisions().accepted : doc.rejectAllRevisions().rejected;
    rest = `, ${other}ed the other ${m}`;
  }
  await doc.toFile(out);
  const left = listChanges(doc).length;
  process.stdout.write(`${mode}ed ${n} revision${n === 1 ? "" : "s"}${rest}; ${left} change${left === 1 ? "" : "s"} still tracked\nwrote ${out}\n`);
}

async function main(): Promise<void> {
  const [cmd, ...rest] = process.argv.slice(2);
  switch (cmd) {
    case "view": return cmdView(rest);
    case "revisions": return cmdRevisions(rest);
    case "comments": return cmdComments(rest);
    case "replace": return cmdReplace(rest);
    case "comment": return cmdComment(rest);
    case "edit": return cmdEdit(rest);
    case "accept": return cmdResolve("accept", rest);
    case "reject": return cmdResolve("reject", rest);
    case "metadata": return cmdMetadata(rest);
    case undefined: case "help": case "-h": case "--help":
      process.stdout.write(USAGE);
      return;
    default:
      fail(`unknown command "${cmd}"\n\n${USAGE}`);
  }
}

// Output piped into head, grep -m or sed -n closes early: stop quietly instead of crashing.
process.stdout.on("error", (e: NodeJS.ErrnoException) => {
  if (e.code === "EPIPE") process.exit(0);
  throw e;
});

main().catch((e: unknown) => fail((e as Error).stack ?? String(e)));
