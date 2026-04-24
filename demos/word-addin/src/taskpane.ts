/// <reference types="office-js" />

import { Docxtor } from '@eazydraft/docxtor';
import type { Patch, PatchResult } from '@eazydraft/docxtor';

// ─────────────────────────────────────────────────────────────────────────────
// UI helpers
// ─────────────────────────────────────────────────────────────────────────────

type LogLevel = 'info' | 'success' | 'error' | 'heading' | 'code';

const consoleEl = () => document.getElementById('console')!;
const projectionEl = () => document.getElementById('projection')!;
const statusEl = () => document.getElementById('status-bar')!;

function log(level: LogLevel, message: string): void {
  const el = consoleEl();
  const empty = el.querySelector('.console-empty');
  if (empty) empty.remove();

  const entry = document.createElement('div');
  entry.className = `log-entry log-${level}`;

  const ts = new Date().toLocaleTimeString('en-US', { hour12: false });
  const prefix = level === 'success' ? '✓' : level === 'error' ? '✗' : level === 'heading' ? '▶' : '·';

  if (level === 'code') {
    entry.innerHTML = `<span class="log-ts">${ts}</span><pre class="log-code">${escHtml(message)}</pre>`;
  } else {
    entry.innerHTML = `<span class="log-ts">${ts}</span><span class="log-prefix">${prefix}</span> ${escHtml(message)}`;
  }

  el.appendChild(entry);
  el.scrollTop = el.scrollHeight;
}

function logHeading(title: string): void { log('heading', title); }

function logPatchResults(results: PatchResult[]): void {
  for (const r of results) {
    const target = 'target' in r.patch ? (r.patch as { target: string }).target : '(document)';
    const desc = describePatch(r.patch);
    if (r.success) log('success', `[${target}] ${desc}`);
    else log('error', `[${target}] ${desc} — ${r.error ?? 'unknown error'}`);
  }
}

function describePatch(patch: Patch): string {
  switch (patch.type) {
    case 'replace_text':      return `replace_text: "${trunc(patch.find, 25)}" → "${trunc(patch.replace, 25)}"`;
    case 'insert_text':       return `insert_text ${patch.position} "${trunc(patch.anchor, 20)}"`;
    case 'delete_text':       return `delete_text: "${trunc(patch.find, 30)}"`;
    case 'append_text':       return `append_text: "${trunc(patch.text, 35)}"`;
    case 'replace_cell_text': return `replace_cell_text: "${trunc(patch.find, 20)}" → "${trunc(patch.replace, 20)}"`;
    case 'add_table':         return `add_table: ${patch.rows.length} row(s)${patch.headers ? ' + headers' : ''}${patch.afterId ? ` after ${patch.afterId}` : ''}`;
    case 'add_table_row':     return `add_table_row: after row ${patch.afterRow} [${patch.cells.map(c => trunc(c, 12)).join(', ')}]`;
    case 'delete_table_row':  return `delete_table_row: row ${patch.row}`;
    case 'apply_formatting':  return `apply_formatting: {${Object.keys(patch.formatting).join(', ')}}${patch.find ? ` on "${trunc(patch.find, 20)}"` : ' (whole para)'}`;
    case 'apply_style':       return `apply_style: "${patch.style}"${patch.find ? ` on "${trunc(patch.find, 20)}"` : ''}`;
    case 'add_header':        return `add_header (${patch.variant}): ${patch.paragraphs.length} para(s)`;
    case 'add_footer':        return `add_footer (${patch.variant}): ${patch.paragraphs.length} para(s)`;
    case 'remove_header':     return `remove_header (${patch.variant})`;
    case 'remove_footer':     return `remove_footer (${patch.variant})`;
  }
}

function trunc(s: string, n: number): string {
  return s.length <= n ? s : `${s.slice(0, n - 1)}…`;
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function clearConsole(): void {
  consoleEl().innerHTML = '<div class="console-empty">Console cleared.</div>';
}

function setStatus(message: string, state: 'idle' | 'working' | 'done' | 'error'): void {
  const el = statusEl();
  el.textContent = message;
  el.className = `status ${state}`;
}

function setProjection(text: string): void {
  projectionEl().textContent = text;
  document.getElementById('projection-panel')!.setAttribute('open', '');
}

function setButtonsDisabled(disabled: boolean): void {
  document.querySelectorAll<HTMLButtonElement>('.btn, .btn-accent').forEach(btn => {
    btn.disabled = disabled;
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Office.js helpers
// ─────────────────────────────────────────────────────────────────────────────

async function getDocOoxml(): Promise<string> {
  // Word.run's runtime does NOT propagate the batch return value — always capture via closure.
  let xml = '';
  await Word.run(async (context) => {
    const result = context.document.body.getOoxml();
    await context.sync();
    xml = result.value ?? '';
  });
  if (!xml) throw new Error('body.getOoxml() returned empty — make sure a document is open.');
  return xml;
}

async function setDocOoxml(xml: string): Promise<void> {
  await Word.run(async (context) => {
    context.document.body.insertOoxml(xml, Word.InsertLocation.replace);
    await context.sync();
  });
}

// Read the current doc, apply a transformation, write it back.
async function withDoc(fn: (doc: Docxtor) => void | Promise<void>): Promise<void> {
  setButtonsDisabled(true);
  setStatus('Reading document…', 'working');
  try {
    const xml = await getDocOoxml();
    const doc = Docxtor.fromOoxml(xml);
    await fn(doc);
    setStatus('Writing document…', 'working');
    await setDocOoxml(doc.toFlatOpc());
    setStatus('Done', 'done');
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    log('error', `Unexpected error: ${msg}`);
    setStatus('Error', 'error');
  } finally {
    setButtonsDisabled(false);
  }
}

async function applyDemo(
  patches: Patch[],
  opts?: { trackRevisions?: boolean; revisionAuthor?: string },
): Promise<void> {
  await withDoc((doc) => {
    const results = doc.applyPatches(patches, opts);
    const ok = results.filter(r => r.success).length;
    const fail = results.filter(r => !r.success).length;

    log('code', JSON.stringify(patches, null, 2));
    logPatchResults(results);

    if (opts?.trackRevisions) {
      log('info', `Tracked changes by "${opts.revisionAuthor ?? 'docxtor'}" — open Word's Review pane to accept/reject.`);
    }

    if (fail === 0) log('success', `${ok}/${results.length} patch(es) applied.`);
    else log('error', `${ok} applied, ${fail} failed. Run "Load Sample Doc" and try again.`);
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Demo: Load Sample Document
//
// Uses Word.js APIs directly to build document structure.
// Docxtor.create() produces a minimal DOCX without styles.xml/settings.xml,
// so inserting its toFlatOpc() output throws GeneralException. Building with
// Word.js APIs preserves the full document infrastructure, making the Flat OPC
// round-trip reliable for all subsequent docxtor demos.
//
// Resulting paragraph IDs (docxtor assigns in body order):
//   p0  — "Docxtor Demo Document"  (title)
//   p1  — Welcome / intro
//   p2  — "Text Editing: …"
//   p3  — "The quick brown fox…"   ← target for text-edit demos
//   p4  — "Important Notice: …"    ← target for format demo
//   tbl0 — capability table
//   p5  — Revision tracking blurb
// ─────────────────────────────────────────────────────────────────────────────

async function loadSampleDocument(): Promise<void> {
  logHeading('Load Sample Document');
  setButtonsDisabled(true);
  setStatus('Building sample document…', 'working');

  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // Clear all existing content (leaves one empty paragraph → p0)
      body.clear();
      await context.sync();

      // Insert paragraphs via Word.js API — guaranteed to produce a valid document
      const title = body.insertParagraph('Docxtor Demo Document', Word.InsertLocation.end);
      title.font.bold = true;
      title.font.size = 24;

      body.insertParagraph(
        'Welcome to the docxtor Word Add-in Demo. This showcase demonstrates surgical DOCX editing — ' +
        'replacing, inserting, and formatting text without breaking run-level formatting, styles, tables, or tracked changes.',
        Word.InsertLocation.end,
      );
      body.insertParagraph(
        'Text Editing: docxtor maps every element to a stable ID. The markdown projection is what an LLM receives to understand and edit the document.',
        Word.InsertLocation.end,
      );
      body.insertParagraph(
        'The quick brown fox jumps over the lazy dog. Replace, insert, or delete any substring with precision.',
        Word.InsertLocation.end,
      );

      const notice = body.insertParagraph(
        'Important Notice: This document was generated by the docxtor add-in demo. All patches target predictable paragraph IDs.',
        Word.InsertLocation.end,
      );
      notice.font.bold = true;

      body.insertTable(5, 3, Word.InsertLocation.end, [
        ['Patch Type', 'Target', 'What It Does'],
        ['replace_text', 'p3', 'Substitute exact text in any paragraph'],
        ['apply_formatting', 'p4', 'Bold, italic, color, font size — per run'],
        ['add_table', '(document)', 'Insert a new OOXML table anywhere'],
        ['add_header', '(section)', 'Create header with field codes'],
      ]);

      body.insertParagraph(
        'Revision tracking produces OOXML w:ins / w:del markup. Open the Review pane in Word to accept or reject individual changes — just as you would with a human editor.',
        Word.InsertLocation.end,
      );

      await context.sync();
    });

    log('success', 'Sample document built via Word.js API.');
    log('info', 'IDs: p0 (title) p1 (welcome) p2 (text-editing) p3 (fox) p4 (notice) tbl0 p5 (revision).');
    log('info', 'Run "Inspect Document" to verify IDs before running demos.');
    setStatus('Sample document ready', 'done');
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    log('error', `Failed to load sample: ${msg}`);
    setStatus('Error', 'error');
  } finally {
    setButtonsDisabled(false);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Demo: Inspect Document
// ─────────────────────────────────────────────────────────────────────────────

async function inspectDocument(): Promise<void> {
  logHeading('Inspect Document');
  setButtonsDisabled(true);
  setStatus('Projecting…', 'working');

  try {
    const xml = await getDocOoxml();
    const doc = Docxtor.fromOoxml(xml);

    const projection = doc.project('markdown', { includeFormatting: false, includeHeadersFooters: true });

    const body = doc.getBody();
    const paraCount = body.filter(e => e.kind === 'paragraph').length;
    const tableCount = body.filter(e => e.kind === 'table').length;

    log('info', `${paraCount} paragraphs, ${tableCount} tables, ${doc.getHeaders().size} headers, ${doc.getFooters().size} footers, ${doc.getStyles().length} styles.`);
    setProjection(projection);
    setStatus('Projection ready', 'done');
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    log('error', `Inspect failed: ${msg}`);
    setStatus('Error', 'error');
  } finally {
    setButtonsDisabled(false);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Demo: Text Editing  (all target p4 — the fox sentence)
// ─────────────────────────────────────────────────────────────────────────────

// ID map (body.clear() leaves NO empty paragraph — insertParagraph starts at p0):
//   p0 = "Docxtor Demo Document"  (title)
//   p1 = "Welcome…"
//   p2 = "Text Editing:…"
//   p3 = "The quick brown fox…"   ← text-edit demos
//   p4 = "Important Notice:…"     ← format demo
//   tbl0 = capability table
//   p5 = "Revision tracking…"

async function demoReplaceText(): Promise<void> {
  logHeading('Replace Text');
  log('info', 'replace_text — "quick brown fox" → "clever red fox" in p3.');
  await applyDemo([
    { type: 'replace_text', target: 'p3', find: 'quick brown fox', replace: 'clever red fox' },
  ]);
}

async function demoInsertText(): Promise<void> {
  logHeading('Insert Text');
  log('info', 'insert_text — " (with sub-run precision)" after "precision" in p3.');
  await applyDemo([
    { type: 'insert_text', target: 'p3', position: 'after', anchor: 'precision', text: ' (with sub-run precision)' },
  ]);
}

async function demoDeleteText(): Promise<void> {
  logHeading('Delete Text');
  log('info', 'delete_text — removing "lazy dog" from p3 (independent of other demos).');
  await applyDemo([
    { type: 'delete_text', target: 'p3', find: 'lazy dog' },
  ]);
}

async function demoFormatText(): Promise<void> {
  logHeading('Format Text');
  log('info', 'apply_formatting — bold+red+underline on "Important Notice:" in p4; large blue bold on p0 title.');
  await applyDemo([
    { type: 'apply_formatting', target: 'p4', find: 'Important Notice:', formatting: { bold: true, color: 'C7003A', underline: 'single' } },
    { type: 'apply_formatting', target: 'p0', formatting: { bold: true, fontSize: 40, color: '0078D4' } },
  ]);
}

// ─────────────────────────────────────────────────────────────────────────────
// Demo: Structure
// ─────────────────────────────────────────────────────────────────────────────

async function demoAddTable(): Promise<void> {
  logHeading('Add Table');
  log('info', 'add_table — inserting a team roster table after p4.');
  await applyDemo([
    {
      type: 'add_table',
      afterId: 'p4',
      headers: ['Name', 'Role', 'Start Date'],
      rows: [
        ['Alice Chen', 'Engineering Lead', '2023-01-15'],
        ['Bob Martínez', 'Product Manager', '2022-08-01'],
        ['Carol Osei', 'Design Lead', '2023-06-12'],
      ],
    },
  ]);
}

async function demoAddTableRow(): Promise<void> {
  logHeading('Add Table Row');
  log('info', 'add_table_row — appending a row to tbl0.');
  await applyDemo([
    { type: 'add_table_row', target: 'tbl0', afterRow: -1, cells: ['track_revisions', '(all patches)', 'OOXML w:ins / w:del tracked changes'] },
  ]);
}

async function demoDeleteTableRow(): Promise<void> {
  logHeading('Delete Table Row');
  log('info', 'delete_table_row — removing row index 1 from tbl0.');
  await applyDemo([
    { type: 'delete_table_row', target: 'tbl0', row: 1 },
  ]);
}

async function demoAppendContent(): Promise<void> {
  logHeading('Append Content');
  log('info', 'append_text — paragraph + markdown pipe-table (auto-converted to w:tbl).');
  await applyDemo([
    { type: 'append_text', text: 'Appended by docxtor:' },
    {
      type: 'append_text',
      text: ['| Language | Popularity | Trend', '| --- | --- | ---', '| TypeScript | ★★★★★ | ↑', '| Rust | ★★★★☆ | ↑↑', '| Python | ★★★★★ | →'].join('\n'),
    },
    { type: 'append_text', text: 'Markdown pipe-tables are auto-converted to real OOXML w:tbl elements.' },
  ]);
}

// ─────────────────────────────────────────────────────────────────────────────
// Demo: Headers & Footers
// ─────────────────────────────────────────────────────────────────────────────

async function demoAddHeader(): Promise<void> {
  logHeading('Add Header');
  log('info', 'add_header — CONFIDENTIAL banner on default section.');
  await applyDemo([{ type: 'add_header', variant: 'default', paragraphs: ['DOCXTOR DEMO · CONFIDENTIAL'] }]);
}

async function demoAddFooter(): Promise<void> {
  logHeading('Add / Update Footer');
  log('info', 'add_footer — PAGE, NUMPAGES, DATE field codes in default footer.');
  await applyDemo([{ type: 'add_footer', variant: 'default', paragraphs: ['docxtor demo  ·  Page {{PAGE}} of {{NUMPAGES}}  ·  {{DATE}}'] }]);
}

async function demoRemoveHeader(): Promise<void> {
  logHeading('Remove Header');
  log('info', 'remove_header — removing default header.');
  await applyDemo([{ type: 'remove_header', variant: 'default' }]);
}

async function demoRemoveFooter(): Promise<void> {
  logHeading('Remove Footer');
  log('info', 'remove_footer — removing default footer.');
  await applyDemo([{ type: 'remove_footer', variant: 'default' }]);
}

// ─────────────────────────────────────────────────────────────────────────────
// Demo: Tracked Changes
// ─────────────────────────────────────────────────────────────────────────────

async function demoTrackedChanges(): Promise<void> {
  logHeading('Tracked Changes Demo');
  log('info', '5 patches with trackRevisions: true — open Review pane to accept/reject redlines.');

  await applyDemo(
    [
      { type: 'replace_text', target: 'p3', find: 'quick brown fox', replace: 'diligent grey wolf' },
      { type: 'insert_text', target: 'p1', position: 'after', anchor: 'tracked changes', text: ' (accepted by reviewers)' },
      { type: 'apply_formatting', target: 'p0', formatting: { color: '107C10', italic: true } },
      { type: 'add_table_row', target: 'tbl0', afterRow: 0, cells: ['replace_cell_text', 'tbl0.r1.c0.p0', 'Replace text inside a table cell'] },
      { type: 'add_footer', variant: 'first', paragraphs: ['DRAFT — For Review Only · {{DATE}}'] },
    ],
    { trackRevisions: true, revisionAuthor: 'docxtor-agent' },
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Copy projection to clipboard
// ─────────────────────────────────────────────────────────────────────────────

async function copyProjection(): Promise<void> {
  const text = projectionEl().textContent ?? '';
  if (!text || text.startsWith('Run')) {
    log('error', 'No projection to copy. Run "Inspect Document" first.');
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    log('success', 'Projection copied to clipboard.');
  } catch {
    log('error', 'Clipboard access denied. Select the text manually.');
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Init
// ─────────────────────────────────────────────────────────────────────────────

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus('Open this add-in in Microsoft Word', 'error');
    log('error', `This add-in requires Microsoft Word. Current host: ${info.host ?? 'unknown'}`);
    return;
  }

  const bind = (id: string, fn: () => void) => {
    document.getElementById(id)?.addEventListener('click', fn);
  };

  bind('btn-load-sample', loadSampleDocument);
  bind('btn-inspect', inspectDocument);
  bind('btn-replace', demoReplaceText);
  bind('btn-insert', demoInsertText);
  bind('btn-delete', demoDeleteText);
  bind('btn-format', demoFormatText);
  bind('btn-add-table', demoAddTable);
  bind('btn-add-row', demoAddTableRow);
  bind('btn-del-row', demoDeleteTableRow);
  bind('btn-append', demoAppendContent);
  bind('btn-add-header', demoAddHeader);
  bind('btn-add-footer', demoAddFooter);
  bind('btn-rm-header', demoRemoveHeader);
  bind('btn-rm-footer', demoRemoveFooter);
  bind('btn-track-changes', demoTrackedChanges);
  bind('btn-clear-console', clearConsole);
  bind('btn-copy-projection', copyProjection);

  setStatus('Ready · load a sample or open any .docx', 'idle');
  log('info', 'docxtor Word Add-in ready. Click "Load Sample Doc" to start.');
  log('info', 'Text/table demos target: p3 (fox sentence), p4 (notice), tbl0 (capability table).');
});
