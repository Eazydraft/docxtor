/**
 * Interactive DOCX editing CLI powered by an LLM.
 *
 * Usage:
 *   npx tsx scripts/chat.ts <input.docx>
 *
 * Requires OPENROUTER_API_KEY environment variable.
 * Optionally set OPENROUTER_MODEL (default: openai/gpt-4o).
 *
 * Commands:
 *   /save [path]   — save the current document (default: <input>.edited.docx)
 *   /show          — print the current document projection
 *   /undo          — undo the last batch of patches
 *   /history       — show edit history
 *   /quit          — save and exit
 *
 * Anything else is sent to the LLM as an editing instruction.
 */

import * as readline from "node:readline";
import OpenAI from "openai";
import { Docxtor } from "../src/index.js";
import type { Patch, PatchResult } from "../src/index.js";

// =============================================================================
// Config
// =============================================================================

const MODEL = process.env.OPENROUTER_MODEL ?? "openai/gpt-4o";

const SYSTEM_PROMPT = `You are a precise document editor. You receive a DOCX document as a structured projection and the user asks you to make edits.

You MUST respond with a JSON object with two fields:
- "explanation": a brief description of what you're changing and why
- "patches": an array of patch objects to apply

Patch types:

1. replace_text — Replace a substring in a paragraph.
   { "type": "replace_text", "target": "<paragraph_id>", "find": "<exact_text>", "replace": "<new_text>", "occurrence": 1 }

2. insert_text — Insert text before or after an anchor string.
   { "type": "insert_text", "target": "<paragraph_id>", "position": "before"|"after", "anchor": "<exact_text>", "text": "<text_to_insert>", "occurrence": 1 }

3. delete_text — Delete a substring from a paragraph.
   { "type": "delete_text", "target": "<paragraph_id>", "find": "<exact_text>", "occurrence": 1 }

4. replace_cell_text — Replace text in a table cell paragraph.
   { "type": "replace_cell_text", "target": "<cell_paragraph_id>", "find": "<exact_text>", "replace": "<new_text>", "occurrence": 1 }

5. add_table_row — Add a row to a table.
   { "type": "add_table_row", "target": "<table_id>", "afterRow": <row_index>, "cells": ["cell1", "cell2", ...] }

6. delete_table_row — Delete a row from a table.
   { "type": "delete_table_row", "target": "<table_id>", "row": <row_index> }

7. add_table — Add a new table to the document.
   { "type": "add_table", "headers": ["col1", "col2"], "rows": [["val1", "val2"], ...] }
   "headers" is optional. "afterId" is optional — if specified, the table is inserted after that element ID; otherwise it's appended to the end.
   { "type": "add_table", "afterId": "<element_id>", "headers": ["col1", "col2"], "rows": [["val1", "val2"]] }

8. append_text — Append text (or a markdown table) to the end of the document.
   { "type": "append_text", "text": "New paragraph text" }
   Supports markdown table syntax — pipe-delimited rows with a separator line are auto-converted to a real table:
   { "type": "append_text", "text": "| Name | Age |\\n| --- | --- |\\n| Alice | 30 |" }

9. apply_formatting — Apply inline formatting to a substring or entire paragraph.
   { "type": "apply_formatting", "target": "<paragraph_id>", "find": "<exact_text>", "occurrence": 1, "formatting": { ... } }
   Omit "find" to format the entire paragraph. The "formatting" object supports:
   - bold: true/false/null (true=add, false/null=remove)
   - italic: true/false/null
   - underline: "single"/"double"/"wave"/null (string=add, null=remove)
   - strikethrough: true/false/null
   - fontSize: number/null (in half-points, e.g. 24 = 12pt)
   - fontFamily: "Arial"/null (string=set, null=remove)
   - color: "FF0000"/null (hex color, null=remove)
   - highlight: "yellow"/"green"/"cyan"/null (color name, null=remove)
   Only properties you include are changed; omitted properties are left unchanged.

8. apply_style — Apply a pre-existing style to a paragraph or substring.
   { "type": "apply_style", "target": "<paragraph_id>", "style": "<style_id>" }
   Omit "find" to apply a paragraph style (e.g. "Heading1", "Title").
   Include "find" to apply a character style (e.g. "Strong", "Emphasis") to matched text:
   { "type": "apply_style", "target": "<paragraph_id>", "find": "<exact_text>", "style": "<character_style_id>", "occurrence": 1 }
   The style must already exist in the document (listed under "Available Styles" in the projection).

Rules:
- "find" must be an EXACT substring of the paragraph text. Copy it character-for-character from the document.
- "occurrence" is 1-based. Default is 1 (first occurrence). Use higher numbers only if there are duplicates within the same paragraph.
- Paragraph IDs look like: p0, p1, tbl0.r0.c0.p0, hdr.default.p0, etc.
- Table IDs look like: tbl0, tbl1, etc.
- When replacing text that could be a substring of other words, include surrounding characters in "find" and "replace" to disambiguate (e.g., use " the " instead of "the" to avoid matching "other").
- For formatting changes, prefer apply_formatting over apply_style unless the user specifically asks for a style.
- For structural changes like "make this a heading", use apply_style with the appropriate paragraph style.
- If the user asks for something you can't do with patches (e.g., add images), explain that in "explanation" and return an empty "patches" array.
- ALWAYS respond with valid JSON. No markdown fences, no extra text.`;

// =============================================================================
// State
// =============================================================================

interface EditBatch {
  instruction: string;
  explanation: string;
  patches: Patch[];
  results: PatchResult[];
  snapshotBuffer: Buffer;
}

let doc: Docxtor;
let inputPath: string;
let defaultOutputPath: string;
const history: EditBatch[] = [];
const conversationMessages: OpenAI.Chat.ChatCompletionMessageParam[] = [];

// =============================================================================
// Helpers
// =============================================================================

function getProjection(): string {
  return doc.project("markdown", { includePageBreaks: true });
}

function dim(s: string): string {
  return `\x1b[2m${s}\x1b[0m`;
}
function bold(s: string): string {
  return `\x1b[1m${s}\x1b[0m`;
}
function green(s: string): string {
  return `\x1b[32m${s}\x1b[0m`;
}
function red(s: string): string {
  return `\x1b[31m${s}\x1b[0m`;
}
function yellow(s: string): string {
  return `\x1b[33m${s}\x1b[0m`;
}
function cyan(s: string): string {
  return `\x1b[36m${s}\x1b[0m`;
}

function printPatches(results: PatchResult[]): void {
  for (const r of results) {
    const status = r.success ? green("OK") : red("FAIL");
    const target = "target" in r.patch ? r.patch.target : "(document)";
    let desc: string;
    switch (r.patch.type) {
      case "replace_text":
      case "replace_cell_text":
        desc = `replace "${truncate(r.patch.find, 30)}" -> "${truncate(r.patch.replace, 30)}"`;
        break;
      case "insert_text":
        desc = `insert "${truncate(r.patch.text, 30)}" ${r.patch.position} "${truncate(r.patch.anchor, 30)}"`;
        break;
      case "delete_text":
        desc = `delete "${truncate(r.patch.find, 30)}"`;
        break;
      case "add_table_row":
        desc = `add row after ${r.patch.afterRow} [${r.patch.cells.map((c) => truncate(c, 15)).join(", ")}]`;
        break;
      case "delete_table_row":
        desc = `delete row ${r.patch.row}`;
        break;
      case "add_table":
        desc = `add table (${r.patch.rows.length} rows, ${r.patch.headers ? "with headers" : "no headers"})${r.patch.afterId ? ` after ${r.patch.afterId}` : ""}`;
        break;
      case "append_text":
        desc = `append "${truncate(r.patch.text, 40)}"`;
        break;
      case "apply_formatting": {
        const props = Object.keys(r.patch.formatting).join(", ");
        const findHint = r.patch.find
          ? ` on "${truncate(r.patch.find, 20)}"`
          : " (whole paragraph)";
        desc = `format {${props}}${findHint}`;
        break;
      }
      case "apply_style": {
        const findHint = r.patch.find
          ? ` on "${truncate(r.patch.find, 20)}"`
          : "";
        desc = `style "${r.patch.style}"${findHint}`;
        break;
      }
    }
    console.log(`  ${status} [${target}] ${desc}`);
    if (!r.success && r.error) {
      console.log(`       ${red(r.error)}`);
    }
  }
}

function truncate(s: string, max: number): string {
  if (s.length <= max) return s;
  return s.slice(0, max - 1) + "\u2026";
}

// =============================================================================
// LLM
// =============================================================================

async function callLLM(
  client: OpenAI,
  userMessage: string,
): Promise<{ explanation: string; patches: Patch[] }> {
  // Always include the latest projection in the user message
  const projection = getProjection();
  const fullMessage = `Current document:\n\n${projection}\n\n---\n\nUser instruction: ${userMessage}`;

  conversationMessages.push({ role: "user", content: fullMessage });

  const response = await client.chat.completions.create({
    model: MODEL,
    messages: [
      { role: "system", content: SYSTEM_PROMPT },
      ...conversationMessages,
    ],
    temperature: 0,
    response_format: { type: "json_object" },
  });

  const content = response.choices[0]?.message?.content ?? "{}";
  conversationMessages.push({ role: "assistant", content });

  // remove "```" and/or "json" if the model wrapped the JSON in a code block
  const cleaned = content
    .trim()
    .replace(/^```json\s*/, "")
    .replace(/```\s*$/, "");

  const parsed = JSON.parse(cleaned);
  return {
    explanation: parsed.explanation ?? "",
    patches: Array.isArray(parsed.patches) ? parsed.patches : [],
  };
}

// =============================================================================
// Commands
// =============================================================================

async function handleSave(args: string): Promise<void> {
  const path = args.trim() || defaultOutputPath;
  await doc.toFile(path);
  console.log(green(`Saved to ${path}`));
}

function handleShow(): void {
  console.log();
  console.log(getProjection());
  console.log();
}

async function handleUndo(): Promise<void> {
  if (history.length === 0) {
    console.log(yellow("Nothing to undo."));
    return;
  }
  const last = history.pop()!;
  // Reload from snapshot
  doc = await Docxtor.load(last.snapshotBuffer);
  // Also pop the conversation pair (user + assistant)
  conversationMessages.pop();
  conversationMessages.pop();
  console.log(green(`Undid: ${last.explanation}`));
}

function handleHistory(): void {
  if (history.length === 0) {
    console.log(dim("No edits yet."));
    return;
  }
  for (let i = 0; i < history.length; i++) {
    const h = history[i]!;
    const succeeded = h.results.filter((r) => r.success).length;
    console.log(
      `  ${dim(`${i + 1}.`)} ${h.explanation} ${dim(`(${succeeded}/${h.patches.length} patches)`)}`,
    );
  }
}

// =============================================================================
// Main
// =============================================================================

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.error("Usage: npx tsx scripts/chat.ts <input.docx>");
    process.exit(1);
  }

  inputPath = args[0]!;
  defaultOutputPath = inputPath.replace(/\.docx$/i, ".edited.docx");

  // Validate API key
  if (!process.env.OPENROUTER_API_KEY) {
    console.error(
      red("Error: OPENROUTER_API_KEY environment variable is not set."),
    );
    process.exit(1);
  }

  const client = new OpenAI({
    baseURL: "https://openrouter.ai/api/v1",
    apiKey: process.env.OPENROUTER_API_KEY,
  });

  // Load document
  console.log(dim(`Loading ${inputPath}...`));
  doc = await Docxtor.load(inputPath);

  const body = doc.getBody();
  const paraCount = body.filter((e) => e.kind === "paragraph").length;
  const tableCount = body.filter((e) => e.kind === "table").length;
  const headerCount = doc.getHeaders().size;
  const footerCount = doc.getFooters().size;
  const styleCount = doc.getStyles().length;

  console.log();
  console.log(bold("docxtor") + dim(" interactive editor"));
  console.log(
    dim(
      `Loaded: ${paraCount} paragraphs, ${tableCount} tables, ${headerCount} headers, ${footerCount} footers, ${styleCount} styles`,
    ),
  );
  console.log(dim(`Model: ${MODEL}`));
  console.log();
  console.log(dim("Commands: /save [path]  /show  /undo  /history  /quit"));
  console.log(dim("Type an editing instruction to get started."));
  console.log();

  // REPL
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  const prompt = (): void => {
    rl.question(cyan("you> "), async (input) => {
      const trimmed = input.trim();
      if (!trimmed) {
        prompt();
        return;
      }

      try {
        // Handle commands
        if (trimmed.startsWith("/")) {
          const [cmd, ...rest] = trimmed.split(" ");
          const cmdArgs = rest.join(" ");

          switch (cmd) {
            case "/save":
              await handleSave(cmdArgs);
              break;
            case "/show":
              handleShow();
              break;
            case "/undo":
              await handleUndo();
              break;
            case "/history":
              handleHistory();
              break;
            case "/quit":
            case "/exit":
              await handleSave(cmdArgs || defaultOutputPath);
              console.log(dim("Goodbye."));
              rl.close();
              return;
            default:
              console.log(yellow(`Unknown command: ${cmd}`));
              console.log(
                dim("Commands: /save [path]  /show  /undo  /history  /quit"),
              );
          }

          prompt();
          return;
        }

        // Take a snapshot before edits (for undo)
        const snapshotBuffer = await doc.toBuffer();

        // Call LLM
        process.stdout.write(dim("Thinking...\r"));
        const { explanation, patches } = await callLLM(client, trimmed);
        // Clear "Thinking..."
        process.stdout.write("            \r");

        console.log(`\n${bold("Agent:")} ${explanation}`);

        if (patches.length === 0) {
          console.log(dim("  (no patches to apply)"));
        } else {
          // Validate first
          const validations = doc.validatePatches(patches);
          const invalid = validations.filter((v) => !v.valid);
          if (invalid.length > 0) {
            console.log(
              yellow(`\n  ${invalid.length} patch(es) failed validation:`),
            );
            for (let i = 0; i < validations.length; i++) {
              if (!validations[i]!.valid) {
                console.log(
                  red(`    - patch ${i + 1}: ${validations[i]!.error}`),
                );
              }
            }
          }

          // Apply
          const results = doc.applyPatches(patches, {
            trackRevisions: true,
            revisionAuthor: "docxtor-agent",
          });
          const succeeded = results.filter((r) => r.success).length;
          const failed = results.filter((r) => !r.success).length;

          console.log();
          printPatches(results);
          console.log();

          if (failed > 0) {
            console.log(
              yellow(
                `${succeeded}/${results.length} patches applied (${failed} failed)`,
              ),
            );
          } else {
            console.log(
              green(`${succeeded}/${results.length} patches applied`),
            );
          }

          // Record in history
          history.push({
            instruction: trimmed,
            explanation,
            patches,
            results,
            snapshotBuffer,
          });
        }

        console.log();
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        console.error(red(`Error: ${msg}`));
        console.log();
      }

      prompt();
    });
  };

  prompt();
}

main().catch((err) => {
  console.error(
    red(`Fatal: ${err instanceof Error ? err.message : String(err)}`),
  );
  process.exit(1);
});
