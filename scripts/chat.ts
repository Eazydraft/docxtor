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

7. add_table — Add a new table to the document body (or a header/footer).
   { "type": "add_table", "headers": ["col1", "col2"], "rows": [["val1", "val2"], ...] }
   "headers" is optional. "afterId" is optional — if specified, the table is inserted after that element ID; otherwise it's appended to the end.
   { "type": "add_table", "afterId": "<element_id>", "headers": ["col1", "col2"], "rows": [["val1", "val2"]] }
   To add a table inside a header/footer, use "container":
   { "type": "add_table", "headers": ["A", "B"], "rows": [["1", "2"]], "container": { "type": "header", "variant": "default" } }

8. append_text — Append text (or a markdown table) to the end of the document body (or a header/footer).
   { "type": "append_text", "text": "New paragraph text" }
   Supports markdown table syntax — pipe-delimited rows with a separator line are auto-converted to a real table:
   { "type": "append_text", "text": "| Name | Age |\\n| --- | --- |\\n| Alice | 30 |" }
   Supports {{FIELD}} placeholders for dynamic content: {{PAGE}}, {{NUMPAGES}}, {{DATE}}, {{TIME}}, {{FILENAME}}.
   Example: { "type": "append_text", "text": "Page {{PAGE}} of {{NUMPAGES}}" }
   To append to a header/footer instead of the body, use "container":
   { "type": "append_text", "text": "Updated: {{DATE}}", "container": { "type": "footer", "variant": "default" } }

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

10. apply_style — Apply a pre-existing style to a paragraph or substring.
   { "type": "apply_style", "target": "<paragraph_id>", "style": "<style_id>" }
   Omit "find" to apply a paragraph style (e.g. "Heading1", "Title").
   Include "find" to apply a character style (e.g. "Strong", "Emphasis") to matched text:
   { "type": "apply_style", "target": "<paragraph_id>", "find": "<exact_text>", "style": "<character_style_id>", "occurrence": 1 }
   The style must already exist in the document (listed under "Available Styles" in the projection).

11. add_header — Create a new header in the document's last section.
   { "type": "add_header", "variant": "default"|"first"|"even", "paragraphs": ["Header text", "Page {{PAGE}}"] }
   "variant" controls which pages: "default" (all/odd pages), "first" (first page only), "even" (even pages).
   Paragraph text supports {{FIELD}} placeholders: {{PAGE}}, {{NUMPAGES}}, {{DATE}}, {{TIME}}, {{FILENAME}}.

12. add_footer — Create a new footer in the document's last section.
   { "type": "add_footer", "variant": "default"|"first"|"even", "paragraphs": ["Page {{PAGE}} of {{NUMPAGES}}"] }
   Same variant and field placeholder support as add_header.

13. remove_header — Remove a header from the document's last section.
   { "type": "remove_header", "variant": "default"|"first"|"even" }

14. remove_footer — Remove a footer from the document's last section.
   { "type": "remove_footer", "variant": "default"|"first"|"even" }

Inline formatting in the projection:
- The projection uses GitHub-flavored Markdown markers to show existing formatting: **bold**, _italic_, ~~strikethrough~~, <u>underline</u>, **_bold+italic_**.
- These markers are NOT part of the actual document text. They are annotations showing how text is formatted.

Using formatting markers in "find" and "anchor":
- You may include GFM markers in "find" and "anchor" fields for readability, but they are automatically stripped before matching. The match is always against the plain text.
- Example: "find": "**important**" matches the plain text "important" (which happens to be bold).

Using formatting markers in "replace" and "text":
- To apply inline formatting when replacing or inserting text, use GFM markers in the "replace" or "text" fields.
- Example: "replace": "**bold text**" will create bold text in the document.
- Example: "text": "This is _italic_ and **bold**" will create a mix of plain, italic, and bold runs.
- Supported markers: **bold**, _italic_, ~~strikethrough~~, <u>underline</u>, and combinations like **_bold+italic_**.
- This is a convenient alternative to using a separate apply_formatting patch.

Rules:
- "find" must be an EXACT substring of the paragraph text (ignoring any formatting markers). Copy it character-for-character from the document.
- "occurrence" is 1-based. Default is 1 (first occurrence). Use higher numbers only if there are duplicates within the same paragraph.
- Paragraph IDs look like: p0, p1, tbl0.r0.c0.p0, hdr.default.p0, ftr.default.p0, etc.
- Table IDs look like: tbl0, tbl1, etc.
- When replacing text that could be a substring of other words, include surrounding characters in "find" and "replace" to disambiguate (e.g., use " the " instead of "the" to avoid matching "other").
- For formatting changes, prefer using inline markers in replace/text fields for simple bold/italic/strikethrough/underline. Use apply_formatting for other properties (color, font size, etc.) or when you need to format existing text without replacing it.
- For structural changes like "make this a heading", use apply_style with the appropriate paragraph style.
- For headers/footers, use add_header/add_footer to create new ones. Existing header/footer paragraphs can be edited with replace_text, insert_text, etc. using their paragraph IDs (e.g., "hdr.default.p0").
- Use {{PAGE}} and {{NUMPAGES}} in header/footer text for page numbers (e.g., "Page {{PAGE}} of {{NUMPAGES}}").
15. add_list — Add a new list (bullet, ordered, or multilevel) to the document body.
   { "type": "add_list", "listType": "bullet"|"ordered"|"multilevel", "items": ["Item 1", "Item 2"] }
   Items can be strings or objects with text and level: { "text": "Sub-item", "level": 1 }
   Optional "afterId" to insert after a specific element. Optional "container" for header/footer.
   To continue an existing list's numbering, pass its "numId" (shown as {list:N} in the projection):
   { "type": "add_list", "listType": "multilevel", "numId": "1", "items": [{"text": "New section", "level": 0}, {"text": "Sub-section", "level": 1}], "afterId": "p10" }
   Without "numId", a new numbering definition is created (counters start at 1).

16. add_list_item — Add a new item to an existing list, before or after a target list paragraph.
   { "type": "add_list_item", "target": "<list_paragraph_id>", "position": "before"|"after", "text": "New item" }
   Optional "level" to set nesting depth (default: same as target).

17. remove_list_item — Remove a list item paragraph from the document.
   { "type": "remove_list_item", "target": "<list_paragraph_id>" }

18. convert_to_list — Convert existing paragraphs into a list.
   { "type": "convert_to_list", "targets": ["p0", "p1", "p2"], "listType": "bullet"|"ordered"|"multilevel" }
   Optional "level" to set all items to a specific nesting depth.

19. remove_list — Strip list formatting from paragraphs (keep the text, remove numbering).
   { "type": "remove_list", "targets": ["p0", "p1", "p2"] }

20. set_list_properties — Change the level or list type of an existing list item.
   { "type": "set_list_properties", "target": "<list_paragraph_id>", "level": 1 }
   { "type": "set_list_properties", "target": "<list_paragraph_id>", "listType": "ordered" }

List items in the projection:
- List items show: "[p0] {list:N LX} prefix text" where N is the numId and X is the indentation level.
- Example: "[p10] {list:5 L0} 1. Position and Duties" — this is numId=5, level 0 (top-level heading).
- Example: "[p11] {list:6 L1}   1.1. The Company hereby..." — this is numId=6, level 1 (sub-item).
- Items sharing the same numId belong to the same list and share counters.
- IMPORTANT: A document may use DIFFERENT numIds for top-level headings vs sub-items (e.g. {list:5} for sections, {list:6} for sub-sections). When adding a new section, you must add items using the CORRECT numId for each level. For example, to add "Section 6" after "Section 5", use numId from the section headings (list:5), and for sub-items use the sub-item numId (list:6).
- To add items to an existing list, pass "numId" in add_list. Without "numId", a new numbering definition is created (counters restart at 1).
- You can use add_list_item to add a single item relative to an existing list item (it inherits the numId automatically).
- List paragraphs have the same IDs as regular paragraphs (p0, p1, etc.) and can be edited with replace_text, delete_text, etc.

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
  return doc.project("markdown", {
    includePageBreaks: true,
    includeFormatting: true,
  });
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
        desc = `add table (${r.patch.rows.length} rows, ${r.patch.headers ? "with headers" : "no headers"})${r.patch.afterId ? ` after ${r.patch.afterId}` : ""}${r.patch.container ? ` in ${r.patch.container.type}.${r.patch.container.variant}` : ""}`;
        break;
      case "append_text":
        desc = `append "${truncate(r.patch.text, 40)}"${r.patch.container ? ` to ${r.patch.container.type}.${r.patch.container.variant}` : ""}`;
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
      case "add_header":
        desc = `add ${r.patch.variant} header (${r.patch.paragraphs.length} paragraphs)`;
        break;
      case "add_footer":
        desc = `add ${r.patch.variant} footer (${r.patch.paragraphs.length} paragraphs)`;
        break;
      case "remove_header":
        desc = `remove ${r.patch.variant} header`;
        break;
      case "remove_footer":
        desc = `remove ${r.patch.variant} footer`;
        break;
      case "add_list":
        desc = `add ${r.patch.listType} list (${r.patch.items.length} items)${r.patch.afterId ? ` after ${r.patch.afterId}` : ""}`;
        break;
      case "add_list_item":
        desc = `add list item "${truncate(r.patch.text, 30)}" ${r.patch.position} target`;
        break;
      case "remove_list_item":
        desc = `remove list item`;
        break;
      case "convert_to_list":
        desc = `convert ${r.patch.targets.length} paragraphs to ${r.patch.listType} list`;
        break;
      case "remove_list":
        desc = `remove list from ${r.patch.targets.length} paragraphs`;
        break;
      case "set_list_properties": {
        const parts: string[] = [];
        if (r.patch.listType) parts.push(`type=${r.patch.listType}`);
        if (r.patch.level != null) parts.push(`level=${r.patch.level}`);
        desc = `set list properties {${parts.join(", ")}}`;
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
