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
 *
 * Architecture:
 *   The LLM receives a compact projection (truncated to ~80 chars per
 *   paragraph) injected fresh each turn — never accumulated in history.
 *   It calls `get_paragraphs` to fetch full text of specific paragraphs,
 *   then `apply_patches` to edit. Conversation history stores only the
 *   user instructions and edit summaries, keeping token usage bounded.
 */

import * as readline from "node:readline";
import OpenAI from "openai";
import type { ChatCompletionMessageParam, ChatCompletionTool } from "openai/resources/chat/completions";
import { Docxtor } from "../src/index.js";
import type { Patch, PatchResult } from "../src/index.js";

// =============================================================================
// Config
// =============================================================================

const MODEL = process.env.OPENROUTER_MODEL ?? "openai/gpt-4o";
const COMPACT_TEXT_LENGTH = 80;
/** Max number of prior turns to keep in conversation history. */
const MAX_HISTORY_TURNS = 10;

const TOOLS: ChatCompletionTool[] = [
  {
    type: "function",
    function: {
      name: "get_paragraphs",
      description:
        "Fetch the full untruncated text of one or more paragraphs by their IDs. " +
        "Use this before writing patches so you can see the exact text to match against. " +
        "You can also pass table IDs to get all cell paragraph texts.",
      parameters: {
        type: "object",
        properties: {
          ids: {
            type: "array",
            items: { type: "string" },
            description:
              'Array of paragraph or table IDs to fetch, e.g. ["p5", "p17", "tbl0.r1.c0.p0"]',
          },
        },
        required: ["ids"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "apply_patches",
      description:
        "Apply an array of patches to the document. Call this once you have the full " +
        "text of relevant paragraphs and are ready to make edits. Include an explanation field.",
      parameters: {
        type: "object",
        properties: {
          explanation: {
            type: "string",
            description: "Brief description of what you're changing and why",
          },
          patches: {
            type: "array",
            description: "Array of patch objects to apply",
            items: {
              type: "object",
              properties: {
                type: {
                  type: "string",
                  enum: [
                    "replace_text",
                    "insert_text",
                    "delete_text",
                    "replace_cell_text",
                    "add_table_row",
                    "delete_table_row",
                  ],
                },
                target: { type: "string", description: "Paragraph or table ID" },
                find: { type: "string", description: "Exact substring to find" },
                replace: { type: "string", description: "Replacement text" },
                position: { type: "string", enum: ["before", "after"] },
                anchor: { type: "string", description: "Anchor text for insert_text" },
                text: { type: "string", description: "Text to insert" },
                occurrence: { type: "number", description: "1-based occurrence (default: 1)" },
                afterRow: { type: "number", description: "Row index to insert after (-1 for beginning)" },
                row: { type: "number", description: "Row index to delete (0-based)" },
                cells: { type: "array", items: { type: "string" }, description: "Cell texts for new row" },
              },
              required: ["type", "target"],
            },
          },
        },
        required: ["explanation", "patches"],
      },
    },
  },
];

const SYSTEM_PROMPT = `You are a precise document editor. You will receive:
1. A compact projection of the DOCX document (each paragraph truncated to ~${COMPACT_TEXT_LENGTH} chars)
2. The user's editing instruction

You have two tools:

1. **get_paragraphs** — Fetch the full untruncated text of paragraphs by ID. ALWAYS call this before writing patches so you can see the exact text to use in "find" fields.

2. **apply_patches** — Apply edits to the document. Call this after you've fetched and reviewed the relevant paragraph texts.

Patch types:
- replace_text: { type, target, find, replace, occurrence? }
- insert_text: { type, target, position ("before"|"after"), anchor, text, occurrence? }
- delete_text: { type, target, find, occurrence? }
- replace_cell_text: { type, target, find, replace, occurrence? }
- add_table_row: { type, target, afterRow, cells }
- delete_table_row: { type, target, row }

Rules:
- "find" / "anchor" must be an EXACT substring copied character-for-character from the full paragraph text.
- "occurrence" is 1-based, default 1.
- Paragraph IDs: p0, p1, tbl0.r0.c0.p0, hdr.default.p0, etc.
- Table IDs: tbl0, tbl1, etc.
- When replacing short words that could be substrings of other words, include surrounding context in "find" and "replace" (e.g. " the " not "the").
- If the user asks for something patches can't do, explain why in apply_patches with an empty patches array.

Workflow for each user request:
1. Read the compact projection to identify relevant paragraph IDs
2. Call get_paragraphs with those IDs to get full text
3. Call apply_patches with precise patches based on the full text`;

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

/**
 * A lightweight record of a past turn for conversation context.
 * Does NOT contain the projection — just the user instruction and outcome.
 */
interface TurnRecord {
  instruction: string;
  outcome: string; // short summary like "Replaced 3 instances of 'foo' with 'bar' in p5, p12, p17"
}

let doc: Docxtor;
let inputPath: string;
let defaultOutputPath: string;
const history: EditBatch[] = [];
const pastTurns: TurnRecord[] = [];

// =============================================================================
// Helpers
// =============================================================================

function getCompactProjection(): string {
  return doc.project("markdown", { maxTextLength: COMPACT_TEXT_LENGTH });
}

function getFullProjection(): string {
  return doc.project("markdown");
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
    const target = r.patch.target;
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
// Tool handlers
// =============================================================================

function handleGetParagraphs(ids: string[]): string {
  const results: Record<string, string> = {};
  for (const id of ids) {
    const para = doc.findParagraphById(id);
    if (para) {
      results[id] = para.text;
      continue;
    }
    const table = doc.findTableById(id);
    if (table) {
      for (const row of table.rows) {
        for (const cell of row.cells) {
          for (const p of cell.paragraphs) {
            results[p.id] = p.text;
          }
        }
      }
      continue;
    }
    results[id] = "(not found)";
  }
  return JSON.stringify(results, null, 2);
}

function handleApplyPatches(
  explanation: string,
  patches: Patch[],
  snapshotBuffer: Buffer,
  userInstruction: string,
): { results: PatchResult[]; succeeded: number; failed: number } {
  const validations = doc.validatePatches(patches);
  const invalid = validations.filter((v) => !v.valid);
  if (invalid.length > 0) {
    console.log(yellow(`\n  ${invalid.length} patch(es) failed validation:`));
    for (let i = 0; i < validations.length; i++) {
      if (!validations[i]!.valid) {
        console.log(red(`    - patch ${i + 1}: ${validations[i]!.error}`));
      }
    }
  }

  const results = doc.applyPatches(patches);
  const succeeded = results.filter((r) => r.success).length;
  const failed = results.filter((r) => !r.success).length;

  history.push({
    instruction: userInstruction,
    explanation,
    patches,
    results,
    snapshotBuffer,
  });

  return { results, succeeded, failed };
}

// =============================================================================
// Message construction
// =============================================================================

/**
 * Build the messages array for an API call.
 *
 * Structure:
 *   1. System prompt (static)
 *   2. Past turn summaries (lightweight, capped at MAX_HISTORY_TURNS)
 *   3. Current turn: projection + user instruction (injected fresh, not stored)
 *   4. Any tool call/response messages from the current turn's tool loop
 */
function buildMessages(
  userInstruction: string,
  currentTurnMessages: ChatCompletionMessageParam[],
): ChatCompletionMessageParam[] {
  const messages: ChatCompletionMessageParam[] = [
    { role: "system", content: SYSTEM_PROMPT },
  ];

  // Add lightweight summaries of recent turns for conversational context
  const recentTurns = pastTurns.slice(-MAX_HISTORY_TURNS);
  if (recentTurns.length > 0) {
    const summaryLines = recentTurns.map(
      (t) => `- User asked: "${t.instruction}" → ${t.outcome}`,
    );
    messages.push({
      role: "user",
      content: `Previous edits in this session:\n${summaryLines.join("\n")}`,
    });
    messages.push({
      role: "assistant",
      content: "Understood. I have context on the prior edits. What would you like to change next?",
    });
  }

  // Current turn: fresh projection + instruction
  const projection = getCompactProjection();
  messages.push({
    role: "user",
    content: `Current document (compact, each paragraph truncated to ~${COMPACT_TEXT_LENGTH} chars):\n\n${projection}\n\n---\n\nUser instruction: ${userInstruction}`,
  });

  // Any in-progress tool call messages from the current turn
  messages.push(...currentTurnMessages);

  return messages;
}

// =============================================================================
// LLM with tool loop
// =============================================================================

async function callLLM(
  client: OpenAI,
  userMessage: string,
): Promise<void> {
  const snapshotBuffer = await doc.toBuffer();
  let editApplied = false;
  let editExplanation = "";
  let editSummary = "";

  // Messages accumulated within this single turn's tool loop only
  const turnMessages: ChatCompletionMessageParam[] = [];

  const MAX_ROUNDS = 10;
  for (let round = 0; round < MAX_ROUNDS; round++) {
    const messages = buildMessages(userMessage, turnMessages);

    const response = await client.chat.completions.create({
      model: MODEL,
      messages,
      tools: TOOLS,
      temperature: 0,
    });

    const choice = response.choices[0]!;
    const message = choice.message;

    // Add to current turn's messages (NOT to persistent history)
    turnMessages.push(message);

    if (message.tool_calls && message.tool_calls.length > 0) {
      for (const toolCall of message.tool_calls) {
        const args = JSON.parse(toolCall.function.arguments);

        if (toolCall.function.name === "get_paragraphs") {
          const ids: string[] = args.ids ?? [];
          process.stdout.write(dim(`  Fetching ${ids.length} paragraph(s): ${ids.join(", ")}\n`));
          const result = handleGetParagraphs(ids);

          turnMessages.push({
            role: "tool",
            tool_call_id: toolCall.id,
            content: result,
          });
        } else if (toolCall.function.name === "apply_patches") {
          const explanation: string = args.explanation ?? "";
          const patches: Patch[] = args.patches ?? [];

          console.log(`\n${bold("Agent:")} ${explanation}`);

          if (patches.length === 0) {
            console.log(dim("  (no patches to apply)"));
            editExplanation = explanation;
            editSummary = "No patches needed";
            turnMessages.push({
              role: "tool",
              tool_call_id: toolCall.id,
              content: JSON.stringify({ success: true, applied: 0 }),
            });
          } else {
            const { results, succeeded, failed } = handleApplyPatches(
              explanation,
              patches,
              snapshotBuffer,
              userMessage,
            );

            console.log();
            printPatches(results);
            console.log();

            if (failed > 0) {
              console.log(yellow(`${succeeded}/${results.length} patches applied (${failed} failed)`));
            } else {
              console.log(green(`${succeeded}/${results.length} patches applied`));
            }

            editApplied = true;
            editExplanation = explanation;

            const targets = [...new Set(results.filter((r) => r.success).map((r) => r.patch.target))];
            editSummary = `${succeeded}/${results.length} patches applied (${explanation}) in ${targets.join(", ")}`;

            const summary = results.map((r) => ({
              target: r.patch.target,
              type: r.patch.type,
              success: r.success,
              error: r.error,
            }));
            turnMessages.push({
              role: "tool",
              tool_call_id: toolCall.id,
              content: JSON.stringify({ success: true, applied: succeeded, failed, details: summary }),
            });
          }
        }
      }
      continue;
    }

    // Plain text response (done with tools)
    if (message.content) {
      if (!editApplied) {
        console.log(`\n${bold("Agent:")} ${message.content}`);
        editSummary = editSummary || message.content.slice(0, 120);
      }
    }
    break;
  }

  // Record a lightweight summary for future turns
  pastTurns.push({
    instruction: userMessage,
    outcome: editSummary || editExplanation || "(no changes)",
  });
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
  console.log(getFullProjection());
  console.log();
}

async function handleUndo(): Promise<void> {
  if (history.length === 0) {
    console.log(yellow("Nothing to undo."));
    return;
  }
  const last = history.pop()!;
  doc = await Docxtor.load(last.snapshotBuffer);
  // Also remove the corresponding turn record
  if (pastTurns.length > 0) {
    pastTurns.pop();
  }
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

  if (!process.env.OPENROUTER_API_KEY) {
    console.error(red("Error: OPENROUTER_API_KEY environment variable is not set."));
    process.exit(1);
  }

  const client = new OpenAI({
    baseURL: "https://openrouter.ai/api/v1",
    apiKey: process.env.OPENROUTER_API_KEY,
  });

  console.log(dim(`Loading ${inputPath}...`));
  doc = await Docxtor.load(inputPath);

  const body = doc.getBody();
  const paraCount = body.filter((e) => e.kind === "paragraph").length;
  const tableCount = body.filter((e) => e.kind === "table").length;
  const headerCount = doc.getHeaders().size;
  const footerCount = doc.getFooters().size;

  const compactProjection = getCompactProjection();
  const estimatedTokens = Math.ceil(compactProjection.length / 4);

  console.log();
  console.log(bold("docxtor") + dim(" interactive editor"));
  console.log(
    dim(
      `Loaded: ${paraCount} paragraphs, ${tableCount} tables, ${headerCount} headers, ${footerCount} footers`,
    ),
  );
  console.log(dim(`Model: ${MODEL}`));
  console.log(dim(`Compact projection: ~${estimatedTokens} tokens`));
  console.log();
  console.log(dim("Commands: /save [path]  /show  /undo  /history  /quit"));
  console.log(dim("Type an editing instruction to get started."));
  console.log();

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
              console.log(dim("Commands: /save [path]  /show  /undo  /history  /quit"));
          }

          prompt();
          return;
        }

        process.stdout.write(dim("Thinking...\r"));
        await callLLM(client, trimmed);
        process.stdout.write("            \r");
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
  console.error(red(`Fatal: ${err instanceof Error ? err.message : String(err)}`));
  process.exit(1);
});
