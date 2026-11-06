<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/Eazydraft/docxtor/main/assets/logo-dark.svg">
  <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/Eazydraft/docxtor/main/assets/logo-light.svg">
  <img src="https://raw.githubusercontent.com/Eazydraft/docxtor/main/assets/logo-light.svg" alt="docxtor" width="360">
</picture>

[![npm](https://img.shields.io/npm/v/@lucio-ai/docxtor)](https://www.npmjs.com/package/@lucio-ai/docxtor)
[![CI](https://github.com/Eazydraft/docxtor/actions/workflows/ci.yml/badge.svg)](https://github.com/Eazydraft/docxtor/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**Precise, revision-aware editing for Microsoft Word documents. Built for AI agents.**

docxtor gives AI agents and applications a structured interface for reading, editing, and reviewing `.docx` files.

Create tracked changes, add comments anchored to specific passages, resolve revisions individually, and modify document content without rewriting the entire file.

Unlike approaches that convert documents to plain text and reconstruct them, docxtor applies targeted patches directly to the underlying document XML. Formatting, styles, numbering, headers, footers, tables, and everything else outside the edited regions are preserved.

Works with Claude, Codex, Cursor, Gemini CLI, GitHub Copilot, and any agent that supports skills.

## What can you do?

Ask your agent to make changes in plain language:

> Change the payment terms to net 60 days and the delivery window to ten business days. Mark both as tracked changes and save the result as a new file.

> Review the counterparty's revisions. Reject changes involving monetary amounts or dates, accept everything else, and summarize what was rejected.

> Add a comment to the liability clause asking the other party to reconsider the proposed cap.

The resulting document opens directly in Microsoft Word, with native tracked changes and comments.

## Installation

### AI agents

docxtor ships with a self-contained agent skill that teaches compatible agents how to inspect and edit Word documents.

**Claude Cowork / Claude.ai**

1. Download `docxtor-skill.zip` from the [latest release](https://github.com/Eazydraft/docxtor/releases/latest).
2. Upload it under Customize → Skills.
3. Enable code execution.

**Codex, Cursor, Gemini CLI, GitHub Copilot, and other skill-compatible agents**

Copy the skill into your agent's skills directory:

```bash
git clone --depth 1 https://github.com/Eazydraft/docxtor.git /tmp/docxtor

mkdir -p ~/.agents/skills
cp -r /tmp/docxtor/plugin/skills/docxtor ~/.agents/skills/
```

Gemini CLI also supports direct installation:

```bash
gemini skills install https://github.com/Eazydraft/docxtor.git \
  --path plugin/skills/docxtor
```

The skill is self-contained and requires Node.js 20 or newer.

### CLI

Install globally:

```bash
npm install -g @lucio-ai/docxtor
```

Or run directly with npx:

```bash
npx @lucio-ai/docxtor
```

## CLI

The `docxtor` CLI provides a structured interface for inspecting and modifying documents.

```bash
# Inspect a document and find a particular clause
docxtor view contract.docx --grep "payment terms"

# Replace text using tracked changes
docxtor replace contract.docx -o redline.docx \
  --author "Acme" \
  --find "within thirty (30) days" \
  --replace "within sixty (60) days"

# Add an anchored comment
docxtor comment contract.docx -o commented.docx \
  --find "shall not exceed" \
  --text "Cap is too low."

# Inspect existing revisions
docxtor revisions counterparty.docx

# Resolve selected revisions
docxtor reject counterparty.docx -o resolved.docx \
  --ids 14,15 --others accept
```

### Command reference

| Command | Description |
|---|---|
| `view` | Inspect document contents, including paragraphs, tables, headers, footers, and inline revision markers. Supports `--grep`. |
| `revisions` | List tracked changes with their IDs, authors, and surrounding context. |
| `replace` | Find and replace text using native tracked changes. Only differing text is marked. |
| `comment` | Add a Word comment anchored to a specific passage. |
| `comments` | Inspect existing comments and their associated passages. |
| `edit` | Apply structured JSON patches to document content. |
| `accept` / `reject` | Resolve all revisions or selected revisions by ID or author. |
| `metadata` | Read or modify document properties and scrub author/company information. |

**Safety guarantees:**
- `replace` and `edit` write tracked changes by default, attributed with `--author`. Use `--untracked` to edit directly.
- `comment` adds a comment, attributed with `--author`; it changes no text.
- `accept` and `reject` resolve existing tracked changes, and `metadata` changes document properties; neither creates tracked changes.
- Failed operations do not produce partially written output.
- Ambiguous text matches are rejected unless the intended match is specified.

See the [CLI reference](docs/cli.md) for all commands and options.

## Benchmark

We evaluated docxtor using [docx-agent-bench](https://github.com/clauseo-ai/docx-agent-bench), a benchmark of realistic document-editing tasks involving real contracts.

The benchmark compares an AI agent equipped with the docxtor skill against the same agent without a document-editing plugin.

**Evaluation setup**

- 7 short tasks, 2 runs each.
- Two models: Claude Haiku 5.5 and Claude Opus 5.5.
- Evaluation date: November 6, 2026.
- Benchmark commit: `e214567`.

Tasks covered tracked-change redlines, bulk renaming, comments, chained edits, revision resolution, multi-turn editing, and document reading.

The benchmark's document-comparison task was excluded because docxtor does not currently implement document comparison.

### Results

| Model | Configuration | Grader-correct | Median time | Median tool calls | Median tool errors |
|---|---|---:|---:|---:|---:|
| Claude Haiku 5.5 | With docxtor | 14/14 | 17 s | 5 | 0 |
| Claude Haiku 5.5 | Without plugin | 14/14 | 30.5 s | 8.5 | 1 |
| Claude Opus 5.5 | With docxtor | 14/14 | 22 s | 5 | 0 |
| Claude Opus 5.5 | Without plugin | 13/13 | 36 s | 5 | 0 |

With docxtor, agents completed tasks approximately **40% faster**, with zero tool errors across the evaluated runs.

## TypeScript API

docxtor is also available as a TypeScript library.

Applications can use it to expose document structure to a model, receive structured editing instructions, and apply those instructions as native Word revisions.

Install:

```bash
npm install @lucio-ai/docxtor
```

### Example

```ts
import { Docxtor } from "@lucio-ai/docxtor";

// Load an existing document
const doc = await Docxtor.load("contract.docx");

// Project the document into a model-readable representation
console.log(doc.project("markdown"));

// Apply structured patches
const results = doc.applyPatches(
  [
    {
      type: "replace_text",
      target: "p1",
      find: "valid",
      replace: "enforceable"
    },
    {
      type: "insert_text",
      target: "p2",
      position: "after",
      anchor: "terms:",
      text: " and conditions:"
    },
    {
      type: "apply_formatting",
      target: "p1",
      find: "enforceable",
      formatting: { bold: true }
    },
    {
      type: "add_comment",
      target: "p1",
      find: "enforceable",
      text: "Confirm this wording."
    }
  ],
  {
    trackRevisions: true,
    revisionAuthor: "AI Editor"
  }
);

// Inspect individual patch results
for (const result of results) {
  if (!result.success) {
    console.error(`Patch failed: ${result.error}`);
  }
}

// Export the edited document
await doc.toFile("contract-redline.docx");
```

### Capabilities

**Structured document representation**

- Stable positional IDs for paragraphs, table cells, headers, and footers.
- Access to native Word `w14:paraId` identifiers, with support for assigning them when absent.
- Markdown and JSON projections for LLM consumption.
- Inline formatting markers, revision markers, page breaks, and run filtering by colour.

**Targeted editing**

- 23 patch types covering text, paragraphs, comments, tables, rows, lists, formatting, styles, headers, and footers.
- Fine-grained text replacement that marks only the differing text.
- Native Word tracked changes, including `w:ins`, `w:del`, `w:rPrChange`, and `w:pPrChange`.
- Accept or reject all revisions, or selectively resolve them by author or ID.

**Execution guarantees**

- Atomic patch batches and validation before application.
- Structured success and failure results for individual patches.
- All-or-nothing execution when requested.
- Unmodified document content remains untouched.

**Runtime**

- Node.js 20.19 or newer.
- ESM and CommonJS builds.
- TypeScript declarations included.
- Two runtime dependencies: [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) and [JSZip](https://stuk.github.io/jszip/).

## Architecture

docxtor is designed around a simple principle:

**Never reconstruct what you can modify in place.**

Traditional document-editing approaches often convert DOCX content into an intermediate representation and regenerate the document afterward. This can introduce unintended changes to formatting, styles, relationships, and other document structures.

docxtor takes a different approach.

```text
DOCX --> XML --> Normalized AST (read-only, with xmlRef pointers) --> LLM projection
                        |                                                   |
                        |                                             model produces
                        |                                             patch objects
                        v                                                   v
               Mutable XML objects <-------- Patch engine (mutates via xmlRef pointers)
                        |
                        v
                   Serialize --> DOCX
```

The normalized AST is a read-only representation of the document. Each node retains an `xmlRef` pointer to its corresponding object in the parsed XML.

Models operate on projections of this AST and produce structured patch objects. The patch engine resolves these patches through their references and modifies the underlying XML directly.

The AST is never serialized back into the document.

This separation allows docxtor to provide a structured, model-friendly editing interface while preserving document content and structure outside the scope of an edit.

## Documentation

| Resource | Description |
|---|---|
| [CLI reference](docs/cli.md) | Commands, arguments, and options |
| [API reference](docs/api.md) | Loading, creating, projecting, patching, and exporting documents |
| [Patch types](docs/patches.md) | Complete patch reference and behaviour |
| [ID scheme](docs/ids.md) | Positional IDs and native `w14:paraId` |
| [Revision tracking](docs/revisions.md) | Tracked-change generation and revision inspection |
| [LLM integration](docs/llm-integration.md) | System prompts and validate-then-apply workflows |
| [Feature coverage](docs/coverage.md) | Supported functionality and known limitations |
| [Examples](examples/) | Interactive chat CLI demonstrating model-driven editing |

## Current limitations

The following are not currently supported:

- Merged cells.
- Images.
- Creating new styles.
- Reading existing comments through the library API (the CLI supports listing comments).
- Document comparison.
- Mathematical equations.

See [feature coverage](docs/coverage.md) for the complete breakdown.

Contributions are welcome.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, architectural invariants, and contribution requirements.

The CLI source lives in [`cli/`](cli/).

```bash
# Build the npm package
pnpm build

# Rebuild the self-contained agent skill
pnpm build:skill
```

Every commit requires DCO sign-off.

## License

docxtor is released under the [MIT License](LICENSE).
