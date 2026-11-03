<img src="logo.png" alt="Docxtor logo" />

[![npm](https://img.shields.io/npm/v/@lucio-ai/docxtor)](https://www.npmjs.com/package/@lucio-ai/docxtor)
[![CI](https://github.com/Eazydraft/docxtor/actions/workflows/ci.yml/badge.svg)](https://github.com/Eazydraft/docxtor/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Surgical, LLM-friendly DOCX editing. Load a `.docx`, get a normalized view an LLM can read, apply the patches the LLM produces, and write out a valid `.docx`, with formatting, styles, numbering, headers and tables left exactly as they were.

Built for workflows where a model proposes edits and a human reviews them in Word: every patch can land as a tracked change, with author and date, that the reviewer accepts or rejects.

## Features

- **Stable ids for every element.** Paragraphs, table cells, headers and footers get positional ids (`p12`, `tbl0.r1.c0.p0`), and Word's native `w14:paraId` is exposed wherever the file carries one (or stamped on load with `assignParaIds`), so patches target exactly one place.
- **Patches, not rewrites.** Replace, insert and delete text, insert and delete paragraphs, add comments, tables, rows, lists, headers and footers, apply formatting and styles. Twenty-three patch types in all.
- **Tracked changes.** Turn on `trackRevisions` and every content edit emits `w:ins`, `w:del`, `w:rPrChange` or `w:pPrChange` markup; comments and header or footer removal are the exceptions, since Word does not track them. Accept or reject all revisions, optionally per author.
- **Projections for models.** Markdown or JSON views with ids, inline formatting markers, revision marks, page breaks and a filter for runs by colour.
- **Inline formatting in patch text.** `**bold**`, `_italic_`, `~~strike~~`, `<u>underline</u>`, `{{PAGE}}` style fields, markdown tables and lists, or verbatim text when you need it.
- **Atomic batches and validation.** Validate a batch before applying it, or run it all-or-nothing.
- **Two runtime dependencies.** [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) and [JSZip](https://stuk.github.io/jszip/). ESM and CommonJS builds with types.

## Install

```bash
npm install @lucio-ai/docxtor
# or
pnpm add @lucio-ai/docxtor
```

Requires Node.js 20.19 or newer.

## Quick start

```ts
import { Docxtor } from "@lucio-ai/docxtor";

const doc = await Docxtor.load("contract.docx");

// A view the model can read. Every line carries the id patches target.
console.log(doc.project("markdown"));
// [p0] (Heading1) Service Agreement
// [p1] This agreement is valid from January 1, 2025...
// [p2] The parties agree to the following terms:

// Apply the patches the model produced, as tracked changes.
const results = doc.applyPatches(
  [
    { type: "replace_text", target: "p1", find: "valid", replace: "enforceable" },
    { type: "insert_text", target: "p2", position: "after", anchor: "terms:", text: " and conditions:" },
    { type: "apply_formatting", target: "p1", find: "enforceable", formatting: { bold: true } },
    { type: "add_comment", target: "p1", find: "enforceable", text: "Confirm with counsel." },
  ],
  { trackRevisions: true, revisionAuthor: "AI Editor" },
);

for (const r of results) {
  if (!r.success) console.error(`Patch failed: ${r.error}`);
}

await doc.toFile("contract-redline.docx");
```

Open the result in Word and the edits show up as redlines from "AI Editor". Drop the options argument to apply the same patches directly.

## How it works

```
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

The AST is never written back. It is a read-only lens with `xmlRef` pointers into the parsed XML. Patches follow those pointers and mutate the XML in place, and serializing the result produces the output document. Anything a patch does not touch survives untouched. See [docs/design.md](docs/design.md) for the invariants this rests on.

## Benchmark

docxtor was run through [docx-agent-bench](https://github.com/clauseo-ai/docx-agent-bench), an agent benchmark of realistic edits to real contracts (tracked-change redlines, a bulk rename, comments, chained edits, resolving a counterparty's tracked changes, a version compare, and a reading task), graded on the `.docx` the agent writes. The agent got docxtor through a thin CLI over the library and a skill describing it; the baseline is the same agent with no document tool. Eight short tasks, two runs each, 3 November 2026, benchmark commit `e214567`.

| Model | Arm | Grader-correct runs | Median wall time | Median tool calls | Median tool errors |
|---|---|---|---|---|---|
| Claude Haiku 5.5 | docxtor | 16/16 | 19 s | 5 | 0 |
| Claude Haiku 5.5 | no tool | 16/16 | 42.5 s | 9 | 1 |
| Claude Opus 5.5 | docxtor | 16/16 | 34 s | 5 | 0 |
| Claude Opus 5.5 | no tool | 16/16 | 41 s | 5.5 | 0 |

Every file docxtor wrote passed the benchmark's schema check with no new findings. The two tasks docxtor has no primitive for, per-revision accept/reject and document compare, were solved by the agent from the existing primitives and are where its wall time goes; the other tasks ran two to four times faster than the baseline. Word-based validity and the benchmark's long tasks were not part of this run.

## Documentation

- [API reference](docs/api.md): `load`, `create`, `project`, `applyPatches`, accept and reject revisions, metadata, export, low-level exports
- [Patch types](docs/patches.md): every patch with its fields and behaviour
- [ID scheme](docs/ids.md): positional ids and native `w14:paraId`
- [Revision tracking](docs/revisions.md): what each patch emits as tracked changes, reading revision marks
- [LLM integration pattern](docs/llm-integration.md): a system prompt and validate-then-apply loop
- [Feature coverage](docs/coverage.md): what is handled and what is not yet
- [Design](docs/design.md): architecture and invariants
- [Examples](examples/): an interactive chat CLI that edits a document with a model

## Not handled yet

Merged cells, images, creating new styles, reading existing comments, document compare, and math equations. The full list is in [feature coverage](docs/coverage.md). Contributions are welcome.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for setup, the invariants to keep, and the DCO sign-off every commit needs.

## License

[MIT](LICENSE)
