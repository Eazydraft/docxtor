# docxtor plugin

Redline Word documents with your coding agent. With this plugin, Claude can read a `.docx`, edit it as tracked changes with an author and date, add comments anchored to a passage, and accept or reject someone else's tracked changes one by one. Everything the agent doesn't touch, such as formatting, styles, numbering, headers and tables, stays as it was.

## Install

In Cowork or Claude.ai, upload `docxtor-skill.zip` from the [latest release](https://github.com/Eazydraft/docxtor/releases/latest) under Customize > Skills, with code execution on. The skill needs Node.js 20 or newer where it runs.

## What's inside

| Path | What it is |
|---|---|
| `skills/docxtor/SKILL.md` | Tells the agent when to use docxtor and how: which command fits each kind of request, and how to check its work. |
| `skills/docxtor/reference/` | Loaded only when needed: every command and option, every patch type, and worked examples. |
| `skills/docxtor/scripts/docxtor.mjs` | The `docxtor` command-line tool. It is the open-source docxtor library and its CLI, built from this repository's `src/` and `cli/` by `pnpm build:skill`, with its two dependencies bundled in: [JSZip](https://stuk.github.io/jszip/) and [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser). Minified, unmodified build output. Licenses are listed in `THIRD_PARTY_NOTICES.md`. |

## What it does on your machine

- It reads the `.docx` files the agent names, and writes only the output file given with `-o`. It never modifies its input unless the output path is the input path.
- It makes no network requests, collects no data, and starts no background processes.
- It runs only when the agent runs `node .../scripts/docxtor.mjs`, under your normal permission settings.

## Commands

`view`, `revisions`, `comments`, `replace`, `comment`, `edit`, `accept`, `reject` and `metadata`. Every write is a tracked change by default and writes nothing if any part fails. The full reference is at https://github.com/Eazydraft/docxtor/blob/main/docs/cli.md.

## License

MIT. Source: https://github.com/Eazydraft/docxtor
