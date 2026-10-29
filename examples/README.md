# Examples

Runnable samples that use docxtor from source. They are not part of the published package.

## `chat.ts`: interactive LLM editing session

Loads a `.docx`, sends the projection to a language model, and applies the patches the model returns. Uses the OpenAI SDK against OpenRouter, so any OpenRouter-hosted model works.

```bash
pnpm install
export OPENROUTER_API_KEY=...
export OPENROUTER_MODEL=openai/gpt-4o   # optional
npx tsx examples/chat.ts path/to/input.docx
```

Commands inside the session:

| Command | Effect |
|---|---|
| `/show` | Print the current projection |
| `/undo` | Undo the last batch of patches |
| `/history` | Show the edit history |
| `/save [path]` | Save the document (default `<input>.edited.docx`) |
| `/quit` | Save and exit |

Anything else is sent to the model as an editing instruction. See [LLM integration pattern](../docs/llm-integration.md) for the prompt shape it uses.
