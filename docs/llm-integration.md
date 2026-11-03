# LLM integration pattern

docxtor is LLM-agnostic. The typical integration:

```ts
import { Docxtor } from "@lucio-ai/docxtor";

const doc = await Docxtor.load("contract.docx");
const projection = doc.project("markdown", { includeFormatting: true });

// Send projection to your LLM with instructions to produce patches
const response = await llm.chat([
  {
    role: "system",
    content: `You are a document editor. The user will give you a document 
and editing instructions. Respond with a JSON array of patches.

Patch types:
- replace_text: { type: "replace_text", target: "<id>", find: "<text>", replace: "<text>" }
- insert_text: { type: "insert_text", target: "<id>", position: "before"|"after", anchor: "<text>", text: "<text>" }
- delete_text: { type: "delete_text", target: "<id>", find: "<text>" }
- append_text: { type: "append_text", text: "<text>" } (supports markdown tables, {{FIELD}} placeholders, and inline formatting)
- insert_paragraph: { type: "insert_paragraph", target: "<id>", position: "before"|"after", text: "<text>" } (inherits the target's style and numbering)
- delete_paragraph: { type: "delete_paragraph", target: "<id>" }
- add_comment: { type: "add_comment", target: "<id>", find?: "<text>", text: "<comment>" }
- add_table: { type: "add_table", headers: ["col1", ...], rows: [["val1", ...], ...], afterId: "<id>" }
- apply_formatting: { type: "apply_formatting", target: "<id>", find: "<text>", formatting: { bold: true, color: "FF0000" } }
- apply_style: { type: "apply_style", target: "<id>", style: "Heading1" }
- add_header: { type: "add_header", variant: "default"|"first"|"even", paragraphs: ["text", ...] }
- add_footer: { type: "add_footer", variant: "default"|"first"|"even", paragraphs: ["text", ...] }
- remove_header: { type: "remove_header", variant: "default"|"first"|"even" }
- remove_footer: { type: "remove_footer", variant: "default"|"first"|"even" }
- add_list: { type: "add_list", listType: "bullet"|"ordered"|"multilevel", items: ["item1", ...], afterId?: "<id>", numId?: "<existing numId>" }
- add_list_item: { type: "add_list_item", target: "<list paragraph id>", position: "before"|"after", text: "<text>", level?: <number> }
- remove_list_item: { type: "remove_list_item", target: "<list paragraph id>" }
- convert_to_list: { type: "convert_to_list", targets: ["<id>", ...], listType: "bullet"|"ordered"|"multilevel" }
- remove_list: { type: "remove_list", targets: ["<id>", ...] }
- set_list_properties: { type: "set_list_properties", target: "<id>", level?: <number>, listType?: "bullet"|"ordered"|"multilevel" }

List paragraphs show {list:N LX} in the projection where N is the numId and X is the nesting level.
To continue an existing list, reuse the numId from the hint.

Document:
${projection}`,
  },
  { role: "user", content: "Change all mentions of 'Provider' to 'Vendor' and make them bold" },
]);

const patches = JSON.parse(response);

// Validate before applying. Stop here if anything is invalid: applyPatches
// validates each patch just before applying it, so an invalid patch in the
// middle of a batch would leave the earlier ones applied.
const validations = doc.validatePatches(patches);
const invalid = validations.filter((v) => !v.valid);
if (invalid.length > 0) {
  throw new Error(`Invalid patches: ${JSON.stringify(invalid)}`);
}

// Or let the engine enforce all-or-nothing with { atomic: true }.
const results = doc.applyPatches(patches, { atomic: true });
if (!results.every((r) => r.success)) {
  throw new Error("Batch failed; the document is unchanged");
}
await doc.toFile("contract-edited.docx");
```
