// Browser stub for node:fs/promises — docxtor's file-path APIs are Node-only.
// In Office.js add-ins, use Docxtor.fromOoxml() and toFlatOpc() instead.
export async function readFile(): Promise<never> {
  throw new Error('readFile is not available in the browser. Use Docxtor.fromOoxml() with body.ooxml instead.');
}
export async function writeFile(): Promise<never> {
  throw new Error('writeFile is not available in the browser. Use doc.toFlatOpc() and body.insertOoxml() instead.');
}
