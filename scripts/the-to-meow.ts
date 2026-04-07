/**
 * Real-world test script: replace every instance of the word "the" with "meow"
 * in a DOCX file using docxtor.
 *
 * Usage:
 *   npx tsx scripts/the-to-meow.ts input.docx [output.docx]
 *
 * If output path is omitted, writes to input.meow.docx
 */

import { Docxtor } from "../src/index.js";
import type { Patch, NormalizedParagraph } from "../src/index.js";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Find all whole-word occurrences of "the" (case-insensitive) in a string.
 *
 * Returns { find, replace } pairs with enough surrounding context to avoid
 * matching "the" inside words like "otherwise" or "together". We grab one
 * character of context on each side (when available) and include that context
 * in both find and replace so only the word itself changes.
 */
function findTheWithContext(
  text: string,
): { find: string; replace: string }[] {
  const regex = /\bthe\b/gi;
  const results: { find: string; replace: string }[] = [];
  let m: RegExpExecArray | null;
  while ((m = regex.exec(text)) !== null) {
    const matchStart = m.index;
    const matchEnd = m.index + m[0].length;

    // Grab one char of surrounding context (space, punctuation, etc.)
    const ctxStart = Math.max(0, matchStart - 1);
    const ctxEnd = Math.min(text.length, matchEnd + 1);

    const prefix = text.slice(ctxStart, matchStart);
    const suffix = text.slice(matchEnd, ctxEnd);

    results.push({
      find: prefix + m[0] + suffix,
      replace: prefix + "meow" + suffix,
    });
  }
  return results;
}

/**
 * Build replace_text patches for every whole-word "the" in a paragraph.
 *
 * Because the patch engine uses exact substring matching (no regex), we
 * include one character of surrounding context so "the" inside "otherwise"
 * or "together" is never matched.
 *
 * Because the engine re-normalizes after each patch, we always target
 * occurrence 1 — after each replacement the next match becomes occurrence 1.
 */
function buildPatchesForParagraph(para: NormalizedParagraph): Patch[] {
  const hits = findTheWithContext(para.text);
  if (hits.length === 0) return [];

  return hits.map(({ find, replace }) => ({
    type: "replace_text" as const,
    target: para.id,
    find,
    replace,
    occurrence: 1,
  }));
}

// ---------------------------------------------------------------------------
// Collect paragraphs from all document locations
// ---------------------------------------------------------------------------

function collectAllParagraphs(doc: Docxtor): NormalizedParagraph[] {
  const allParas: NormalizedParagraph[] = [];

  // Body paragraphs and table cell paragraphs
  for (const el of doc.getBody()) {
    if (el.kind === "paragraph") {
      allParas.push(el);
    } else if (el.kind === "table") {
      for (const row of el.rows) {
        for (const cell of row.cells) {
          allParas.push(...cell.paragraphs);
          // Nested tables
          for (const nested of cell.nestedTables) {
            for (const nRow of nested.rows) {
              for (const nCell of nRow.cells) {
                allParas.push(...nCell.paragraphs);
              }
            }
          }
        }
      }
    }
  }

  // Headers
  for (const [, hdr] of doc.getHeaders()) {
    allParas.push(...hdr.paragraphs);
  }

  // Footers
  for (const [, ftr] of doc.getFooters()) {
    allParas.push(...ftr.paragraphs);
  }

  return allParas;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.error("Usage: npx tsx scripts/the-to-meow.ts <input.docx> [output.docx]");
    process.exit(1);
  }

  const inputPath = args[0]!;
  const outputPath = args[1] ?? inputPath.replace(/\.docx$/i, ".meow.docx");

  console.log(`Loading: ${inputPath}`);
  const doc = await Docxtor.load(inputPath);

  // Collect all paragraphs
  const allParas = collectAllParagraphs(doc);
  console.log(`Found ${allParas.length} paragraphs across body, tables, headers, and footers.`);

  // Build patches
  const patches: Patch[] = [];
  for (const para of allParas) {
    patches.push(...buildPatchesForParagraph(para));
  }

  console.log(`Generated ${patches.length} replacement patches.`);

  if (patches.length === 0) {
    console.log("No instances of 'the' found. Nothing to do.");
    return;
  }

  // Apply patches
  const results = doc.applyPatches(patches);

  // Report
  const succeeded = results.filter((r) => r.success).length;
  const failed = results.filter((r) => !r.success);

  console.log(`Applied: ${succeeded}/${results.length} patches succeeded.`);

  if (failed.length > 0) {
    console.error(`\nFailed patches:`);
    for (const f of failed) {
      console.error(`  - [${f.patch.target}] ${f.error}`);
    }
  }

  // Save
  await doc.toFile(outputPath);
  console.log(`Saved: ${outputPath}`);
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
