/**
 * Issue #34: replace_text where the match starts at offset 0 of a run that
 * sits inside another author's `w:ins`. The replacement must land where the
 * matched text was, not after the surrounding insertion.
 */

import { describe, it, expect } from "vitest";
import { Docxtor } from "../../src/index.js";
import { createBodyDocx } from "../fixtures/create-docx.js";

const TRACKED = { trackRevisions: true, revisionAuthor: "T", revisionDate: "2026-10-17T00:00:00Z" };

const PREMISES = (inner: string): string => `<w:p><w:pPr><w:pStyle w:val="FlushLeft"/></w:pPr>
  <w:r><w:t xml:space="preserve">PREMISES – </w:t></w:r>
  ${inner}
  <w:r><w:t xml:space="preserve"> – 13,848 SF </w:t></w:r>
</w:p>`;

const PLAIN_RUN = `<w:r><w:t>Suite [SUITE NUMBER]</w:t></w:r>`;
const OTHER_INS = `<w:ins w:id="7" w:author="Other" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Suite [SUITE NUMBER]</w:t></w:r></w:ins>`;
const OWN_INS = `<w:ins w:id="7" w:author="T" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Suite [SUITE NUMBER]</w:t></w:r></w:ins>`;

const EXPECTED = "PREMISES – REVISED [SUITE NUMBER] – 13,848 SF ";

describe("replace_text at offset 0 of a run (issue #34)", () => {
  it("plain run, untracked and tracked", async () => {
    for (const opts of [undefined, TRACKED]) {
      const doc = await Docxtor.load(await createBodyDocx(PREMISES(PLAIN_RUN)));
      const [r] = doc.applyPatches([{ type: "replace_text", target: "p0", find: "Suite", replace: "REVISED", occurrence: 1 }], opts);
      expect(r!.success).toBe(true);
      expect(doc.getParagraphs()[0]!.text).toBe(EXPECTED);
    }
  });

  it("run inside another author's w:ins, untracked", async () => {
    const doc = await Docxtor.load(await createBodyDocx(PREMISES(OTHER_INS)));
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Suite", replace: "REVISED" }]);
    expect(doc.getParagraphs()[0]!.text).toBe(EXPECTED);
  });

  it("run inside another author's w:ins, tracked: insertion is split around the replacement", async () => {
    const doc = await Docxtor.load(await createBodyDocx(PREMISES(OTHER_INS)));
    const [r] = doc.applyPatches([{ type: "replace_text", target: "p0", find: "Suite", replace: "REVISED" }], TRACKED);
    expect(r!.success).toBe(true);
    expect(doc.getParagraphs()[0]!.text).toBe(EXPECTED);

    const xml = doc.toString().replace(/>\s+</g, "><");
    // No w:del: the matched text was an unaccepted insertion, so it is simply gone
    expect(xml).not.toContain("<w:del ");
    // Our w:ins sits before the remainder of the other author's insertion
    expect(xml).toMatch(
      /<w:ins w:id="\d+" w:author="T"[^>]*><w:r><w:t>REVISED<\/w:t><\/w:r><\/w:ins><w:ins w:id="\d+" w:author="Other"[^>]*><w:r><w:t xml:space="preserve"> \[SUITE NUMBER\]<\/w:t><\/w:r><\/w:ins>/,
    );

    // Revision marks in the projection reflect both authors in order
    const revs = doc.getParagraphs()[0]!.revisions!;
    expect(revs.map((m) => [m.type, m.text, m.author])).toEqual([
      ["insert", "REVISED", "T"],
      ["insert", " [SUITE NUMBER]", "Other"],
    ]);

    // Accepting everything yields the clean text; rejecting restores the other author's text
    const accepted = await Docxtor.load(await doc.toBuffer());
    accepted.acceptAllRevisions();
    expect(accepted.getParagraphs()[0]!.text).toBe(EXPECTED);
    const rejected = await Docxtor.load(await doc.toBuffer());
    rejected.rejectAllRevisions();
    expect(rejected.getParagraphs()[0]!.text).toBe("PREMISES –  – 13,848 SF ");
  });

  it("match in the middle of another author's w:ins keeps both halves attributed to them", async () => {
    const doc = await Docxtor.load(await createBodyDocx(PREMISES(OTHER_INS)));
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "[SUITE", replace: "[UNIT" }], TRACKED);
    expect(doc.getParagraphs()[0]!.text).toBe("PREMISES – Suite [UNIT NUMBER] – 13,848 SF ");
    const xml = doc.toString().replace(/>\s+</g, "><");
    expect(xml).toMatch(
      /<w:ins [^>]*w:author="Other"[^>]*><w:r><w:t xml:space="preserve">Suite <\/w:t><\/w:r><\/w:ins><w:ins [^>]*w:author="T"[^>]*><w:r><w:t>\[UNIT<\/w:t><\/w:r><\/w:ins><w:ins [^>]*w:author="Other"[^>]*><w:r><w:t xml:space="preserve"> NUMBER\]<\/w:t><\/w:r><\/w:ins>/,
    );
    // The split-off half got its own revision id
    const ids = [...xml.matchAll(/<w:ins w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("match at the end of another author's w:ins", async () => {
    const doc = await Docxtor.load(await createBodyDocx(PREMISES(OTHER_INS)));
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "NUMBER]", replace: "NO.]" }], TRACKED);
    expect(doc.getParagraphs()[0]!.text).toBe("PREMISES – Suite [SUITE NO.] – 13,848 SF ");
  });

  it("whole w:ins replaced: the other author's insertion disappears, ours takes its place", async () => {
    const doc = await Docxtor.load(await createBodyDocx(PREMISES(OTHER_INS)));
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Suite [SUITE NUMBER]", replace: "Suite 100" }], TRACKED);
    expect(doc.getParagraphs()[0]!.text).toBe("PREMISES – Suite 100 – 13,848 SF ");
    const xml = doc.toString();
    expect(xml).not.toContain('w:author="Other"');
  });

  it("run inside the same author's w:ins, tracked: replacement joins that insertion in place", async () => {
    const doc = await Docxtor.load(await createBodyDocx(PREMISES(OWN_INS)));
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Suite", replace: "REVISED" }], TRACKED);
    expect(doc.getParagraphs()[0]!.text).toBe(EXPECTED);
    const xml = doc.toString().replace(/>\s+</g, "><");
    expect((xml.match(/<w:ins /g) ?? []).length).toBe(1);
    expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t>REVISED<\/w:t><\/w:r><w:r><w:t xml:space="preserve"> \[SUITE NUMBER\]<\/w:t><\/w:r><\/w:ins>/);
  });

  it("empty replacement inside another author's w:ins just removes the text", async () => {
    const doc = await Docxtor.load(await createBodyDocx(PREMISES(OTHER_INS)));
    doc.applyPatches([{ type: "replace_text", target: "p0", find: "Suite ", replace: "" }], TRACKED);
    expect(doc.getParagraphs()[0]!.text).toBe("PREMISES – [SUITE NUMBER] – 13,848 SF ");
    expect((doc.toString().match(/<w:ins /g) ?? []).length).toBe(1);
  });

  it("run inside a hyperlink, untracked and tracked", async () => {
    const inner = `<w:hyperlink r:id="rId9"><w:r><w:t>Suite [SUITE NUMBER]</w:t></w:r></w:hyperlink>`;
    for (const opts of [undefined, TRACKED]) {
      const doc = await Docxtor.load(await createBodyDocx(PREMISES(inner)));
      const [r] = doc.applyPatches([{ type: "replace_text", target: "p0", find: "Suite", replace: "REVISED" }], opts);
      expect(r!.success).toBe(true);
      expect(doc.getParagraphs()[0]!.text).toBe(EXPECTED);
    }
  });
});
