# Changelog

All notable changes to this project are documented here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses [Semantic Versioning](https://semver.org). Until 1.0, minor versions may contain breaking changes; they are called out when they happen.

## [Unreleased]

## [0.4.0] - 2026-11-03

First public release, published to npm as `@lucio-ai/docxtor`.

### Changed

- Relicensed under MIT and prepared for public release: contribution guide, DCO, code of conduct, security policy, CI workflow and npm publishing through trusted publishing with provenance.
- Package renamed from `@eazydraft/docxtor` to `@lucio-ai/docxtor`.
- Documentation split out of the README into `docs/`, with a benchmark section in the README.
- The sample chat CLI moved from `scripts/` to `examples/`.

### Fixed

- Page and column breaks stay on their original side of the text when a run is rebuilt by a patch.

## [0.3.2] - 2026-10-19

### Fixed

- Soft line breaks (`w:br`) now count as one character in projections and patch offsets, so offsets computed from a projection line up.
- `replace_text` and `delete_text` inside another author's unaccepted insertion split the `w:ins` around the edit instead of double-marking it.
- Verbatim text: `inlineMarkers: false` on a batch and `markers: false` on a patch write text exactly as given. Backslash escapes for marker characters.
- Field placeholders use a collision-free internal marker, so documents that already contain the old placeholder text are not corrupted.
- Runs that hold a soft break, page break or column break are rebuilt with the break on its original side of the edit.

## [0.3.1] - 2026-10-13

### Fixed

- CI: approve the esbuild build script so `pnpm install` succeeds non-interactively.

## [0.3.0] - 2026-10-13

### Added

- Native paragraph identity: `w14:paraId` is read, exposed in projections, accepted as a `para:` patch target, and assigned on load with `assignParaIds`.
- `insert_paragraph` and `delete_paragraph` patches, tracked or untracked, including inside table cells and headers/footers.
- `acceptAllRevisions` and `rejectAllRevisions`, optionally filtered by author.
- Document metadata: `getCoreProperties`, `setCoreProperties` and `scrubMetadata`.
- Atomic batches (`atomic: true`) that leave the document untouched if any patch fails.
- `add_comment` patch, creating `word/comments.xml` on demand.
- `runFilter` on the JSON projection for reading runs by formatting, such as answers typed in red.
- `revisionDate` option for byte-stable tracked output.

## [0.2.0] - 2026-04-25

### Added

- List operations: `add_list`, `add_list_item`, `remove_list_item`, `convert_to_list`, `remove_list`, `set_list_properties`, with numbering definitions managed in `word/numbering.xml`.
- GFM inline formatting markers (`**bold**`, `_italic_`, `~~strike~~`, `<u>underline</u>`) in projections and in patch text.
- Headers and footers: normalization, `add_header`, `add_footer`, `remove_header`, `remove_footer`, container targeting for `append_text` and `add_table`, and `{{FIELD}}` placeholders.

## [0.1.5] - 2026-04-20

### Added

- `add_table` patch with validation, and markdown table conversion in `append_text`.

## [0.1.4] - 2026-04-15

### Changed

- `append_text` splits on `\n` into separate paragraphs.

## [0.1.3] - 2026-04-14

### Added

- `doc.toString()` returns the current `word/document.xml`.

### Changed

- Build moved to tsdown (ESM, CJS and type declarations).

## [0.1.2] - 2026-04-11

### Added

- `Docxtor.create()` builds a valid empty document from scratch.
- `append_text` patch.
- Page break detection in projections with positional `pageBreakOffset`.

## [0.1.1] - 2026-04-09

### Fixed

- Package files include the logo; README alt text.

## [0.1.0] - 2026-04-09

### Added

- Initial release: ZIP and XML round trip, normalized AST with `xmlRef` pointers, markdown and JSON projections, text patches (`replace_text`, `insert_text`, `delete_text`), table row patches, `apply_formatting`, `apply_style`, hyperlinks, nested tables, and tracked changes (`w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange`).

[Unreleased]: https://github.com/Eazydraft/docxtor/compare/v0.4.0...HEAD
[0.4.0]: https://github.com/Eazydraft/docxtor/compare/v0.3.2...v0.4.0
[0.3.2]: https://github.com/Eazydraft/docxtor/compare/v0.3.1...v0.3.2
[0.3.1]: https://github.com/Eazydraft/docxtor/compare/v0.3.0...v0.3.1
[0.3.0]: https://github.com/Eazydraft/docxtor/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/Eazydraft/docxtor/compare/v0.1.5...v0.2.0
[0.1.5]: https://github.com/Eazydraft/docxtor/compare/v0.1.4...v0.1.5
[0.1.4]: https://github.com/Eazydraft/docxtor/compare/v0.1.3...v0.1.4
[0.1.3]: https://github.com/Eazydraft/docxtor/compare/v0.1.2...v0.1.3
[0.1.2]: https://github.com/Eazydraft/docxtor/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/Eazydraft/docxtor/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/Eazydraft/docxtor/releases/tag/v0.1.0
