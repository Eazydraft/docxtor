/**
 * Markdown formatting parser and stripper.
 *
 * Parses GFM inline formatting markers into structured segments
 * and provides a utility to strip markers for plain-text matching.
 *
 * Supported markers:
 * - **bold**
 * - _italic_
 * - ~~strikethrough~~
 * - <u>underline</u>
 * - Nesting: **_bold+italic_**, _**bold+italic**_
 *
 * Escapes: a backslash before `*`, `_`, `~`, `<` or `\` writes that
 * character literally (`\_` → `_`). Any other backslash is kept as-is.
 *
 * Underscore rules: a `_` is an italic marker only when it stands alone
 * (not next to another `_`) at a word boundary. Runs of underscores such as
 * signature blanks (`Name: ______`) and intra-word underscores
 * (`snake_case`) are literal text.
 *
 * Every entry point takes a `markers` flag; with `false` the text is
 * returned verbatim — no markers, no escapes.
 */

/** Characters a backslash may escape when markers are enabled. */
const ESCAPABLE = new Set(["*", "_", "~", "<", "\\"]);

/**
 * A segment of text with its inline formatting flags.
 */
export interface FormattedSegment {
  text: string;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  underline?: boolean;
}

/**
 * Strip all GFM inline formatting markers from a string,
 * returning the plain text content.
 *
 * Used to normalize `find` / `anchor` fields in patches so they
 * match against the plain paragraph text.
 */
export function stripFormattingMarkers(text: string, markers: boolean = true): string {
  if (!markers) return text;
  // Process by walking the string and removing marker characters.
  // We use the same parser and concatenate the text segments.
  const segments = parseFormattedText(text);
  return segments.map((s) => s.text).join("");
}

/**
 * Parse a GFM-formatted string into an array of segments, each
 * carrying the formatting flags that were active when the text
 * appeared.
 *
 * Handles:
 * - **bold**
 * - _italic_  (underscore only, to avoid conflicts with intra-word asterisks)
 * - ~~strikethrough~~
 * - Nesting in any order: **_text_**, _**text**_
 *
 * Plain text without markers returns a single segment with no flags.
 * With `markers: false` the input is returned as one verbatim segment.
 */
export function parseFormattedText(input: string, markers: boolean = true): FormattedSegment[] {
  if (!markers) return input.length > 0 ? [{ text: input }] : [];

  const segments: FormattedSegment[] = [];
  let pos = 0;
  let currentText = "";
  let bold = false;
  let italic = false;
  let strikethrough = false;
  let underline = false;

  function flush(): void {
    if (currentText.length > 0) {
      const seg: FormattedSegment = { text: currentText };
      if (bold) seg.bold = true;
      if (italic) seg.italic = true;
      if (strikethrough) seg.strikethrough = true;
      if (underline) seg.underline = true;
      segments.push(seg);
      currentText = "";
    }
  }

  while (pos < input.length) {
    // Backslash escape: the next character is literal
    if (input[pos] === "\\" && pos + 1 < input.length && ESCAPABLE.has(input[pos + 1]!)) {
      currentText += input[pos + 1];
      pos += 2;
      continue;
    }

    // Check for <u> (underline open)
    if (input.slice(pos, pos + 3) === "<u>") {
      flush();
      underline = true;
      pos += 3;
      continue;
    }

    // Check for </u> (underline close)
    if (input.slice(pos, pos + 4) === "</u>") {
      flush();
      underline = false;
      pos += 4;
      continue;
    }

    // Check for ~~ (strikethrough toggle)
    if (input[pos] === "~" && input[pos + 1] === "~") {
      flush();
      strikethrough = !strikethrough;
      pos += 2;
      continue;
    }

    // Check for ** (bold toggle)
    if (input[pos] === "*" && input[pos + 1] === "*") {
      flush();
      bold = !bold;
      pos += 2;
      continue;
    }

    // Check for _ (italic toggle)
    if (input[pos] === "_" && isItalicMarker(input, pos, italic)) {
      flush();
      italic = !italic;
      pos += 1;
      continue;
    }

    currentText += input[pos];
    pos++;
  }

  flush();

  // Merge adjacent segments with identical formatting
  return mergeSegments(segments);
}

/**
 * Decide whether the `_` at `pos` is an italic marker.
 *
 * - Never when it touches another `_` (underscore blanks, `__init__`).
 * - Opening: at the start of the string or after a non-word character, and
 *   followed by something that is not whitespace.
 * - Closing (italic already open): preceded by something that is not
 *   whitespace.
 * - Otherwise only at the end of the string or next to a non-word
 *   character — the word-boundary rule that keeps `snake_case` literal.
 */
function isItalicMarker(input: string, pos: number, italicOpen: boolean): boolean {
  const prevChar = pos > 0 ? input[pos - 1]! : "";
  const nextChar = pos + 1 < input.length ? input[pos + 1]! : "";
  if (prevChar === "_" || nextChar === "_") return false;

  if (italicOpen) {
    return prevChar !== "" && !/\s/.test(prevChar);
  }
  const opensAtBoundary = pos === 0 || /\W/.test(prevChar);
  if (opensAtBoundary) {
    return nextChar !== "" && !/\s/.test(nextChar);
  }
  return pos === input.length - 1 || /\W/.test(nextChar);
}

/**
 * Merge adjacent segments with identical formatting flags.
 */
function mergeSegments(segments: FormattedSegment[]): FormattedSegment[] {
  if (segments.length <= 1) return segments;

  const merged: FormattedSegment[] = [segments[0]!];

  for (let i = 1; i < segments.length; i++) {
    const prev = merged[merged.length - 1]!;
    const curr = segments[i]!;

    if (
      !!prev.bold === !!curr.bold &&
      !!prev.italic === !!curr.italic &&
      !!prev.strikethrough === !!curr.strikethrough &&
      !!prev.underline === !!curr.underline
    ) {
      prev.text += curr.text;
    } else {
      merged.push(curr);
    }
  }

  return merged;
}

/**
 * Check whether a string contains any GFM formatting markers or escapes,
 * i.e. whether {@link parseFormattedText} would change it.
 */
export function hasFormattingMarkers(text: string): boolean {
  return stripFormattingMarkers(text) !== text;
}
