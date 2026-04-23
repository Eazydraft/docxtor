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
 */

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
export function stripFormattingMarkers(text: string): string {
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
 */
export function parseFormattedText(input: string): FormattedSegment[] {
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
    // Only treat as marker if:
    // - at start/end of input, or
    // - preceded/followed by a non-word character or is at a word boundary
    if (input[pos] === "_") {
      const prevChar = pos > 0 ? input[pos - 1]! : "";
      const nextChar = pos + 1 < input.length ? input[pos + 1]! : "";

      // Treat _ as italic marker when it's at a boundary:
      // - start or end of string
      // - preceded or followed by whitespace/punctuation
      // - or when italic is already open (closing)
      const atBoundary =
        pos === 0 ||
        pos === input.length - 1 ||
        /\W/.test(prevChar) ||
        /\W/.test(nextChar) ||
        italic;

      if (atBoundary) {
        flush();
        italic = !italic;
        pos += 1;
        continue;
      }
    }

    currentText += input[pos];
    pos++;
  }

  flush();

  // Merge adjacent segments with identical formatting
  return mergeSegments(segments);
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
 * Check whether a string contains any GFM formatting markers.
 */
export function hasFormattingMarkers(text: string): boolean {
  return /\*\*|~~|<\/?u>/.test(text) || /_/.test(text);
}
