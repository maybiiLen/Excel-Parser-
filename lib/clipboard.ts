// Word-output step: wrap the renderer's HTML fragment as Word-flavored HTML for
// the clipboard ("Copy for Word"). NOTHING is a Word heading -- the whole pivot
// is plain, directly-formatted body paragraphs, so it drops under an existing
// document heading without entering Word's outline or going blue. renderPivotTree
// returns a bare fragment (no <html>/<head>/<body>); the browser's ClipboardItem
// writes the Windows CF_HTML header for us.

/**
 * How one heading level should look -- one entry per depth (index 0 = level 1).
 * Shared across all tables; defaults are all the same until the user diverges a
 * level.
 */
export type LevelStyle = {
  color: string; // hex
  font: string; // font-family name
  size: number; // pt
  bold: boolean;
};

/**
 * The single styling source, shared across all tables.
 * - `levels` -- per-depth direct look (index 0 = level 1). Always used for the
 *   body (nested rows), and for the title when `headingStyleName` is blank.
 * - `indentStep` -- left-indent added per nesting level (inches).
 * - `headingStyleName` -- optional Word style name for the TITLE only. When set,
 *   the title gets `mso-style-name:"<headingStyleName>"`, so a "Use Destination
 *   Styles" paste adopts the destination document's heading style. Blank =
 *   the app's direct level-1 look. The body is always the app's direct look.
 */
export type HeadingStyle = {
  levels: LevelStyle[];
  indentStep: number; // inches of left-indent per nesting level
  headingStyleName: string; // Word style for the title ("" = direct look)
};

/**
 * Make a string safe inside `mso-style-name:"..."`. A Word style name is short
 * plain text; strip the quote/angle-bracket characters that could break out of
 * the attribute (keep spaces/dots so names like "Heading 1" still match). The
 * inputs are form-controlled; this is belt-and-suspenders.
 */
function sanitizeStyleName(s: string): string {
  return s.replaceAll('"', "").replaceAll("<", "").replaceAll(">", "").trim();
}

// Guard if a level entry is ever absent (callers pass a full 9-entry array).
const FALLBACK_LEVEL: LevelStyle = {
  color: "#000000",
  font: "Arial",
  size: 11,
  bold: false,
};

/**
 * Wrap a rendered pivot fragment as a minimal Word-flavored HTML document for the
 * clipboard. The title (`<p class="ws-title">`) becomes `MsoTitle` and the nested
 * rows (`<p class="ws-lvl" data-level="N">`) become `MsoPiv*`.
 *
 * The TITLE maps to a Word heading style when `headingStyleName` is set (it
 * carries `mso-style-name:"<name>";mso-outline-level:1`, so a "Use Destination
 * Styles" paste adopts the destination document's heading + its numbering);
 * blank = the app's direct level-1 look.
 *
 * The BODY always uses the app's direct per-level look (color/font/size/bold) +
 * indent + compact spacing, emitted as INLINE formatting on each `<p>` (and `<b>`
 * for bold) rather than CSS classes. A "Use Destination Styles" paste discards
 * class/style-name formatting it can't map (it resets unknown classes like
 * `MsoPiv2` to Normal -- dropping bold and bringing Normal's space-after) but
 * KEEPS inline direct formatting, so the per-level look + tight spacing survive
 * and match the live preview.
 */
export function buildWordHtml(
  fragment: string,
  heading: HeadingStyle,
  bodyFont: string,
): string {
  const lvl = (i: number) => heading.levels[i] ?? FALLBACK_LEVEL;
  const step = heading.indentStep ?? 0.2;
  const headingName = sanitizeStyleName(heading.headingStyleName ?? "");

  // Inline direct formatting for an app-styled paragraph: level index `i`, indented
  // for render-level `n` (1-based). Inline (not a CSS class) so it survives a
  // "Use Destination Styles" paste. Compact spacing = no space before/after +
  // a 1.15 line, so paragraphs sit tight like the preview.
  const directStyle = (i: number, n: number) => {
    const s = lvl(i);
    const indent = ((n - 1) * step).toFixed(2);
    return (
      `margin-top:0in;margin-bottom:0in;margin-left:${indent}in;line-height:115%;` +
      `color:${s.color};font-family:'${s.font}';font-size:${s.size}pt`
    );
  };
  // Bold is a <b> run (direct character formatting), which a "Use Destination
  // Styles" paste keeps -- unlike a class-level font-weight, which it discards.
  const wrapBold = (i: number, content: string) =>
    lvl(i).bold ? `<b>${content}</b>` : content;

  const body = fragment
    // Title: a mapped Word heading (class + mso rule) when named, else the app's
    // direct level-1 look inline.
    .replace(/<p class="ws-title">([\s\S]*?)<\/p>/g, (_m, content: string) =>
      headingName
        ? `<p class="MsoTitle">${content}</p>`
        : `<p style="${directStyle(0, 1)}">${wrapBold(0, content)}</p>`,
    )
    // Nested rows: always the app's per-level look, inline (+ <b> when bold).
    .replace(
      /<p class="ws-lvl" data-level="([1-9])">([\s\S]*?)<\/p>/g,
      (_m, d: string, content: string) => {
        const i = Number(d) - 1;
        return `<p style="${directStyle(i, Number(d))}">${wrapBold(i, content)}</p>`;
      },
    );

  // The only CSS rule needed is the mapped-title style; everything else is inline.
  const titleRule = headingName
    ? `p.MsoTitle{mso-style-name:"${headingName}";mso-outline-level:1}`
    : "";
  return (
    `<html xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:w="urn:schemas-microsoft-com:office:word" ` +
    `xmlns="http://www.w3.org/TR/REC-html40">` +
    `<head><meta charset="utf-8">` +
    `<style>` +
    `@page{size:8.5in 11in;margin:1in}` +
    // Body font fallback (so Word doesn't drop to Times New Roman); each paragraph
    // also sets its own font inline.
    `body{overflow-wrap:break-word;font-family:"${bodyFont}"}` +
    titleRule +
    `</style>` +
    `</head><body>${body}</body></html>`
  );
}

/**
 * Readable plain-text fallback for the `text/plain` clipboard flavor.
 *
 * Operates on the machine-generated fragment shape (all `<p>` plus escaped
 * `& < >`). Entities are unescaped lt/gt BEFORE amp -- the inverse of the
 * renderer's "amp first" escape order -- so `&amp;lt;` decodes to `&lt;`, never `<`.
 */
export function htmlToPlainText(fragment: string): string {
  return fragment
    .replace(/<\/p>/g, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/\n{2,}/g, "\n")
    .trim();
}
