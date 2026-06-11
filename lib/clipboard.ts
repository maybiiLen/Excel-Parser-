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
 * carries `mso-style-name:"<name>";mso-outline-level:1` and NO direct font/color,
 * so a "Use Destination Styles" paste adopts the destination document's heading);
 * blank = the app's direct level-1 look. The BODY always uses the app's direct
 * per-level look (color/font/size/bold), so what the preview shows is what Word
 * gets. Body paragraphs are compactly spaced; the app adds the left-indent (and
 * the markers live in the text).
 */
export function buildWordHtml(
  fragment: string,
  heading: HeadingStyle,
  bodyFont: string,
): string {
  const body = fragment
    // Title -> MsoTitle; nested rows -> MsoPiv{N}. The N rewrite is anchored to a
    // 1-9 digit so it never matches a bare <p>.
    .replace(/<p class="ws-title">([\s\S]*?)<\/p>/g, '<p class="MsoTitle">$1</p>')
    .replace(
      /<p class="ws-lvl" data-level="([1-9])">([\s\S]*?)<\/p>/g,
      '<p class="MsoPiv$1">$2</p>',
    );
  const lookOf = (lv: LevelStyle) =>
    `color:${lv.color};font-family:"${lv.font}";font-size:${lv.size}pt;font-weight:${lv.bold ? "bold" : "normal"}`;
  const lvl = (i: number) => heading.levels[i] ?? FALLBACK_LEVEL;
  const step = heading.indentStep ?? 0.2;
  const headingName = sanitizeStyleName(heading.headingStyleName ?? "");
  // Compact spacing: a 1.25 line (matching the preview's leading-tight) with no
  // space before/after, so paragraphs sit tight under one another.
  const single = "line-height:1.25;margin-top:0;margin-bottom:0";
  // Title: a Word heading style when named (template controls the look + the one
  // outline entry), else the direct level-1 look.
  const titleRule = headingName
    ? `p.MsoTitle{mso-style-name:"${headingName}";mso-outline-level:1}`
    : `p.MsoTitle{${lookOf(lvl(0))};${single}}`;
  // Nested rows (levels 1-9): always the app's per-level direct look (color/font/
  // size/bold) + the growing left indent + compact spacing. No Word-style mapping,
  // so the bold and spacing the preview shows are exactly what Word receives.
  const pivotRules = Array.from({ length: 9 }, (_, i) => {
    const n = i + 1;
    const pivLeft = `margin-left:${((n - 1) * step).toFixed(2)}in`;
    return `p.MsoPiv${n}{${lookOf(lvl(i))};${pivLeft};${single}}`;
  }).join("");
  return (
    `<html xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:w="urn:schemas-microsoft-com:office:word" ` +
    `xmlns="http://www.w3.org/TR/REC-html40">` +
    `<head><meta charset="utf-8">` +
    `<style>` +
    `@page{size:8.5in 11in;margin:1in}` +
    // Body font (so Word doesn't fall back to Times New Roman). Per-paragraph
    // spacing/look is set by the title + pivot rules below.
    `body{overflow-wrap:break-word;font-family:"${bodyFont}"}` +
    titleRule +
    pivotRules +
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
