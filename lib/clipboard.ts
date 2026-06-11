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
 * - `levels` -- per-depth direct look (index 0 = level 1), used when no Word
 *   style name maps that paragraph.
 * - `indentStep` -- left-indent added per nesting level (inches).
 * - `headingStyleName` / `bodyStyleName` -- optional Word style names. When set,
 *   the title gets `mso-style-name:"<headingStyleName>"` and every body paragraph
 *   `mso-style-name:"<bodyStyleName>"`, so a "Use Destination Styles" paste
 *   adopts the destination document's styles (the template controls font/size/
 *   color/spacing; the app only adds indent + markers). Blank = direct look.
 */
export type HeadingStyle = {
  levels: LevelStyle[];
  indentStep: number; // inches of left-indent per nesting level
  headingStyleName: string; // Word style for the title ("" = direct look)
  bodyStyleName: string; // Word style for the nested rows + details ("" = direct)
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
 * clipboard. The title (`<p class="ws-title">`) becomes `MsoTitle`, the nested
 * rows (`<p class="ws-lvl" data-level="N">`) `MsoPiv*`, and the detail lines
 * (`<p class="ws-detail" data-detail="N">`) `MsoDet*`.
 *
 * When a Word style NAME is given, that class carries `mso-style-name:"<name>"`
 * (title also `mso-outline-level:1`) and NO direct font/color, so a
 * "Use Destination Styles" paste adopts the destination document's style (the
 * template controls font/size/color). When the name is blank, the app's direct
 * per-level look is used instead. Body paragraphs are single-spaced; the title's
 * spacing comes from its heading style (mapped) or the direct look (blank). The
 * app always adds the left-indent (and the markers live in the text).
 */
export function buildWordHtml(
  fragment: string,
  heading: HeadingStyle,
  bodyFont: string,
): string {
  const body = fragment
    // Title -> MsoTitle; nested rows -> MsoPiv{N}; detail lines -> MsoDet{N}.
    // The N rewrites are anchored to a 1-9 digit so they never match a bare <p>.
    .replace(/<p class="ws-title">([\s\S]*?)<\/p>/g, '<p class="MsoTitle">$1</p>')
    .replace(
      /<p class="ws-lvl" data-level="([1-9])">([\s\S]*?)<\/p>/g,
      '<p class="MsoPiv$1">$2</p>',
    )
    .replace(
      /<p class="ws-detail" data-detail="([1-9])">([\s\S]*?)<\/p>/g,
      '<p class="MsoDet$1">$2</p>',
    );
  const lookOf = (lv: LevelStyle) =>
    `color:${lv.color};font-family:"${lv.font}";font-size:${lv.size}pt;font-weight:${lv.bold ? "bold" : "normal"}`;
  const lvl = (i: number) => heading.levels[i] ?? FALLBACK_LEVEL;
  const step = heading.indentStep ?? 0.2;
  const headingName = sanitizeStyleName(heading.headingStyleName ?? "");
  const bodyName = sanitizeStyleName(heading.bodyStyleName ?? "");
  const single = "line-height:1.0;margin-top:0;margin-bottom:0"; // single spacing
  // Title: a Word heading style when named (template controls the look + the one
  // outline entry), else the direct level-1 look.
  const titleRule = headingName
    ? `p.MsoTitle{mso-style-name:"${headingName}";mso-outline-level:1}`
    : `p.MsoTitle{${lookOf(lvl(0))};${single}}`;
  // Nested rows + details (levels 1-9). When `bodyName` is set they map to that
  // Word style (no direct font/color -> template wins); otherwise the per-level
  // direct look. The app always adds the growing left indent + single spacing.
  const pivotRules = Array.from({ length: 9 }, (_, i) => {
    const n = i + 1;
    const pivLeft = `margin-left:${((n - 1) * step).toFixed(2)}in`;
    const detLeft = `margin-left:${(n * step).toFixed(2)}in`;
    const pivLook = bodyName ? `mso-style-name:"${bodyName}"` : lookOf(lvl(i));
    const detLook = bodyName ? `mso-style-name:"${bodyName}"` : "";
    const piv = `p.MsoPiv${n}{${pivLook};${pivLeft};${single}}`;
    const det = `p.MsoDet${n}{${detLook ? detLook + ";" : ""}${detLeft};${single}}`;
    return piv + det;
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
