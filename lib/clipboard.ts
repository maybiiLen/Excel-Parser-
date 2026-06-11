// Word-output step: wrap the renderer's HTML fragment as Word-flavored HTML for
// the clipboard ("Copy for Word"). The pivot title lands as a native Word Heading
// (so it shows in the outline); the nested rows + details are styled body
// paragraphs. renderPivotTree returns a bare fragment (no <html>/<head>/<body>);
// the browser's ClipboardItem writes the Windows CF_HTML header for us.

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
 * The single styling source, shared across all tables. Index 0 = level 1 styles
 * the pivot title (`<h2>` -> MsoHeading1, the one outline heading); indices 1-8
 * style the nested pivot rows (MsoPiv2..9 -- styled BODY paragraphs, not
 * headings). buildWordHtml reads every look from here.
 */
export type HeadingStyle = {
  levels: LevelStyle[];
};

// Guard if a level entry is ever absent (callers pass a full 9-entry array).
const FALLBACK_LEVEL: LevelStyle = {
  color: "#2F5496",
  font: "Calibri Light",
  size: 13,
  bold: false,
};

/**
 * Wrap a rendered pivot fragment as a minimal Word-flavored HTML document for the
 * clipboard. The `<h2>` pivot title becomes a `MsoHeading1` STYLE PARAGRAPH (only
 * the `mso-style-name`/`mso-outline-level`, no visual props Word would keep as
 * direct formatting), so on a "Use Destination Styles" paste Word applies its own
 * Heading 1 -- and it's the one entry in the navigation outline. The nested rows
 * (`<p class="ws-lvl" data-level="N">`) become non-heading `MsoPiv*` styled body
 * paragraphs. Detail `<p style=...>` lines pass through untouched as body text.
 *
 * Caveat: a browser can only place HTML/text on the clipboard (never Word's
 * native/RTF), so a default Ctrl+V may keep source formatting -- pasting with
 * "Use Destination Styles" applies the document's Heading styles reliably.
 */
export function buildWordHtml(
  fragment: string,
  heading: HeadingStyle,
  bodyFont: string,
): string {
  const body = fragment
    .replace(/<h2>([\s\S]*?)<\/h2>/g, '<p class="MsoHeading1">$1</p>')
    // Pivot (nested-rows) levels: data-level="N" -> MsoPiv{N}. Anchored to a 1-9
    // digit so it never matches a bare body <p>. See renderPivotTree.
    .replace(
      /<p class="ws-lvl" data-level="([1-9])">([\s\S]*?)<\/p>/g,
      '<p class="MsoPiv$1">$2</p>',
    );
  // Every heading's look comes from heading.levels (the single style source).
  const lookOf = (lv: LevelStyle) =>
    `color:${lv.color};font-family:"${lv.font}";font-size:${lv.size}pt;font-weight:${lv.bold ? "bold" : "normal"}`;
  const lvl = (i: number) => heading.levels[i] ?? FALLBACK_LEVEL;
  // Pivot levels 1-9. Deliberately NOT Word headings: they carry no
  // mso-style-name / mso-outline-level, so the nested rows are plain styled
  // BODY paragraphs and don't flood Word's navigation outline (only the pivot
  // title -- emitted as <h2> -> MsoHeading1 -- is a heading). Look comes from
  // heading.levels[n-1]; the growing left indent shows the nesting.
  const pivotRules = Array.from({ length: 9 }, (_, i) => {
    const n = i + 1;
    const indent = ((n - 1) * 0.2).toFixed(1);
    return `p.MsoPiv${n}{${lookOf(lvl(i))};margin-left:${indent}in}`;
  }).join("");
  return (
    `<html xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:w="urn:schemas-microsoft-com:office:word" ` +
    `xmlns="http://www.w3.org/TR/REC-html40">` +
    `<head><meta charset="utf-8">` +
    `<style>` +
    `@page{size:8.5in 11in;margin:1in}` +
    // Set a body font so Word doesn't fall back to Times New Roman for unstyled
    // HTML body text. Headings override this with their own font below.
    `body{overflow-wrap:break-word;font-family:"${bodyFont}"}` +
    `p.MsoHeading1{mso-style-name:"heading 1";mso-outline-level:1;${lookOf(lvl(0))}}` +
    pivotRules +
    `</style>` +
    `</head><body>${body}</body></html>`
  );
}

/**
 * Readable plain-text fallback for the `text/plain` clipboard flavor.
 *
 * Operates on the machine-generated fragment shape (h2/p plus escaped `& < >`).
 * Entities are unescaped lt/gt BEFORE amp -- the inverse of the renderer's "amp
 * first" escape order -- so `&amp;lt;` decodes to the literal `&lt;`, never `<`.
 */
export function htmlToPlainText(fragment: string): string {
  return fragment
    .replace(/<\/(h2|p)>/g, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/\n{2,}/g, "\n")
    .trim();
}
