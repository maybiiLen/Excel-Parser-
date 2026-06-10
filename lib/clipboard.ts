// Word-output step: wrap the renderer's HTML fragment as Word-flavored HTML for
// the clipboard ("Copy for Word") and the downloadable .doc ("Download for Word")
// so headings land as native Word Heading styles, not just bold text.
//
// `renderTree` returns a bare fragment (no <html>/<head>/<body>) and hands the
// MIME/charset framing to this module. The browser's ClipboardItem writes the
// Windows CF_HTML header for us; we never hand-build it.

/**
 * Wrap a rendered fragment in a minimal Word-flavored HTML document. Used for
 * BOTH the "Copy for Word" clipboard write and the "Download for Word" .doc file
 * (which opens in Word, applying these styles reliably via file-open import).
 *
 * Headings are emitted as Word "MsoHeading" STYLE PARAGRAPHS, not <h1>/<h2>: a
 * bare <h1> carries intrinsic bold/size that Word keeps as DIRECT formatting on a
 * default "keep source" paste (masking the document's Heading style -- the cause
 * of the "bold not styled" symptom). A <p> that carries only the style NAME (via
 * `mso-style-name`, no visual props) gives Word nothing to keep, so it applies
 * the destination doc's built-in Heading style. The office/word XML namespaces
 * put Word into "read the mso-* hints" mode. Wrapper section -> "heading 1",
 * children -> "heading 2" (matching a numbered report: "5 Fruit Database" peers
 * with "4. Data Entry Section").
 *
 * Coupled to renderTree's attribute-less <h2>/<h3>: each is matched whole and
 * heading text is escaped (so it can't contain the literal tags); <p>/<ul>/<li>/
 * <table> are never matched. If renderTree emits heading attributes or a new
 * level, update these replacements.
 *
 * Caveat: a browser can only place HTML/text on the clipboard (never Word's
 * native/RTF formats), so a default Ctrl+V from another program may still keep
 * source formatting -- pasting with "Use Destination Styles", or opening the
 * downloaded .doc, applies the document's Heading styles reliably.
 */
/**
 * How the section headings should look, so a default "keep source" paste matches
 * the destination document. The `mso-style-name` link is kept regardless, so the
 * headings stay tagged as Word Heading 1/2 (visible in the outline); these props
 * are the source formatting a keep-source paste preserves.
 */
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
 * The single styling source for every heading, shared across all tables. Index
 * 0 = level 1 styles every top heading (`<h2>`/MsoHeading1 for the
 * grouped/list/sections views AND pivot level 1); index 1 = level 2 styles
 * subsections (`<h3>`/MsoHeading2 AND pivot level 2); indices 2-8 style the
 * deeper pivot levels (MsoPiv3..9).
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

export function buildWordHtml(
  fragment: string,
  heading: HeadingStyle,
  bodyFont: string,
): string {
  const body = fragment
    .replace(/<h2>([\s\S]*?)<\/h2>/g, '<p class="MsoHeading1">$1</p>')
    .replace(/<h3>([\s\S]*?)<\/h3>/g, '<p class="MsoHeading2">$1</p>')
    // Pivot (nested-rows) headings: data-level="N" -> Word Heading N. Anchored to
    // a 1-9 digit so it never matches a bare body <p>. See renderPivotTree.
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
    `p.MsoHeading2{mso-style-name:"heading 2";mso-outline-level:2;${lookOf(lvl(1))}}` +
    pivotRules +
    `</style>` +
    `</head><body>${body}</body></html>`
  );
}

/**
 * Readable plain-text fallback for the `text/plain` clipboard flavor.
 *
 * Operates on the machine-generated fragment shape (h2/h3/p/ul/li plus escaped
 * `& < >`). Entities are unescaped lt/gt BEFORE amp -- the inverse of
 * `renderTree`'s "amp first" escape order -- so `&amp;lt;` decodes to the literal
 * `&lt;`, never to `<`.
 */
export function htmlToPlainText(fragment: string): string {
  return fragment
    .replace(/<li>/g, "- ")
    .replace(/<\/(h2|h3|p|li|ul)>/g, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/\n{2,}/g, "\n")
    .trim();
}
