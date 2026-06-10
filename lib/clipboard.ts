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
 * Per-nesting-level look for the pivot (nested-rows) view -- one entry per depth
 * (index 0 = level 1). Shared across all pivot tables; defaults are all the same
 * so nesting reads by indent alone (Excel-like) until the user diverges a level.
 */
export type LevelStyle = {
  color: string; // hex
  font: string; // font-family name
  size: number; // pt
  bold: boolean;
};

export type HeadingStyle = {
  color: string; // hex, e.g. "#2F5496"
  font: string; // font-family name, e.g. "Calibri Light"
  h1Size: number; // pt, wrapper / Heading 1 (non-pivot views)
  h2Size: number; // pt, children / Heading 2 (non-pivot views)
  bold: boolean;
  /** Per-pivot-level look, index 0 = level 1; drives the MsoPiv* rules below. */
  levels: LevelStyle[];
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
  const weight = heading.bold ? "bold" : "normal";
  const look = (size: number) =>
    `color:${heading.color};font-family:"${heading.font}";font-size:${size}pt;font-weight:${weight}`;
  // Pivot heading levels 1-9. Their own class (not MsoHeading*) so the per-level
  // look + indent never leak onto grouped/list subsections, which reuse
  // MsoHeading2. Each still maps to Word's built-in "heading N" via
  // mso-style-name, so the document outline is correct; the left indent is direct
  // formatting kept on a keep-source paste. Each level's color/font/size/bold
  // comes from heading.levels[n-1] (falls back to the base look); indent grows.
  const pivotRules = Array.from({ length: 9 }, (_, i) => {
    const n = i + 1;
    const lv = heading.levels[i] ?? {
      color: heading.color,
      font: heading.font,
      size: n === 1 ? heading.h1Size : heading.h2Size,
      bold: heading.bold,
    };
    const lvLook = `color:${lv.color};font-family:"${lv.font}";font-size:${lv.size}pt;font-weight:${lv.bold ? "bold" : "normal"}`;
    const indent = ((n - 1) * 0.2).toFixed(1);
    return `p.MsoPiv${n}{mso-style-name:"heading ${n}";mso-outline-level:${n};${lvLook};margin-left:${indent}in}`;
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
    `p.MsoHeading1{mso-style-name:"heading 1";mso-outline-level:1;${look(heading.h1Size)}}` +
    `p.MsoHeading2{mso-style-name:"heading 2";mso-outline-level:2;${look(heading.h2Size)}}` +
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
