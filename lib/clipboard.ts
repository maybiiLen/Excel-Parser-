// Clipboard-output step: wrap the renderer's HTML fragment for a clean Word paste.
//
// `renderTree` deliberately returns a bare fragment (no <html>/<head>/<body>) and
// hands the MIME/charset framing to this module. We add ONLY the doc wrapper here
// -- no Tailwind classes, no inline heading styles -- so the bare <h2>/<h3>/<ul>
// tags map onto Word's native Heading/List styles on paste. The browser's
// ClipboardItem writes the Windows CF_HTML header for us; we never hand-build it.

/**
 * Wrap a rendered fragment in a minimal Word-friendly HTML document.
 *
 * `<meta charset="utf-8">` is the load-bearing line (keeps Word from mis-decoding
 * entities / non-ASCII text). The `@page` rule is an intent signal for the
 * "open as .htm" path -- on paste-into-an-existing-doc Word keeps its own page
 * setup, so the 8.5x11 fit really comes from the content being narrow block flow
 * (headings + wrapping bullet lists, no wide tables). `overflow-wrap` is a
 * best-effort guard for a pathological long unbroken value.
 */
export function buildWordHtml(fragment: string): string {
  return (
    `<!DOCTYPE html><html><head><meta charset="utf-8">` +
    `<style>@page{size:8.5in 11in;margin:1in}body{overflow-wrap:break-word}</style>` +
    `</head><body>${fragment}</body></html>`
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
