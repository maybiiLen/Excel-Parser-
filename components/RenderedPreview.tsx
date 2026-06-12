import type { HeadingStyle } from "@/lib/clipboard";

/**
 * Read-only display of the rendered pivot HTML produced by `renderPivotTree`.
 *
 * The HTML is injected via `dangerouslySetInnerHTML`, which is safe here:
 * `renderPivotTree` escapes every user-derived value (titles, field labels), so
 * the only raw markup is machine-generated tags, numbering markers, and constant
 * `style` attributes -- no user value lands in an attribute, so there is no
 * injection surface.
 *
 * Tailwind v4 preflight strips heading sizes. We restore spacing with
 * arbitrary-descendant variants, and the heading look (color/font/size/weight)
 * from a scoped `<style>` so the preview matches what a "Copy for Word" +
 * keep-source paste will produce. The style values are sanitized in the form
 * (hex color, allow-listed font, clamped sizes).
 */
const previewClasses =
  "ws-preview rounded-lg border border-foreground/15 bg-foreground/[0.03] p-4 text-sm " +
  "[&_.ws-title]:mt-1 [&_[data-level]]:mt-1 [&_p]:my-0.5 [&_p]:leading-tight";

export function RenderedPreview({
  html,
  emptyHint,
  headingStyle,
  bodyFont,
  numbered = false,
}: {
  html: string;
  emptyHint?: string;
  headingStyle: HeadingStyle;
  bodyFont: string;
  /** Whether any level is Word-numbered (shows a "#" placeholder + a hint). */
  numbered?: boolean;
}) {
  if (html === "") {
    return (
      <p className="text-sm text-foreground/60">
        {emptyHint ??
          "Pasted, but nothing was rendered. Switch to JSON to inspect the raw grid."}
      </p>
    );
  }

  // Approximate look from headingStyle.levels (level 1 = the title row; levels
  // 2-9 the nested rows by depth). When a Word style name is mapped the real
  // template look only shows on paste; this preview uses the per-level look.
  const FALLBACK = { color: "#000000", font: "Arial", size: 11, bold: false };
  const step = headingStyle.indentStep ?? 0.2;
  const rule = (sel: string, i: number, indentIn: string) => {
    const lv = headingStyle.levels[i] ?? FALLBACK;
    const margin = indentIn ? `;margin-left:${indentIn}in` : "";
    return `.ws-preview ${sel}{color:${lv.color};font-family:'${lv.font}';font-size:${lv.size}pt;font-weight:${lv.bold ? 700 : 400}${margin}}`;
  };
  // Title (level 1, no indent); nested rows (per-level look + (n-1)*step indent).
  // Numbered lines get a muted "#" placeholder where Word will insert the real
  // number on paste (the browser can't know the document's "5.1").
  const css =
    rule(".ws-title", 0, "") +
    Array.from({ length: 9 }, (_, i) =>
      rule(`[data-level="${i + 1}"]`, i, (i * step).toFixed(2)),
    ).join("") +
    `.ws-preview [data-heading]::before{content:"# ";opacity:0.4}`;

  return (
    <div>
      <div
        aria-label="Rendered section preview"
        className={previewClasses}
        style={{ fontFamily: bodyFont }}
      >
        <style>{css}</style>
        <div dangerouslySetInnerHTML={{ __html: html }} />
      </div>
      {numbered && (
        <p className="mt-1 text-xs text-foreground/50">
          <span className="opacity-40"># </span>= Word fills in the live number
          (5.1, 5.1.1) on paste with <strong>Use Destination Styles</strong>.
        </p>
      )}
    </div>
  );
}
