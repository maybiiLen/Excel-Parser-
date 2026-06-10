import type { HeadingStyle } from "@/lib/clipboard";

/**
 * Read-only display of the rendered section HTML produced by `renderTree`.
 *
 * The HTML is injected via `dangerouslySetInnerHTML`, which is safe here:
 * `renderTree` escapes every user-derived value (titles, content, bullet items,
 * table cells), so the only raw markup is the machine-generated tags, dotted
 * numbers, and constant `style` attributes -- no user value lands in an
 * attribute, so there is no injection surface.
 *
 * Tailwind v4 preflight strips heading sizes and list bullets. We restore
 * spacing/lists with arbitrary-descendant variants, and the heading look
 * (color/font/size/weight) from a scoped `<style>` so the preview matches what a
 * "Copy for Word" + keep-source paste will produce. The heading-style values are
 * sanitized in the form (hex color, allow-listed font, clamped sizes).
 */
const previewClasses =
  "ws-preview rounded-lg border border-foreground/15 bg-foreground/[0.03] p-4 text-sm " +
  "[&_h2]:mt-4 [&_h3]:mt-3 [&_[data-level]]:mt-2 " +
  "[&_p]:my-2 [&_ul]:my-2 [&_ul]:list-disc [&_ul]:pl-6 " +
  "[&_table]:my-2 [&_td]:align-top";

export function RenderedPreview({
  html,
  emptyHint,
  headingStyle,
  bodyFont,
}: {
  html: string;
  emptyHint?: string;
  headingStyle: HeadingStyle;
  bodyFont: string;
}) {
  if (html === "") {
    return (
      <p className="text-sm text-foreground/60">
        {emptyHint ??
          "Pasted, but nothing was rendered. Switch to JSON to inspect the raw grid."}
      </p>
    );
  }

  const h = headingStyle;
  const weight = h.bold ? 700 : 400;
  // Pivot (nested-rows) levels 1-9: each level's own color/font/size/bold, plus a
  // growing left indent. Mirrors the MsoPiv* rules in buildWordHtml so the
  // preview matches the Word output. Per-level look comes from h.levels[n-1].
  const levelCss = Array.from({ length: 9 }, (_, i) => {
    const n = i + 1;
    const lv = h.levels[i] ?? {
      color: h.color,
      font: h.font,
      size: n === 1 ? h.h1Size : h.h2Size,
      bold: h.bold,
    };
    const lvWeight = lv.bold ? 700 : 400;
    const indent = ((n - 1) * 0.2).toFixed(1);
    return `.ws-preview [data-level="${n}"]{color:${lv.color};font-family:'${lv.font}';font-size:${lv.size}pt;font-weight:${lvWeight};margin-left:${indent}in}`;
  }).join("");
  // Body text inherits the container font; headings override it below.
  const css =
    `.ws-preview h2{color:${h.color};font-family:'${h.font}';font-size:${h.h1Size}pt;font-weight:${weight}}` +
    `.ws-preview h3{color:${h.color};font-family:'${h.font}';font-size:${h.h2Size}pt;font-weight:${weight}}` +
    levelCss;

  return (
    <div
      aria-label="Rendered section preview"
      className={previewClasses}
      style={{ fontFamily: bodyFont }}
    >
      <style>{css}</style>
      <div dangerouslySetInnerHTML={{ __html: html }} />
    </div>
  );
}
