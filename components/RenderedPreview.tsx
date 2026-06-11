import type { HeadingStyle } from "@/lib/clipboard";

/**
 * Read-only display of the rendered pivot HTML produced by `renderPivotTree`.
 *
 * The HTML is injected via `dangerouslySetInnerHTML`, which is safe here:
 * `renderPivotTree` escapes every user-derived value (titles, field labels,
 * detail lines), so the only raw markup is machine-generated tags, numbering
 * markers, and constant `style` attributes -- no user value lands in an
 * attribute, so there is no injection surface.
 *
 * Tailwind v4 preflight strips heading sizes. We restore spacing with
 * arbitrary-descendant variants, and the heading look (color/font/size/weight)
 * from a scoped `<style>` so the preview matches what a "Copy for Word" +
 * keep-source paste will produce. The style values are sanitized in the form
 * (hex color, allow-listed font, clamped sizes).
 */
const previewClasses =
  "ws-preview rounded-lg border border-foreground/15 bg-foreground/[0.03] p-4 text-sm " +
  "[&_h2]:mt-4 [&_[data-level]]:mt-2 [&_p]:my-2";

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

  // Every heading's look comes from headingStyle.levels (the single style
  // source). Level 1 (index 0) styles the title (h2) + pivot level 1; levels 2-9
  // style the nested rows by depth. Indent applies to the pivot [data-level]
  // rules only.
  const FALLBACK = { color: "#2F5496", font: "Calibri Light", size: 13, bold: false };
  const rule = (sel: string, i: number, indentIn: string) => {
    const lv = headingStyle.levels[i] ?? FALLBACK;
    const margin = indentIn ? `;margin-left:${indentIn}in` : "";
    return `.ws-preview ${sel}{color:${lv.color};font-family:'${lv.font}';font-size:${lv.size}pt;font-weight:${lv.bold ? 700 : 400}${margin}}`;
  };
  const levelCss = Array.from({ length: 9 }, (_, i) =>
    rule(`[data-level="${i + 1}"]`, i, (i * 0.2).toFixed(1)),
  ).join("");
  // Body text inherits the container font; the title (h2) overrides it below.
  const css = rule("h2", 0, "") + levelCss;

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
