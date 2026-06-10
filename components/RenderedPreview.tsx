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
  "[&_h2]:mt-4 [&_h3]:mt-3 " +
  "[&_p]:my-2 [&_ul]:my-2 [&_ul]:list-disc [&_ul]:pl-6 " +
  "[&_table]:my-2 [&_td]:align-top";

export function RenderedPreview({
  html,
  emptyHint,
  headingStyle,
}: {
  html: string;
  emptyHint?: string;
  headingStyle: HeadingStyle;
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
  const css =
    `.ws-preview h2{color:${h.color};font-family:'${h.font}';font-size:${h.h1Size}pt;font-weight:${weight}}` +
    `.ws-preview h3{color:${h.color};font-family:'${h.font}';font-size:${h.h2Size}pt;font-weight:${weight}}`;

  return (
    <div aria-label="Rendered section preview" className={previewClasses}>
      <style>{css}</style>
      <div dangerouslySetInnerHTML={{ __html: html }} />
    </div>
  );
}
