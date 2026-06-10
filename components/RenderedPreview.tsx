/**
 * Read-only display of the rendered section HTML produced by `renderTree`.
 *
 * The HTML is injected via `dangerouslySetInnerHTML`, which is safe here:
 * `renderTree` escapes every user-derived value (titles, content, bullet items,
 * table cells), so the only raw markup is the machine-generated tags, dotted
 * numbers, and constant `style` attributes -- no user value lands in an
 * attribute, so there is no injection surface.
 *
 * Tailwind v4 preflight strips heading sizes and list bullets, so we restore
 * basic typography with arbitrary-descendant variants on the wrapper. The table
 * borders come from inline styles in `renderTree` and need no help here.
 */
const previewClasses =
  "rounded-lg border border-foreground/15 bg-foreground/[0.03] p-4 text-sm " +
  "[&_h2]:mt-4 [&_h2]:text-xl [&_h2]:font-bold " +
  "[&_h3]:mt-3 [&_h3]:text-base [&_h3]:font-semibold " +
  "[&_p]:my-2 [&_ul]:my-2 [&_ul]:list-disc [&_ul]:pl-6 " +
  "[&_table]:my-2 [&_td]:align-top";

export function RenderedPreview({
  html,
  emptyHint,
}: {
  html: string;
  emptyHint?: string;
}) {
  if (html === "") {
    return (
      <p className="text-sm text-foreground/60">
        {emptyHint ??
          "Pasted, but nothing was rendered. Switch to JSON to inspect the raw grid."}
      </p>
    );
  }

  return (
    <div
      aria-label="Rendered section preview"
      className={previewClasses}
      dangerouslySetInnerHTML={{ __html: html }}
    />
  );
}
