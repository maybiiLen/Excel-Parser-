import { DEFAULT_FIELD_LABEL, type FieldLabel, type PivotNode } from "./types";

/**
 * Escape the three markup-significant characters in user-derived text.
 *
 * `&` is replaced first so the entities introduced for `<`/`>` are not
 * double-escaped (e.g. `<` -> `&lt;`, never `&amp;lt;`).
 *
 * `"` / `'` are intentionally NOT escaped: every value we emit lands in element
 * text content, never inside an attribute (all `style="..."` attributes are
 * machine-generated constants), so the 3-character set is sufficient and keeps
 * the Word paste clean.
 */
function escapeHtml(s: string): string {
  return s.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

/**
 * Build the label prefix `Field name: ` for one field, escaped and optionally
 * wrapped in `<b>`/`<u>` per its `FieldLabel`. The `<b>`/`<u>` are inline runs
 * (they survive a Word "Use Destination Styles" paste); the name is escaped, the
 * tags + separator are machine constants.
 */
function wrapLabel(name: string, lf: FieldLabel): string {
  let html = `${escapeHtml(name)}: `;
  if (lf.underline) html = `<u>${html}</u>`;
  if (lf.bold) html = `<b>${html}</b>`;
  return html;
}

/** Bijective base-26: 0 -> "a", 25 -> "z", 26 -> "aa". */
function toAlpha(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(97 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const ROMAN: [number, string][] = [
  [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
  [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
];
/** Lowercase roman numeral for n >= 1. */
function toRoman(n: number): string {
  let s = "";
  for (const [v, sym] of ROMAN) {
    while (n >= v) {
      s += sym;
      n -= v;
    }
  }
  return s;
}

/** A per-level marker style the user can pick for each nesting depth. */
export type MarkerKind =
  | "decimal" // 1.
  | "paren" // 1)
  | "upperAlpha" // A.
  | "lowerAlpha" // a.
  | "upperRoman" // I.
  | "lowerRoman" // i.
  | "bullet" // •
  | "dash" // –
  | "none"; // (no marker)

/** The marker text for a 0-based sibling `index` in the given `kind`. */
export function markerText(kind: MarkerKind, index: number): string {
  switch (kind) {
    case "decimal":
      return `${index + 1}.`;
    case "paren":
      return `${index + 1})`;
    case "upperAlpha":
      return `${toAlpha(index).toUpperCase()}.`;
    case "lowerAlpha":
      return `${toAlpha(index)}.`;
    case "upperRoman":
      return `${toRoman(index + 1).toUpperCase()}.`;
    case "lowerRoman":
      return `${toRoman(index + 1)}.`;
    case "bullet":
      return "•";
    case "dash":
      return "–";
    case "none":
      return "";
  }
}

/** Legacy default marker for a nesting depth: 1. -> a. -> i. -> cycle. */
export function defaultMarker(depth: number): MarkerKind {
  return (["decimal", "lowerAlpha", "lowerRoman"] as const)[(depth - 1) % 3];
}

/**
 * Render a pivot (nested-rows) tree as one HTML document fragment.
 *
 * The optional `title` is emitted as a distinct `<p class="ws-title">` so
 * `buildWordHtml` can map it to a Word heading style; the nested rows
 * (`<p class="ws-lvl" data-level="N">`) map to a body style. When a title is
 * present the nested data starts at level 2 beneath it, otherwise at level 1.
 *
 * Each node carries one or more `lines` (`PivotLine` fields stacked at that
 * indent level). For each line the label (`name: `) is shown/bolded/underlined
 * per `fieldLabels[col]` (default: shown, plain) and the value follows; `name`
 * and `value` are escaped separately. Only the FIRST line of a node gets the
 * level's marker (`markers[depth-1]`, falling back to the legacy cycle, counting
 * up per parent); the rest read as plain body lines. The title is never marked.
 *
 * NUMBERING: when `numberDepth > 0`, the FIRST line of each bucket at body depth
 * `<= numberDepth` is tagged `data-heading="K"` (K = `(title?1:0) + depth`,
 * clamped 9) so `buildWordHtml` maps it to `Heading K` and Word supplies the live
 * number (e.g. 5.1, 5.1.1). The app's own text marker is suppressed on those
 * lines (Word owns the number); deeper levels keep their markers.
 *
 * The nesting depth rides in a `data-level` ATTRIBUTE, not the tag name (HTML
 * only has h1-h6, and a class like `pl-3` would collide with Tailwind padding
 * utilities); `RenderedPreview` styles `[data-level="N"]`. Level is 1-based and
 * clamped at 9; deeper nodes still render at level 9.
 *
 * Only user-derived text is escaped; the `class`/`data-*` values are machine
 * constants (single digits), so there is no attribute-injection surface. Returns
 * a bare fragment. Pure: the input is never mutated.
 */
export function renderPivotTree(
  nodes: PivotNode[],
  title?: string,
  markers: MarkerKind[] = [],
  fieldLabels: Record<number, FieldLabel> = {},
  numberDepth = 0,
): string {
  const blocks: string[] = [];
  const walk = (list: PivotNode[], level: number, depth: number) => {
    const lvl = Math.min(level, 9);
    const kind = markers[depth - 1] ?? defaultMarker(depth);
    // Word-numbered levels: the bucket's first line becomes a Heading; Word makes
    // the number, so the app's text marker is suppressed there.
    const numbered = depth <= numberDepth;
    const headingK = Math.min((title ? 1 : 0) + depth, 9);
    list.forEach((node, i) => {
      const m = markerText(kind, i);
      node.lines.forEach((line, j) => {
        const lf = fieldLabels[line.col] ?? DEFAULT_FIELD_LABEL;
        const label =
          lf.show && line.name !== "" ? wrapLabel(line.name, lf) : "";
        const value = escapeHtml(line.value);
        const marker = j === 0 && m && !numbered ? `${m} ` : "";
        const headingAttr =
          j === 0 && numbered ? ` data-heading="${headingK}"` : "";
        blocks.push(
          `<p class="ws-lvl" data-level="${lvl}"${headingAttr}>${marker}${label}${value}</p>`,
        );
      });
      if (node.children.length > 0) walk(node.children, level + 1, depth + 1);
    });
  };
  if (title) {
    // Title is a distinct paragraph (ws-title) so buildWordHtml can map it to a
    // heading style; the data nests beneath it starting at level 2.
    blocks.push(`<p class="ws-title">${escapeHtml(title)}</p>`);
    walk(nodes, 2, 1);
  } else {
    walk(nodes, 1, 1);
  }
  return blocks.join("\n");
}
