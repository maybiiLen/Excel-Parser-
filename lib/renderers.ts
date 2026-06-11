import type { PivotNode } from "./types";

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
 * (`<p class="ws-lvl" data-level="N">`) and leaf details (`<p class="ws-detail"
 * data-detail="N">`) map to a body style. When a title is present the nested
 * data starts at level 2 beneath it, otherwise at level 1.
 *
 * A node's `details` (leaf "Field: value" lines) render as `<p class="ws-detail"
 * data-detail="N">` body paragraphs -- not nesting levels, not headings; their
 * indent (one step past the leaf) is applied in the style layer, like the rows.
 * Each nested node gets a per-level marker (`markers[depth-1]`, falling back to
 * the legacy cycle), restarting per parent; the title and detail lines are never
 * marked.
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
): string {
  const blocks: string[] = [];
  const walk = (list: PivotNode[], level: number, depth: number) => {
    const lvl = Math.min(level, 9);
    const kind = markers[depth - 1] ?? defaultMarker(depth);
    list.forEach((node, i) => {
      const m = markerText(kind, i);
      const marker = m ? `${m} ` : "";
      blocks.push(
        `<p class="ws-lvl" data-level="${lvl}">${marker}${escapeHtml(node.title)}</p>`,
      );
      if (node.details) {
        for (const line of node.details) {
          blocks.push(
            `<p class="ws-detail" data-detail="${lvl}">${escapeHtml(line)}</p>`,
          );
        }
      }
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
