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

/** Multilevel marker by nesting depth: 1./2. -> a./b. -> i./ii. -> cycle. */
function pivotMarker(depth: number, index: number): string {
  const kind = (depth - 1) % 3;
  if (kind === 1) return `${toAlpha(index)}.`;
  if (kind === 2) return `${toRoman(index + 1)}.`;
  return `${index + 1}.`;
}

/**
 * Render a pivot (nested-rows) tree as one HTML document fragment.
 *
 * Only the optional `title` is a real Word heading: it is emitted as `<h2>` (->
 * MsoHeading1), the single outline entry for the section. The nested data nodes
 * are emitted as `<p class="ws-lvl" data-level="N">` -- styled, indented BODY
 * paragraphs (`buildWordHtml` maps them to non-heading `MsoPiv{N}` classes), so
 * they do NOT clutter Word's navigation outline. When a title is present the
 * data starts at level 2 (the title occupies level 1); otherwise it starts at
 * level 1 and there is no heading at all.
 *
 * A node's `details` (leaf "Field: value" lines) render as plain `<p>` body
 * paragraphs indented one step past the leaf -- not nesting levels, not headings.
 * When `numbered`, each nested node gets a depth-based marker (`1.`/`a.`/`i.`,
 * restarting per parent); the title and detail lines are never numbered.
 *
 * The nesting depth rides in a `data-level` ATTRIBUTE, not the tag name (HTML
 * only has h1-h6, and a class like `pl-3` would collide with Tailwind padding
 * utilities); `RenderedPreview` styles `[data-level="N"]`. Level is 1-based and
 * clamped at 9; deeper nodes still render at level 9.
 *
 * Only user-derived text is escaped; the `class`/`data-level`/`style` values are
 * machine constants (digits, a computed indent), so there is no attribute-
 * injection surface. Returns a bare fragment. Pure: the input is never mutated.
 */
export function renderPivotTree(
  nodes: PivotNode[],
  title?: string,
  numbered = false,
): string {
  const blocks: string[] = [];
  const walk = (list: PivotNode[], level: number, depth: number) => {
    const lvl = Math.min(level, 9);
    list.forEach((node, i) => {
      const marker = numbered ? `${pivotMarker(depth, i)} ` : "";
      blocks.push(
        `<p class="ws-lvl" data-level="${lvl}">${marker}${escapeHtml(node.title)}</p>`,
      );
      if (node.details) {
        const indent = (lvl * 0.2).toFixed(1);
        for (const line of node.details) {
          blocks.push(
            `<p style="margin-left:${indent}in">${escapeHtml(line)}</p>`,
          );
        }
      }
      if (node.children.length > 0) walk(node.children, level + 1, depth + 1);
    });
  };
  if (title) {
    blocks.push(`<h2>${escapeHtml(title)}</h2>`);
    walk(nodes, 2, 1); // title is level 1; data nests beneath it
  } else {
    walk(nodes, 1, 1);
  }
  return blocks.join("\n");
}
