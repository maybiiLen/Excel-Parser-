import type { Body, PivotNode, Section } from "./types";

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
 * Body -> HTML renderer.
 *
 * Emits markup that pastes cleanly into Word as native elements:
 *   - { type: "text" }    -> <p>...</p>
 *   - { type: "bullets" } -> <ul><li>...</li></ul>
 *   - { type: "table" }   -> <table>...</table>  (inline borders so Word renders
 *                            it as a real table; all rows treated alike, no <th>)
 *
 * Pure: reads `body`, returns a string, mutates nothing. All user-derived text
 * is escaped. A later sprint adds the wide-table width strategy (transpose or
 * split) so tables fit an 8.5x11 Word page -- not handled here.
 */
export function renderBody(body: Body): string {
  switch (body.type) {
    case "text":
      return `<p>${escapeHtml(body.content)}</p>`;
    case "bullets":
      return `<ul>${body.items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
    case "table":
      return `<table style="border-collapse:collapse">${body.rows
        .map(
          (row) =>
            `<tr>${row
              .map((cell) => `<td style="border:1px solid #000;padding:4px">${escapeHtml(cell)}</td>`)
              .join("")}</tr>`,
        )
        .join("")}</table>`;
  }
}

/**
 * Render a numbered section tree as one HTML document fragment.
 *
 * Walks the tree in order, emitting a heading per node and the rendered body
 * per subsection:
 *   - section    -> <h2>{number} {title}</h2> (+ renderBody(body) if the section
 *                   carries one, e.g. the transpose layout's attribute table)
 *   - subsection -> <h3>{number} {title}</h3> followed by renderBody(body)
 *
 * A blank `number` is omitted (along with its trailing space), so an un-numbered
 * tree renders title-only headings. When present, the dotted `number` is
 * machine-generated (digits and dots) and left unescaped; only user-derived text
 * (titles, plus the body content) is escaped.
 *
 * Returns a fragment -- no <html>/<head>/<meta charset>/<body> wrapper. The
 * clipboard MIME/charset framing belongs to the later clipboard-output step.
 *
 * Pure: builds a local array and returns its join; the input tree is never
 * mutated.
 */
export function renderTree(sections: Section[]): string {
  const blocks: string[] = [];
  for (const section of sections) {
    const sNum = section.number ? `${section.number} ` : "";
    blocks.push(`<h2>${sNum}${escapeHtml(section.title)}</h2>`);
    if (section.body) blocks.push(renderBody(section.body));
    for (const sub of section.children) {
      const subNum = sub.number ? `${sub.number} ` : "";
      blocks.push(`<h3>${subNum}${escapeHtml(sub.title)}</h3>`);
      blocks.push(renderBody(sub.body));
    }
  }
  return blocks.join("\n");
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
 * The nesting depth rides in a `data-level` ATTRIBUTE, not the tag name (HTML
 * only has h1-h6, and a class like `pl-3` would collide with Tailwind padding
 * utilities); `RenderedPreview` styles `[data-level="N"]`. Level is 1-based and
 * clamped at 9; deeper nodes still render at level 9.
 *
 * Only user-derived text is escaped; the `class`/`data-level` values are machine
 * constants (a single digit), so there is no attribute-injection surface.
 * Returns a bare fragment, like `renderTree`. Pure: the input is never mutated.
 */
export function renderPivotTree(nodes: PivotNode[], title?: string): string {
  const blocks: string[] = [];
  const walk = (list: PivotNode[], level: number) => {
    const lvl = Math.min(level, 9);
    for (const node of list) {
      blocks.push(
        `<p class="ws-lvl" data-level="${lvl}">${escapeHtml(node.title)}</p>`,
      );
      if (node.children.length > 0) walk(node.children, level + 1);
    }
  };
  if (title) {
    blocks.push(`<h2>${escapeHtml(title)}</h2>`);
    walk(nodes, 2); // title is level 1; data nests beneath it
  } else {
    walk(nodes, 1);
  }
  return blocks.join("\n");
}
