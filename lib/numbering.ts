import type { Section } from "./types";

/**
 * Wrap mapper output (a flat list of Sections, each carrying a bullets body) as
 * the children of ONE numbered top-level section, returning a fresh tree.
 *
 * The chosen `sectionNumber`/`sectionTitle` become the top heading (e.g.
 * `5 Fruit Database`); each input item becomes a numbered subsection
 * `${sectionNumber}.${i + 1}` (5.1, 5.2, ...) carrying that item's body. This
 * lets the grouped and per-item views slot their items into a larger Word
 * document at a chosen section number.
 *
 * Empty input yields `[]` (renders as no HTML, so the empty hint still shows).
 * Pure: the input is never mutated; titles are left unescaped (renderTree
 * escapes them). Not used by the A/B/C/D view, which is already two levels deep.
 */
export function wrapInNumberedSection(
  items: Section[],
  sectionNumber: number,
  sectionTitle: string,
): Section[] {
  if (items.length === 0) return [];
  return [
    {
      number: String(sectionNumber),
      title: sectionTitle,
      children: items.map((item, i) => ({
        number: `${sectionNumber}.${i + 1}`,
        title: item.title,
        // grouped/list items always set a bullets body; `??` only satisfies the type.
        body: item.body ?? { type: "text", content: "" },
      })),
    },
  ];
}
