import type { Section } from "./types";

/**
 * Assign dotted numbers from a configurable root, returning a NEW tree.
 *
 * Root 6 yields: 6, 6.1, 6.2, 7, 7.1, ...
 * Each section gets a single integer (starting at `root`); each subsection gets
 * `${section.number}.${index}` with the index starting at 1.
 *
 * The input is never mutated -- fresh section objects, `children` arrays, and
 * subsection objects are returned, so the mapper's output (with its `""` numbers)
 * is left untouched. The tree is exactly two levels deep; no grandchildren.
 *
 * TODO: `root` is assumed a positive integer; no validation for negative or
 * fractional roots today.
 */
export function numberTree(sections: Section[], root = 6): Section[] {
  return sections.map((section, i) => {
    const number = String(root + i);
    return {
      ...section,
      number,
      children: section.children.map((sub, j) => ({
        ...sub,
        number: `${number}.${j + 1}`,
      })),
    };
  });
}
