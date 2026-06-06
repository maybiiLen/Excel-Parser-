import type { Section } from "./types";

/**
 * Numbering walk (NOT YET IMPLEMENTED -- deferred past Day 1).
 *
 * Assigns dotted numbers from a configurable root. Root 6 yields:
 *   6, 6.1, 6.2, 7, 7.1, ...
 * Each section gets a single integer (starting at `root`); each subsection gets
 * `${section.number}.${index}`.
 *
 * Returns the same tree with every `number` field populated.
 */
export function numberTree(sections: Section[], root = 6): Section[] {
  // TODO: walk sections, assign section.number then subsection.number.
  void sections;
  void root;
  throw new Error("numberTree: not implemented yet");
}
