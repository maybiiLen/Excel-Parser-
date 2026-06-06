import type { Grid, Section } from "./types";

/**
 * Convention mapper (NOT YET IMPLEMENTED -- deferred past Day 1).
 *
 * MVP input convention, where column position defines role:
 *   - Column A filled  -> start a new Section (its title)
 *   - Column A blank   -> the row is a Subsection of the Section above
 *   - Column B         -> Subsection title
 *   - Column C         -> Body content
 *   - Column D         -> Body type flag: "text" | "bullet" | "table"
 *
 * Walks the raw Grid and produces a Section tree. Numbers are assigned later by
 * `numberTree` in ./numbering.
 */
export function rowsToTree(rows: Grid): Section[] {
  // TODO: group rows into sections by column-A presence, build Subsection bodies.
  void rows;
  throw new Error("rowsToTree: not implemented yet");
}
