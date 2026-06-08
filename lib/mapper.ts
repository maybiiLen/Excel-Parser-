import type { Grid, Section, Subsection, Body } from "./types";

/**
 * Coerce a raw grid cell to a string.
 *
 * Cells are typed `string | number | boolean | null` (and a short row may have
 * `undefined` past its last column), so we normalize to a string before any
 * `trim`/`split`. Blank-ish cells (`null`/`undefined`) become `""`.
 */
function cellToString(cell: unknown): string {
  return cell == null ? "" : String(cell);
}

/**
 * Build a Subsection Body from the content cell (C) and the type flag (D).
 *
 * The flag is matched case-insensitively after trimming. Anything empty or
 * unrecognized falls back to `text` -- this never throws.
 */
function buildBody(content: string, flag: string): Body {
  switch (flag.trim().toLowerCase()) {
    case "bullet":
      // Split on newline, trim each item, drop empties.
      return {
        type: "bullets",
        items: content
          .split(/\r?\n/)
          .map((item) => item.trim())
          .filter((item) => item !== ""),
      };
    case "table":
      // Newlines delimit rows, tabs delimit cells. Kept structural -- no filtering.
      return {
        type: "table",
        rows: content.split(/\r?\n/).map((row) => row.split("\t")),
      };
    case "text":
    default:
      // Empty or unrecognized flag defaults to text; content kept as-is.
      return { type: "text", content };
  }
}

/**
 * Convention mapper: walk the raw Grid and produce a Section tree.
 *
 * MVP input convention, where column position defines role:
 *   - Column A filled  -> start a new Section (its title)
 *   - Column A blank   -> the row is a Subsection of the Section above
 *   - Column B         -> Subsection title
 *   - Column C         -> Body content
 *   - Column D         -> Body type flag: "text" | "bullet" | "table"
 *
 * This is resilient by design: it never throws on malformed input. The `number`
 * field is left as `""` on every node -- numbering is owned by `numberTree` in
 * ./numbering and is intentionally not touched here.
 */
export function rowsToTree(rows: Grid): Section[] {
  const sections: Section[] = [];
  let current: Section | null = null;

  for (const row of rows) {
    const a = cellToString(row?.[0]);

    if (a.trim() !== "") {
      // Non-empty column A starts a new section.
      current = { number: "", title: a.trim(), children: [] };
      sections.push(current);
      continue;
    }

    // Blank column A => subsection of the most recent section.
    // TODO: orphan rows (blank A before any section) are skipped for MVP.
    // Alternative considered: attach them to an implicit untitled section. We
    // chose to skip so the tree only contains explicitly-titled sections.
    if (!current) continue;

    const title = cellToString(row?.[1]).trim();
    const content = cellToString(row?.[2]);
    const flag = cellToString(row?.[3]);

    const subsection: Subsection = {
      number: "",
      title,
      body: buildBody(content, flag),
    };
    current.children.push(subsection);
  }

  return sections;
}
