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

/**
 * Header-aware transpose mapper for wide catalogs (header row + one row per item).
 *
 * Each data row becomes a top-level Section titled by the chosen `titleColumn`,
 * with the selected `fieldColumns` rendered as a bulleted "Field: value" list
 * held on the section body. Row 0 supplies the field names. The caller picks and
 * orders `fieldColumns` (excluding the title column and any unchecked columns).
 *
 * Resilient like `rowsToTree`: never throws. Rows whose title cell is blank are
 * skipped (same "only explicitly-titled sections" policy), so an only-header or
 * empty grid yields `[]`. Empty `fieldColumns` yields sections with no bullets.
 * The `number` field is left `""` for `numberTree`.
 */
export function rowsToAttributeSections(
  rows: Grid,
  titleColumn: number,
  fieldColumns: number[],
): Section[] {
  const [header = [], ...dataRows] = rows;
  const headers = header.map(cellToString);

  const sections: Section[] = [];
  for (const row of dataRows) {
    const title = cellToString(row?.[titleColumn]).trim();
    if (title === "") continue;

    const items = fieldColumns.map(
      (c) => `${headers[c]}: ${cellToString(row?.[c])}`,
    );

    sections.push({
      number: "",
      title,
      children: [],
      body: { type: "bullets", items },
    });
  }

  return sections;
}

/**
 * Group-by mapper (pivot-style): one Section per distinct value of `groupColumn`,
 * with the rows that share it listed as bullets under that heading.
 *
 * Row 0 supplies field names. Each data row is bucketed by the trimmed value of
 * its `groupColumn` cell; a blank group cell buckets under "(blank)" and a blank
 * label becomes "(untitled)", so NO row is silently dropped (unlike the per-item
 * mappers, where a blank title leaves nothing renderable -- here the group
 * section already exists). First-seen group order and within-group row order are
 * preserved (a Map keeps insertion order). Each member bullet is the `labelColumn`
 * value plus, when `memberFieldColumns` is non-empty, a parenthetical of those
 * values joined by ", " (values only).
 *
 * Resilient like the other mappers: never throws. An empty or header-only grid
 * yields `[]`. The `number` field is left "" (numbering omitted by renderTree).
 */
export function rowsToGroupedSections(
  rows: Grid,
  groupColumn: number,
  labelColumn: number,
  memberFieldColumns: number[],
): Section[] {
  const [, ...dataRows] = rows;

  const groups = new Map<string, string[]>();
  for (const row of dataRows) {
    const groupValue = cellToString(row?.[groupColumn]).trim() || "(blank)";
    const label = cellToString(row?.[labelColumn]).trim() || "(untitled)";
    const extras = memberFieldColumns.map((c) => cellToString(row?.[c]));
    const item = extras.length ? `${label} (${extras.join(", ")})` : label;

    const bucket = groups.get(groupValue);
    if (bucket) bucket.push(item);
    else groups.set(groupValue, [item]);
  }

  const sections: Section[] = [];
  for (const [title, items] of groups) {
    sections.push({
      number: "",
      title,
      children: [],
      body: { type: "bullets", items },
    });
  }

  return sections;
}
