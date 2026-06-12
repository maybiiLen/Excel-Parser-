import type { Grid, PivotLine, PivotNode } from "./types";

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
 * The sort field for one bucket: the FIRST column in bucket order whose
 * `sortDirs` entry is set. Lets a stacked level be sorted by any one of its
 * fields; `null` (no column set) means leave that level's order untouched.
 */
function findSortCol(
  bucket: number[],
  sortDirs: Record<number, "asc" | "desc">,
): { col: number; dir: "asc" | "desc" } | null {
  for (const col of bucket) {
    const dir = sortDirs[col];
    if (dir) return { col, dir };
  }
  return null;
}

/** A node's value for a given column (the matching `PivotLine`, else ""). */
function lineValue(node: PivotNode, col: number): string {
  return node.lines.find((l) => l.col === col)?.value ?? "";
}

/**
 * Recursive post-pass: reorder each node's children (and the roots) by the sort
 * field for that child level. One comparator (`localeCompare` with `numeric:true,
 * sensitivity:"base"`) auto-handles numbers and text -- "2" before "10", text
 * alphabetical + case-insensitive -- reversed for "desc". A level with no sort
 * column set keeps first-seen order; the language-guaranteed STABLE sort keeps
 * equal-keyed siblings in first-seen order too. Mutates the freshly-built private
 * tree in place (safe: it is not shared); `levels`/`sortDirs` stay read-only.
 */
function sortTree(
  nodes: PivotNode[],
  levels: number[][],
  sortDirs: Record<number, "asc" | "desc">,
  depth: number,
): void {
  const sort = findSortCol(levels[depth] ?? [], sortDirs);
  if (sort) {
    const sign = sort.dir === "desc" ? -1 : 1;
    nodes.sort(
      (a, b) =>
        sign *
        lineValue(a, sort.col).localeCompare(lineValue(b, sort.col), undefined, {
          numeric: true,
          sensitivity: "base",
        }),
    );
  }
  for (const node of nodes) sortTree(node.children, levels, sortDirs, depth + 1);
}

/**
 * Pivot mapper (Excel "Rows area"): nest rows into an arbitrary-depth tree by an
 * ORDERED list of INDENT BUCKETS. `levels[0]` is the outermost indent level;
 * within each group rows nest by `levels[1]`; and so on. Each bucket holds one or
 * more columns: a bucket's group key is the COMBINATION of all its columns'
 * values, so rows merge into the same node only when they match across every
 * field in that bucket. A node's `lines` are the bucket's fields, each a
 * `PivotLine { col, name, value }` stacked at the same indent.
 *
 * Row 0 is the header (used as each field's label `name`). The label `name` (the
 * trimmed header) and the cell `value` are kept SEPARATE and raw; the renderer
 * joins/escapes/formats them ("Field name: value", or just the value when the
 * label is hidden or the header is blank). A blank value -> "(blank)".
 * First-seen order is preserved at every level: a JS array keeps push order, and
 * a per-level `Map` (held in a `WeakMap` keyed by node, discarded after the
 * build) is only a dedup index, never the ordered output.
 *
 * `sortDirs` (keyed by grid column) then drives an OPTIONAL post-pass that
 * reorders SIBLING groups at each level by the bucket's first sort-enabled field
 * (numeric + text aware, case-insensitive; "desc" reverses). A level with no
 * sort column keeps first-seen order; ties keep first-seen order (stable sort).
 *
 * Resilient like the other mappers: never throws (`row?.[col]` + `cellToString`).
 * Empty `levels`, or an empty/header-only grid, yields `[]`.
 */
export function rowsToPivotTree(
  rows: Grid,
  levels: number[][],
  sortDirs: Record<number, "asc" | "desc"> = {},
): PivotNode[] {
  if (levels.length === 0) return [];
  const [header = [], ...dataRows] = rows;
  const headers = header.map(cellToString);

  const roots: PivotNode[] = [];
  const rootIndex = new Map<string, PivotNode>();
  // Per-node child dedup index; kept off to the side so PivotNode stays clean,
  // and garbage-collected once the build finishes.
  const childIndex = new WeakMap<PivotNode, Map<string, PivotNode>>();

  for (const row of dataRows) {
    let siblings = roots;
    let index = rootIndex;
    for (const bucket of levels) {
      // A bucket's identity is the combination of all its fields' RAW trimmed
      // values. JSON.stringify gives a collision-free key (distinct value tuples
      // always serialize differently, quotes/brackets escaped). Keying on the raw
      // value -- empty string included -- keeps a truly-empty cell distinct from a
      // cell that literally contains the text "(blank)"; the sentinel is only a
      // DISPLAY value, applied to `lines` below. The key is internal; `lines`
      // carry the human labels.
      const raw = bucket.map((col) => cellToString(row?.[col]).trim());
      const key = JSON.stringify(raw);
      let node = index.get(key);
      if (!node) {
        // Keep label (header name) and value separate + raw; the renderer escapes
        // and formats them (show/hide/bold/underline the label) per field. A blank
        // cell shows as "(blank)".
        const lines: PivotLine[] = bucket.map((col, i) => ({
          col,
          name: headers[col]?.trim() ?? "",
          value: raw[i] || "(blank)",
        }));
        node = { lines, children: [] };
        siblings.push(node); // first sight -> preserve insertion order
        index.set(key, node);
        childIndex.set(node, new Map());
      }
      siblings = node.children;
      index = childIndex.get(node)!;
    }
  }

  // Reorder siblings by the per-level sort field (no-op when no dir is set).
  sortTree(roots, levels, sortDirs, 0);
  return roots;
}
