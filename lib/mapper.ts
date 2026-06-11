import type { Grid, PivotNode } from "./types";

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
 * Pivot mapper (Excel "Rows area"): nest rows into an arbitrary-depth tree by an
 * ORDERED list of INDENT BUCKETS. `levels[0]` is the outermost indent level;
 * within each group rows nest by `levels[1]`; and so on. Each bucket holds one or
 * more columns: a bucket's group key is the COMBINATION of all its columns'
 * values, so rows merge into the same node only when they match across every
 * field in that bucket. A node's `lines` are the bucket's fields rendered as
 * stacked "Field: value" lines at the same indent.
 *
 * Row 0 is the header (used to label each field). Each line reads as
 * `Field name: value` (e.g. `Origin: Brazil`), with the field name taken from
 * the header of that field's column; a blank value -> "(blank)" and a blank
 * header drops the prefix (just the value). First-seen order is preserved at
 * every level: a JS array keeps push order, and a per-level `Map` (held in a
 * `WeakMap` keyed by node, discarded after the build) is only a dedup index,
 * never the ordered output.
 *
 * Resilient like the other mappers: never throws (`row?.[col]` + `cellToString`).
 * Empty `levels`, or an empty/header-only grid, yields `[]`.
 */
export function rowsToPivotTree(rows: Grid, levels: number[][]): PivotNode[] {
  if (levels.length === 0) return [];
  const [header = [], ...dataRows] = rows;
  const headers = header.map(cellToString);
  const labelOf = (col: number) => {
    const name = headers[col]?.trim() ?? "";
    return (value: string) => (name ? `${name}: ${value}` : value);
  };

  const roots: PivotNode[] = [];
  const rootIndex = new Map<string, PivotNode>();
  // Per-node child dedup index; kept off to the side so PivotNode stays clean,
  // and garbage-collected once the build finishes.
  const childIndex = new WeakMap<PivotNode, Map<string, PivotNode>>();

  for (const row of dataRows) {
    let siblings = roots;
    let index = rootIndex;
    for (const bucket of levels) {
      // A bucket's identity is the combination of all its fields' values.
      // JSON.stringify gives a collision-free key (distinct value tuples always
      // serialize differently, quotes/brackets escaped). The key is internal;
      // `lines` below carry the human labels.
      const values = bucket.map(
        (col) => cellToString(row?.[col]).trim() || "(blank)",
      );
      const key = JSON.stringify(values);
      let node = index.get(key);
      if (!node) {
        const lines = bucket.map((col, i) => labelOf(col)(values[i]));
        node = { lines, children: [] };
        siblings.push(node); // first sight -> preserve insertion order
        index.set(key, node);
        childIndex.set(node, new Map());
      }
      siblings = node.children;
      index = childIndex.get(node)!;
    }
  }

  return roots;
}
