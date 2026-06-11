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
 * ORDERED list of columns. `orderedColumns[0]` is the outermost grouping; within
 * each group rows nest by `orderedColumns[1]`; and so on. Rows that share a full
 * value-path MERGE into the same branch (a pivot with only Row fields, no
 * Values), so duplicate paths collapse.
 *
 * Row 0 is the header (used to label each level). Each node reads as
 * `Field name: value` (e.g. `Origin: Brazil`), with the field name taken from
 * the header of that level's column; a blank value -> "(blank)" and a blank
 * header drops the prefix (just the value). First-seen order is preserved at
 * every level: a JS array keeps push order, and a per-level `Map` (held in a
 * `WeakMap` keyed by node, discarded after the build) is only a dedup index,
 * never the ordered output. Rows sharing the same labelled path MERGE.
 *
 * `detailColumns` (optional) are NOT nesting levels: each row's values for those
 * columns are attached as `Field: value` lines on the leaf node it reaches
 * (`PivotNode.details`). When several rows merge into one leaf, their blocks
 * stack in row order.
 *
 * Resilient like the other mappers: never throws (`row?.[col]` + `cellToString`).
 * Empty `orderedColumns`, or an empty/header-only grid, yields `[]`.
 */
export function rowsToPivotTree(
  rows: Grid,
  orderedColumns: number[],
  detailColumns: number[] = [],
): PivotNode[] {
  if (orderedColumns.length === 0) return [];
  const [header = [], ...dataRows] = rows;
  const headers = header.map(cellToString);
  const labelOf = (col: number) => {
    const name = headers[col]?.trim() ?? "";
    return (value: string) => (name ? `${name}: ${value}` : value);
  };
  const detailLabels = detailColumns.map(labelOf);

  const roots: PivotNode[] = [];
  const rootIndex = new Map<string, PivotNode>();
  // Per-node child dedup index; kept off to the side so PivotNode stays clean,
  // and garbage-collected once the build finishes.
  const childIndex = new WeakMap<PivotNode, Map<string, PivotNode>>();

  for (const row of dataRows) {
    let siblings = roots;
    let index = rootIndex;
    let leaf: PivotNode | null = null;
    for (const col of orderedColumns) {
      const value = cellToString(row?.[col]).trim() || "(blank)";
      // Label each level with its field name, e.g. "Fruit Name: Apple".
      const label = labelOf(col)(value);
      let node = index.get(label);
      if (!node) {
        node = { title: label, children: [] };
        siblings.push(node); // first sight -> preserve insertion order
        index.set(label, node);
        childIndex.set(node, new Map());
      }
      siblings = node.children;
      index = childIndex.get(node)!;
      leaf = node;
    }
    // Hang this row's detail fields off the leaf it reached. Rows that merged
    // into the same leaf append their blocks, in row order.
    if (leaf && detailLabels.length > 0) {
      const lines = detailColumns.map((col, i) =>
        detailLabels[i](cellToString(row?.[col]).trim()),
      );
      (leaf.details ??= []).push(...lines);
    }
  }

  return roots;
}
