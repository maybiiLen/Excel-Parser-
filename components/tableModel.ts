// Per-table model + pure derivation, shared by PasteInput (the parent) and
// TableCard. Lives in its own module so both can import the type and helpers
// without a parent <-> child import cycle.

import { rowsToPivotTree } from "@/lib/mapper";
import { renderPivotTree, type MarkerKind } from "@/lib/renderers";
import type { Grid } from "@/lib/types";

/** One pasted table: its grid plus the per-table pivot config the user tweaks. */
export type TableState = {
  /** Stable id from a monotonic counter; used as the React key. */
  id: string;
  grid: Grid;
  /**
   * The pivot structure as ordered INDENT BUCKETS. `pivotLevels[b]` is indent
   * level b+1 and holds one or more grid column indices, in display (stacking)
   * order. Fields stacked in one bucket render at the same indent and merge as a
   * composite group. Columns in no bucket are unused (hidden).
   */
  pivotLevels: number[][];
  /** Marker style per indent level (index = level − 1; sparse → default cycle). */
  markers: MarkerKind[];
  /** Optional title; the one Word heading, above the nested rows. */
  sectionTitle: string;
};

// ---------------------------------------------------------------------------
// Pure helpers on the bucket structure (`number[][]`). All return a fresh array
// and never mutate. A "flattened index" `fi` is a field's position in
// `levels.flat()`. Bucket invariant: every bucket is non-empty.
// ---------------------------------------------------------------------------

/** All placed grid columns, in flat (top-to-bottom) order. */
export function placedColumns(levels: number[][]): number[] {
  return levels.flat();
}

/** Grid columns NOT placed in any bucket (the "unused" pool), in column order. */
export function unusedColumns(width: number, levels: number[][]): number[] {
  const placed = new Set(placedColumns(levels));
  const cols: number[] = [];
  for (let i = 0; i < width; i++) if (!placed.has(i)) cols.push(i);
  return cols;
}

/** Locate a field by flattened index: its bucket `b` and within-bucket index `k`. */
function locate(levels: number[][], fi: number): { b: number; k: number } | null {
  let n = 0;
  for (let b = 0; b < levels.length; b++) {
    if (fi < n + levels[b].length) return { b, k: fi - n };
    n += levels[b].length;
  }
  return null;
}

/** Append a column as a new deepest single-field bucket. */
export function addField(levels: number[][], col: number): number[][] {
  return [...levels, [col]];
}

/** Remove the field at flat index `fi`; drop its bucket if it becomes empty. */
export function removeField(levels: number[][], fi: number): number[][] {
  const at = locate(levels, fi);
  if (!at) return levels;
  const { b, k } = at;
  const bucket = levels[b].filter((_, i) => i !== k);
  if (bucket.length === 0) return levels.filter((_, i) => i !== b);
  return levels.map((bk, i) => (i === b ? bucket : bk));
}

/** Whether ► (indent) is allowed for the field at `fi`: not first in its bucket. */
export function canIndent(levels: number[][], fi: number): boolean {
  const at = locate(levels, fi);
  return at != null && at.k > 0;
}

/**
 * Indent the field at `fi` one level deeper: split its bucket at `k` so the field
 * and everything after it in that bucket move into a NEW bucket one level deeper.
 * No-op unless `canIndent`.
 */
export function indentField(levels: number[][], fi: number): number[][] {
  const at = locate(levels, fi);
  if (!at || at.k === 0) return levels;
  const { b, k } = at;
  const head = levels[b].slice(0, k);
  const tail = levels[b].slice(k);
  return [...levels.slice(0, b), head, tail, ...levels.slice(b + 1)];
}

/** Whether ◄ (outdent) is allowed for `fi`: first of its bucket and bucket > 0. */
export function canOutdent(levels: number[][], fi: number): boolean {
  const at = locate(levels, fi);
  return at != null && at.k === 0 && at.b > 0;
}

/**
 * Outdent the field at `fi` one level shallower by merging its bucket into the
 * previous one (the field and the rest of its bucket join the previous level).
 * No-op unless `canOutdent`.
 */
export function outdentField(levels: number[][], fi: number): number[][] {
  const at = locate(levels, fi);
  if (!at || at.k !== 0 || at.b === 0) return levels;
  const { b } = at;
  const merged = [...levels[b - 1], ...levels[b]];
  return [...levels.slice(0, b - 1), merged, ...levels.slice(b + 1)];
}

/**
 * Move the field at `fi` up (`dir = -1`) or down (`dir = +1`) by swapping it with
 * its flat-order neighbor while keeping the bucket SHAPE fixed: the two fields
 * trade slots, so a field moved up adopts the upper slot's indent. Always valid
 * (a swap can't empty a bucket). No-op when the neighbor doesn't exist.
 */
export function moveField(levels: number[][], fi: number, dir: -1 | 1): number[][] {
  const flat = placedColumns(levels);
  const fj = fi + dir;
  if (fj < 0 || fj >= flat.length) return levels;
  const a = locate(levels, fi);
  const b = locate(levels, fj);
  if (!a || !b) return levels;
  return levels.map((bucket, bi) =>
    bucket.map((col, ki) => {
      if (bi === a.b && ki === a.k) return flat[fj];
      if (bi === b.b && ki === b.k) return flat[fi];
      return col;
    }),
  );
}

/**
 * parse -> nest -> render for one table. An optional Section title is the ONLY
 * heading (`<p class="ws-title">`); the nested rows are styled body paragraphs
 * beneath it. Empty-guard first so a title never renders over nothing. Pure: the
 * single source used by both the card preview and the combined export.
 */
export function tableToHtml(t: TableState): string {
  const tree = rowsToPivotTree(t.grid, t.pivotLevels);
  if (tree.length === 0) return "";
  return renderPivotTree(tree, t.sectionTitle.trim() || undefined, t.markers);
}
