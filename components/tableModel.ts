// Per-table model + pure derivation, shared by PasteInput (the parent) and
// TableCard. Lives in its own module so both can import the type and helpers
// without a parent <-> child import cycle.

import { rowsToPivotTree } from "@/lib/mapper";
import { renderPivotTree } from "@/lib/renderers";
import type { Grid } from "@/lib/types";

/** One pasted table: its grid plus the per-table pivot config the user tweaks. */
export type TableState = {
  /** Stable id from a monotonic counter; used as the React key. */
  id: string;
  grid: Grid;
  /** Columns checked to show as detail lines under each leaf item. */
  selectedCols: Set<number>;
  /** Ordered columns the pivot nests rows by (selection order). */
  pivotOrder: number[];
  /** Number the nested levels (1./a./i. …). */
  pivotNumbered: boolean;
  /** Optional title; the one Word heading, above the nested rows. */
  sectionTitle: string;
};

/**
 * Pivot detail columns, in column order: checked (`selectedCols`) but NOT one of
 * the nesting levels (`pivotOrder`). These render as "Field: value" lines under
 * each leaf item rather than as deeper nesting levels.
 */
export function pivotDetailColumnsOf(t: TableState): number[] {
  const width = t.grid[0]?.length ?? 0;
  const cols: number[] = [];
  for (let i = 0; i < width; i++) {
    if (t.selectedCols.has(i) && !t.pivotOrder.includes(i)) cols.push(i);
  }
  return cols;
}

/**
 * parse -> nest -> render for one table. An optional Section title is the ONLY
 * heading (rendered as `<h2>`); the nested rows + details are styled body
 * paragraphs beneath it. Empty-guard first so a title never renders over nothing.
 * Pure: the single source used by both the card preview and the combined export.
 */
export function tableToHtml(t: TableState): string {
  const tree = rowsToPivotTree(t.grid, t.pivotOrder, pivotDetailColumnsOf(t));
  if (tree.length === 0) return "";
  return renderPivotTree(
    tree,
    t.sectionTitle.trim() || undefined,
    t.pivotNumbered,
  );
}
