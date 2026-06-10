// Per-table model + pure derivation, shared by PasteInput (the parent) and
// TableCard. Lives in its own module so both can import the type and helpers
// without a parent <-> child import cycle.

import {
  rowsToTree,
  rowsToAttributeSections,
  rowsToGroupedSections,
  rowsToPivotTree,
} from "@/lib/mapper";
import { wrapInNumberedSection } from "@/lib/numbering";
import { renderTree, renderPivotTree } from "@/lib/renderers";
import type { Grid } from "@/lib/types";

export type Layout = "list" | "grouped" | "sections" | "pivot";

/** One pasted table: its grid plus the per-table config the user can tweak. */
export type TableState = {
  /** Stable id from a monotonic counter; used as the React key. */
  id: string;
  grid: Grid;
  layout: Layout;
  titleCol: number;
  groupCol: number;
  selectedCols: Set<number>;
  /** Ordered columns the pivot view nests rows by (selection order). */
  pivotOrder: number[];
  /** Wrapping section number, held as a raw string so it can be backspaced. */
  sectionNumberInput: string;
  sectionTitle: string;
};

/** First header containing "name" (case-insensitive), else the first column. */
export function pickDefaultTitleCol(rows: Grid): number {
  const header = rows[0] ?? [];
  const i = header.findIndex(
    (c) => typeof c === "string" && c.toLowerCase().includes("name"),
  );
  return i >= 0 ? i : 0;
}

const GROUP_KEYWORDS = [
  "origin", "season", "type", "category", "status", "risk", "level",
  "color", "region", "class", "group", "storage", "method",
];

/** First header matching a category-ish keyword, else first column != titleCol. */
export function pickDefaultGroupCol(rows: Grid, titleCol: number): number {
  const header = (rows[0] ?? []).map((c) =>
    typeof c === "string" ? c.toLowerCase() : "",
  );
  for (let i = 0; i < header.length; i++) {
    if (i !== titleCol && GROUP_KEYWORDS.some((k) => header[i].includes(k))) {
      return i;
    }
  }
  for (let i = 0; i < header.length; i++) {
    if (i !== titleCol) return i;
  }
  return titleCol;
}

/** Section number used for output (empty/invalid -> 1), derived from the input. */
export function sectionNumberOf(t: TableState): number {
  const n = parseInt(t.sectionNumberInput, 10);
  return Number.isFinite(n) && n > 0 ? n : 1;
}

/**
 * Columns shown as bullets / parenthetical extras, in column order: checked,
 * minus the title (label) column, and -- in grouped mode -- the group column.
 */
export function fieldColumnsOf(t: TableState): number[] {
  const width = t.grid[0]?.length ?? 0;
  const cols: number[] = [];
  for (let i = 0; i < width; i++) {
    if (
      i !== t.titleCol &&
      (t.layout !== "grouped" || i !== t.groupCol) &&
      t.selectedCols.has(i)
    ) {
      cols.push(i);
    }
  }
  return cols;
}

/**
 * parse -> map -> (wrap in a numbered section) -> render for one table. The
 * grouped and per-item views are wrapped under one numbered, titled heading
 * (5 Fruit Database -> 5.1, 5.2, ...); A/B/C/D is left as-is. Pure: the single
 * source of truth used by both the card preview and the combined export.
 */
export function tableToHtml(t: TableState): string {
  // Pivot: plain, un-numbered nested rows (not wrapped in a numbered section like
  // the grouped/per-item views). An optional Section title is the ONLY heading
  // (rendered as <h2>); the nested rows are styled body paragraphs beneath it.
  // Guard the empty tree first so a title never renders over nothing.
  if (t.layout === "pivot") {
    const tree = rowsToPivotTree(t.grid, t.pivotOrder);
    if (tree.length === 0) return "";
    return renderPivotTree(tree, t.sectionTitle.trim() || undefined);
  }
  const fieldColumns = fieldColumnsOf(t);
  const num = sectionNumberOf(t);
  const title = t.sectionTitle.trim();
  let tree;
  if (t.layout === "list") {
    tree = wrapInNumberedSection(
      rowsToAttributeSections(t.grid, t.titleCol, fieldColumns),
      num,
      title,
    );
  } else if (t.layout === "grouped") {
    tree = wrapInNumberedSection(
      rowsToGroupedSections(t.grid, t.groupCol, t.titleCol, fieldColumns),
      num,
      title,
    );
  } else {
    tree = rowsToTree(t.grid);
  }
  return renderTree(tree);
}
