// Core data model for the Excel -> Word pivot pipeline.
//
// The product is a PIVOT TREE (Excel "Rows area"). Everything downstream of the
// parser (mapper -> renderers -> clipboard output) is built around the PivotNode
// shape below.

/**
 * A node in the pivot (nested-rows) tree -- the Excel "Rows area" model, where
 * an ordered list of fields nests rows into an arbitrary-depth hierarchy and
 * shared value-paths merge. A leaf is a node with no children; it may carry
 * `details` (flat "Field: value" lines rendered as body text under the item).
 */
export interface PivotNode {
  title: string;
  children: PivotNode[];
  /**
   * Flat "Field: value" detail lines shown under a leaf item (body text, not a
   * nesting level). Set only on leaves; when rows merge into one leaf, each
   * contributing row appends its block, in row order.
   */
  details?: string[];
}

// ---------------------------------------------------------------------------
// Raw clipboard grid -- the output of today's parser, before any mapping.
//
// A Grid is an array of rows; each row is an array of cell values keyed by
// column position. Blank cells are preserved as "" so column positions stay
// aligned for the pivot's field selection.
// ---------------------------------------------------------------------------

export type Cell = string | number | boolean | null;
export type Row = Cell[];
export type Grid = Row[];
