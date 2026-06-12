// Core data model for the Excel -> Word pivot pipeline.
//
// The product is a PIVOT TREE (Excel "Rows area"). Everything downstream of the
// parser (mapper -> renderers -> clipboard output) is built around the PivotNode
// shape below.

/**
 * One field within a node's bucket: its source column plus the label `name`
 * (the trimmed header, `""` when blank) and the cell `value`, kept SEPARATE and
 * raw (unescaped). The renderer escapes each independently and decides whether
 * to show/format the label (`name: `) per the field's `FieldLabel`.
 */
export interface PivotLine {
  col: number;
  name: string;
  value: string;
}

/**
 * A node in the pivot (nested-rows) tree -- the Excel "Rows area" model, where
 * an ordered list of INDENT BUCKETS nests rows into an arbitrary-depth hierarchy
 * and shared value-paths merge. Each bucket is one indent level holding one or
 * more fields; rows whose values match across all of a bucket's fields merge
 * into the same node (a composite group key).
 *
 * A node carries `lines` -- one field (`PivotLine`) per field in its bucket
 * (always >= 1), rendered stacked at the same indent. The FIRST line carries the
 * level's marker (or, when numbered, the Word heading); the rest are plain.
 * Children nest one level deeper.
 */
export interface PivotNode {
  lines: PivotLine[];
  children: PivotNode[];
}

/**
 * Per-field control over the label prefix (`Field name: `): whether it shows at
 * all, and whether it is bold / underlined for emphasis. The value is always
 * shown plain. Stored per grid column in `TableState.fieldLabels`.
 */
export interface FieldLabel {
  show: boolean;
  bold: boolean;
  underline: boolean;
}

/** Default label look: shown, no emphasis. */
export const DEFAULT_FIELD_LABEL: FieldLabel = {
  show: true,
  bold: false,
  underline: false,
};

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
