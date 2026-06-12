# Roadmap

The app is **pivot-only**: paste → nest → render → Copy for Word. It began with four view modes and a Download option; those were removed to focus on the pivot. Status of the current feature set below.

| Step | What | Status |
| ---- | ---- | ------ |
| **Data model** | `PivotNode { lines: PivotLine[], children }`, `PivotLine { col, name, value }`, `FieldLabel` ([`lib/types.ts`](../lib/types.ts)) | Done |
| **Paste → raw rows** | SheetJS clipboard parse + JSON preview ([`lib/parser.ts`](../lib/parser.ts)) | Done |
| **Pivot mapper** | `rowsToPivotTree(rows, levels)` — ordered indent buckets; each bucket is one level of ≥1 field merged by a composite key ([`lib/mapper.ts`](../lib/mapper.ts)) | Done |
| **Structure picker** | Add-fields pool + placed list with ◄ outdent / ► indent (stack = same indent) / ▲ ▼ reorder / ✕ remove (`pivotLevels: number[][]`) | Done |
| **Per-field label format** | Per-field toggles (`fieldLabels`): show/hide the `Field name:` label, bold it, underline it (label-only emphasis) | Done |
| **Markers** | Per-level `<select>` (`1./1)/A./a./I./i./•/–/None`); first field of a multi-field level marked, restarting per parent | Done |
| **Title heading** | Optional Section title → Word `Heading 1` (`mso-style-name`, "5.0") on a Use-Destination-Styles paste | Done |
| **Word heading numbering** | `numberDepth` maps the top N body levels → `Heading 2/3/…` (`MsoHeadingK`) so Word numbers them live (`5.1`, `5.1.1`). Number is template-driven; numbered items enter the Nav pane/TOC | Done |
| **Renderer** | `renderPivotTree(nodes, title?, markers?, fieldLabels?, numberDepth?)` → HTML fragment ([`lib/renderers.ts`](../lib/renderers.ts)) | Done |
| **Clipboard output** | `text/html` (+ `text/plain`) for Word ([`lib/clipboard.ts`](../lib/clipboard.ts) `buildWordHtml`) | Done |
| **Multiple tables + tab strip** | Paste appends a table; cards in a horizontal tab strip (cap 100) | Done |
| **Shared per-level styling** | One "Heading levels" panel → `HeadingStyle = { levels: LevelStyle[] }`: Level 1 = title, Levels 2-9 = nested rows (color/font/size/bold each, default all the same; "Reset levels") + one Body font | Done |
| **Combined export** | **Copy all** — every table as one Word doc | Done |
| **Malformed-paste handling** | Graceful empty/parse-error states | Done |

## Removed (was built, then cut to focus on pivot)

- The other three layouts: **Grouped by field**, **Fields as bullets**, **A/B/C/D sections** (with `lib/numbering.ts`, the `Section`/`Subsection`/`Body` model, `rowsToTree`/`rowsToAttributeSections`/`rowsToGroupedSections`, `renderTree`/`renderBody`).
- **Download for Word** / **Download all** (`.doc` export) — clipboard-only now.

## Possible next steps

- **Pivot aggregation** — counts/sums per group (e.g. `Brazil (3)` or a totals summary).
- **Per-table heading styling** — currently styling is global; per-table would need per-instance class scoping or inline styles.
- **`.docx` export** (currently out of scope).
