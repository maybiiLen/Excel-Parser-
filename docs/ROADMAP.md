# Roadmap

The app is **pivot-only**: paste → nest → render → Copy for Word. It began with four view modes and a Download option; those were removed to focus on the pivot. Status of the current feature set below.

| Step | What | Status |
| ---- | ---- | ------ |
| **Data model** | `PivotNode { lines: PivotLine[], children }`, `PivotLine { col, name, value }`, `FieldLabel` ([`lib/types.ts`](../lib/types.ts)) | Done |
| **Paste → raw rows** | SheetJS clipboard parse + JSON preview ([`lib/parser.ts`](../lib/parser.ts)) | Done |
| **Pivot mapper** | `rowsToPivotTree(rows, levels, sortDirs?)` — ordered indent buckets; each bucket is one level of ≥1 field merged by a composite key; a post-pass reorders sibling groups per `sortDirs` ([`lib/mapper.ts`](../lib/mapper.ts)) | Done |
| **Structure picker** | Add-fields pool + placed list with ◄ outdent / ► indent (stack = same indent) / ▲ ▼ reorder / ✕ remove (`pivotLevels: number[][]`) | Done |
| **Per-field label format** | Per-field toggles (`fieldLabels`): show/hide the `Field name:` label, bold it, underline it (label-only emphasis) | Done |
| **Per-column sorting** | Per-field ↕/↑/↓ cycle (`sortDirs: Record<col, "asc"|"desc">`) orders the sibling groups at that field's level; numeric + text aware, case-insensitive (`localeCompare`, `numeric:true`); off keeps first-seen order | Done |
| **Markers** | Per-level `<select>` (`1./1)/A./a./I./i./•/–/None`); first field of a multi-field level marked, restarting per parent | Done |
| **Blank line after group** | Per-level checkboxes (`breakAfter: boolean[]`); when on, a nbsp spacer paragraph follows that level's whole subtree (survives a Word paste; no marker/number/label) | Done |
| **Title heading** | Optional Section title → Word `Heading 1` (`mso-style-name`, "5.0") on a Use-Destination-Styles paste | Done |
| **Multilevel numbering (app-drawn static numbers)** | Per-table `numbering: { mode, start, levels }` prefixes each node's first line with a static number (top level `5.0`, then `5.1`, `5.1.1`) as plain body text and suppresses that node's marker; a per-level checkbox (`numbering.levels`) hides a level and makes it transparent, so the numbers follow the gap (`1.0` → `1.1`) with no gaps or collisions. NOT Word auto-numbering — nothing becomes a Word heading, so the Nav pane / TOC stay clean and the preview shows the real numbers | Done |
| **Renderer** | `renderPivotTree(nodes, title?, markers?, fieldLabels?, breakAfter?, numbering?)` → HTML fragment ([`lib/renderers.ts`](../lib/renderers.ts)) | Done |
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
