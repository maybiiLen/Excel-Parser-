# Roadmap

The app is **pivot-only**: paste → nest → render → Copy for Word. It began with four view modes and a Download option; those were removed to focus on the pivot. Status of the current feature set below.

| Step | What | Status |
| ---- | ---- | ------ |
| **Data model** | `PivotNode { title, children, details? }` ([`lib/types.ts`](../lib/types.ts)) | Done |
| **Paste → raw rows** | SheetJS clipboard parse + JSON preview ([`lib/parser.ts`](../lib/parser.ts)) | Done |
| **Pivot mapper** | `rowsToPivotTree(rows, nestCols, detailCols)` — ordered nested group-by; shared paths merge; leaf `details` ([`lib/mapper.ts`](../lib/mapper.ts)) | Done |
| **Nest-by picker** | Ordered field picker w/ numbered badges + legend + ▲/▼ reorder | Done |
| **Detail fields** | Checklist of non-nested columns → `Field: value` lines under each leaf | Done |
| **Number levels** | Toggle: multilevel `1./a./i.` markers by depth, restarting per parent | Done |
| **Title heading** | Optional Section title → the one Word heading (`<h2>` → Heading 1); nested rows are non-heading body paragraphs | Done |
| **Renderer** | `renderPivotTree(nodes, title?, numbered?)` → HTML fragment ([`lib/renderers.ts`](../lib/renderers.ts)) | Done |
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
