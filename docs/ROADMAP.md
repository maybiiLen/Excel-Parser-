# Roadmap

The original MVP build order, with current status. The core pipeline (paste → map → render → copy for Word) is **done**; the remaining items are deliberately not wired in.

| # | Step | Status |
| - | ---- | ------ |
| 1 | **Data model** — `Section` / `Subsection` / `Body` ([`lib/types.ts`](../lib/types.ts)) | Done (added optional `Section.body`) |
| 2 | **Paste → raw rows** — SheetJS clipboard parse + JSON preview ([`lib/parser.ts`](../lib/parser.ts), [`components/`](../components/)) | Done |
| 3 | **Convention mapper** — rows → section tree, A/B/C/D ([`lib/mapper.ts`](../lib/mapper.ts) `rowsToTree`) | Done |
| 4 | **Body-type detection** — text / bullets / table from column D | Done |
| 5 | **Numbering** — wrap grouped/per-item output in a numbered, titled section ([`lib/numbering.ts`](../lib/numbering.ts) `wrapInNumberedSection`) | Done — children numbered N.1, N.2, …; A/B/C/D stays un-numbered |
| 6 | **Section # + title inputs** — choose the wrapper number (default 1) and heading text | Done |
| 7 | **Body renderers** — body → Word-friendly HTML ([`lib/renderers.ts`](../lib/renderers.ts) `renderTree` / `renderBody`) | Done |
| 8 | **Wide-table strategy** — transpose or split so tables fit a Word page | Not built (largely moot — active views emit narrow bullet lists, not tables) |
| 9 | **Clipboard output** — `text/html` (+ `text/plain`) for Word ([`lib/clipboard.ts`](../lib/clipboard.ts), `PasteInput.copyForWord`) | Done |
| 10 | **Malformed-paste handling** — graceful errors on bad / empty input | Done (friendly empty/parse-error states per view) |
| 11 | **Cleanup & polish** | Ongoing |

## Added beyond the original plan

| Step | What | Status |
| ---- | ---- | ------ |
| **Per-item view** | `rowsToAttributeSections` — one section per row, chosen fields as `Field: value` bullets, with a field checklist | Done |
| **Grouped view** (default) | `rowsToGroupedSections` — group rows by a chosen field; members as bullets (label + optional fields) | Done |
| **Pivot (nested rows) view** | `rowsToPivotTree` → `PivotNode[]` — ordered-field nested group-by (Excel "Rows area"); ordered picker w/ badges + ▲/▼; optional plain title (synthetic root) | Done |
| **View selector + JSON toggle** | Per-table layout select (Grouped / Fields-as-bullets / Pivot / A/B/C/D) + rendered-vs-JSON toggle | Done |
| **Multiple tables + tab strip** | Paste appends a table; cards managed in a horizontal tab strip (one edited at a time, cap 100) | Done |
| **Shared heading styling** | One panel (color, heading/body font, H1/H2 pt, bold) → `HeadingStyle`, applied to every table; headings emitted as Word `MsoHeading`/`MsoPiv` styles | Done |
| **Per-pivot-level styling** | Shared `LevelStyle[]` — color/font/size/bold per nesting depth, defaulting to all the same; "Reset levels" | Done |
| **Copy + Download for Word** | Per-table clipboard write and `.doc` download, plus combined **Copy all / Download all** | Done |

## Possible next steps

- **Pivot aggregation** — counts/sums per group (e.g. `Brazil (3)` or a totals summary).
- **Per-table heading styling** — currently styling is global; per-table would need per-instance class scoping or inline styles.
- **Wide-table width strategy** for the A/B/C/D `table` body, if needed.
- **`.docx` export** (currently out of scope).
