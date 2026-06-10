# Roadmap

The original MVP build order, with current status. The core pipeline (paste → map → render → copy for Word) is **done**; the remaining items are deliberately not wired in.

| # | Step | Status |
| - | ---- | ------ |
| 1 | **Data model** — `Section` / `Subsection` / `Body` ([`lib/types.ts`](../lib/types.ts)) | Done (added optional `Section.body`) |
| 2 | **Paste → raw rows** — SheetJS clipboard parse + JSON preview ([`lib/parser.ts`](../lib/parser.ts), [`components/`](../components/)) | Done |
| 3 | **Convention mapper** — rows → section tree, A/B/C/D ([`lib/mapper.ts`](../lib/mapper.ts) `rowsToTree`) | Done |
| 4 | **Body-type detection** — text / bullets / table from column D | Done |
| 5 | **Numbering** — tree walk assigns dotted numbers ([`lib/numbering.ts`](../lib/numbering.ts) `numberTree`) | Built, then **un-wired** — output is un-numbered by product decision; `numberTree` remains available |
| 6 | **Root-number input** — configurable starting number | Built, then removed with numbering |
| 7 | **Body renderers** — body → Word-friendly HTML ([`lib/renderers.ts`](../lib/renderers.ts) `renderTree` / `renderBody`) | Done |
| 8 | **Wide-table strategy** — transpose or split so tables fit a Word page | Not built (largely moot — active views emit narrow bullet lists, not tables) |
| 9 | **Clipboard output** — `text/html` (+ `text/plain`) for Word ([`lib/clipboard.ts`](../lib/clipboard.ts), `PasteInput.copyForWord`) | Done |
| 10 | **Malformed-paste handling** — graceful errors on bad / empty input | Done (friendly empty/parse-error states per view) |
| 11 | **Cleanup & polish** | Ongoing |

## Added beyond the original plan

The product grew two header-aware view modes (default is the grouped one):

| Step | What | Status |
| ---- | ---- | ------ |
| **Per-item view** | `rowsToAttributeSections` — one section per row, chosen fields as `Field: value` bullets, with a field checklist | Done |
| **Grouped (pivot-style) view** | `rowsToGroupedSections` — group rows by a chosen field; members as bullets (label + optional fields) | Done |
| **View selector** | Switch between Grouped / Fields-as-bullets / A/B/C/D, plus a rendered-vs-JSON toggle | Done |

## Possible next steps

- **Pivot aggregation** — counts/sums per group (e.g. `Brazil (3)` or a totals summary), building on the grouped view.
- **Re-enable numbering** as an optional toggle (the `numberTree` utility and `renderTree`'s number support already exist).
- **Wide-table width strategy** for the A/B/C/D `table` body, if needed.
- **`.docx` export** (currently out of scope).
