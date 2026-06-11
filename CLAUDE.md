@AGENTS.md
# Excel → Word Pivot

Client-side Next.js app. User pastes one or more Excel/Sheets tables; the app restructures each into an Excel-pivot-style **nested hierarchy** ("Rows area"), renders Word-ready HTML, and copies it to the clipboard for pasting into Word.

The problem this solves: wide Excel tables do not fit an 8.5x11 Word page. This app turns one wide table into a narrow, nested outline (group by an ordered list of fields; remaining fields become detail lines under each item) that flows down the page. An optional title is the one Word heading; the nested rows are styled body text.

Multiple pasted tables are managed as cards in a horizontal **tab strip** (one table edited at a time). Heading appearance is a **shared** per-level style applied to every table; each table keeps its own field choices and title. Export is per-table **Copy for Word** plus a combined **Copy all** (clipboard only — no download).

## Commands
- `npm run dev` — dev server on localhost:3000
- `npm run build` — production build, must pass with no type errors
- `npm run lint` — ESLint check

## Stack
- Next.js 16, App Router, TypeScript
- Tailwind CSS v4 (CSS-first, no tailwind.config.js)
- SheetJS for parsing

## Hard Rules
- SheetJS install is the official CDN tarball: `https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`. NEVER install `xlsx` from npm. The npm version is stale (0.18.5).
- Client-side only. No API routes, no backend, no database. Parsing runs in the browser.
- Files at repo root. No `src/` directory. Import alias `@/*` maps to `./*`.
- SheetJS clipboard parse: `getData('text/html')` first, `getData('text/plain')` TSV fallback. Use `sheet_to_json(ws, { header: 1, blankrows: false, defval: "", raw: false })`. Keep `defval: ""` so blank cells stay as empty strings and column positions stay aligned.
- Keep docs in sync with code. When a change alters the pivot behavior, the data model, the pipeline, or a user-facing control, update the affected docs in the SAME change: this `CLAUDE.md` (Architecture / Data Model / Pivot view / Current status), `README.md`, and the relevant `docs/` file (OVERVIEW / ARCHITECTURE / ROADMAP). Stale docs are a bug.

## Architecture
Pipeline (per table): paste → `parseClipboard` (Grid) → `rowsToPivotTree` (`PivotNode[]`) → `renderPivotTree` (HTML fragment) → live preview + per-table/combined "Copy for Word" clipboard write.
- `app/page.tsx` — renders the paste view
- `components/PasteInput.tsx` — parent client component; captures paste (appends a table), holds the `tables[]` array + shared per-level heading style + body font, manages the tab strip, and does combined **Copy all**
- `components/TableCard.tsx` — one table's editor: Section-title input, the pivot **Nest by** ordered picker, **Detail fields** checklist, **Number levels** toggle, per-card View-JSON toggle, per-card **Copy for Word**, and its `RenderedPreview`
- `components/tableModel.ts` — `TableState` type, `tableToHtml(t)` (the per-table map→render pipeline, the single source used by the card preview and combined export), and `pivotDetailColumnsOf(t)`
- `components/RenderedPreview.tsx` — injects the rendered HTML via `dangerouslySetInnerHTML` (input is escaped by the renderer); scoped `<style>` styles the `h2` title and the pivot `[data-level]` rows
- `components/JsonPreview.tsx` — displays the raw parsed Grid as JSON
- `lib/types.ts` — `PivotNode` and raw Grid types
- `lib/parser.ts` — SheetJS clipboard → Grid (`parseClipboard`)
- `lib/mapper.ts` — `cellToString` + `rowsToPivotTree(rows, nestCols, detailCols)` (ordered nested group-by → `PivotNode[]`; detail cols → leaf `details`)
- `lib/renderers.ts` — `renderPivotTree(nodes, title?, numbered?)` (optional title → one `<h2>`; PivotNode tree → `<p data-level="N">` body paragraphs + leaf detail `<p>`; optional `1./a./i.` multilevel markers; depth clamped at 9); escapes user text. Private `escapeHtml` + numbering helpers
- `lib/clipboard.ts` — `HeadingStyle`/`LevelStyle` types; `buildWordHtml` (Word doc wrapper: `<h2>` title → `MsoHeading1` heading, pivot `data-level`→`MsoPiv1..9` **non-heading** styled paragraphs, all looks from `levels`) + `htmlToPlainText`
- `docs/` — OVERVIEW, ARCHITECTURE, ROADMAP

## Data Model
The product is a recursive pivot tree, not a flat table:
- `PivotNode { title, children: PivotNode[], details? }` — `title` reads `Field name: value`; a leaf has `children: []` and optional `details` (flat `Field: value` lines). Shared value-paths merge; merged leaves stack each row's detail block.
- Rendered as nested, indented body paragraphs (`<p data-level="N">`) with details as plain indented `<p>`; the optional title is passed separately and rendered as the single `<h2>` heading.

## Pivot view
The one view (`rowsToPivotTree`) replicates Excel's "Rows area":
- **Nest by** — pick an ordered list of fields (numbered badges + legend + ▲/▼ reorder); rows nest by that order (field 1 = outermost), shared value-paths merge. Each nested row is labelled `Field name: value` from that level's column header.
- **Detail fields** — a checklist of the columns NOT in the nest order; checked ones render as flat `Field: value` body lines under each leaf item (not nesting levels). Merged leaves stack each contributing row's block.
- **Number levels** (`pivotNumbered`, default on) — prefixes multilevel markers `1.`/`a.`/`i.` by depth, restarting per parent. The title and detail lines stay un-numbered.
- **Section title** (optional) — the ONLY Word heading (`<h2>` → `MsoHeading1`, the one outline entry). With a title, the nested data starts at level 2; without one, level 1 and no heading.

## Styling
One **shared per-level** styling panel ("Heading levels") drives the pivot look across all tables — `HeadingStyle = { levels: LevelStyle[] }`. Each level row sets color/font/size/bold; defaults are all the same (distinguished by indent), with a "Reset levels" button. **Level 1** styles the title (`<h2>`/`MsoHeading1`); **Levels 2-9** style the nested rows by depth (`MsoPiv2..9`). The panel shows one row per depth actually in use. A single **Body font** control sits beside it (drives the detail lines + the document body font). All style inputs are form-controlled (hex color, allow-listed font, clamped pt), so no free user text lands in a `style`/`class` attribute.

**Pivot headings:** ONLY the title is a real Word heading. The nested rows are emitted as `MsoPiv1..9` paragraphs that deliberately carry **no** `mso-style-name`/`mso-outline-level` — styled, indented body text, so they don't flood Word's navigation outline.

## Current status
Full pivot pipeline implemented for multiple tables: paste (append) → parse → nest → render → live preview → per-table **Copy for Word** + combined **Copy all** (`text/html` + `text/plain`). Tables are managed in a tab strip (cap 100). Export is clipboard-only (Download-for-Word and the other layouts were removed). `.docx` generation is out of scope.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing the pivot convention or the data model.
- Never: install `xlsx` from npm, add a backend.
