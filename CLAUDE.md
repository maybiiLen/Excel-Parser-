@AGENTS.md
# Excel → Word Sections

Client-side Next.js app. User pastes one or more Excel/Sheets tables; the app restructures each into a section tree, renders Word-ready HTML (native headings + bullet lists), and copies it to the clipboard for pasting into Word.

The problem this solves: wide Excel tables do not fit an 8.5x11 Word page. This app restructures spreadsheet data into readable, narrow sections (grouped, per-item, an A/B/C/D convention, or an Excel-pivot-style nested hierarchy) that fit the page. The grouped and per-item views wrap their output in a numbered, titled section (e.g. `5 Fruit Database` → 5.1, 5.2, …) via `lib/numbering.ts`; the A/B/C/D and pivot views are left un-numbered (the pivot may carry a plain title).

Multiple pasted tables are managed as cards in a horizontal **tab strip** (one table edited at a time). Heading appearance is a **shared** style applied to every table; each table keeps its own layout, column choices, and title. Export is per-table (**Copy/Download for Word**) plus a combined **Copy all / Download all**.

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
- Keep docs in sync with code. When a change adds or alters a view mode, the data model, the pipeline, or a user-facing control, update the affected docs in the SAME change: this `CLAUDE.md` (Architecture / Data Model / View modes / Current status), `README.md`, and the relevant `docs/` file (OVERVIEW / ARCHITECTURE / ROADMAP). Stale docs are a bug.

## Architecture
Pipeline (per table): paste → `parseClipboard` (Grid) → a mapper (chosen by layout) → `Section[]` or `PivotNode[]` → `renderTree`/`renderPivotTree` (HTML fragment) → live preview + per-table/combined "Copy for Word" clipboard write.
- `app/page.tsx` — renders the paste view
- `components/PasteInput.tsx` — parent client component; captures paste (appends a table), holds the `tables[]` array + shared heading style + per-pivot-level styles, manages the tab strip, and does combined Copy/Download all
- `components/TableCard.tsx` — one table's editor: layout + column/title/pivot-order controls, per-card View-JSON toggle, per-card Copy/Download for Word, and its `RenderedPreview`
- `components/tableModel.ts` — `TableState` type, `Layout`, and `tableToHtml(t)` (the per-table map→render pipeline, the single source used by the card preview and combined export); `pickDefault*` / `sectionNumberOf` / `fieldColumnsOf` helpers
- `components/RenderedPreview.tsx` — injects the rendered HTML via `dangerouslySetInnerHTML` (input is escaped by the renderer); scoped `<style>` styles `h2`/`h3` and pivot `[data-level]`
- `components/JsonPreview.tsx` — displays the raw parsed Grid as JSON
- `lib/types.ts` — Section, Subsection, Body, `PivotNode`, and raw Grid types
- `lib/parser.ts` — SheetJS clipboard → Grid (`parseClipboard`)
- `lib/mapper.ts` — Grid → tree; `rowsToTree` (A/B/C/D), `rowsToAttributeSections` (per-item), `rowsToGroupedSections` (group-by), `rowsToPivotTree(rows, nestCols, detailCols)` (ordered nested group-by → `PivotNode[]`, detail cols → leaf `details`)
- `lib/numbering.ts` — `wrapInNumberedSection`: wraps the grouped/per-item output under one user-numbered, titled section (children numbered N.1, N.2, …)
- `lib/renderers.ts` — `renderTree` (Section tree → `<h2>/<h3>` fragment) + `renderBody`; `renderPivotTree(nodes, title?, numbered?)` (optional title → one `<h2>`; PivotNode tree → `<p data-level="N">` body paragraphs + leaf detail `<p>`; optional `1./a./i.` multilevel markers; depth clamped at 9); escapes user text, omits blank numbers
- `lib/clipboard.ts` — `HeadingStyle`/`LevelStyle` types; `buildWordHtml` (Word doc wrapper: `<h2>/<h3>`→`MsoHeading1/2` headings, pivot `data-level`→`MsoPiv1..9` **non-heading** styled paragraphs, all looks from `levels`) + `htmlToPlainText`
- `docs/` — OVERVIEW, ARCHITECTURE, ROADMAP

## Data Model
The product is a tree, not a flat table. The four non-pivot views build a 2-level `Section` tree; the pivot view builds an arbitrary-depth `PivotNode` tree:
- `Section { number, title, children: Subsection[], body? }` — `body?` carries content directly under a section heading (used by the per-item and grouped views)
- `Subsection { number, title, body: Body }`
- `Body = { type: "text", content } | { type: "bullets", items } | { type: "table", rows }`
- `PivotNode { title, children: PivotNode[], details? }` — recursive, no number; a leaf has `children: []` and optional `details` (flat `Field: value` lines). Rendered as nested, indented body paragraphs (`<p data-level="N">`) with details as plain indented `<p>`; the pivot's optional title is passed separately and rendered as the single `<h2>` heading.

## View modes
Each table picks its own layout (per-table; default **Grouped by field**):
- **Grouped by field** (`rowsToGroupedSections`) — row 0 is headers; rows are grouped by a chosen field (each distinct value → a heading), members listed as bullets (a label column + optional checked fields in parens).
- **Fields as bullets** (`rowsToAttributeSections`) — header row + one row per item; each row → a section titled by a chosen column, other selected fields as "Field: value" bullets.
- **Pivot (nested rows)** (`rowsToPivotTree`) — Excel "Rows area": pick an ordered list of **nesting** fields; rows nest by that order (field 1 = outermost), shared value-paths merge. Each nested row is labelled `Field name: value` (e.g. `Origin: Brazil`), from that level's column header. A separate **Detail fields** checklist picks columns shown as flat `Field: value` lines under each leaf item (body text, not nesting levels); when rows merge into one leaf, each row's detail block stacks (`PivotNode.details`). A **Number levels** toggle (`pivotNumbered`, default on) prefixes multilevel markers `1.`/`a.`/`i.` by depth, restarting per parent. The ordered field picker records selection order (numbered badges + legend + ▲/▼ reorder). An optional **Section title** is the ONLY Word heading (`<h2>`); the nested rows + details are styled/indented body paragraphs (not in Word's outline). With a title, data starts at level 2; without one, level 1 and no heading.
- **A/B/C/D sections** (`rowsToTree`, the original position convention) — Column A filled = section title; A blank = subsection of the section above; B = subsection title; C = body content; D = body type flag ("text" | "bullet" | "table").

The grouped/per-item views are header-aware and share a field checklist; they wrap their items under a chosen **Section #** + **Section title** (children numbered N.1, N.2, …). A/B/C/D stays un-numbered.

## Styling
One **shared per-level** styling panel ("Heading levels") drives every heading across all tables — `HeadingStyle = { levels: LevelStyle[] }`. Each level row sets color/font/size/bold; defaults are all the same (distinguished by indent), with a "Reset levels" button. Mapping: **Level 1** styles every top heading (`<h2>`/`MsoHeading1` — grouped/list/sections section headings AND the pivot title); **Level 2** styles subsections (`<h3>`/`MsoHeading2` AND pivot data level 2); **Levels 3-9** style deeper pivot levels. The panel shows one row per level actually in use (non-pivot tables use 2; a pivot uses its depth + 1 for a title). A single **Body font** control sits beside it. All style inputs are form-controlled (hex color, allow-listed font, clamped pt), so no free user text lands in a `style`/`class` attribute.

**Pivot headings:** in the pivot view, ONLY the title is a real Word heading (`<h2>` → `MsoHeading1`, the one outline entry). The nested rows are emitted as `MsoPiv1..9` paragraphs that deliberately carry **no** `mso-style-name`/`mso-outline-level` — they're styled, indented body text, so they don't flood Word's navigation outline.

## Current status
Full pipeline implemented for multiple tables: paste (append) → parse → map (4 layouts) → render → live preview → per-table **Copy/Download for Word** + combined **Copy all / Download all** (`text/html` + `text/plain`; `.doc` via an HTML-doc blob). Tables are managed in a tab strip (cap 100). A wide-table width strategy (transpose/split) is NOT wired in; true `.docx` generation is out of scope. The active views emit headings + bullet lists / nested headings (narrow, page-fitting); only the A/B/C/D view can emit a raw `table` body (via the D flag), which is not yet width-managed.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing a view-mode convention or the data model.
- Never: install `xlsx` from npm, add a backend.