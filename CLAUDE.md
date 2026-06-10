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
- `lib/mapper.ts` — Grid → tree; `rowsToTree` (A/B/C/D), `rowsToAttributeSections` (per-item), `rowsToGroupedSections` (group-by), `rowsToPivotTree` (ordered nested group-by → `PivotNode[]`)
- `lib/numbering.ts` — `wrapInNumberedSection`: wraps the grouped/per-item output under one user-numbered, titled section (children numbered N.1, N.2, …)
- `lib/renderers.ts` — `renderTree` (Section tree → `<h2>/<h3>` fragment) + `renderBody`; `renderPivotTree` (PivotNode tree → `<p data-level="N">` fragment, depth clamped at 9); escapes user text, omits blank numbers
- `lib/clipboard.ts` — `HeadingStyle`/`LevelStyle` types; `buildWordHtml` (Word doc wrapper: `<h2>/<h3>`→`MsoHeading1/2`, pivot `data-level`→`MsoPiv1..9`, per-level look + indent) + `htmlToPlainText`
- `docs/` — OVERVIEW, ARCHITECTURE, ROADMAP

## Data Model
The product is a tree, not a flat table. The four non-pivot views build a 2-level `Section` tree; the pivot view builds an arbitrary-depth `PivotNode` tree:
- `Section { number, title, children: Subsection[], body? }` — `body?` carries content directly under a section heading (used by the per-item and grouped views)
- `Subsection { number, title, body: Body }`
- `Body = { type: "text", content } | { type: "bullets", items } | { type: "table", rows }`
- `PivotNode { title, children: PivotNode[] }` — recursive, no body/number; a leaf has `children: []`. Rendered as nested Word headings (depth → Heading 1..9). The pivot's optional title is a synthetic root `PivotNode`.

## View modes
Each table picks its own layout (per-table; default **Grouped by field**):
- **Grouped by field** (`rowsToGroupedSections`) — row 0 is headers; rows are grouped by a chosen field (each distinct value → a heading), members listed as bullets (a label column + optional checked fields in parens).
- **Fields as bullets** (`rowsToAttributeSections`) — header row + one row per item; each row → a section titled by a chosen column, other selected fields as "Field: value" bullets.
- **Pivot (nested rows)** (`rowsToPivotTree`) — Excel "Rows area": pick an ordered list of fields; rows nest by that order (field 1 = outermost), shared value-paths merge. An ordered field picker records selection order (numbered badges + legend + ▲/▼ reorder). Plain (un-numbered); an optional **Section title** becomes the top heading (synthetic root; groups shift one level deeper).
- **A/B/C/D sections** (`rowsToTree`, the original position convention) — Column A filled = section title; A blank = subsection of the section above; B = subsection title; C = body content; D = body type flag ("text" | "bullet" | "table").

The grouped/per-item views are header-aware and share a field checklist; they wrap their items under a chosen **Section #** + **Section title** (children numbered N.1, N.2, …). A/B/C/D stays un-numbered.

## Styling
Heading appearance is **shared** across all tables (one panel: heading color, heading font, body font, H1 pt, H2 pt, bold) → `HeadingStyle`. It drives `<h2>/<h3>` (Word Heading 1/2) for the non-pivot views. The **pivot** view additionally has shared **per-nesting-level styles** (`LevelStyle[]`, one row per level in use): color/font/size/bold per depth, defaulting to all the same (distinguished by indent), with a "Reset levels" button. Per-level styles feed the `MsoPiv1..9` Word classes and the preview `[data-level]` CSS. All style inputs are form-controlled (hex color, allow-listed font, clamped pt), so no free user text lands in a `style`/`class` attribute.

## Current status
Full pipeline implemented for multiple tables: paste (append) → parse → map (4 layouts) → render → live preview → per-table **Copy/Download for Word** + combined **Copy all / Download all** (`text/html` + `text/plain`; `.doc` via an HTML-doc blob). Tables are managed in a tab strip (cap 100). A wide-table width strategy (transpose/split) is NOT wired in; true `.docx` generation is out of scope. The active views emit headings + bullet lists / nested headings (narrow, page-fitting); only the A/B/C/D view can emit a raw `table` body (via the D flag), which is not yet width-managed.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing a view-mode convention or the data model.
- Never: install `xlsx` from npm, add a backend.