@AGENTS.md
# Excel → Word Pivot

Client-side Next.js app. User pastes one or more Excel/Sheets tables; the app restructures each into an Excel-pivot-style **nested hierarchy** ("Rows area"), renders Word-ready HTML, and copies it to the clipboard for pasting into Word.

The problem this solves: wide Excel tables do not fit an 8.5x11 Word page. This app turns one wide table into a narrow, nested outline (group by an ordered list of fields; remaining fields become detail lines under each item) that flows down the page. Output can **map to the destination document's Word styles** (e.g. the title → `Heading 1`, the body → `paratext`) via `mso-style-name`, so a "Use Destination Styles" paste adopts the template's look; or, with the style names left blank, it uses the app's own per-level direct styling (Arial 11 default).

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
- `lib/renderers.ts` — `renderPivotTree(nodes, title?, markers?)` (optional title → a plain level-1 row; nested rows → `<p data-level="N">`; leaf detail → `<p class="ws-detail" data-detail="N">`; per-level markers; depth clamped at 9); `MarkerKind`/`markerText`/`defaultMarker`; escapes user text. Private `escapeHtml` + `toAlpha`/`toRoman`
- `lib/clipboard.ts` — `HeadingStyle` (`{ levels, indentStep, headingStyleName, bodyStyleName }`)/`LevelStyle` types; `buildWordHtml` (Word doc wrapper: `ws-title`→`MsoTitle`, `data-level`→`MsoPiv1..9`, `data-detail`→`MsoDet1..9`; when a style name is set the class carries `mso-style-name:"<name>"` (title also `mso-outline-level:1`) + indent + single spacing and NO direct font, so the destination style wins — else the direct per-level look; `sanitizeStyleName`) + `htmlToPlainText`
- `docs/` — OVERVIEW, ARCHITECTURE, ROADMAP

## Data Model
The product is a recursive pivot tree, not a flat table:
- `PivotNode { title, children: PivotNode[], details? }` — `title` reads `Field name: value`; a leaf has `children: []` and optional `details` (flat `Field: value` lines). Shared value-paths merge; merged leaves stack each row's detail block.
- Rendered as a title row (`<p class="ws-title">`) + nested rows (`<p data-level="N">`) + leaf details (`<p class="ws-detail" data-detail="N">`); `buildWordHtml` maps these to Word styles (title → heading style, body → body style) when style names are set.

## Pivot view
The one view (`rowsToPivotTree`) replicates Excel's "Rows area":
- **Nest by** — pick an ordered list of fields (numbered badges + legend + ▲/▼ reorder); rows nest by that order (field 1 = outermost), shared value-paths merge. Each nested row is labelled `Field name: value` from that level's column header.
- **Detail fields** — a checklist of the columns NOT in the nest order; checked ones render as flat `Field: value` body lines under each leaf item (not nesting levels). Merged leaves stack each contributing row's block.
- **Markers** (per-table, `markers: MarkerKind[]` indexed by nesting depth) — a `<select>` per nesting level picks its marker style: `1.` / `1)` / `A.` / `a.` / `I.` / `i.` / `•` / `–` / None. Sparse → `defaultMarker(depth)` (the legacy `1./a./i.` cycle), so a fresh pivot looks numbered until you change it. Title + detail lines are never marked.
- **Section title** (optional) — the title row (`MsoTitle`). Maps to the **Heading style** name when set (a real Word heading, `mso-outline-level:1`); else a plain level-1 row. With a title the nested data starts at level 2; without one, level 1.

## Styling
**Word-style mapping (primary path):** two shared inputs — **Heading style** (default `Heading 1`) and **Body style** (default `paratext`) — are Word style names. When set, the title carries `mso-style-name:"<heading>"`+`mso-outline-level:1` and every body paragraph `mso-style-name:"<body>"`, so a **Use Destination Styles** paste adopts the destination document's styles (the template controls font/size/color). The app then only adds the left-indent + the `1./a./i.` markers + **single line spacing**. Leave a name blank to fall back to the app's direct per-level look for that part. Names are sanitized (`sanitizeStyleName`) before landing in the `mso-style-name` attribute.

One **shared per-level** styling panel ("Level styles") drives the look across all tables — `HeadingStyle = { levels: LevelStyle[]; indentStep: number }`. Each level row sets color/font/size/bold; **defaults are plain Arial 11 black** (to match a document body), with a "Reset levels" button. **Level 1** styles the title row; **Levels 2-9** style the nested rows by depth. The panel shows one row per depth in use. Beside it: a **Body font** control (detail lines + document body font) and an **Indent/level (in)** control (`indentStep`, clamped 0–2; nested rows `(n-1)×step`, detail lines `n×step`). Every paragraph also carries `margin-bottom:6pt;line-height:1.15` for comfortable spacing. All style inputs are form-controlled (hex color, allow-listed font, clamped numbers), so no free user text lands in a `style`/`class` attribute.

## Current status
Full pivot pipeline implemented for multiple tables: paste (append) → parse → nest → render → live preview → per-table **Copy for Word** + combined **Copy all** (`text/html` + `text/plain`). Tables are managed in a tab strip (cap 100). Export is clipboard-only (Download-for-Word and the other layouts were removed). `.docx` generation is out of scope.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing the pivot convention or the data model.
- Never: install `xlsx` from npm, add a backend.
