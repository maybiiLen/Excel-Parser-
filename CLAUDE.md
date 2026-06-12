@AGENTS.md
# Excel → Word Pivot

Client-side Next.js app. User pastes one or more Excel/Sheets tables; the app restructures each into an Excel-pivot-style **nested hierarchy** ("Rows area"), renders Word-ready HTML, and copies it to the clipboard for pasting into Word.

The problem this solves: wide Excel tables do not fit an 8.5x11 Word page. This app turns one wide table into a narrow, nested outline. You arrange fields into ordered **indent levels** (each level can hold one or more fields stacked at the same indentation); rows nest and merge by those levels. The **title** can **map to the destination document's Heading style** (e.g. `Heading 1`) via `mso-style-name`, so a "Use Destination Styles" paste adopts the template's heading; the **body** always uses the app's own per-level direct styling (Arial 11 default, with per-level bold).

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
- `components/TableCard.tsx` — one table's editor: Section-title input, the **Add fields** pool (unused columns), the **Structure** list (placed fields with ◄ outdent / ► indent / ▲ ▼ reorder / ✕ remove, plus per-field label toggles **Aa** show/hide / **B** bold / **U** underline), the per-level **Markers** pickers, per-card View-JSON toggle, per-card **Copy for Word**, and its `RenderedPreview`
- `components/tableModel.ts` — `TableState` type, `tableToHtml(t, numberDepth)` (the per-table map→render pipeline, the single source used by the card preview and combined export — both pass the shared `numberDepth`), and the pure bucket helpers (`addField`/`removeField`/`indentField`/`outdentField`/`moveField`/`canIndent`/`canOutdent`/`placedColumns`/`unusedColumns`)
- `components/RenderedPreview.tsx` — injects the rendered HTML via `dangerouslySetInnerHTML` (input is escaped by the renderer); scoped `<style>` styles the `ws-title` title and the pivot `[data-level]` rows
- `components/JsonPreview.tsx` — displays the raw parsed Grid as JSON
- `lib/types.ts` — `PivotNode`/`PivotLine`, `FieldLabel` + `DEFAULT_FIELD_LABEL`, and raw Grid types
- `lib/parser.ts` — SheetJS clipboard → Grid (`parseClipboard`)
- `lib/mapper.ts` — `cellToString` + `rowsToPivotTree(rows, levels)` (ordered indent buckets → `PivotNode[]`; each bucket is one level holding ≥1 field as `PivotLine{col,name,value}`, merged by a composite key)
- `lib/renderers.ts` — `renderPivotTree(nodes, title?, markers?, fieldLabels?, numberDepth?)` (optional title row; each node's `lines` → `<p class="ws-lvl" data-level="N">` with the label `name:` shown/bolded/underlined per `fieldLabels[col]` then the value; marker on the first line only; the top `numberDepth` levels' first line gets `data-heading="K"`, marker suppressed; depth clamped at 9); `MarkerKind`/`markerText`/`defaultMarker`; escapes user text. Private `escapeHtml`/`wrapLabel` + `toAlpha`/`toRoman`
- `lib/clipboard.ts` — `HeadingStyle` (`{ levels, indentStep, headingStyleName }`)/`LevelStyle` types; `buildWordHtml` (Word doc wrapper: the title → `<p class="MsoTitle">` + a `mso-style-name:"<heading>"`+`mso-outline-level:1` rule when `headingStyleName` is set (else inline level-1 look); each `data-heading="K"` line → `<p class="MsoHeadingK">` + a `mso-style-name:"Heading K"`+`mso-outline-level:K` rule (Word numbers it); every other body row → a `<p>` with **inline** direct formatting (per-level color/font/size + indent + compact spacing) and `<b>`/`<u>` label runs — inline so it survives a "Use Destination Styles" paste; `sanitizeStyleName`) + `htmlToPlainText`
- `docs/` — OVERVIEW, ARCHITECTURE, ROADMAP

## Data Model
The product is a recursive pivot tree, not a flat table:
- `PivotNode { lines: PivotLine[], children: PivotNode[] }`; `PivotLine { col, name, value }` keeps each field's label `name` (header) and `value` SEPARATE + raw (the renderer escapes/formats them). A node's `lines` are the fields at one indent level (≥1), rendered stacked; the first line carries the level's marker (or the Word number). Rows whose values match across all of a level's fields **merge** (composite group key).
- The pivot structure is `pivotLevels: number[][]` — ordered **indent buckets**; bucket index = indent depth, each bucket a non-empty ordered list of grid columns. Columns in no bucket are unused (hidden).
- `FieldLabel { show, bold, underline }` (per grid column, in `TableState.fieldLabels`) controls each field's label prefix (`name: `): hide it, bold it, underline it. Default = shown, plain.
- Rendered as a title row (`<p class="ws-title">`) + nested rows (`<p class="ws-lvl" data-level="N">`, numbered ones also `data-heading="K"`); `buildWordHtml` maps the title + numbered levels to Word heading styles, and the rest to the app's direct per-level look.

## Pivot view
The one view (`rowsToPivotTree`) replicates Excel's "Rows area" as ordered indent buckets:
- **Add fields** — a pool of the columns not yet placed; clicking one appends it as a new deepest single-field level.
- **Structure** — the placed fields as an indented list. Per field: ◄ **outdent** / ► **indent** (set the indent level; stacking fields at one level shows them together and merges by the composite of that level's values), ▲/▼ **reorder** (swap with the flat-order neighbour, keeping the bucket shape), ✕ **remove**, and three label toggles — **Aa** (show/hide the `name:` label prefix → "Item Category: Fruit" vs "Fruit"), **B** (bold the label), **U** (underline the label). Emphasis targets the label only; the value stays plain. Stored per grid column in `fieldLabels` (persists across remove/re-add).
- **Markers** (per-table, `markers: MarkerKind[]` indexed by indent level) — a `<select>` per level picks its marker style: `1.` / `1)` / `A.` / `a.` / `I.` / `i.` / `•` / `–` / None. Sparse → `defaultMarker(depth)` (the `1./a./i.` cycle). Only the FIRST field of a multi-field level is marked; the title and Word-numbered levels are never app-marked.
- **Number top N levels** (shared, `numberDepth`) — maps the top N body indent levels to `Heading 2/3/…` so Word numbers them live (`5.1`, `5.1.1`) from the destination document's heading numbering (only on a **Use Destination Styles** paste; the app suppresses its text marker there). Numbered items also land in Word's Nav pane / TOC. The browser preview can't show the real number (shows a `#` placeholder + a hint).
- **Section title** (optional) — the title row (`MsoTitle`). Maps to the **Heading style** name when set (a real Word heading, `mso-outline-level:1` → "5.0"); else a plain level-1 row. With a title the nested data starts at level 2; without one, level 1.

## Styling
**Title + numbered levels → Word heading mapping:** the **Heading style** input (default `Heading 1`) maps the title via `mso-style-name`+`mso-outline-level:1`, and **Number top N levels** (`numberDepth`) maps the top N body levels to `Heading 2/3/…` (`MsoHeadingK` + `mso-style-name:"Heading K"`+`mso-outline-level:K`). On a **Use Destination Styles** paste these adopt the destination document's heading styles + their live numbering (`5.0 / 5.1 / 5.1.1`). The exact number comes from the template, not the app; numbered items also enter Word's Nav pane / TOC. **Every non-numbered body row is direct** — the app's per-level look, emitted inline (so it survives the paste and matches the live preview), with per-field label `<b>`/`<u>` runs on top.

One **shared per-level** styling panel ("Level styles") drives the look across all tables — `HeadingStyle = { levels: LevelStyle[]; indentStep: number; headingStyleName }`. Each level row sets color/font/size/bold; **defaults are plain Arial 11 black** (to match a document body), with a "Reset levels" button. **Level 1** styles the title row (when not mapped); **Levels 2-9** style the nested rows by depth. The panel shows one row per depth in use. Beside it: a **Body font** control (document body font) and an **Indent/level (in)** control (`indentStep`, clamped 0–2; nested rows indent `(n-1)×step`). Body paragraphs use compact spacing (`line-height:1.25`, no space before/after). All style inputs are form-controlled (hex color, allow-listed font, clamped numbers), so no free user text lands in a `style`/`class` attribute.

## Current status
Full pivot pipeline implemented for multiple tables: paste (append) → parse → nest → render → live preview → per-table **Copy for Word** + combined **Copy all** (`text/html` + `text/plain`). Tables are managed in a tab strip (cap 100). Export is clipboard-only (Download-for-Word and the other layouts were removed). `.docx` generation is out of scope.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing the pivot convention or the data model.
- Never: install `xlsx` from npm, add a backend.
