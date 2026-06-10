@AGENTS.md
# Excel → Word Sections

Client-side Next.js app. User pastes an Excel/Sheets table; the app restructures it into a section tree, renders Word-ready HTML (native headings + bullet lists), and copies it to the clipboard for pasting into Word.

The problem this solves: wide Excel tables do not fit an 8.5x11 Word page. This app restructures spreadsheet data into readable, narrow sections (grouped, per-item, or by an A/B/C/D convention) that fit the page. Section numbering exists as a utility (`lib/numbering.ts`) but is not currently wired into the UI.

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

## Architecture
Pipeline: paste → `parseClipboard` (Grid) → a mapper (chosen by view mode) → `Section[]` → `renderTree` (HTML fragment) → live preview + "Copy for Word" clipboard write.
- `app/page.tsx` — renders the paste view
- `components/PasteInput.tsx` — client component; captures paste, holds state, switches view modes, copies for Word
- `components/RenderedPreview.tsx` — injects the rendered HTML via `dangerouslySetInnerHTML` (input is escaped by the renderer)
- `components/JsonPreview.tsx` — displays the raw parsed Grid as JSON
- `lib/types.ts` — Section, Subsection, Body, and raw Grid types
- `lib/parser.ts` — SheetJS clipboard → Grid (`parseClipboard`)
- `lib/mapper.ts` — Grid → Section tree; three mappers: `rowsToTree` (A/B/C/D), `rowsToAttributeSections` (per-item), `rowsToGroupedSections` (group-by)
- `lib/numbering.ts` — `numberTree` dotted numbering; implemented but NOT wired into the UI
- `lib/renderers.ts` — `renderTree` (tree → HTML fragment) + `renderBody`; escapes user text, omits blank numbers
- `lib/clipboard.ts` — `buildWordHtml` (Word-friendly doc wrapper) + `htmlToPlainText`
- `docs/` — OVERVIEW, ARCHITECTURE, ROADMAP

## Data Model
The product is a section tree, not a flat table:
- `Section { number, title, children: Subsection[], body? }` — `body?` carries content directly under a section heading (used by the per-item and grouped views)
- `Subsection { number, title, body: Body }`
- `Body = { type: "text", content } | { type: "bullets", items } | { type: "table", rows }`

## View modes
The paste UI offers three layouts (selectable; default **Grouped by field**):
- **Grouped by field** (`rowsToGroupedSections`) — row 0 is headers; rows are grouped by a chosen field (each distinct value → a heading), members listed as bullets (a label column + optional checked fields in parens).
- **Fields as bullets** (`rowsToAttributeSections`) — header row + one row per item; each row → a section titled by a chosen column, other selected fields as "Field: value" bullets.
- **A/B/C/D sections** (`rowsToTree`, the original position convention) — Column A filled = section title; A blank = subsection of the section above; B = subsection title; C = body content; D = body type flag ("text" | "bullet" | "table").

The first two are header-aware (row 0 = field names) and share a field checklist for picking which columns appear. Numbering is not applied to any view's output.

## Current status
The full pipeline is implemented: paste → parse → map (3 view modes) → render → live preview → **Copy for Word** (writes `text/html` + `text/plain` to the clipboard). Numbering (`numberTree`) and a wide-table width strategy (transpose/split) are NOT wired in; `.docx` generation is out of scope. The active views emit headings + bullet lists (narrow, page-fitting); only the A/B/C/D view can emit a raw `table` body (via the D flag), which is not yet width-managed.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing a view-mode convention or the data model.
- Never: install `xlsx` from npm, add a backend.