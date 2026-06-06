@AGENTS.md
# Excel → Word Sections

Client-side Next.js app. User pastes an Excel table, app converts it into numbered Word-document sections (a section tree), outputs HTML that pastes into Word as native tables and headings.

The problem this solves: wide Excel tables do not fit an 8.5x11 Word page. This app restructures spreadsheet data into readable numbered sections.

## Commands
- `npm run dev` — dev server on localhost:3000
- `npm run build` — production build, must pass with no type errors
- `npm run lint` — ESLint check

## Stack
- Next.js 15, App Router, TypeScript
- Tailwind CSS v4 (CSS-first, no tailwind.config.js)
- SheetJS for parsing

## Hard Rules
- SheetJS install is the official CDN tarball: `https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`. NEVER install `xlsx` from npm. The npm version is stale (0.18.5).
- Client-side only. No API routes, no backend, no database. Parsing runs in the browser.
- Files at repo root. No `src/` directory. Import alias `@/*` maps to `./*`.
- SheetJS clipboard parse: `getData('text/html')` first, `getData('text/plain')` TSV fallback. Use `sheet_to_json(ws, { header: 1, blankrows: false, defval: "", raw: false })`. Keep `defval: ""` so blank cells stay as empty strings and column positions stay aligned.

## Architecture
- `app/page.tsx` — renders the paste view
- `components/PasteInput.tsx` — client component, captures paste
- `components/JsonPreview.tsx` — displays parsed grid
- `lib/types.ts` — Section, Subsection, Body, and raw Grid types
- `lib/parser.ts` — SheetJS clipboard to rows (implemented)
- `lib/mapper.ts` — rows to tree (stub, throws)
- `lib/numbering.ts` — tree-walk numbering (stub, throws)
- `lib/renderers.ts` — body to HTML (stub, throws)
- `docs/` — OVERVIEW, ARCHITECTURE, ROADMAP

## Data Model
The product is a section tree, not a flat table:
- `Section { number, title, children: Subsection[] }`
- `Subsection { number, title, body: Body }`
- `Body = { type: "text", content } | { type: "bullets", items } | { type: "table", rows }`

## Input Convention (MVP)
Column position defines role:
- Column A filled = new section title
- Column A blank = subsection belongs to the section above
- Column B = subsection title
- Column C = body content
- Column D = body type flag ("text" | "bullet" | "table")

## Scope: Today
Build paste-to-JSON only. Capture paste, parse with SheetJS, display the raw grid as JSON on screen. Nothing else.

Do NOT build today: section mapping, numbering, body renderers, clipboard output, .docx generation. These stay as stubs that throw `Error("not implemented")`. Do not implement a stub unless the task explicitly names it.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing the input convention or data model.
- Never: install `xlsx` from npm, add a backend, build past today's scope without being asked.