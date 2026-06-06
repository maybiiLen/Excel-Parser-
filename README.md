# Excel &rarr; Word Sections

Convert pasted Excel data into formatted, **numbered Word-document sections**. Paste a table copied from Excel or Google Sheets; the app restructures it into a hierarchical, automatically-numbered section tree and (eventually) outputs HTML that pastes cleanly into Microsoft Word as native headings and tables.

## Why

Wide Excel tables don't fit an 8.5" x 11" Word page — columns bleed off the edge and the table becomes unreadable. Rather than shrink or split the grid, this app turns one wide table into a sequence of readable numbered sections (`6`, `6.1`, `6.2 ...`), each with its own short table, bullet list, or paragraph.

## Status

**Day-1 proof working:** paste &rarr; parse with SheetJS &rarr; raw JSON on screen. The mapper, numbering, body renderers, wide-table strategy, and Word clipboard export are scaffolded as typed stubs and scheduled in the [roadmap](./docs/ROADMAP.md).

## Stack

- Next.js (App Router) + TypeScript
- SheetJS (`xlsx`) for parsing
- Tailwind CSS v4
- Client-side only — no backend, no database, nothing uploaded or persisted.

## Getting started

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000), click the paste zone, and press **Ctrl/Cmd + V** with a cell range copied from Excel or Google Sheets. The parsed rows appear as JSON.

## Docs

- [docs/OVERVIEW.md](./docs/OVERVIEW.md) — problem, goals, stack
- [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md) — data model, input convention, pipeline diagram
- [docs/ROADMAP.md](./docs/ROADMAP.md) — build order

## Project layout

```
app/                 App Router pages (home renders the paste view)
components/
  PasteInput.tsx     captures the paste event, parses, holds state
  JsonPreview.tsx    shows the parsed Grid as JSON
lib/
  types.ts           Section / Subsection / Body model + raw Grid
  parser.ts          SheetJS clipboard -> rows   (implemented)
  mapper.ts          rows -> section tree         (stub)
  numbering.ts       dotted-number walk            (stub)
  renderers.ts       body -> HTML                  (stub)
docs/                OVERVIEW, ARCHITECTURE, ROADMAP
```
