# Excel &rarr; Word Pivot

Convert pasted Excel data into a **Word-ready nested outline**. Paste one or more tables copied from Excel or Google Sheets; the app restructures each into an Excel-pivot-style hierarchy and copies it to the clipboard so it pastes into Microsoft Word.

## Why

Wide Excel tables don't fit an 8.5" x 11" Word page — columns bleed off the edge and the table becomes unreadable. Rather than shrink or split the grid, this app turns one wide table into a narrow, nested outline (group by an ordered list of fields; the rest become detail lines under each item) that flows down the page.

## Status

**Working end to end:** paste → parse (SheetJS) → nest → render → live preview → **Copy for Word**. Paste several tables (managed in a tab strip), configure each table's pivot, style the levels once for all tables, and copy each table or all of them at once. See the [roadmap](./docs/ROADMAP.md) for what's out of scope (`.docx`).

## The pivot view

Pick an ordered list of **Nest by** fields; rows nest by that order (field 1 = outermost) and shared value-paths **merge**. Each nested row reads `Field name: value`. A **Detail fields** checklist picks the remaining columns to show as flat `Field: value` lines under each item. A **Number levels** toggle prefixes `1.`/`a.`/`i.` markers by depth. An optional **Section title** is the only Word heading (Heading 1); the nested rows + details are indented body text, so they don't clutter Word's navigation outline. A **View JSON** toggle inspects the raw parsed grid.

## Multiple tables

Paste table after table — each becomes a card in a horizontal **tab strip** (one edited at a time, cap 100). One shared **Heading levels** styling panel (Level 1 = the title, Levels 2-9 = the nested rows by depth) + a Body font apply to every table; each table keeps its own fields and title. Export per-table (**Copy for Word**) or everything at once (**Copy all**).

## Stack

- Next.js 16 (App Router) + TypeScript
- SheetJS (`xlsx`, official CDN tarball) for parsing
- Tailwind CSS v4
- Client-side only — no backend, no database, nothing uploaded or persisted.

## Getting started

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000), click the paste zone, and press **Ctrl/Cmd + V** with a cell range copied from Excel or Google Sheets. Pick your nest + detail fields; the live preview updates. Click **Copy for Word** and paste into a Word document (with *Use Destination Styles* so the title maps to your Heading 1).

## Docs

- [docs/OVERVIEW.md](./docs/OVERVIEW.md) — problem, goals, stack, status
- [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md) — data model, pivot, pipeline diagram
- [docs/ROADMAP.md](./docs/ROADMAP.md) — build order & status

## Project layout

```
app/                 App Router pages (home renders the paste view)
components/
  PasteInput.tsx     parent: paste/append, tables[] + shared styles, tab strip, Copy all
  TableCard.tsx      one table's pivot editor + preview + per-table Copy for Word
  tableModel.ts      TableState + tableToHtml (per-table nest->render) + pivotDetailColumnsOf
  RenderedPreview.tsx renders the pivot HTML (live preview; h2 title + [data-level] CSS)
  JsonPreview.tsx    shows the raw parsed Grid as JSON
lib/
  types.ts           PivotNode model + raw Grid
  parser.ts          SheetJS clipboard -> Grid
  mapper.ts          rowsToPivotTree (Grid -> PivotNode[]) + cellToString
  renderers.ts       renderPivotTree (tree -> HTML fragment) + numbering helpers
  clipboard.ts       Word-friendly clipboard wrapper (buildWordHtml / htmlToPlainText; HeadingStyle / LevelStyle)
docs/                OVERVIEW, ARCHITECTURE, ROADMAP
```
