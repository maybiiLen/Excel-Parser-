# Excel &rarr; Word Sections

Convert pasted Excel data into **Word-ready document sections**. Paste one or more tables copied from Excel or Google Sheets; the app restructures each into a section tree, renders it as HTML, and copies it to the clipboard so it pastes into Microsoft Word as native headings and bullet lists.

## Why

Wide Excel tables don't fit an 8.5" x 11" Word page — columns bleed off the edge and the table becomes unreadable. Rather than shrink or split the grid, this app turns one wide table into a sequence of readable sections that flow down the page (a heading, then a short bullet list), so it fits.

## Status

**Working end to end:** paste → parse (SheetJS) → map → render → live preview → **Copy/Download for Word**. Paste several tables (managed in a tab strip), choose each table's layout and columns, style headings once for all tables, and export each table or all of them at once. See the [roadmap](./docs/ROADMAP.md) for what's intentionally not wired in (wide-table transpose, `.docx`).

## Views

Each table picks its own layout:

- **Grouped by field** (default) — group rows by a chosen field; each value is a heading, members are bullets.
- **Fields as bullets** — one section per row; chosen fields as `Field: value` bullets.
- **Pivot (nested rows)** — Excel "Rows area": pick fields in order; rows nest by that order (shared paths merge). Optional plain title on top; shared per-level styling (color/font/size per depth).
- **A/B/C/D sections** — the original position convention (A = section, B = subsection, C = body, D = type).

A field checklist controls which columns show, and a **View JSON** toggle inspects the raw parsed grid. Grouped/per-item tables can be slotted into a document under a numbered, titled section (e.g. `5 Fruit Database` → 5.1, 5.2, …).

## Multiple tables

Paste table after table — each becomes a card in a horizontal **tab strip** (one edited at a time, cap 100). Heading appearance is shared across all tables; each table keeps its own layout, columns, and title. Export per-table (**Copy/Download for Word**) or everything at once (**Copy all / Download all**).

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

Open [http://localhost:3000](http://localhost:3000), click the paste zone, and press **Ctrl/Cmd + V** with a cell range copied from Excel or Google Sheets. The data renders as sections in the live preview; click **Copy for Word** and paste into a Word document.

## Docs

- [docs/OVERVIEW.md](./docs/OVERVIEW.md) — problem, goals, stack, status
- [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md) — data model, view modes, pipeline diagram
- [docs/ROADMAP.md](./docs/ROADMAP.md) — build order & status

## Project layout

```
app/                 App Router pages (home renders the paste view)
components/
  PasteInput.tsx     parent: paste/append, tables[] + shared styles, tab strip, Copy/Download all
  TableCard.tsx      one table's editor + preview + per-table Copy/Download
  tableModel.ts      TableState / Layout + tableToHtml (per-table map->render)
  RenderedPreview.tsx renders the section HTML (live preview; h2/h3 + pivot data-level CSS)
  JsonPreview.tsx    shows the raw parsed Grid as JSON
lib/
  types.ts           Section / Subsection / Body / PivotNode model + raw Grid
  parser.ts          SheetJS clipboard -> Grid
  mapper.ts          Grid -> tree (rowsToTree / rowsToAttributeSections / rowsToGroupedSections / rowsToPivotTree)
  numbering.ts       wrapInNumberedSection: wrap grouped/per-item output in a numbered, titled section
  renderers.ts       tree -> HTML fragment (renderTree / renderBody / renderPivotTree)
  clipboard.ts       Word-friendly clipboard wrapper (buildWordHtml / htmlToPlainText; HeadingStyle / LevelStyle)
docs/                OVERVIEW, ARCHITECTURE, ROADMAP
```
