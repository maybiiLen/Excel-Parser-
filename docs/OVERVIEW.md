# Overview

## What this is

A web app that converts **pasted Excel data into a Word-ready nested outline**. You paste one or more tables copied from Excel (or Google Sheets); the app restructures each into an Excel-pivot-style hierarchy ("Rows area") and copies it to the clipboard so it pastes into Microsoft Word.

## The problem it solves

Wide Excel tables do not fit on an 8.5" x 11" Word page and become unreadable once columns bleed off the right edge. Instead of shrinking or splitting the grid, this app **restructures** spreadsheet data into a narrow, nested outline — you arrange fields into ordered indent levels (each level can stack several fields at the same indentation), and rows nest and merge by those levels — so it flows down a Word page instead of off the side.

## Tech stack

- **Next.js 16** (App Router) + **TypeScript**
- **SheetJS** (`xlsx`, official CDN tarball) for parsing clipboard / Excel data
- **Tailwind CSS** (v4) for styling
- **Client-side only** — no backend, no API routes, no database. All parsing runs in the browser; nothing is uploaded or persisted.

## What works today

The full pivot pipeline is implemented end to end, for multiple tables:

1. **Paste → parse.** A client component captures each paste and parses it with SheetJS into a raw Grid, appending a new table. Tables are managed as cards in a horizontal tab strip (one edited at a time, cap 100).
2. **Nest → tree.** `rowsToPivotTree` nests rows by an ordered list of **indent buckets** (`pivotLevels`) into an arbitrary-depth `PivotNode` tree. Each bucket is one indent level holding one or more fields; fields stacked in a bucket render at the same indent and rows merge by the composite of that level's values.
3. **Render → preview.** The tree is rendered to HTML and shown in each card's live preview (or toggle to inspect the raw Grid as JSON). A per-level **Markers** picker adds `1./a./i.`/etc.; per field you can hide/bold/underline the `Field name:` label.
4. **Copy for Word.** Each card writes `text/html` (+ a `text/plain` fallback) to the clipboard; a combined **Copy all** exports every table as one document. The **Section title** maps to your document's Heading 1 ("5.0"), and **Number top N levels** maps the top body levels to `Heading 2/3/…` so Word numbers them live (`5.1`, `5.1.1`) on a *Use Destination Styles* paste; the rest are styled body text.

**Styling.** One shared **per-level** panel ("Heading levels") styles the pivot across all tables: Level 1 = the title, Levels 2-9 = the nested rows by depth — color/font/size/bold each. The title and the top-N numbered levels become real Word headings (they appear in the Nav pane / TOC); the rest stay out of Word's outline.

See [ARCHITECTURE.md](./ARCHITECTURE.md) for the data model, the pivot, and the pipeline, and [ROADMAP.md](./ROADMAP.md) for status.

## Out of scope

- **`.docx` generation** — export is HTML-on-clipboard only ("Copy for Word"). There is no download.
