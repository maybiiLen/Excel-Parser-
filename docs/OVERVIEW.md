# Overview

## What this is

A web app that converts **pasted Excel data into Word-ready document sections**. You paste a table copied from Excel (or Google Sheets); the app restructures it into a section tree, renders it as HTML, and copies it to the clipboard so it pastes into Microsoft Word as native headings and bullet lists.

## The problem it solves

Wide Excel tables do not fit on an 8.5" x 11" Word page and become unreadable once columns bleed off the right edge. Instead of shrinking or splitting the grid, this app **restructures** spreadsheet data into readable, narrow sections — turning one unwieldy wide table into a sequence of headings, each followed by a short bullet list, so it flows down a Word page instead of off the side.

## Tech stack

- **Next.js 16** (App Router) + **TypeScript**
- **SheetJS** (`xlsx`, official CDN tarball) for parsing clipboard / Excel data
- **Tailwind CSS** (v4) for styling
- **Client-side only** — no backend, no API routes, no database. All parsing runs in the browser; nothing is uploaded or persisted.

## What works today

The full pipeline is implemented end to end:

1. **Paste → parse.** A client component captures a paste event and parses it with SheetJS into a raw Grid.
2. **Map → tree.** One of three view modes turns the Grid into a section tree:
   - **Grouped by field** (default) — group rows by a chosen field; members listed as bullets.
   - **Fields as bullets** — one section per row; chosen fields as `Field: value` bullets.
   - **A/B/C/D sections** — the original position-based convention (column A = section, B = subsection, C = body, D = type).
3. **Render → preview.** The tree is rendered to HTML and shown as a live preview (or toggle to inspect the raw Grid as JSON).
4. **Copy for Word.** A button writes `text/html` (+ a `text/plain` fallback) to the clipboard; pasting into Word yields native headings and bullet lists that fit a Letter page.

See [ARCHITECTURE.md](./ARCHITECTURE.md) for the data model, view modes, and pipeline, and [ROADMAP.md](./ROADMAP.md) for status.

## Numbering

The grouped and per-item views wrap their output in one numbered, titled section — you pick a **Section #** (default 1) and a **Section title**, e.g. `5 Fruit Database`, and the items beneath are numbered `5.1`, `5.2`, … (`lib/numbering.ts` `wrapInNumberedSection`). The A/B/C/D view is left un-numbered.

## Not wired in

- **Wide-table transpose/split** is not implemented (largely moot, since the active views emit narrow bullet lists rather than tables).
- **`.docx` generation** is out of scope; export is HTML-on-clipboard only.
