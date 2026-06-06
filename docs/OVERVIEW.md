# Overview

## What this is

A web app that converts **pasted Excel data into formatted, numbered Word-document sections**. You paste a table copied from Excel (or Google Sheets); the app maps it into a hierarchical section tree, applies automatic section numbering, and outputs HTML that pastes cleanly into Microsoft Word as native tables and headings.

## The problem it solves

Wide Excel tables do not fit on an 8.5" x 11" Word page and become unreadable once columns bleed off the right edge. Instead of shrinking or splitting the grid, this app **restructures** spreadsheet data into readable, numbered report sections — turning one unwieldy wide table into a sequence of `6`, `6.1`, `6.2 ...` sections, each with its own short table, bullet list, or paragraph.

## Tech stack

- **Next.js** (App Router) + **TypeScript**
- **SheetJS** (`xlsx`) for parsing clipboard / Excel data
- **Tailwind CSS** (v4) for styling
- **Client-side only** — no backend, no API routes, no database. All parsing runs in the browser; nothing is uploaded or persisted.

## Today's scope (Day 1)

A deliberately small proof:

1. Scaffold the project, folder structure, and docs.
2. A single client component captures a paste event, parses the pasted Excel data with SheetJS, and displays the resulting JSON on screen.
3. Runs on localhost — paste-to-JSON confirmed working.

Everything else (convention mapping, numbering, body renderers, wide-table strategy, clipboard output, Word-friendly export) is **stubbed** and scheduled — see [ROADMAP.md](./ROADMAP.md). The data flow and diagram live in [ARCHITECTURE.md](./ARCHITECTURE.md).
