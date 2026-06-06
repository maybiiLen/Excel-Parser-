# Roadmap

Build order for the MVP. Today's scope (steps 1-2) is **done**; everything below is sequenced but not yet built.

| # | Step | Status |
| - | ---- | ------ |
| 1 | **Data model** — `Section` / `Subsection` / `Body` ([`lib/types.ts`](../lib/types.ts)) | Done |
| 2 | **Paste -> raw rows** — SheetJS clipboard parse + JSON preview ([`lib/parser.ts`](../lib/parser.ts), [`components/`](../components/)) | Done |
| 3 | **Convention mapper** — rows -> section tree (A=section, B=sub, C=body, D=type) ([`lib/mapper.ts`](../lib/mapper.ts)) | Todo |
| 4 | **Body-type detection** — text / bullets / table from column D | Todo |
| 5 | **Numbering** — tree walk assigns dotted numbers ([`lib/numbering.ts`](../lib/numbering.ts)) | Todo |
| 6 | **Root-number input** — configurable starting number (default 6) | Todo |
| 7 | **Body renderers** — body -> Word-friendly HTML ([`lib/renderers.ts`](../lib/renderers.ts)) | Todo |
| 8 | **Wide-table strategy** — transpose or split so tables fit a Word page | Todo |
| 9 | **Clipboard output** — serialize tree to `text/html`, copy for Word | Todo |
| 10 | **Malformed-paste handling** — graceful errors on bad / empty input | Todo |
| 11 | **Cleanup & polish** | Todo |

## Today (Day 1) — done

- Project scaffolded (Next.js + TypeScript + Tailwind v4 + SheetJS).
- Folder structure and docs in place.
- `PasteInput` captures a paste event, `parser.ts` parses it with SheetJS, `JsonPreview` shows the raw rows as JSON.
- Verified on localhost.
