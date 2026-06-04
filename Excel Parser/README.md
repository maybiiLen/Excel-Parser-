# Table Copy-Paste Optimizer

A small local web app that fixes a very specific pain: **pasting a wide table into
an 8.5" × 11" Portrait Word document where the columns bleed off the right edge.**

Workflow in one line: **paste your table in → see if it fits a Word page → merge
columns if it's too wide → copy a clean, Word-ready table back out.**

You never have to hand over a whole file — copy a cell range in Excel/Google Sheets
(Ctrl/Cmd+C), paste it into the app (Ctrl/Cmd+V), and go.

---

## Features

- **Paste-to-go ingestion** — paste a range copied from Excel/Sheets. The app reads
  the clipboard's HTML table, with a tab/comma plain-text fallback.
- **Optional file upload** of `.xlsx` / `.csv` as a secondary path.
- **Two-panel workflow** — configuration on the left, a live, true-scale output
  table on the right that updates instantly.
- **Word-compatibility indicator** — tells you up front whether the table is
  **page-ready** (green) or **too cramped** (amber). The check is readability-based:
  each column must be wide enough that words don't break; those minimums are summed
  and compared to the 7.5" printable width. Too wide → you're prompted to merge.
- **Data densification (column merging):**
  - **Quick presets** (toggle independently):
    - **Strategy A — Item Properties:** `ItemCode` + `ItemName` →
      `Item Details`, formatted `[ItemCode] - ItemName`.
    - **Strategy B — Operational State:** `Assignee` + `CurrentStatus` →
      `Personnel (Status)`, formatted `Assignee (CurrentStatus)`.
  - **Custom merges** — pick *any* columns, name the merged column, choose a
    separator. Add as many as you like; remove anytime. Columns already used by a
    preset/merge are locked to prevent conflicts.
- **The golden rule (fits 8.5 × 11 portrait):** the output table is pinned to the
  exact **7.5" printable width** (Letter minus 0.5" margins) with explicit per-column
  inch widths, so it cannot bleed off the page. The preview renders inside a
  **true-scale page frame** with margin guides so you see the real footprint.
- **Copy that survives the paste** — the **Copy for Word** button writes the table to
  the clipboard with **every style inlined** (borders, header shading, fonts, and the
  per-column inch widths), exactly the way Excel does it. Word ignores external CSS,
  so inline styling is what keeps the grid intact on paste.
- **Graceful null handling** — no stray `[]` or `()` when a value is missing.
- **Non-target columns are left completely untouched.**
- **Defensive header checks** — preset toggles auto-disable when their columns are
  absent; a warning appears if none of the preset columns exist (custom merges still
  work).
- **Nothing is persisted to disk** — uploads/pastes are parsed and held **in-memory
  only**, in a small, size-bounded store keyed by a random token.

---

## Tech stack

- **Backend:** Python + [Flask](https://flask.palletsprojects.com/)
- **Parsing:** [pandas](https://pandas.pydata.org/) (+ `openpyxl` for `.xlsx`)
- **Frontend:** vanilla HTML/CSS/JS (no build step)

---

## Setup

```bash
python3 -m venv venv
./venv/bin/pip install -r requirements.txt
```

## Run

```bash
./venv/bin/python app.py
```

Then open **http://127.0.0.1:5001** in your browser.

> The default port is `5001` because port `5000` is commonly taken by macOS Control
> Center / AirPlay Receiver. Override it with `PORT=8000 ./venv/bin/python app.py`.

---

## Usage

1. In Excel/Sheets, select your table and copy it (**Ctrl/Cmd+C**).
2. Click the paste box in the app and paste (**Ctrl/Cmd+V**).
   _(Keep "First row contains column headers" checked if your copy includes a header
   row. Or use the file-upload link; a `sample_data.csv` is included.)_
3. Check the compatibility banner. If it's amber (**"Too cramped for Word"**), toggle
   the quick presets and/or build custom merges on the left until it turns green.
4. Click **Copy for Word** (or **Select table**, then Ctrl/Cmd+C) and paste into your
   Word document — borders, shading, and the 7.5" width come across natively.

---

## How it works

```
┌──────────────┐   paste / upload    ┌───────────────┐   options    ┌──────────────┐
│  Clipboard   │ ──────────────────▶ │  Flask store   │ ───────────▶ │  transform   │
│ (HTML/TSV) / │   /ingest, /upload  │ (in-memory,    │   /preview   │  + layout    │
│   .xlsx/.csv │                     │  token-keyed)  │              │  (7.5" fit)  │
└──────────────┘                     └───────────────┘              └──────┬───────┘
                                                                           │ JSON
                                                                           ▼
                                                          live preview + inline-styled
                                                          "Copy for Word" output
```

- **Ingestion** parses the table into `columns` + `rows`, stores it in-memory under a
  random token, and reports which presets are available.
- **Transform** applies the selected preset/custom merges. Each merge is a list of
  parts `(column, prefix, suffix)` joined by a separator; blank values are skipped so
  no empty brackets/parens appear. A merged column takes the position of its earliest
  source column; other consumed columns are dropped; everything else is left intact.
- **Layout** estimates a readable minimum width per column (longest unbreakable word,
  floored at a comfortable minimum), decides whether the total fits 7.5", and computes
  explicit per-column inch widths used for both the preview and the copy output.

### HTTP endpoints

| Method | Path       | Purpose                                                        |
| ------ | ---------- | ------------------------------------------------------------- |
| `GET`  | `/`        | Serve the app.                                                 |
| `POST` | `/ingest`  | Ingest a pasted grid `{columns, rows}` → `{token, ...}`.       |
| `POST` | `/upload`  | Ingest an uploaded `.xlsx`/`.csv` file → `{token, ...}`.       |
| `POST` | `/preview` | `{token, options}` → transformed `{columns, rows, layout}`.    |

`options` shape: `{ "presets": { "A": true, "B": false }, "custom": [ { "name": "...", "columns": ["A","B"], "sep": " - " } ] }`

---

## Project layout

```
app.py                 # Flask backend: ingestion, merge engine, layout/fit, endpoints
templates/index.html   # UI shell (paste zone, config panel, page-frame preview)
static/style.css       # Styling, incl. true-scale page frame + table styles
static/app.js          # Clipboard parsing, live preview, fit indicator, copy-to-Word
sample_data.csv        # Example dataset matching the preset schema
requirements.txt       # Flask, pandas, openpyxl
```

---

## Troubleshooting

- **`ERR_CONNECTION_REFUSED` / page won't load:** the server isn't running. Check with
  `lsof -ti tcp:5001` (empty = down), then restart with `./venv/bin/python app.py`.
- **"Address already in use" on start:** something else holds the port. Use another
  port (`PORT=8000 ./venv/bin/python app.py`) — note macOS uses **5000** for AirPlay.
- **Pasted output looks like plain text in Word:** make sure you used **Copy for Word**
  (it inlines all styles). Also hard-refresh the app (Cmd/Ctrl+Shift+R) so the latest
  `app.js` is loaded.
- **Couldn't read the pasted table:** copy an actual cell *range* from the spreadsheet
  (not a screenshot), and keep the header checkbox consistent with your selection.

---

## Notes & limitations

- Data lives in memory only and is dropped when the server restarts (max ~25 recent
  sessions are retained). This is a local utility, not a multi-user service.
- Width/fit is a heuristic tuned for ~11pt Calibri; real Word rendering may vary
  slightly by font and content.
- Upload size is capped at 10 MB.
```
