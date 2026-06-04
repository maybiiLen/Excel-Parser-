"""Table Copy-Paste Optimizer.

A local Flask utility that ingests a wide spreadsheet (.xlsx / .csv), lets the
user collapse related columns into single cells ("Data Densification"), and
renders a clean HTML table that copy-pastes natively into Microsoft Word.

Nothing is written to disk: uploaded files are parsed in-memory and the parsed
records are held in a short-lived in-memory store keyed by a random token.
"""

from __future__ import annotations

import io
import math
import uuid

import pandas as pd
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)

# Max upload size: 10 MB (keeps everything comfortably in memory).
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

# In-memory store: token -> list[dict] (original rows) + original column order.
# This is intentionally volatile; nothing is persisted to disk.
_STORE: dict[str, dict] = {}

# --- Page geometry (the "golden rule") ----------------------------------------
# US Letter portrait, 0.5" margins all around -> 7.5" of printable width.
PRINTABLE_WIDTH_IN = 7.5
# Rough width budget for 11pt Calibri text, plus per-cell horizontal padding.
_CHAR_WIDTH_IN = 0.083
_CELL_PADDING_IN = 0.20
_MIN_COL_IN = 0.6
# A column narrower than this is hard to read once text wraps.
_COMFORT_MIN_IN = 0.9
# Cap how much a single fat column can demand when estimating display width.
_MAX_NATURAL_CHARS = 60

# --- Preset strategy definitions ----------------------------------------------
# Presets are just pre-built merge operations. A "part" is (column, prefix,
# suffix); blank values cause the whole part (prefix+value+suffix) to be skipped,
# so no stray "[]" or "()" artifacts are produced. Parts are joined by `sep`.

STRATEGIES = {
    "A": {
        "label": "Strategy A - Item Properties",
        "description": "Merge ItemCode + ItemName -> 'Item Details'  ([ItemCode] - ItemName)",
        "sources": ["ItemCode", "ItemName"],
        "new_column": "Item Details",
        "sep": " - ",
        "parts": [("ItemCode", "[", "]"), ("ItemName", "", "")],
    },
    "B": {
        "label": "Strategy B - Operational State",
        "description": "Merge Assignee + CurrentStatus -> 'Personnel (Status)'  (Assignee (CurrentStatus))",
        "sources": ["Assignee", "CurrentStatus"],
        "new_column": "Personnel (Status)",
        "sep": " ",
        "parts": [("Assignee", "", ""), ("CurrentStatus", "(", ")")],
    },
}


def _is_blank(value) -> bool:
    """True for None, NaN, or strings that are empty/whitespace only."""
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _clean(value) -> str:
    """Render a cell value to a trimmed string ('' when blank)."""
    if _is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        # Avoid '5.0' style artifacts from spreadsheet numeric coercion.
        return str(int(value))
    return str(value).strip()


def _available_strategies(columns: list[str]) -> dict[str, bool]:
    """Which presets can run given the present headers (all sources needed)."""
    present = set(columns)
    return {
        key: all(src in present for src in cfg["sources"])
        for key, cfg in STRATEGIES.items()
    }


def _build_operations(columns: list[str], options: dict) -> list[dict]:
    """Turn selected presets + user-defined custom merges into merge operations.

    Each operation: {"name", "sep", "parts": [(column, prefix, suffix), ...]}.
    Presets are evaluated first (in A, B order), then custom merges in the order
    the user added them. A column is only ever consumed by one operation.
    """
    avail = _available_strategies(columns)
    operations: list[dict] = []

    presets = options.get("presets", {}) or {}
    for key, cfg in STRATEGIES.items():
        if presets.get(key) and avail.get(key, False):
            operations.append(
                {"name": cfg["new_column"], "sep": cfg["sep"], "parts": list(cfg["parts"])}
            )

    present = set(columns)
    for merge in options.get("custom", []) or []:
        cols = [c for c in (merge.get("columns") or []) if c in present]
        if len(cols) < 2:
            continue  # a merge needs at least two real columns to be meaningful
        sep = merge.get("sep")
        if sep is None:
            sep = " - "
        name = (merge.get("name") or "").strip() or " / ".join(cols)
        operations.append(
            {"name": name, "sep": sep, "parts": [(c, "", "") for c in cols]}
        )

    return operations


def transform(records: list[dict], columns: list[str], options: dict) -> dict:
    """Apply preset + custom densification merges.

    Returns {'columns': [...], 'rows': [[...], ...]} preserving original order.
    A merged column is placed where its earliest source column appeared; the
    other consumed columns are dropped. Non-target columns are left intact.
    """
    operations = _build_operations(columns, options)

    col_index = {c: i for i, c in enumerate(columns)}
    consumed: set[str] = set()
    anchor_op: dict[str, dict] = {}  # earliest original column -> operation

    for op in operations:
        valid_parts = [
            (c, p, s) for (c, p, s) in op["parts"] if c in col_index and c not in consumed
        ]
        if len(valid_parts) < 2:
            continue
        op["parts"] = valid_parts
        for c, _p, _s in valid_parts:
            consumed.add(c)
        anchor = min((c for c, _p, _s in valid_parts), key=lambda c: col_index[c])
        anchor_op[anchor] = op

    out_columns: list[str] = []
    for col in columns:
        if col in consumed:
            if col in anchor_op:
                out_columns.append(anchor_op[col]["name"])
        else:
            out_columns.append(col)

    out_rows: list[list[str]] = []
    for rec in records:
        row: list[str] = []
        for col in columns:
            if col in consumed:
                if col in anchor_op:
                    op = anchor_op[col]
                    pieces = []
                    for c, prefix, suffix in op["parts"]:
                        value = _clean(rec.get(c))
                        if value:
                            pieces.append(f"{prefix}{value}{suffix}")
                    row.append(op["sep"].join(pieces))
            else:
                row.append(_clean(rec.get(col)))
        out_rows.append(row)

    layout = _compute_layout(out_columns, out_rows)
    return {"columns": out_columns, "rows": out_rows, "layout": layout}


def _longest_word(text: str) -> int:
    """Length of the longest whitespace-delimited token (the unbreakable unit)."""
    longest = 0
    for token in str(text).split():
        longest = max(longest, len(token))
    return longest


def _compute_layout(columns: list[str], rows: list[list[str]]) -> dict:
    """Decide per-column widths and whether the table is Word-page friendly.

    The table is always pinned to the 7.5" printable area, so it never bleeds off
    the page; text simply wraps. The meaningful question is therefore *readability*:
    each column needs to be at least as wide as its longest unbreakable word (so
    words don't snap mid-token) and at least a comfortable minimum. If the sum of
    those required widths exceeds 7.5", the table is too cramped -> merge columns.
    """
    n = len(columns)
    if n == 0:
        return {
            "printable_in": PRINTABLE_WIDTH_IN,
            "widths_in": [],
            "needed_in": 0.0,
            "fits": True,
        }

    required = []   # minimum readable width per column
    display = []     # preferred (content-proportional) width per column
    for i, col in enumerate(columns):
        longest_word = _longest_word(col)
        longest_cell = len(str(col))
        for row in rows:
            if i < len(row):
                cell = row[i]
                longest_word = max(longest_word, _longest_word(cell))
                longest_cell = max(longest_cell, len(cell))
        word_in = longest_word * _CHAR_WIDTH_IN + _CELL_PADDING_IN
        required.append(max(word_in, _COMFORT_MIN_IN))
        display.append(min(longest_cell, _MAX_NATURAL_CHARS) * _CHAR_WIDTH_IN + _CELL_PADDING_IN)

    needed_total = sum(required)
    fits = needed_total <= PRINTABLE_WIDTH_IN + 1e-6

    # Display widths: distribute the 7.5" proportionally to content, but never
    # below each column's required minimum. Pinned so the total is exactly 7.5".
    disp_total = sum(display) or 1.0
    widths = [w * PRINTABLE_WIDTH_IN / disp_total for w in display]
    widths = [max(w, max(_MIN_COL_IN, req)) for w, req in zip(widths, required)]
    total = sum(widths)
    if total > PRINTABLE_WIDTH_IN:
        widths = [w * PRINTABLE_WIDTH_IN / total for w in widths]

    return {
        "printable_in": PRINTABLE_WIDTH_IN,
        "widths_in": [round(w, 3) for w in widths],
        "needed_in": round(needed_total, 2),
        "fits": fits,
    }


def _parse_upload(file_storage) -> pd.DataFrame:
    """Read an uploaded .xlsx or .csv into a DataFrame (strings preserved)."""
    filename = (file_storage.filename or "").lower()
    data = file_storage.read()
    if not data:
        raise ValueError("The uploaded file is empty.")

    buf = io.BytesIO(data)
    if filename.endswith(".csv"):
        df = pd.read_csv(buf, dtype=object, keep_default_na=True)
    elif filename.endswith((".xlsx", ".xlsm")):
        df = pd.read_excel(buf, dtype=object, engine="openpyxl")
    else:
        raise ValueError("Unsupported file type. Please upload a .xlsx or .csv file.")

    # Normalize header whitespace.
    df.columns = [str(c).strip() for c in df.columns]
    return df


@app.route("/")
def index():
    preset_sources = {key: cfg["sources"] for key, cfg in STRATEGIES.items()}
    return render_template("index.html", strategies=STRATEGIES, preset_sources=preset_sources)


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files or request.files["file"].filename == "":
        return jsonify({"error": "No file was provided."}), 400

    try:
        df = _parse_upload(request.files["file"])
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:  # pragma: no cover - defensive parse guard
        return jsonify({"error": "Could not read the file. Is it a valid Excel/CSV?"}), 400

    columns = list(df.columns)
    if not columns or df.empty:
        return jsonify({"error": "The file has no readable rows or headers."}), 400

    # Defensive header check: if NONE of the expected strategy columns exist,
    # the densification feature cannot be applied to this file.
    avail = _available_strategies(columns)
    expected_any = any(col in columns for cfg in STRATEGIES.values() for col in cfg["sources"])

    records = df.where(pd.notnull(df), None).to_dict(orient="records")

    token = uuid.uuid4().hex
    _STORE[token] = {"records": records, "columns": columns}

    # Bound memory usage to the most recent few uploads.
    if len(_STORE) > 25:
        for old in list(_STORE.keys())[:-25]:
            _STORE.pop(old, None)

    warning = None
    if not expected_any:
        warning = (
            "None of the preset columns (ItemCode/ItemName/Assignee/CurrentStatus) "
            "were found, so the quick presets are disabled. You can still build your "
            "own custom merges below."
        )

    return jsonify(
        {
            "token": token,
            "columns": columns,
            "available": avail,
            "row_count": len(records),
            "warning": warning,
        }
    )


@app.route("/ingest", methods=["POST"])
def ingest():
    """Ingest a table pasted from the clipboard (parsed into columns + rows on
    the client). No file is uploaded; the grid is held in-memory only.
    """
    payload = request.get_json(silent=True) or {}
    columns = payload.get("columns") or []
    rows = payload.get("rows") or []

    columns = [str(c).strip() for c in columns]
    if not columns or all(c == "" for c in columns):
        return jsonify({"error": "Could not find any columns in the pasted data."}), 400
    if not rows:
        return jsonify({"error": "No data rows were found in the pasted table."}), 400

    # De-duplicate blank/identical headers so they remain addressable.
    seen: dict[str, int] = {}
    norm_columns = []
    for i, c in enumerate(columns):
        name = c or f"Column {i + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 0
        norm_columns.append(name)
    columns = norm_columns

    records = []
    for row in rows:
        rec = {}
        for i, col in enumerate(columns):
            rec[col] = row[i] if i < len(row) else None
        records.append(rec)

    avail = _available_strategies(columns)
    expected_any = any(col in columns for cfg in STRATEGIES.values() for col in cfg["sources"])

    token = uuid.uuid4().hex
    _STORE[token] = {"records": records, "columns": columns}
    if len(_STORE) > 25:
        for old in list(_STORE.keys())[:-25]:
            _STORE.pop(old, None)

    warning = None
    if not expected_any:
        warning = (
            "None of the preset columns (ItemCode/ItemName/Assignee/CurrentStatus) "
            "were found, so the quick presets are disabled. You can still build your "
            "own custom merges below."
        )

    return jsonify(
        {
            "token": token,
            "columns": columns,
            "available": avail,
            "row_count": len(records),
            "warning": warning,
        }
    )


@app.route("/preview", methods=["POST"])
def preview():
    payload = request.get_json(silent=True) or {}
    token = payload.get("token")
    entry = _STORE.get(token)
    if not entry:
        return jsonify({"error": "Session expired. Please re-upload your file."}), 404

    options = payload.get("options", {})
    result = transform(entry["records"], entry["columns"], options)
    result["available"] = _available_strategies(entry["columns"])
    return jsonify(result)


if __name__ == "__main__":
    # Port 5001 by default: macOS Control Center (AirPlay Receiver) often
    # occupies port 5000. Override with the PORT env var if needed.
    import os

    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 5001)), debug=True)
