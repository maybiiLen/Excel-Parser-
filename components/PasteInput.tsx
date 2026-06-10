"use client";

import { useMemo, useState, type ClipboardEvent } from "react";
import { parseClipboard } from "@/lib/parser";
import {
  rowsToTree,
  rowsToAttributeSections,
  rowsToGroupedSections,
} from "@/lib/mapper";
import { wrapInNumberedSection } from "@/lib/numbering";
import { renderTree } from "@/lib/renderers";
import { buildWordHtml, htmlToPlainText } from "@/lib/clipboard";
import type { Grid } from "@/lib/types";
import { JsonPreview } from "./JsonPreview";
import { RenderedPreview } from "./RenderedPreview";

type Layout = "list" | "grouped" | "sections";

/** First header containing "name" (case-insensitive), else the first column. */
function pickDefaultTitleCol(rows: Grid): number {
  const header = rows[0] ?? [];
  const i = header.findIndex(
    (c) => typeof c === "string" && c.toLowerCase().includes("name"),
  );
  return i >= 0 ? i : 0;
}

const GROUP_KEYWORDS = [
  "origin", "season", "type", "category", "status", "risk", "level",
  "color", "region", "class", "group", "storage", "method",
];

/** First header matching a category-ish keyword, else first column != titleCol. */
function pickDefaultGroupCol(rows: Grid, titleCol: number): number {
  const header = (rows[0] ?? []).map((c) =>
    typeof c === "string" ? c.toLowerCase() : "",
  );
  for (let i = 0; i < header.length; i++) {
    if (i !== titleCol && GROUP_KEYWORDS.some((k) => header[i].includes(k))) {
      return i;
    }
  }
  for (let i = 0; i < header.length; i++) {
    if (i !== titleCol) return i;
  }
  return titleCol;
}

export function PasteInput() {
  const [grid, setGrid] = useState<Grid | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [view, setView] = useState<"rendered" | "json">("rendered");
  const [layout, setLayout] = useState<Layout>("grouped");
  const [titleCol, setTitleCol] = useState<number>(0);
  const [groupCol, setGroupCol] = useState<number>(0);
  const [selectedCols, setSelectedCols] = useState<Set<number>>(new Set());
  const [copyState, setCopyState] = useState<"idle" | "copied" | "error">("idle");
  // Section number + title for the wrapping heading (e.g. "5 Fruit Database").
  // Sticky: set to match the target document, not reset on paste/clear.
  const [sectionNumber, setSectionNumber] = useState<number>(1);
  const [sectionTitle, setSectionTitle] = useState<string>("");

  // Header row, for the title-column picker and field checklist.
  const headers = useMemo(
    () => (grid?.[0] ? grid[0].map((c) => (c == null ? "" : String(c))) : []),
    [grid],
  );

  // Columns shown as bullets / parenthetical extras, in column order: checked,
  // minus the title (label) column, and -- in grouped mode -- the group column.
  const fieldColumns = useMemo(
    () =>
      headers
        .map((_, i) => i)
        .filter(
          (i) =>
            i !== titleCol &&
            (layout !== "grouped" || i !== groupCol) &&
            selectedCols.has(i),
        ),
    [headers, titleCol, groupCol, layout, selectedCols],
  );

  // parse -> map -> (wrap in a numbered section) -> render, derived from the
  // stored grid so layout/column/field/section changes re-render without a
  // re-paste. The grouped and per-item views are wrapped under one numbered,
  // titled heading (5 Fruit Database -> 5.1, 5.2, ...); A/B/C/D is left as-is.
  const tree = useMemo(() => {
    if (!grid) return null;
    if (layout === "list")
      return wrapInNumberedSection(
        rowsToAttributeSections(grid, titleCol, fieldColumns),
        sectionNumber,
        sectionTitle.trim(),
      );
    if (layout === "grouped")
      return wrapInNumberedSection(
        rowsToGroupedSections(grid, groupCol, titleCol, fieldColumns),
        sectionNumber,
        sectionTitle.trim(),
      );
    return rowsToTree(grid);
  }, [grid, layout, titleCol, groupCol, fieldColumns, sectionNumber, sectionTitle]);
  const html = useMemo(() => (tree ? renderTree(tree) : ""), [tree]);

  function handlePaste(e: ClipboardEvent<HTMLDivElement>) {
    e.preventDefault();
    try {
      const rows = parseClipboard(e.clipboardData);
      if (rows.length === 0) {
        setGrid(null);
        setError(
          "Couldn't read a table from the clipboard. Copy a cell range from Excel or Google Sheets, then paste here.",
        );
        return;
      }
      setError(null);
      setGrid(rows);
      const t = pickDefaultTitleCol(rows);
      setTitleCol(t);
      setGroupCol(pickDefaultGroupCol(rows, t));
      setSelectedCols(new Set((rows[0] ?? []).map((_, i) => i))); // all fields on
    } catch (err) {
      setGrid(null);
      setError(
        err instanceof Error ? err.message : "Failed to parse the pasted data.",
      );
    }
  }

  function toggleCol(i: number) {
    setSelectedCols((prev) => {
      const next = new Set(prev);
      if (next.has(i)) next.delete(i);
      else next.add(i);
      return next;
    });
  }

  // Write the rendered sections to the clipboard as text/html (+ a text/plain
  // fallback) so they paste into Word as native headings + bullet lists. The
  // ClipboardItem is built synchronously from the in-memory `html` before any
  // await, preserving the user-gesture requirement.
  async function copyForWord() {
    if (typeof ClipboardItem === "undefined" || !navigator.clipboard?.write) {
      setCopyState("error");
      setTimeout(() => setCopyState("idle"), 2000);
      return;
    }
    try {
      const item = new ClipboardItem({
        "text/html": new Blob([buildWordHtml(html)], { type: "text/html" }),
        "text/plain": new Blob([htmlToPlainText(html)], { type: "text/plain" }),
      });
      await navigator.clipboard.write([item]);
      setCopyState("copied");
    } catch {
      setCopyState("error");
    }
    setTimeout(() => setCopyState("idle"), 2000);
  }

  function clear() {
    setGrid(null);
    setError(null);
    setView("rendered");
  }

  return (
    <div className="flex w-full flex-col gap-4">
      <div
        role="textbox"
        tabIndex={0}
        aria-label="Paste Excel data here"
        onPaste={handlePaste}
        className="flex min-h-[8rem] cursor-text items-center justify-center rounded-xl border-2 border-dashed border-foreground/25 bg-foreground/[0.02] p-6 text-center text-foreground/60 outline-none transition-colors focus:border-foreground/50 focus:bg-foreground/[0.04]"
      >
        <span>
          Click here and press{" "}
          <kbd className="rounded border border-foreground/30 px-1.5 py-0.5 font-mono text-xs">
            Ctrl
          </kbd>{" "}
          +{" "}
          <kbd className="rounded border border-foreground/30 px-1.5 py-0.5 font-mono text-xs">
            V
          </kbd>{" "}
          to paste a table copied from Excel or Google Sheets.
        </span>
      </div>

      {error && (
        <p className="rounded-lg border border-amber-500/40 bg-amber-500/10 px-4 py-3 text-sm text-amber-700 dark:text-amber-300">
          {error}
        </p>
      )}

      {grid && (
        <div className="flex flex-col gap-3">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-4">
              <label className="flex items-center gap-2 text-sm text-foreground/60">
                Layout
                <select
                  value={layout}
                  onChange={(e) => setLayout(e.target.value as Layout)}
                  className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                >
                  <option value="grouped">Grouped by field</option>
                  <option value="list">Fields as bullets</option>
                  <option value="sections">A/B/C/D sections</option>
                </select>
              </label>
              {layout === "grouped" && headers.length > 0 && (
                <label className="flex items-center gap-2 text-sm text-foreground/60">
                  Group by
                  <select
                    value={groupCol}
                    onChange={(e) => setGroupCol(Number(e.target.value))}
                    className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                  >
                    {headers.map((h, i) => (
                      <option key={i} value={i}>
                        {h || `Column ${i + 1}`}
                      </option>
                    ))}
                  </select>
                </label>
              )}
              {layout !== "sections" && headers.length > 0 && (
                <label className="flex items-center gap-2 text-sm text-foreground/60">
                  {layout === "grouped" ? "Label column" : "Title column"}
                  <select
                    value={titleCol}
                    onChange={(e) => setTitleCol(Number(e.target.value))}
                    className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                  >
                    {headers.map((h, i) => (
                      <option key={i} value={i}>
                        {h || `Column ${i + 1}`}
                      </option>
                    ))}
                  </select>
                </label>
              )}
              {layout !== "sections" && (
                <>
                  <label className="flex items-center gap-2 text-sm text-foreground/60">
                    Section #
                    <input
                      type="number"
                      min={1}
                      step={1}
                      value={sectionNumber}
                      onChange={(e) => {
                        const n = Math.floor(Number(e.target.value));
                        if (Number.isFinite(n) && n > 0) setSectionNumber(n);
                      }}
                      className="w-16 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                    />
                  </label>
                  <label className="flex items-center gap-2 text-sm text-foreground/60">
                    Section title
                    <input
                      type="text"
                      value={sectionTitle}
                      onChange={(e) => setSectionTitle(e.target.value)}
                      placeholder="e.g. Fruit Database"
                      className="w-48 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                    />
                  </label>
                </>
              )}
              <button
                type="button"
                onClick={() =>
                  setView((v) => (v === "rendered" ? "json" : "rendered"))
                }
                className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5"
              >
                {view === "rendered" ? "View JSON" : "View rendered"}
              </button>
              <button
                type="button"
                onClick={copyForWord}
                disabled={html === ""}
                className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {copyState === "copied"
                  ? "Copied!"
                  : copyState === "error"
                    ? "Copy failed"
                    : "Copy for Word"}
              </button>
            </div>
            <button
              type="button"
              onClick={clear}
              className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5"
            >
              Clear
            </button>
          </div>

          {(layout === "list" || layout === "grouped") &&
            headers.length > 0 && (
              <div className="flex flex-wrap items-center gap-x-4 gap-y-2 text-sm text-foreground/70">
                <span className="text-foreground/60">Show fields:</span>
                {headers.map((h, i) =>
                  i === titleCol ||
                  (layout === "grouped" && i === groupCol) ? null : (
                    <label key={i} className="flex items-center gap-1.5">
                      <input
                        type="checkbox"
                        checked={selectedCols.has(i)}
                        onChange={() => toggleCol(i)}
                      />
                      {h || `Column ${i + 1}`}
                    </label>
                  ),
                )}
              </div>
            )}

          {view === "json" ? (
            <JsonPreview grid={grid} />
          ) : (
            <RenderedPreview
              html={html}
              emptyHint={
                layout === "grouped"
                  ? "Pasted, but no groups were produced. The first row is read as headers; you need at least one data row. Pick a 'Group by' column and a label column above, or switch to JSON to inspect the raw grid."
                  : layout === "list"
                    ? "Pasted, but no rows produced a section. The first row is read as headers; each later row needs a non-blank value in the chosen title column. Tick fields above to show them. Switch to JSON to inspect the raw grid."
                    : "Pasted, but no sections were found. Column A should hold section titles; rows with a blank column A become subsections of the section above. Switch to JSON to inspect the raw grid."
              }
            />
          )}
        </div>
      )}
    </div>
  );
}
