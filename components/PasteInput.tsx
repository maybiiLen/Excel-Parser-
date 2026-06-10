"use client";

import {
  useMemo,
  useState,
  type ChangeEvent,
  type ClipboardEvent,
} from "react";
import { parseClipboard } from "@/lib/parser";
import { rowsToTree, rowsToAttributeSections } from "@/lib/mapper";
import { numberTree } from "@/lib/numbering";
import { renderTree } from "@/lib/renderers";
import type { Grid } from "@/lib/types";
import { JsonPreview } from "./JsonPreview";
import { RenderedPreview } from "./RenderedPreview";

type Layout = "sections" | "list";

/** First header containing "name" (case-insensitive), else the first column. */
function pickDefaultTitleCol(rows: Grid): number {
  const header = rows[0] ?? [];
  const i = header.findIndex(
    (c) => typeof c === "string" && c.toLowerCase().includes("name"),
  );
  return i >= 0 ? i : 0;
}

export function PasteInput() {
  const [grid, setGrid] = useState<Grid | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [root, setRoot] = useState<number>(6);
  const [view, setView] = useState<"rendered" | "json">("rendered");
  const [layout, setLayout] = useState<Layout>("list");
  const [titleCol, setTitleCol] = useState<number>(0);
  const [selectedCols, setSelectedCols] = useState<Set<number>>(new Set());

  // Header row, for the title-column picker and field checklist.
  const headers = useMemo(
    () => (grid?.[0] ? grid[0].map((c) => (c == null ? "" : String(c))) : []),
    [grid],
  );

  // Columns rendered as bullets, in column order: checked, minus the title column.
  const fieldColumns = useMemo(
    () =>
      headers
        .map((_, i) => i)
        .filter((i) => i !== titleCol && selectedCols.has(i)),
    [headers, titleCol, selectedCols],
  );

  // parse -> map -> number -> render, derived from the stored grid so layout,
  // title-column, field, and root changes re-render without a re-paste. The
  // mapper re-runs only on grid/layout/titleCol/fieldColumns; numberTree/
  // renderTree on a root change.
  const tree = useMemo(() => {
    if (!grid) return null;
    return layout === "list"
      ? rowsToAttributeSections(grid, titleCol, fieldColumns)
      : rowsToTree(grid);
  }, [grid, layout, titleCol, fieldColumns]);
  const html = useMemo(
    () => (tree ? renderTree(numberTree(tree, root)) : ""),
    [tree, root],
  );

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
      setTitleCol(pickDefaultTitleCol(rows));
      setSelectedCols(new Set((rows[0] ?? []).map((_, i) => i))); // all fields on
    } catch (err) {
      setGrid(null);
      setError(
        err instanceof Error ? err.message : "Failed to parse the pasted data.",
      );
    }
  }

  // Keep `root` out of UI-validation reach of numberTree: ignore empty/NaN/
  // non-positive/fractional input so the preview never renders "NaN" numbers.
  function handleRootChange(e: ChangeEvent<HTMLInputElement>) {
    const next = Number(e.target.value);
    if (Number.isInteger(next) && next > 0) setRoot(next);
  }

  function toggleCol(i: number) {
    setSelectedCols((prev) => {
      const next = new Set(prev);
      if (next.has(i)) next.delete(i);
      else next.add(i);
      return next;
    });
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
                  <option value="list">Fields as bullets</option>
                  <option value="sections">A/B/C/D sections</option>
                </select>
              </label>
              {layout === "list" && headers.length > 0 && (
                <label className="flex items-center gap-2 text-sm text-foreground/60">
                  Title column
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
              <label className="flex items-center gap-2 text-sm text-foreground/60">
                Start numbering at
                <input
                  type="number"
                  min={1}
                  step={1}
                  value={root}
                  onChange={handleRootChange}
                  className="w-16 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                />
              </label>
              <button
                type="button"
                onClick={() =>
                  setView((v) => (v === "rendered" ? "json" : "rendered"))
                }
                className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5"
              >
                {view === "rendered" ? "View JSON" : "View rendered"}
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

          {layout === "list" && headers.length > 0 && (
            <div className="flex flex-wrap items-center gap-x-4 gap-y-2 text-sm text-foreground/70">
              <span className="text-foreground/60">Show fields:</span>
              {headers.map((h, i) =>
                i === titleCol ? null : (
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
                layout === "list"
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
