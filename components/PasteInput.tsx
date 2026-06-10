"use client";

import { useMemo, useRef, useState, type ClipboardEvent } from "react";
import { parseClipboard } from "@/lib/parser";
import { buildWordHtml, htmlToPlainText } from "@/lib/clipboard";
import type { HeadingStyle } from "@/lib/clipboard";
import {
  pickDefaultTitleCol,
  pickDefaultGroupCol,
  sectionNumberOf,
  tableToHtml,
  type TableState,
} from "./tableModel";
import { TableCard } from "./TableCard";

const HEADING_FONTS = [
  "Calibri Light",
  "Calibri",
  "Arial",
  "Times New Roman",
  "Georgia",
  "Cambria",
];

const MAX_TABLES = 100;

export function PasteInput() {
  // Each paste appends one table; tables stack down the page in paste order.
  const [tables, setTables] = useState<TableState[]>([]);
  const [error, setError] = useState<string | null>(null);
  // Which table's tab is open; only the active table is rendered (no scrolling
  // through the rest). New pastes become active; removing the active tab falls
  // to a neighbor.
  const [activeId, setActiveId] = useState<string | null>(null);
  const [copyAllState, setCopyAllState] = useState<"idle" | "copied" | "error">(
    "idle",
  );
  // Monotonic id source: never reused, so React keys stay stable when a middle
  // table is removed (a ref, so bumping it never triggers a render).
  const idRef = useRef(0);

  // Heading appearance, shared across every table so a "keep source" paste
  // matches the target document. Sticky (template metadata), even across "Clear
  // all". Sizes held as strings so they can be cleared.
  const [headingColor, setHeadingColor] = useState<string>("#2F5496");
  const [headingFont, setHeadingFont] = useState<string>("Calibri Light");
  const [h1SizeInput, setH1SizeInput] = useState<string>("16");
  const [h2SizeInput, setH2SizeInput] = useState<string>("13");
  const [headingBold, setHeadingBold] = useState<boolean>(false);
  // Body font is separate (default Calibri) so unstyled body text doesn't fall
  // back to Times New Roman on a Word paste; headings keep their own font.
  const [bodyFont, setBodyFont] = useState<string>("Calibri");
  const headingStyle = useMemo<HeadingStyle>(() => {
    const clampPt = (s: string, fallback: number) => {
      const n = parseInt(s, 10);
      return Number.isFinite(n) && n >= 1 && n <= 72 ? n : fallback;
    };
    return {
      color: headingColor,
      font: headingFont,
      h1Size: clampPt(h1SizeInput, 16),
      h2Size: clampPt(h2SizeInput, 13),
      bold: headingBold,
    };
  }, [headingColor, headingFont, h1SizeInput, h2SizeInput, headingBold]);

  function handlePaste(e: ClipboardEvent<HTMLDivElement>) {
    e.preventDefault();
    if (tables.length >= MAX_TABLES) {
      setError(
        `Reached the ${MAX_TABLES}-table limit. Remove a table to add another.`,
      );
      return;
    }
    try {
      const rows = parseClipboard(e.clipboardData);
      if (rows.length === 0) {
        setError(
          "Couldn't read a table from the clipboard. Copy a cell range from Excel or Google Sheets, then paste here.",
        );
        return;
      }
      const t = pickDefaultTitleCol(rows);
      const id = `t${++idRef.current}`;
      setTables((prev) => {
        const last = prev[prev.length - 1];
        const nextNumber = last ? String(sectionNumberOf(last) + 1) : "1";
        const next: TableState = {
          id,
          grid: rows,
          layout: "grouped",
          titleCol: t,
          groupCol: pickDefaultGroupCol(rows, t),
          selectedCols: new Set((rows[0] ?? []).map((_, i) => i)), // all fields on
          pivotOrder: [],
          sectionNumberInput: nextNumber,
          sectionTitle: "",
        };
        return [...prev, next];
      });
      setActiveId(id); // open the just-pasted table
      setError(null);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Failed to parse the pasted data.",
      );
    }
  }

  function patchTable(id: string, patch: Partial<TableState>) {
    setTables((ts) => ts.map((t) => (t.id === id ? { ...t, ...patch } : t)));
  }

  function removeTable(id: string) {
    // If the open tab is removed, fall to the next tab (or the previous one if
    // it was the last). Computed from the pre-removal order.
    setActiveId((curr) => {
      if (curr !== id) return curr;
      const idx = tables.findIndex((t) => t.id === id);
      const remaining = tables.filter((t) => t.id !== id);
      if (remaining.length === 0) return null;
      return remaining[Math.min(idx, remaining.length - 1)].id;
    });
    setTables((ts) => ts.filter((t) => t.id !== id));
  }

  function clearAll() {
    setTables([]);
    setActiveId(null);
    setError(null);
  }

  // Combined export: concatenate every table's fragment and wrap ONCE. Valid
  // because buildWordHtml's heading rewrites are global and it emits a single
  // @page rule, so the whole stack becomes one Word doc (sections in paste
  // order). Built synchronously before any await to keep the user gesture.
  async function copyAll() {
    if (typeof ClipboardItem === "undefined" || !navigator.clipboard?.write) {
      setCopyAllState("error");
      setTimeout(() => setCopyAllState("idle"), 2000);
      return;
    }
    const combined = tables.map(tableToHtml).join("\n");
    try {
      const item = new ClipboardItem({
        "text/html": new Blob([buildWordHtml(combined, headingStyle, bodyFont)], {
          type: "text/html",
        }),
        "text/plain": new Blob([htmlToPlainText(combined)], {
          type: "text/plain",
        }),
      });
      await navigator.clipboard.write([item]);
      setCopyAllState("copied");
    } catch {
      setCopyAllState("error");
    }
    setTimeout(() => setCopyAllState("idle"), 2000);
  }

  function downloadAll() {
    const combined = tables.map(tableToHtml).join("\n");
    const blob = new Blob([buildWordHtml(combined, headingStyle, bodyFont)], {
      type: "application/msword",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "word-sections.doc";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  const atLimit = tables.length >= MAX_TABLES;
  // The open tab; fall back to the first table if the id ever goes stale.
  const activeTable = tables.find((t) => t.id === activeId) ?? tables[0] ?? null;
  const activeIndex = activeTable
    ? tables.findIndex((t) => t.id === activeTable.id)
    : -1;

  return (
    <div className="flex w-full flex-col gap-4">
      <div
        role="textbox"
        tabIndex={0}
        aria-label="Paste Excel data here"
        onPaste={handlePaste}
        className="flex min-h-[8rem] cursor-text items-center justify-center rounded-xl border-2 border-dashed border-foreground/25 bg-foreground/[0.02] p-6 text-center text-foreground/60 outline-none transition-colors focus:border-foreground/50 focus:bg-foreground/[0.04]"
      >
        {atLimit ? (
          <span>
            Reached the {MAX_TABLES}-table limit. Remove a table to add another.
          </span>
        ) : tables.length === 0 ? (
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
        ) : (
          <span>
            Paste another table to add it below (click here, then{" "}
            <kbd className="rounded border border-foreground/30 px-1.5 py-0.5 font-mono text-xs">
              Ctrl
            </kbd>{" "}
            +{" "}
            <kbd className="rounded border border-foreground/30 px-1.5 py-0.5 font-mono text-xs">
              V
            </kbd>
            ).
          </span>
        )}
      </div>

      {error && (
        <p className="rounded-lg border border-amber-500/40 bg-amber-500/10 px-4 py-3 text-sm text-amber-700 dark:text-amber-300">
          {error}
        </p>
      )}

      {tables.length > 0 && (
        <div className="flex flex-col gap-3">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={copyAll}
                title="Copies every table as one Word doc (sections stacked in paste order)."
                className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5"
              >
                {copyAllState === "copied"
                  ? "Copied all!"
                  : copyAllState === "error"
                    ? "Copy failed"
                    : "Copy all"}
              </button>
              <button
                type="button"
                onClick={downloadAll}
                title="Downloads one .doc containing every table, stacked in paste order."
                className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5"
              >
                Download all
              </button>
            </div>
            <button
              type="button"
              onClick={clearAll}
              className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5"
            >
              Clear all
            </button>
          </div>

          <p className="text-xs text-foreground/50">
            Getting into Word: paste with <strong>Use Destination Styles</strong>{" "}
            so headings match your document (set it as the default under Word →
            Options → Advanced → &ldquo;Pasting from other programs&rdquo;), or{" "}
            <strong>Download for Word</strong>, open the file, and copy it in
            (Word&nbsp;→&nbsp;Word keeps formatting).
          </p>

          <div className="flex flex-wrap items-center gap-x-4 gap-y-2 rounded-lg border border-foreground/15 bg-foreground/[0.02] p-3 text-sm text-foreground/70">
            <span className="text-foreground/60">
              Heading style{" "}
              <span className="text-foreground/40">(applies to all tables)</span>:
            </span>
            <label className="flex items-center gap-1.5">
              Color
              <input
                type="color"
                value={headingColor}
                onChange={(e) => setHeadingColor(e.target.value)}
                aria-label="Heading color"
                className="h-6 w-8 cursor-pointer rounded border border-foreground/20"
              />
            </label>
            <label className="flex items-center gap-1.5">
              Heading font
              <select
                value={headingFont}
                onChange={(e) => setHeadingFont(e.target.value)}
                className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
              >
                {HEADING_FONTS.map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
            </label>
            <label className="flex items-center gap-1.5">
              Body font
              <select
                value={bodyFont}
                onChange={(e) => setBodyFont(e.target.value)}
                className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
              >
                {HEADING_FONTS.map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
            </label>
            <label className="flex items-center gap-1.5">
              H1 pt
              <input
                type="text"
                inputMode="numeric"
                value={h1SizeInput}
                onChange={(e) =>
                  setH1SizeInput(e.target.value.replace(/[^0-9]/g, ""))
                }
                aria-label="Heading 1 size in points"
                className="w-12 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
              />
            </label>
            <label className="flex items-center gap-1.5">
              H2 pt
              <input
                type="text"
                inputMode="numeric"
                value={h2SizeInput}
                onChange={(e) =>
                  setH2SizeInput(e.target.value.replace(/[^0-9]/g, ""))
                }
                aria-label="Heading 2 size in points"
                className="w-12 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
              />
            </label>
            <label className="flex items-center gap-1.5">
              <input
                type="checkbox"
                checked={headingBold}
                onChange={(e) => setHeadingBold(e.target.checked)}
              />
              Bold
            </label>
          </div>

          <div className="flex items-center gap-1.5 overflow-x-auto pb-1">
            {tables.map((t, i) => {
              const isActive = activeTable?.id === t.id;
              const label = t.sectionTitle.trim() || `Table ${i + 1}`;
              return (
                <div
                  key={t.id}
                  className={`flex shrink-0 items-center gap-1.5 rounded-md border px-3 py-1.5 text-xs transition-colors ${
                    isActive
                      ? "border-foreground/30 bg-foreground/[0.06] font-medium text-foreground"
                      : "border-foreground/15 text-foreground/60 hover:bg-foreground/5"
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => setActiveId(t.id)}
                    className="max-w-[12rem] truncate"
                    title={label}
                  >
                    {label}
                  </button>
                  <button
                    type="button"
                    onClick={() => removeTable(t.id)}
                    aria-label={`Remove ${label}`}
                    title="Remove this table"
                    className="leading-none text-foreground/40 transition-colors hover:text-foreground"
                  >
                    &times;
                  </button>
                </div>
              );
            })}
          </div>

          {activeTable && (
            <TableCard
              key={activeTable.id}
              table={activeTable}
              index={activeIndex}
              headingStyle={headingStyle}
              bodyFont={bodyFont}
              onChange={(patch) => patchTable(activeTable.id, patch)}
            />
          )}
        </div>
      )}
    </div>
  );
}
