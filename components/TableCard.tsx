"use client";

import { memo, useMemo, useState } from "react";
import { buildWordHtml, htmlToPlainText } from "@/lib/clipboard";
import type { HeadingStyle } from "@/lib/clipboard";
import {
  tableToHtml,
  type Layout,
  type TableState,
} from "./tableModel";
import { JsonPreview } from "./JsonPreview";
import { RenderedPreview } from "./RenderedPreview";

type Props = {
  table: TableState;
  index: number;
  headingStyle: HeadingStyle;
  bodyFont: string;
  onChange: (patch: Partial<TableState>) => void;
};

/** Return a copy of `arr` with the items at i and j swapped. */
function swap(arr: number[], i: number, j: number): number[] {
  const next = [...arr];
  [next[i], next[j]] = [next[j], next[i]];
  return next;
}

function TableCardInner({
  table,
  index,
  headingStyle,
  bodyFont,
  onChange,
}: Props) {
  const [view, setView] = useState<"rendered" | "json">("rendered");
  const [copyState, setCopyState] = useState<"idle" | "copied" | "error">("idle");

  const { grid, layout, titleCol, groupCol, selectedCols, pivotOrder } = table;

  // Header row, for the title-column picker and field checklist.
  const headers = useMemo(
    () => (grid[0] ? grid[0].map((c) => (c == null ? "" : String(c))) : []),
    [grid],
  );

  // The whole derive-from-state chain for this one table; recomputed only when
  // this table's record changes (the parent keeps unedited records identical).
  const html = useMemo(() => tableToHtml(table), [table]);

  function toggleCol(i: number) {
    const next = new Set(selectedCols);
    if (next.has(i)) next.delete(i);
    else next.add(i);
    onChange({ selectedCols: next });
  }

  // Write this table's sections to the clipboard as text/html (+ text/plain
  // fallback). Built synchronously from `html` before any await, preserving the
  // user-gesture requirement.
  async function copyForWord() {
    if (typeof ClipboardItem === "undefined" || !navigator.clipboard?.write) {
      setCopyState("error");
      setTimeout(() => setCopyState("idle"), 2000);
      return;
    }
    try {
      const item = new ClipboardItem({
        "text/html": new Blob([buildWordHtml(html, headingStyle, bodyFont)], {
          type: "text/html",
        }),
        "text/plain": new Blob([htmlToPlainText(html)], { type: "text/plain" }),
      });
      await navigator.clipboard.write([item]);
      setCopyState("copied");
    } catch {
      setCopyState("error");
    }
    setTimeout(() => setCopyState("idle"), 2000);
  }

  // Download a .doc (HTML that Word opens, applying the Heading styles via its
  // file-open import -- more reliable than a browser->desktop clipboard paste).
  function downloadForWord() {
    const safeName =
      table.sectionTitle.trim().replace(/[^a-z0-9 _-]/gi, "").trim() ||
      `table-${index + 1}`;
    const blob = new Blob([buildWordHtml(html, headingStyle, bodyFont)], {
      type: "application/msword",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${safeName}.doc`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="flex flex-col gap-3 rounded-xl border border-foreground/15 p-4">
      <div className="flex flex-wrap items-center gap-4">
        <label className="flex items-center gap-2 text-sm text-foreground/60">
          Layout
          <select
            value={layout}
            onChange={(e) => onChange({ layout: e.target.value as Layout })}
            className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
          >
            <option value="grouped">Grouped by field</option>
            <option value="list">Fields as bullets</option>
            <option value="pivot">Pivot (nested rows)</option>
            <option value="sections">A/B/C/D sections</option>
          </select>
        </label>
        {layout === "grouped" && headers.length > 0 && (
          <label className="flex items-center gap-2 text-sm text-foreground/60">
            Group by
            <select
              value={groupCol}
              onChange={(e) => onChange({ groupCol: Number(e.target.value) })}
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
        {(layout === "list" || layout === "grouped") && headers.length > 0 && (
          <label className="flex items-center gap-2 text-sm text-foreground/60">
            {layout === "grouped" ? "Label column" : "Title column"}
            <select
              value={titleCol}
              onChange={(e) => onChange({ titleCol: Number(e.target.value) })}
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
        {(layout === "list" || layout === "grouped") && (
          <>
            <label className="flex items-center gap-2 text-sm text-foreground/60">
              Section #
              <input
                type="text"
                inputMode="numeric"
                pattern="[0-9]*"
                value={table.sectionNumberInput}
                onChange={(e) =>
                  onChange({
                    sectionNumberInput: e.target.value.replace(/[^0-9]/g, ""),
                  })
                }
                aria-label="Section number"
                className="w-16 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
              />
            </label>
            <label className="flex items-center gap-2 text-sm text-foreground/60">
              Section title
              <input
                type="text"
                value={table.sectionTitle}
                onChange={(e) => onChange({ sectionTitle: e.target.value })}
                placeholder="e.g. Fruit Database"
                className="w-48 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
              />
            </label>
          </>
        )}
        <button
          type="button"
          onClick={() => setView((v) => (v === "rendered" ? "json" : "rendered"))}
          className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5"
        >
          {view === "rendered" ? "View JSON" : "View rendered"}
        </button>
        <button
          type="button"
          onClick={copyForWord}
          disabled={html === ""}
          title="Copies this table's HTML to paste into Word. For headings to match your document, paste with 'Use Destination Styles' (Word → Options → Advanced → Pasting from other programs), or use Download for Word."
          className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5 disabled:cursor-not-allowed disabled:opacity-50"
        >
          {copyState === "copied"
            ? "Copied!"
            : copyState === "error"
              ? "Copy failed"
              : "Copy for Word"}
        </button>
        <button
          type="button"
          onClick={downloadForWord}
          disabled={html === ""}
          title="Downloads a .doc of this table you can open in Word with the right Heading styles, then copy into your document (Word → Word keeps formatting)."
          className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5 disabled:cursor-not-allowed disabled:opacity-50"
        >
          Download for Word
        </button>
      </div>

      {(layout === "list" || layout === "grouped") && headers.length > 0 && (
        <div className="flex flex-wrap items-center gap-x-4 gap-y-2 text-sm text-foreground/70">
          <span className="text-foreground/60">Show fields:</span>
          {headers.map((h, i) =>
            i === titleCol || (layout === "grouped" && i === groupCol) ? null : (
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

      {layout === "pivot" && headers.length > 0 && (
        <div className="flex flex-col gap-2 text-sm text-foreground/70">
          <span className="text-foreground/60">
            Nest rows by{" "}
            <span className="text-foreground/40">(click fields in order)</span>:
          </span>
          <div className="flex flex-wrap items-center gap-1.5">
            {headers.map((h, i) => {
              const pos = pivotOrder.indexOf(i);
              const selected = pos >= 0;
              const name = h || `Column ${i + 1}`;
              return (
                <button
                  key={i}
                  type="button"
                  aria-pressed={selected}
                  onClick={() =>
                    onChange({
                      pivotOrder: selected
                        ? pivotOrder.filter((c) => c !== i)
                        : [...pivotOrder, i],
                    })
                  }
                  className={`flex items-center gap-1.5 rounded-md border px-2 py-1 text-xs transition-colors ${
                    selected
                      ? "border-foreground/30 bg-foreground/[0.06] font-medium text-foreground"
                      : "border-foreground/15 text-foreground/60 hover:bg-foreground/5"
                  }`}
                >
                  {selected && (
                    <span className="flex h-4 w-4 items-center justify-center rounded-full bg-foreground/20 text-[10px] font-semibold tabular-nums">
                      {pos + 1}
                    </span>
                  )}
                  {name}
                </button>
              );
            })}
          </div>
          {pivotOrder.length > 0 && (
            <div className="flex flex-wrap items-center gap-1.5 text-xs text-foreground/60">
              <span className="text-foreground/50">Order:</span>
              {pivotOrder.map((col, idx) => {
                const name = headers[col] || `Column ${col + 1}`;
                return (
                  <span key={col} className="flex items-center gap-1">
                    {idx > 0 && <span className="text-foreground/30">→</span>}
                    <span className="rounded bg-foreground/5 px-1.5 py-0.5">
                      {idx + 1}. {name}
                    </span>
                    <span className="flex flex-col leading-none">
                      <button
                        type="button"
                        aria-label={`Move ${name} earlier`}
                        disabled={idx === 0}
                        onClick={() =>
                          onChange({ pivotOrder: swap(pivotOrder, idx, idx - 1) })
                        }
                        className="px-0.5 text-foreground/40 transition-colors hover:text-foreground disabled:opacity-25"
                      >
                        ▲
                      </button>
                      <button
                        type="button"
                        aria-label={`Move ${name} later`}
                        disabled={idx === pivotOrder.length - 1}
                        onClick={() =>
                          onChange({ pivotOrder: swap(pivotOrder, idx, idx + 1) })
                        }
                        className="px-0.5 text-foreground/40 transition-colors hover:text-foreground disabled:opacity-25"
                      >
                        ▼
                      </button>
                    </span>
                  </span>
                );
              })}
            </div>
          )}
        </div>
      )}

      {view === "json" ? (
        <JsonPreview grid={grid} />
      ) : (
        <RenderedPreview
          html={html}
          headingStyle={headingStyle}
          bodyFont={bodyFont}
          emptyHint={
            layout === "grouped"
              ? "Pasted, but no groups were produced. The first row is read as headers; you need at least one data row. Pick a 'Group by' column and a label column above, or switch to JSON to inspect the raw grid."
              : layout === "list"
                ? "Pasted, but no rows produced a section. The first row is read as headers; each later row needs a non-blank value in the chosen title column. Tick fields above to show them. Switch to JSON to inspect the raw grid."
                : layout === "pivot"
                  ? "Pick one or more fields above, in order, to nest rows by (like an Excel pivot's Row Labels)."
                  : "Pasted, but no sections were found. Column A should hold section titles; rows with a blank column A become subsections of the section above. Switch to JSON to inspect the raw grid."
          }
        />
      )}
    </div>
  );
}

export const TableCard = memo(TableCardInner);
