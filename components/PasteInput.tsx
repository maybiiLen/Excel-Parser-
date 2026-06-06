"use client";

import { useState, type ClipboardEvent } from "react";
import { parseClipboard } from "@/lib/parser";
import type { Grid } from "@/lib/types";
import { JsonPreview } from "./JsonPreview";

export function PasteInput() {
  const [grid, setGrid] = useState<Grid | null>(null);
  const [error, setError] = useState<string | null>(null);

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
    } catch (err) {
      setGrid(null);
      setError(
        err instanceof Error ? err.message : "Failed to parse the pasted data.",
      );
    }
  }

  function clear() {
    setGrid(null);
    setError(null);
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
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-semibold">Parsed JSON</h2>
            <button
              type="button"
              onClick={clear}
              className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5"
            >
              Clear
            </button>
          </div>
          <JsonPreview grid={grid} />
        </div>
      )}
    </div>
  );
}
