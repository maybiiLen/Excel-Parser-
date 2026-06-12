"use client";

import { memo, useMemo, useState } from "react";
import { buildWordHtml, htmlToPlainText } from "@/lib/clipboard";
import type { HeadingStyle } from "@/lib/clipboard";
import { defaultMarker, type MarkerKind } from "@/lib/renderers";
import { DEFAULT_FIELD_LABEL, type FieldLabel } from "@/lib/types";
import {
  addField,
  canIndent,
  canOutdent,
  indentField,
  moveField,
  outdentField,
  removeField,
  tableToHtml,
  unusedColumns,
  type TableState,
} from "./tableModel";
import { JsonPreview } from "./JsonPreview";
import { RenderedPreview } from "./RenderedPreview";

/** Marker styles offered per indent level, with a sample label. */
const MARKER_OPTIONS: { kind: MarkerKind; label: string }[] = [
  { kind: "decimal", label: "1." },
  { kind: "paren", label: "1)" },
  { kind: "upperAlpha", label: "A." },
  { kind: "lowerAlpha", label: "a." },
  { kind: "upperRoman", label: "I." },
  { kind: "lowerRoman", label: "i." },
  { kind: "bullet", label: "• bullet" },
  { kind: "dash", label: "– dash" },
  { kind: "none", label: "None" },
];

type Props = {
  table: TableState;
  headingStyle: HeadingStyle;
  bodyFont: string;
  numberDepth: number;
  onChange: (patch: Partial<TableState>) => void;
};

function TableCardInner({
  table,
  headingStyle,
  bodyFont,
  numberDepth,
  onChange,
}: Props) {
  const [view, setView] = useState<"rendered" | "json">("rendered");
  const [copyState, setCopyState] = useState<"idle" | "copied" | "error">("idle");

  const { grid, pivotLevels } = table;

  // Header row, for the field names.
  const headers = useMemo(
    () => (grid[0] ? grid[0].map((c) => (c == null ? "" : String(c))) : []),
    [grid],
  );

  // Fields not yet placed in any indent bucket (the "Add fields" pool).
  const unused = useMemo(
    () => unusedColumns(headers.length, pivotLevels),
    [headers.length, pivotLevels],
  );

  // The placed fields flattened, each tagged with its bucket b, within-bucket k,
  // and flat index fi (the argument the structure helpers expect).
  const placed = useMemo(() => {
    const out: { col: number; b: number; fi: number }[] = [];
    let fi = 0;
    pivotLevels.forEach((bucket, b) =>
      bucket.forEach((col) => {
        out.push({ col, b, fi });
        fi++;
      }),
    );
    return out;
  }, [pivotLevels]);

  // The whole derive-from-state chain for this one table; recomputed only when
  // this table's record (or the shared numberDepth) changes.
  const html = useMemo(() => tableToHtml(table, numberDepth), [table, numberDepth]);

  // Toggle part of one field's label look (keyed by grid column).
  function patchLabel(col: number, patch: Partial<FieldLabel>) {
    const cur = table.fieldLabels[col] ?? DEFAULT_FIELD_LABEL;
    onChange({
      fieldLabels: { ...table.fieldLabels, [col]: { ...cur, ...patch } },
    });
  }

  // Write this table's pivot to the clipboard as text/html (+ text/plain
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

  const btn =
    "px-1 text-foreground/40 transition-colors hover:text-foreground disabled:opacity-20 disabled:hover:text-foreground/40";
  // Per-field label toggle (Aa / B / U): highlighted when active.
  const tgl = (active: boolean) =>
    `flex h-5 w-5 items-center justify-center rounded border text-[10px] transition-colors disabled:opacity-25 ${
      active
        ? "border-foreground/40 bg-foreground/10 text-foreground"
        : "border-foreground/15 text-foreground/50 hover:bg-foreground/5"
    }`;

  return (
    <div className="flex flex-col gap-3 rounded-xl border border-foreground/15 p-4">
      <div className="flex flex-wrap items-center gap-4">
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
          title="Copies this table's HTML to paste into Word. For the title to match your document's Heading 1, paste with 'Use Destination Styles' (Word → Options → Advanced → Pasting from other programs)."
          className="rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5 disabled:cursor-not-allowed disabled:opacity-50"
        >
          {copyState === "copied"
            ? "Copied!"
            : copyState === "error"
              ? "Copy failed"
              : "Copy for Word"}
        </button>
      </div>

      {headers.length > 0 && (
        <div className="flex flex-col gap-3 text-sm text-foreground/70">
          {/* Add fields: the unused pool. Clicking appends a new deepest level. */}
          <div className="flex flex-col gap-1.5">
            <span className="text-foreground/60">
              Add fields{" "}
              <span className="text-foreground/40">
                (click to add to the outline)
              </span>
              :
            </span>
            <div className="flex flex-wrap items-center gap-1.5">
              {unused.length === 0 ? (
                <span className="text-xs text-foreground/40">
                  All fields added.
                </span>
              ) : (
                unused.map((col) => (
                  <button
                    key={col}
                    type="button"
                    onClick={() =>
                      onChange({ pivotLevels: addField(pivotLevels, col) })
                    }
                    className="rounded-md border border-foreground/15 px-2 py-1 text-xs text-foreground/60 transition-colors hover:bg-foreground/5"
                  >
                    + {headers[col] || `Column ${col + 1}`}
                  </button>
                ))
              )}
            </div>
          </div>

          {/* Structure: ordered indent buckets. ◄ ► change level (stack = same
              indent), ▲ ▼ reorder, ✕ remove. */}
          {placed.length > 0 && (
            <div className="flex flex-col gap-1.5">
              <span className="text-foreground/60">
                Structure{" "}
                <span className="text-foreground/40">
                  (◄ ► set indent — stack fields at one level to show them
                  together)
                </span>
                :
              </span>
              <div className="flex flex-col gap-1">
                {placed.map(({ col, b, fi }) => {
                  const name = headers[col] || `Column ${col + 1}`;
                  const lf = table.fieldLabels[col] ?? DEFAULT_FIELD_LABEL;
                  return (
                    <div
                      key={col}
                      className="flex items-center gap-1.5"
                      style={{ marginLeft: `${b * 1.25}rem` }}
                    >
                      <span className="flex h-4 w-4 items-center justify-center rounded bg-foreground/10 text-[10px] font-semibold tabular-nums text-foreground/60">
                        {b + 1}
                      </span>
                      <span className="rounded bg-foreground/5 px-2 py-0.5 text-xs text-foreground">
                        {name}
                      </span>
                      {/* Label look: show/hide "Field:", bold it, underline it. */}
                      <span className="flex items-center gap-0.5">
                        <button
                          type="button"
                          aria-pressed={lf.show}
                          aria-label={`Show the "${name}:" label`}
                          title={`Show/hide the "${name}:" label`}
                          onClick={() => patchLabel(col, { show: !lf.show })}
                          className={tgl(lf.show)}
                        >
                          Aa
                        </button>
                        <button
                          type="button"
                          aria-pressed={lf.bold}
                          disabled={!lf.show}
                          aria-label={`Bold the "${name}:" label`}
                          title="Bold the label"
                          onClick={() => patchLabel(col, { bold: !lf.bold })}
                          className={`${tgl(lf.bold)} font-bold`}
                        >
                          B
                        </button>
                        <button
                          type="button"
                          aria-pressed={lf.underline}
                          disabled={!lf.show}
                          aria-label={`Underline the "${name}:" label`}
                          title="Underline the label"
                          onClick={() =>
                            patchLabel(col, { underline: !lf.underline })
                          }
                          className={`${tgl(lf.underline)} underline`}
                        >
                          U
                        </button>
                      </span>
                      <span className="flex items-center leading-none">
                        <button
                          type="button"
                          aria-label={`Outdent ${name}`}
                          disabled={!canOutdent(pivotLevels, fi)}
                          onClick={() =>
                            onChange({
                              pivotLevels: outdentField(pivotLevels, fi),
                            })
                          }
                          className={btn}
                        >
                          ◄
                        </button>
                        <button
                          type="button"
                          aria-label={`Indent ${name}`}
                          disabled={!canIndent(pivotLevels, fi)}
                          onClick={() =>
                            onChange({
                              pivotLevels: indentField(pivotLevels, fi),
                            })
                          }
                          className={btn}
                        >
                          ►
                        </button>
                        <button
                          type="button"
                          aria-label={`Move ${name} up`}
                          disabled={fi === 0}
                          onClick={() =>
                            onChange({
                              pivotLevels: moveField(pivotLevels, fi, -1),
                            })
                          }
                          className={btn}
                        >
                          ▲
                        </button>
                        <button
                          type="button"
                          aria-label={`Move ${name} down`}
                          disabled={fi === placed.length - 1}
                          onClick={() =>
                            onChange({
                              pivotLevels: moveField(pivotLevels, fi, 1),
                            })
                          }
                          className={btn}
                        >
                          ▼
                        </button>
                        <button
                          type="button"
                          aria-label={`Remove ${name}`}
                          onClick={() =>
                            onChange({
                              pivotLevels: removeField(pivotLevels, fi),
                            })
                          }
                          className={`${btn} ml-0.5`}
                        >
                          ✕
                        </button>
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Markers: one picker per indent level (bucket). */}
          {pivotLevels.length > 0 && (
            <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5 text-sm text-foreground/70">
              <span className="text-foreground/60">
                Markers{" "}
                <span className="text-foreground/40">(per indent level)</span>:
              </span>
              {pivotLevels.map((_, i) => (
                <label
                  key={i}
                  className="flex items-center gap-1.5 text-xs text-foreground/60"
                >
                  Lv {i + 1}
                  <select
                    value={table.markers[i] ?? defaultMarker(i + 1)}
                    onChange={(e) => {
                      const next = [...table.markers];
                      next[i] = e.target.value as MarkerKind;
                      onChange({ markers: next });
                    }}
                    aria-label={`Marker for indent level ${i + 1}`}
                    className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                  >
                    {MARKER_OPTIONS.map((o) => (
                      <option key={o.kind} value={o.kind}>
                        {o.label}
                      </option>
                    ))}
                  </select>
                </label>
              ))}
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
          numbered={numberDepth > 0}
          emptyHint="Add fields above to build the outline. Stack fields at one indent level to show them together (like an Excel pivot's Row Labels)."
        />
      )}
    </div>
  );
}

export const TableCard = memo(TableCardInner);
