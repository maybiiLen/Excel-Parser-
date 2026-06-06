import type { Grid } from "@/lib/types";

/** Read-only display of the parsed clipboard Grid as pretty-printed JSON. */
export function JsonPreview({ grid }: { grid: Grid }) {
  const rows = grid.length;
  const cols = grid.reduce((max, row) => Math.max(max, row.length), 0);

  return (
    <div className="flex flex-col gap-2">
      <div className="text-sm text-foreground/60">
        Parsed {rows} row{rows === 1 ? "" : "s"} &times; {cols} column
        {cols === 1 ? "" : "s"}
      </div>
      <pre className="max-h-[28rem] overflow-auto rounded-lg border border-foreground/15 bg-foreground/[0.03] p-4 font-mono text-xs leading-relaxed">
        {JSON.stringify(grid, null, 2)}
      </pre>
    </div>
  );
}
