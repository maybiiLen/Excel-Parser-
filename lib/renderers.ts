import type { Body } from "./types";

/**
 * Body -> HTML renderer (NOT YET IMPLEMENTED -- deferred past Day 1).
 *
 * Converts a Body to an HTML string that pastes cleanly into Word:
 *   - { type: "text" }    -> <p>...</p>
 *   - { type: "bullets" } -> <ul><li>...</li></ul>
 *   - { type: "table" }   -> <table>...</table>
 *
 * A later sprint adds the wide-table width strategy (transpose or split) so
 * tables fit an 8.5x11 Word page.
 */
export function renderBody(body: Body): string {
  // TODO: switch on body.type and emit Word-friendly HTML.
  void body;
  throw new Error("renderBody: not implemented yet");
}
