export const meta = {
  name: 'review-excel-word-app',
  description: 'Multi-agent review of the Excel-to-Word pivot app: fan out by dimension, adversarially verify each finding',
  phases: [
    { title: 'Review' },
    { title: 'Verify' },
  ],
}

const SHARED = [
  'You are reviewing a finalize-stage client-side Next.js 16 app at the repo root (cwd is the project root). It turns a pasted Excel or Google-Sheets table into an Excel-pivot-style nested outline, renders Word-ready HTML, and copies it to the clipboard for pasting into Microsoft Word.',
  '',
  'PIPELINE (per table): paste, then parseClipboard (lib/parser.ts, SheetJS), then rowsToPivotTree (lib/mapper.ts), then renderPivotTree (lib/renderers.ts), then buildWordHtml (lib/clipboard.ts), then clipboard write. UI: components/PasteInput.tsx (parent: tables array, shared per-level heading style, body font, numberDepth, copyAll), components/TableCard.tsx (one table editor), components/tableModel.ts (TableState plus pure bucket helpers plus tableToHtml), components/RenderedPreview.tsx (dangerouslySetInnerHTML preview), components/JsonPreview.tsx. Types in lib/types.ts.',
  '',
  'KEY MODEL: pivotLevels is an array of arrays of grid-column indices, i.e. ordered indent buckets (bucket index equals depth, each bucket holds one or more grid columns). Rows merge by the composite key of a bucket values. The Title optionally maps to a Word Heading style via mso-style-name on a Use Destination Styles paste; the body uses inline direct per-level formatting so it survives that paste.',
  '',
  'IN-FLIGHT UNCOMMITTED WORK (review this closely, it is the newest and least-tested code; run git --no-pager diff HEAD to see it):',
  '1. Word-heading NUMBERING: a shared numberDepth picks the top N body indent levels; renderPivotTree tags each such bucket first line with a data-heading attribute equal to K, where K is one-if-there-is-a-title-else-zero plus depth, clamped to 9; buildWordHtml maps that data-heading K to a class MsoHeadingK carrying mso-style-name of Heading K and mso-outline-level K, so Word supplies live numbers like 5.1 and 5.1.1. The app text marker is suppressed on numbered lines.',
  '2. Per-field LABEL styling: fieldLabels is a record keyed by grid column whose value has show, bold, underline flags; renderPivotTree wrapLabel renders the field name plus colon-space, escaped, optionally wrapped in bold or underline inline runs; PivotNode.lines changed from an array of strings to an array of PivotLine objects each carrying col, name, value.',
  '',
  'HARD PROJECT RULES (from CLAUDE.md): client-side only, no backend or API; SheetJS from the official CDN tarball (never the npm xlsx package); files at repo root (no src directory); docs MUST stay in sync with code (CLAUDE.md, README.md, and docs OVERVIEW, ARCHITECTURE, ROADMAP) in the SAME change when pivot behavior, the data model, the pipeline, or a user-facing control changes. Stale docs are a bug.',
  '',
  'YOUR JOB: Read the ACTUAL code (use Read, Grep, Glob; run Bash if useful) and report concrete, real findings for YOUR dimension only. Every finding must cite a real file and location (file and line, or function name) and include the specific evidence (quote the code) and a concrete suggested fix. Prefer a few high-confidence, genuinely actionable findings over a long speculative list. Do NOT invent issues; if the code is correct, say so with fewer findings. Severity: critical means data loss or broken output or security; high means wrong output in common cases or a real bug; medium means an edge-case bug or notable gap; low means minor; nit means style. Be specific to THIS codebase, not generic advice.',
].join('\n')

const FINDINGS_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    dimension: { type: 'string' },
    summary: { type: 'string', description: 'one-sentence overall assessment of this dimension' },
    findings: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        properties: {
          title: { type: 'string' },
          severity: { type: 'string', enum: ['critical', 'high', 'medium', 'low', 'nit'] },
          category: { type: 'string', enum: ['bug', 'correctness', 'security', 'performance', 'ux', 'a11y', 'docs', 'cleanup'] },
          file: { type: 'string' },
          location: { type: 'string', description: 'file and line, or function name' },
          description: { type: 'string' },
          evidence: { type: 'string', description: 'the specific code or behavior that proves it' },
          suggestedFix: { type: 'string' },
          confidence: { type: 'string', enum: ['high', 'medium', 'low'] },
        },
        required: ['title', 'severity', 'category', 'file', 'location', 'description', 'evidence', 'suggestedFix', 'confidence'],
      },
    },
  },
  required: ['dimension', 'summary', 'findings'],
}

const VERDICT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    isReal: { type: 'boolean', description: 'true only if you independently confirmed the issue is real by reading the actual code' },
    correctedSeverity: { type: 'string', enum: ['critical', 'high', 'medium', 'low', 'nit', 'invalid'] },
    reasoning: { type: 'string', description: 'what you checked in the real code and why you confirm or refute' },
    fixCritique: { type: 'string', description: 'is the suggested fix correct and complete, or is there a better one' },
  },
  required: ['isReal', 'correctedSeverity', 'reasoning', 'fixCritique'],
}

const DIMENSIONS = [
  {
    key: 'core-pipeline',
    focus: [
      'DIMENSION: Core pipeline correctness (lib/parser.ts, lib/mapper.ts, lib/renderers.ts, components/tableModel.ts).',
      'Hunt real bugs and edge cases: blank cells, blank headers, and the (blank) placeholder handling; the JSON.stringify composite merge key (could distinct value tuples ever collide, or non-distinct ones fail to merge); marker counting per parent (the index passed to markerText); depth clamping at 9 (what happens to the data-level attribute versus the real nesting beyond 9); the bucket helpers addField, removeField, indentField, outdentField, moveField, locate (off-by-one, the non-empty-bucket invariant, moveField swap correctness when the two fields live in different buckets); the tableToHtml empty-guard; and the SheetJS sheet_to_json options. Verify the merge, order, and dedup logic actually preserves first-seen order and merges correctly.',
    ].join('\n'),
  },
  {
    key: 'inflight-diff',
    focus: [
      'DIMENSION: The in-flight uncommitted feature work ONLY (numbering plus field labels). Run git --no-pager diff HEAD to see exactly what changed, then scrutinize:',
      'the numberDepth to data-heading tagging and the buildWordHtml regex that now has an OPTIONAL data-heading group. Can the two replacement paths ever double-match or mis-match? The COUPLING between the user-editable headingStyleName (the title style, default Heading 1) and the hardcoded Heading K body mapping: what breaks if headingStyleName is set to Heading 2 or Title or left blank while numberDepth is greater than 0? Is the title style and the numbered body assumed to be Heading 1 then Heading 2 and so on, and is that assumption ever violated? Marker suppression on numbered lines. The field-label path: the new PivotLine shape, the wrapLabel bold and underline run nesting and order, the label shown only when show is true and the name is non-empty. Does fieldLabels keyed by grid column survive a remove then re-add of the same field? Edge cases: numberDepth greater than 0 with no title; numberDepth greater than 0 but headingStyleName blank.',
    ].join('\n'),
  },
  {
    key: 'word-fidelity',
    focus: [
      'DIMENSION: Word and clipboard output fidelity (lib/clipboard.ts buildWordHtml and htmlToPlainText, components/PasteInput.tsx copyAll, components/TableCard.tsx copyForWord).',
      'Does the emitted HTML actually produce the claimed result in Word? Check: the ClipboardItem path that writes text/html with no manual CF_HTML header (is relying on the browser to add it correct on Windows Word); inline direct formatting surviving a Use Destination Styles paste; the mso-style-name and mso-outline-level correctness for the title and for MsoHeadingK; whether Heading K with K up to 9 maps to real Word styles (Word ships Heading 1 through 9; does the numbered K start at 2, and could it ever exceed 9 and map to a non-existent style); the combined copyAll that concatenates fragments then wraps ONCE (single page rule, global regex) and whether that is correct when one fragment has a title and another does not; htmlToPlainText unescape order (amp last) and tag stripping; line-height 115 percent in the output versus the preview leading-tight; bold via a bold run versus a class. Identify fidelity mismatches between the live preview and the actual Word paste.',
    ].join('\n'),
  },
  {
    key: 'react-state-perf',
    focus: [
      'DIMENSION: React state management, performance, and numeric clamping (components/PasteInput.tsx, components/TableCard.tsx, components/RenderedPreview.tsx).',
      'Look hard at: the clampPt fallback, whose invalid-input fallback is 13 even though the DEFAULT_LEVEL size is 11 (an inconsistency when a size is cleared or invalid); the Number-top-levels select that uses a clamped value for display but stores and uses the raw numberDepth (a possible stale-clamp bug when tables are removed and the computed maxDepth shrinks); maxDepth being computed across all tables versus the per-table heading K; the removeTable activeId fallback logic that reads the pre-removal tables array; the copyAll and copyForWord building the payload synchronously before the first await to preserve the user-gesture requirement (is the build actually before the first await in both); the useMemo dependency arrays (html depends on table and numberDepth); the effectiveness of memo on the table card; the sparse levelStyles array and setLevel; and the indentInput and sizeInput string-state clamping. Flag real correctness or performance issues, not nitpicks.',
    ].join('\n'),
  },
  {
    key: 'security-xss',
    focus: [
      'DIMENSION: Security, XSS, and injection. The app injects renderer HTML via dangerouslySetInnerHTML (components/RenderedPreview.tsx) and writes text/html to the clipboard.',
      'Audit: escapeHtml, which escapes only ampersand, less-than, and greater-than. Is that sufficient given the claim that NO user value ever lands inside an attribute? Independently verify that claim across every path, including the new data-heading and field-label paths. Check wrapLabel. Check sanitizeStyleName for the mso-style-name value: it strips double-quote, less-than, and greater-than, but could a crafted Word style name still break out of the CSS rule p.MsoTitle open-brace mso-style-name colon quoted-name, for example via a close brace, a newline, or a CSS comment sequence? Check the font allow-list versus free text (bodyFont and the per-level font come from a fixed list; confirm no free-text font ever reaches a style attribute). Check whether pasted clipboard HTML is ever reflected back unescaped. Distinguish real exploitable issues from defense-in-depth notes, but report both, labeled.',
    ].join('\n'),
  },
  {
    key: 'ux-product-a11y',
    focus: [
      'DIMENSION: UX, product gaps, and accessibility. Consider the real user flow: paste a wide table, build the outline, copy to Word.',
      'Findings on: the paste target (a div with role textbox and an onPaste handler; keyboard, screen-reader, and focus behavior; can a user paste via the right-click menu or only Ctrl+V); error and empty states; discoverability of the Use Destination Styles requirement; the clarity of the numbering hint; the lack of undo and the lack of persistence across a page reload (is that a gap worth noting for a finalize tool); the 100-table cap; the aria-pressed toggles; the color contrast of the muted foreground text; the hash placeholder in the preview; and whether the preview truly matches Word. Prioritize gaps that would frustrate a real user finalizing a document. Avoid generic add-tests filler unless tied to a concrete risk.',
    ].join('\n'),
  },
  {
    key: 'docs-consistency',
    focus: [
      'DIMENSION: Docs and code sync (a HARD project rule). Compare CLAUDE.md, README.md, AGENTS.md, and everything under docs against the ACTUAL current code INCLUDING the uncommitted diff (git --no-pager diff HEAD).',
      'Find concrete drift: does any doc still describe removed behavior, for example MsoPiv-star body classes, or wording like the title is the only heading, or nothing is a Word heading, or a Download-for-Word feature, that the code no longer does? Are the new numberDepth numbering feature and the per-field label styling documented in the CLAUDE.md Data Model, Pivot view, and Styling sections and in the docs folder? List each stale or missing passage with the file, the quoted stale text, and the corrected text. This dimension findings should read as precise doc edits.',
    ].join('\n'),
  },
]

const shorten = (s) => (s || '').slice(0, 40)

const verifyPrompt = (f) => [
  SHARED,
  '',
  'You are an ADVERSARIAL VERIFIER. A reviewer reported the finding below. Your default stance is SKEPTICAL: independently read the ACTUAL code at the cited location (use Read and Grep; run git --no-pager diff HEAD if it concerns the in-flight work) and decide whether it is genuinely real. Refute it if: the code does not actually do what is claimed, the case cannot occur given how callers invoke it, it is already handled elsewhere, or it is vague or generic. Confirm it only if you can point to the specific real code that proves it. Also assess whether the suggested fix is correct and complete, or propose a better one. Set correctedSeverity to invalid if refuted.',
  '',
  'FINDING:',
  '- dimension: ' + f.dimension,
  '- title: ' + f.title,
  '- severity (claimed): ' + f.severity,
  '- category: ' + f.category,
  '- file: ' + f.file,
  '- location: ' + f.location,
  '- description: ' + f.description,
  '- evidence (claimed): ' + f.evidence,
  '- suggestedFix (claimed): ' + f.suggestedFix,
  '- confidence (claimed): ' + f.confidence,
].join('\n')

phase('Review')
log('Reviewing ' + DIMENSIONS.length + ' dimensions, verifying each finding adversarially...')

const results = await pipeline(
  DIMENSIONS,
  (d) => agent(SHARED + '\n\n' + d.focus, { label: 'review:' + d.key, phase: 'Review', schema: FINDINGS_SCHEMA }),
  (review, d) => {
    const findings = (review && review.findings ? review.findings : []).map((f) => ({ ...f, dimension: d.key }))
    if (findings.length === 0) return []
    return parallel(
      findings.map((f) => () =>
        agent(verifyPrompt(f), { label: 'verify:' + d.key + ':' + shorten(f.title), phase: 'Verify', schema: VERDICT_SCHEMA })
          .then((v) => ({ ...f, verdict: v }))
          .catch(() => ({ ...f, verdict: null })),
      ),
    )
  },
)

const all = results.flat().filter(Boolean)
const confirmed = all.filter((f) => f.verdict && f.verdict.isReal && f.verdict.correctedSeverity !== 'invalid')
const refuted = all.filter((f) => !(f.verdict && f.verdict.isReal && f.verdict.correctedSeverity !== 'invalid'))

log('Done: ' + all.length + ' raw findings, ' + confirmed.length + ' confirmed, ' + refuted.length + ' refuted or unverified.')

return {
  confirmed,
  refuted: refuted.map((f) => ({ dimension: f.dimension, title: f.title, claimedSeverity: f.severity, verdict: f.verdict })),
  counts: { total: all.length, confirmed: confirmed.length, refuted: refuted.length },
}
