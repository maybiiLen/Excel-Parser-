export const meta = {
  name: 'review-heading-feature',
  description: 'Adversarial review of the uncommitted Word-heading-rows feature + numbering-base fix before commit',
  phases: [
    { title: 'Review' },
    { title: 'Verify' },
  ],
}

const CONTEXT = [
  'You are reviewing the UNCOMMITTED changes in a client-side Next.js 16 app at the repo root (cwd is the project root). Run `git --no-pager diff HEAD` to see exactly what is under review; also read the surrounding code (some related numbering code is already committed in HEAD, so read the full files, not just the diff).',
  '',
  'THE APP: pastes an Excel/Sheets table, restructures it into an Excel-pivot-style nested outline, renders Word-ready HTML, copies it to the clipboard for Microsoft Word. Pipeline per table: parseClipboard (lib/parser.ts) -> rowsToPivotTree (lib/mapper.ts) -> renderPivotTree (lib/renderers.ts) -> buildWordHtml (lib/clipboard.ts) -> clipboard write. UI: components/PasteInput.tsx (parent, shared style, copyAll), components/TableCard.tsx (one table editor), components/tableModel.ts (TableState + tableToHtml), components/RenderedPreview.tsx (dangerouslySetInnerHTML preview).',
  '',
  'WHAT THIS UNCOMMITTED DIFF ADDS: a per-level "Word heading" feature. The user can mark an indent level (via a per-level checkbox in TableCard, labelled by field name) so that level\'s rows map to a destination Word Heading style (Heading K) for the Navigation pane + collapsibility, flush-left. The plumbing: TableState gained headingLevels (a boolean array per indent level); tableToHtml gained a titleIsHeading argument (whether a Heading style name is set for the title, via isHeadingStyleSet in lib/clipboard.ts); renderPivotTree gained headingLevels and titleIsHeading parameters. On a heading row the renderer SUPPRESSES the app number/marker (Word supplies the number) and tags the first line with a data-heading attribute equal to K, where K is one-if-the-title-is-a-heading-else-zero plus depth, clamped to 9. buildWordHtml maps each data-heading K row to a paragraph classed MsoHeadingK with a rule mapping it to mso-style-name Heading K and mso-outline-level K. RenderedPreview shows heading rows flush-left and bold with a muted hash placeholder where Word inserts the number, plus a one-line hint.',
  '',
  'THE NUMBERING ENGINE (multilevelNumbers in lib/renderers.ts) and a FIX in this diff: numbering is app-drawn static multilevel numbers. The Start value is a dotted-decimal string that is the exact number of the first top-level item; the top shown level advances the last component per sibling (Start 5.1 gives 5.1, 5.2) and deeper shown levels append dot-one-based-index (5.1 nests 5.1.1 then 5.1.1.1). numbering.levels is a per-level show/hide; a hidden non-heading level is TRANSPARENT (its children continue the nearest shown ancestor sequence, so numbers form a contiguous outline with no gaps and no collisions). THE FIX in this diff: a level mapped to a Word heading now ALSO establishes a numbered BASE (it is treated as shown-for-path even if its number is hidden), so deeper rows RESET under each heading group (TYPE heading Word-numbered 5.1, 5.2; the app numbers its TITLEs 5.1.1, 5.1.2, then 5.2.1 under the next type). Before the fix the heading level was transparent and TITLEs numbered continuously across types (5.1.5, 5.1.6, 5.1.7) which was the bug being fixed.',
  '',
  'HARD RULES: client-side only, no backend; never install xlsx from npm; keep docs in sync (CLAUDE.md, README.md, docs OVERVIEW/ARCHITECTURE/ROADMAP) when behavior or a control or the data model changes; all user-derived text must stay escaped (no XSS); style inputs are form-controlled.',
  '',
  'YOUR JOB: read the ACTUAL code (Read/Grep/Glob; run git --no-pager diff HEAD; run npm commands if useful) and report concrete, real findings for YOUR dimension only. Every finding cites a real file and location and quotes the specific evidence and gives a concrete fix. The numbering/heading interaction is the riskiest area -- TRACE concrete trees by hand. Prefer a few high-confidence findings over speculation; if the code is correct, say so with fewer findings. Severity: critical (data loss / broken output / security), high (wrong output in common cases), medium (edge-case bug / notable gap), low (minor), nit (style). Be specific to THIS codebase.',
].join('\n')

const FINDINGS_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    dimension: { type: 'string' },
    summary: { type: 'string' },
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
          location: { type: 'string' },
          description: { type: 'string' },
          evidence: { type: 'string' },
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
    isReal: { type: 'boolean' },
    correctedSeverity: { type: 'string', enum: ['critical', 'high', 'medium', 'low', 'nit', 'invalid'] },
    reasoning: { type: 'string' },
    fixCritique: { type: 'string' },
  },
  required: ['isReal', 'correctedSeverity', 'reasoning', 'fixCritique'],
}

const DIMENSIONS = [
  {
    key: 'numbering-engine',
    focus: 'DIMENSION: The numbering engine and the heading-base fix (lib/renderers.ts multilevelNumbers + the renderPivotTree walk). TRACE concrete trees by hand. Verify: a heading level establishes a base so children RESET under each heading group (5.1.1, 5.1.2 then 5.2.1), and this holds whether numbering.levels for that level is on OR off; a plain hidden (non-heading) level is still TRANSPARENT (children continue, no gaps, no collisions); the decimal Start and bumpLast (advance the last component per sibling) are correct, including multi-part starts and an invalid/empty start; the interaction when a level is BOTH a heading and has numbering.levels false; deep nesting and the clamp; the case where the TOP level is a heading vs a deeper level; whether the heading base aligns with what Word would number (the app cannot know Word numbering, so confirm the doc tells the user to set Start to match, and flag if mis-set Start silently misaligns). Look for any sibling-counter sharing bug, off-by-one, or a node getting the wrong base.',
  },
  {
    key: 'heading-word-fidelity',
    focus: 'DIMENSION: Heading rows + Word/clipboard output fidelity (lib/clipboard.ts buildWordHtml + isHeadingStyleSet, lib/renderers.ts heading-row emission, components/PasteInput.tsx copyAll, components/TableCard.tsx copyForWord). Verify: the data-heading attribute is emitted only on the first line of a heading node and suppresses BOTH the app number and the marker; K = (titleIsHeading?1:0)+depth is the right Word heading level (a body heading nests one under a Heading-1 title) and is clamped so it never exceeds Word\'s 9 heading styles or maps to a non-existent style; the buildWordHtml regex with the OPTIONAL data-heading group never double-matches or mis-matches the plain-row path or the nbsp spacer; the MsoHeadingK rule + mso-outline-level are correct; the combined copyAll (concatenate fragments, wrap once) is correct when tables differ; titleIsHeading is derived consistently in the card preview and both export paths so the preview matches the paste; whether a heading row that is ALSO a numbered/hidden level behaves right. Identify any Word-paste fidelity mismatch.',
  },
  {
    key: 'react-ui',
    focus: 'DIMENSION: React state + the new UI controls (components/TableCard.tsx, components/PasteInput.tsx, components/RenderedPreview.tsx). Check: the per-level Word-heading checkbox row (stored in table.headingLevels, labelled by headers[bucket[0]]) -- correct keying, sparse-array handling, persistence; the Start input string-state (startInput) sync, clamping, and the commit-on-valid / restore-on-blur; the blank-line dropdown (findIndex on breakAfter); titleIsHeading computed in the card vs the parent copyAll (consistency); the html useMemo deps (does it include titleIsHeading); RenderedPreview heading-row CSS (flush-left + bold + the hash ::before) and the hasHeadings hint via html.includes; new-table init includes headingLevels; memoization correctness. Flag real correctness/UX issues, not nitpicks.',
  },
  {
    key: 'security-xss',
    focus: 'DIMENSION: Security / XSS / injection for the new code. The app injects renderer HTML via dangerouslySetInnerHTML and writes text/html to the clipboard. Audit the NEW paths: the data-heading attribute value (is it always a clamped integer, never user text); the MsoHeadingK class + the Heading-K mso-style-name rule in buildWordHtml (K is a parsed integer -- confirm; is there any way user text reaches the class or the style block); isHeadingStyleSet / sanitizeStyleName; the RenderedPreview added CSS for [data-heading] (constant, no user value -- confirm); whether the heading-row content (label/value) is still escaped. Distinguish real exploitable issues from defense-in-depth notes, label both.',
  },
  {
    key: 'docs-build',
    focus: 'DIMENSION: Docs sync (a HARD rule) + build/lint ground truth. RUN npm run build and npm run lint and report any failure as a finding with the exact error. Then compare the uncommitted diff against CLAUDE.md, README.md, and docs OVERVIEW/ARCHITECTURE/ROADMAP: are the new Word-heading control, the headingLevels data-model field, the renderPivotTree/tableToHtml signature changes (now with headingLevels/titleIsHeading), and the heading-base numbering behavior all documented accurately, with no stale text describing the old behavior? Quote each stale or missing passage with the file and the corrected text.',
  },
]

const shorten = (s) => (s || '').slice(0, 40)

const verifyPrompt = (f) => [
  CONTEXT,
  '',
  'You are an ADVERSARIAL VERIFIER. A reviewer reported the finding below. Default to SKEPTICAL: independently read the ACTUAL code at the cited location (Read/Grep; run git --no-pager diff HEAD; trace concrete trees for numbering claims) and decide whether it is genuinely real. Refute it if the code does not do what is claimed, the case cannot occur given how callers invoke it, it is already handled, or it is vague. Confirm only if you can point to the specific real code that proves it. Assess whether the fix is correct and complete. Set correctedSeverity to invalid if refuted.',
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
log('Reviewing ' + DIMENSIONS.length + ' dimensions of the Word-heading feature, verifying each finding adversarially...')

const results = await pipeline(
  DIMENSIONS,
  (d) => agent(CONTEXT + '\n\n' + d.focus, { label: 'review:' + d.key, phase: 'Review', schema: FINDINGS_SCHEMA }),
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

log('Done: ' + all.length + ' raw findings, ' + confirmed.length + ' confirmed, ' + refuted.length + ' refuted/unverified.')

return {
  confirmed,
  refuted: refuted.map((f) => ({ dimension: f.dimension, title: f.title, claimedSeverity: f.severity, verdict: f.verdict })),
  counts: { total: all.length, confirmed: confirmed.length, refuted: refuted.length },
}
