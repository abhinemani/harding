# CLAUDE.md

Guidance for Claude (via claude.ai, Claude Code, or any agent) working in this repo.

## What this repo is

`index.html` is the entire app: a single self-contained HTML file with React, PapaParse, and SheetJS inlined, and no build step. It reads a county division's line-item budget export and calls the Anthropic API directly from the browser to draft a program layer to Frederick County's Program Standard.

`docs/program_layer_builder.artifact.jsx` is the same app in its original form, written to run as a Claude.ai Artifact (ES module JSX, imports React/Papa/XLSX as packages, API key injected by the host rather than typed by the user). It is kept for convenience when iterating inside claude.ai, where printing and downloads work differently than in a normal browser. **The two files are not automatically kept in sync** — after editing one, port the change to the other by hand, or ask an agent to do it. See "Keeping the two builds in sync" below.

## Editing `index.html`

There is no bundler. The file is: a `<style>` block, then four inlined vendor scripts (React, ReactDOM, PapaParse, SheetJS, each unminified-adjacent and pasted verbatim from npm), then one `<script>` block containing the whole app as plain (non-JSX, non-ESM) JavaScript using `React.createElement` — actually, JSX is compiled out ahead of time; do not hand-edit the compiled `React.createElement(...)` calls if you can avoid it. The practical workflow:

1. Edit `docs/program_layer_builder.artifact.jsx` (real JSX, easy to read and edit).
2. Recompile it to plain JS with Babel standalone (`@babel/preset-react`, classic runtime — not the automatic runtime, which emits an `import` statement that breaks in a plain `<script>` tag).
3. Reassemble `index.html`: vendor scripts + the compiled app script + a `ReactDOM.createRoot(...).render(...)` call.
4. Confirm the result has zero `import`/`export` statements and zero external `<script src="https://...">` references — both are load-bearing constraints (module syntax fails in a plain script tag; external CDN references fail when the file is opened as `file://` with no network, or when CSP blocks cross-origin scripts).

The version history of this repo has the exact build script used each time (search commit messages / the diff in `index.html` for the compiled block). If you don't have that script, the shape is: strip `import`/`export default`, add `const apiKey`/model/effort state, add real `download()`/`window.print()` calls in place of the Artifact's clipboard-copy fallbacks, transform with Babel's React preset (classic runtime), concatenate after the vendor bundles.

## Keeping the two builds in sync

The two files differ in exactly these ways, and no others should creep in:

- **Module syntax**: `index.html` has none; the Artifact uses ES imports/export.
- **API auth**: `index.html` has a password-type input for the user's own API key, stored in `localStorage`, sent as `x-api-key` with `anthropic-dangerous-direct-browser-access: true`. The Artifact has no key input; the host injects credentials.
- **Model/effort choice**: `index.html` exposes a model dropdown (Haiku 4.5 / Sonnet 4.6 / Sonnet 5 / Opus 5) and an effort dropdown, both persisted to `localStorage`. The Artifact is hardcoded to `claude-sonnet-4-6` with no effort parameter, since that's what's available inside claude.ai.
- **Exports**: `index.html` uses real `Blob` + `<a download>` for CSV/Markdown/HTML and a real `window.print()` for PDF. The Artifact copies the same content to the clipboard and to an on-screen `<textarea>` instead, because the Artifact sandbox blocks both downloads and the print dialog.

Everything else — the prompt, the Program Standard logic, the costing/allocation math, the card layout, the responsive breakpoints, the budget-office checks — should be identical between the two. If you change the prompt or the UI, change it in the Artifact first (it's readable JSX) and port the same diff into the compiled block in `index.html`.

## Saved runs (local build only)

`index.html` persists each completed run to `localStorage` under the key `harding_runs` (an array, newest first, capped at 40). A run record holds everything needed to re-render without the original file: `items`, `recoveries`, `emptyCenters`, `zeroCount`, `pools`, `sendCount`, and the parsed `result`, plus metadata. Opening a run sets a `loaded` override that takes precedence over the file-derived memos (`computedItems` → `items`, etc.); uploading a new file clears it. The Artifact build has none of this, since Artifacts cannot use browser storage.

## The prompt

The prompt lives in `buildPrompt()` near the top of the app logic. It encodes:

- The Program Standard's definition, four boundary rules, statement grammar, and measure requirements, reproduced close to verbatim from the source document (`The Program Standard`, Public Works LLC / Funkhouser & Associates).
- A list of misdescription patterns observed across the four Frederick County pilot divisions tested (HR, Energy & Environment, IIT, Aging & Independence) — divisions list visible initiatives and omit core operating work, carry one cost figure across several self-described rows, list obligations funded elsewhere at $0, name programs for statutes rather than services, offer activity counts or "remain in compliance" as measures.
- Costing rules for staff-time allocation, grant-structured rollups, dominant-line portfolios (e.g., a single $12M software line), and permit/mandate-organized divisions.

If you add a new division type and learn something new about how its budget misleads a naive read, add it to this list rather than special-casing it in code — the pattern is more durable than the fix.

## Testing without an API key

There's no offline/mock mode. To test UI changes without spending API credits, either (a) hardcode a `result` object matching the JSON schema in `buildPrompt()` and skip the fetch, or (b) run against Haiku 4.5 at low effort, which is inexpensive.

## What NOT to change casually

- The program record shape (name, facing, statement, beneficiary, fte, measures.{result,volume,unit_cost}, mandate, strategic_linkage, alignment, twenty_percent, is_admin) mirrors the Program Standard document's section 4, plus two fields from the Phase 3 kickoff deck: `facing` (internal → performance/service-level result measure; external → outcome result measure) and `alignment` (direct / contributes / none, for the program-to-priority crosswalk). If the Standard changes, update the prompt schema and every place that reads `p.<field>` (the card renderer, the CSV/Markdown/HTML exporters, and the budget-office `checks` array) together — they're currently kept in sync by hand, not by a shared type.
- The compression step (pooling personnel and small operating lines per costing center before sending to the model) is load-bearing for large divisions like Aging (1,900+ raw rows). Don't remove it without re-testing against that file; token costs and latency scale with sent-line count, not raw row count.
