# Net Cash Flow Popup — Design

## Problem

The Monthly Cash Flow insights section shows four metric tiles (Debit, Credit,
Net Cash Flow, Current Balance), each with a small sparkline. The sparklines
are too small to read exact monthly values or compare Credit vs. Debit vs. Net
Cash Flow side by side. Clicking the Net Cash Flow tile should open a large
popup with a proper multi-line chart of Credit, Debit, and Net Cash Flow per
month, scoped to whatever filters are currently active.

## Scope

- Only the **Net Cash Flow** tile is clickable. Debit, Credit, and Current
  Balance tiles are unchanged.
- The popup shows monthly Credit, Debit, and Net Cash Flow as three lines on
  one chart, plus a legend and hover tooltips with exact values.
- The popup reflects `state.filteredTransactions` and re-renders live if the
  user changes filters (search, account, date range, quick filters) while it
  is open — it does not snapshot data at open time.
- No new dependencies. The chart is hand-rolled SVG, following the existing
  `renderTrendlineChart` pattern already used for the Year-End Outlook chart.

## Data flow

`renderMonthlySummary()` currently builds `monthlySeries` (an array of
`[month, { debit, credit }]`, sorted ascending) inline, purely to derive the
four tile sparklines. This gets extracted into a standalone function:

```js
function computeMonthlyCashFlowSeries(transactions) {
  // returns [{ month, debit, credit }], sorted ascending by month
}
```

Both `renderMonthlySummary()` (for the tile sparklines) and the new
`renderCashFlowChart()` (for the popup) call this function on
`state.filteredTransactions`, so the popup and the tiles always agree —
there's a single source of truth for the monthly aggregation instead of
duplicated logic.

## Trigger & lifecycle

- The Net Cash Flow tile's rendered markup gets `data-tile="net-cash-flow"`
  and a `cursor: pointer` affordance.
- A delegated click handler on `els.monthlySummaryMetrics` opens the popup
  when the click target is inside a tile with `data-tile="net-cash-flow"`.
- `state.cashFlowModalOpen` (boolean, default `false`) tracks whether the
  popup is open.
- Opening sets the flag, unhides the panel, adds `overlay-open` to
  `document.body` (same as the import panel), and calls
  `renderCashFlowChart()` once immediately.
- `renderAll()` calls `renderCashFlowChart()` whenever
  `state.cashFlowModalOpen` is true, so the chart updates on every filter
  change while the popup is open. When the flag is false, `renderAll()`
  skips the chart render (no wasted work while hidden).
- Closing (via close button or backdrop click) clears the flag, hides the
  panel, and removes `overlay-open` from `document.body`.

## UI structure

New markup in `index.html`, following the existing `#import-panel` overlay
pattern exactly:

```html
<div id="cash-flow-modal" class="overlay-panel" hidden>
  <div class="overlay-backdrop" data-close-cash-flow-modal></div>
  <section class="glass-card overlay-card cash-flow-modal-shell" role="dialog" aria-modal="true" aria-labelledby="cash-flow-modal-title">
    <div class="overlay-header">
      <div>
        <p class="section-kicker">Insights</p>
        <h2 id="cash-flow-modal-title">Cash Flow Trend</h2>
      </div>
      <button id="close-cash-flow-modal" class="icon-btn icon-btn-subtle" type="button" aria-label="Close cash flow trend">Close</button>
    </div>
    <div class="cash-flow-legend">
      <span class="legend-swatch legend-credit"></span>Credit
      <span class="legend-swatch legend-debit"></span>Debit
      <span class="legend-swatch legend-net"></span>Net Cash Flow
    </div>
    <svg id="cash-flow-chart" viewBox="0 0 960 360" role="img" aria-label="Monthly credit, debit, and net cash flow"></svg>
  </section>
</div>
```

Sizing: `.cash-flow-modal-shell` uses the same `overlay-card` sizing as the
import panel (`width: min(1240px, calc(100vw - 32px))`), giving a chart
noticeably larger than the trend/outlook chart (760×280) — proposed at
960×360 viewBox, scaling to the card width.

## Chart rendering

`renderCashFlowChart()` mirrors `renderTrendlineChart()`'s structure:

- Empty state: if `computeMonthlyCashFlowSeries(state.filteredTransactions)`
  is empty, render the same `trend-empty` message pattern used elsewhere
  ("No monthly insights for the current filters...").
- Axes: horizontal gridlines at min/mid/max Y value (across all three series,
  so all lines share one scale), a zero line if the range crosses zero, and
  month labels along the X axis (reusing `trend-label` / `trend-gridline` /
  `trend-axis` / `trend-zero-line` CSS classes).
- Three `<path>` lines (Credit, Debit, Net Cash Flow), each with its own
  color via new CSS classes (`cash-flow-credit-line`, `cash-flow-debit-line`,
  `cash-flow-net-line`), same stroke width/rounding as `.trend-actual`.
- Per-point `<circle>` markers on each line, each containing a native SVG
  `<title>` child (e.g. `Mar 2026 · Credit: 45,200`) so hovering shows exact
  values via the browser's built-in tooltip — same mechanism already used in
  `renderTrendlineChart`, no custom JS tooltip tracking needed.
- Legend swatches are static HTML/CSS (fixed colors), not generated per
  render.

## Styling

New CSS additions, following existing `.trend-*` conventions:

- `.cash-flow-modal-shell` — sizing tweaks if the default `overlay-card`
  padding doesn't fit a wider chart.
- `.cash-flow-credit-line`, `.cash-flow-debit-line`, `.cash-flow-net-line` —
  distinct stroke colors (credit reuses the existing "positive" green,
  debit reuses the existing "negative" red, net cash flow gets a distinct
  third color, e.g. the existing `--gold` used for projections elsewhere).
- `.cash-flow-legend`, `.legend-swatch`, `.legend-credit/.legend-debit/.legend-net`
  — small inline color dots + labels, colors matching the line classes above.

## Out of scope

- Debit / Credit / Current Balance tiles remain non-clickable.
- No account-level breakdown inside the popup — it uses the same
  filtered-but-account-aggregated totals as the existing tile sparklines.
- No Escape-key close handler — matches the existing import-panel pattern,
  which only closes via button or backdrop click.
