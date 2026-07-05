# Cash Flow Trend Chart — Scrollable, Fully-Labeled Design

## Problem

The Cash Flow Trend popup (added in the Net Cash Flow popup feature) renders
its 3-line chart at a fixed width. A prior fix capped x-axis labels at 12 and
spaced them evenly so they wouldn't overlap when many months were in view,
but with a wide date filter (years of data) the actual data points and lines
still get crushed into that fixed width — the chart becomes an unreadable,
overlapping blob on the left side, even though the (thinned) labels are
legible.

The chart needs to actually give each month room to breathe. That means the
plot area must grow with the amount of data and become horizontally
scrollable, rather than always squeezing into one fixed width.

## Scope

- Only the Cash Flow Trend popup's chart (`renderCashFlowChart`,
  `#cash-flow-chart`) is affected. No other chart in the app changes.
- The chart becomes two synced SVG panels: a small fixed value-axis panel
  that never scrolls, and a wider plot panel (lines, points, tooltips,
  per-month labels) inside a horizontally scrollable container.
- Every month gets its own x-axis label — the 12-label cap added in the
  previous fix is removed, since it was a workaround for fixed-width
  squeezing that real scrolling now solves properly.
- The plot auto-scrolls to its right edge (most recent months) every time
  the popup opens or live-updates from a filter change.
- No new dependencies. Scrolling uses a plain `overflow-x: auto` div and the
  browser's native scrollbar (already themed globally via this app's
  existing `*::-webkit-scrollbar` rules) — no custom scrollbar UI.

## Architecture

`renderCashFlowChart()` currently renders one SVG (`#cash-flow-chart`) at a
fixed logical width (960, scaled via `viewBox` to fill its container). It
becomes two elements, rendered by the same function, sharing one y-scale:

- **`#cash-flow-axis`** — a small SVG, fixed width (64px), fixed height
  (360px), never scrolls. Contains only the three gridline value labels
  (e.g. "2.6M", "507.6K", "-1.6M") and a short tick mark next to each,
  vertically positioned by the same `yForValue()` used by the plot.
- **`#cash-flow-chart`** — the plot SVG, fixed height (360px), but its
  `width` attribute is set in JS on every render based on how many months
  are in the series: `64px per month`, with an `872px` floor (matching
  today's usable plot width) so short ranges still fill the space. This SVG
  sits inside `#cash-flow-scroll`, a `div` with `overflow-x: auto` — when
  the plot's width exceeds the div's visible width, a horizontal scrollbar
  appears automatically.
- The plot renders: horizontal gridlines (spanning the full plot width),
  the zero line if the range crosses zero, the axis baseline, the three
  series `<path>` lines, one `<circle>` per series per month (with the
  existing native SVG `<title>` tooltip), and now **one `<text>` x-axis
  label per month, unconditionally** — no thinning.
- After every render, `els.cashFlowScroll.scrollLeft = els.cashFlowScroll.scrollWidth;`
  snaps the view to the right edge (most recent months), matching the
  behavior of today's fixed-width chart (which always showed the latest
  months on the right).

Markup shape (both panels live inside one bordered frame, replacing the
current single `.trendline-chart-wrap`):

```html
<div class="cash-flow-chart-frame">
  <svg id="cash-flow-axis" class="cash-flow-axis" width="64" height="360" role="presentation" aria-hidden="true"></svg>
  <div id="cash-flow-scroll" class="cash-flow-scroll">
    <svg id="cash-flow-chart" class="cash-flow-plot" height="360" role="img" aria-label="Monthly credit, debit, and net cash flow"></svg>
  </div>
</div>
```

`#cash-flow-chart` keeps its existing id (tests already select it), but
drops the `trendline-chart` class and `viewBox` attribute — it no longer
scales to its container; it renders at its true pixel width and scrolls
instead.

## Sizing constants

- Axis panel width: 64px (matches the old `pad.left`, already known to fit
  values like "-1.6M" at the existing 12px label font).
- Plot per-month width: 64px.
- Plot minimum width: 872px (today's actual usable plot width: 960 - 64
  left pad - 24 right pad), so a 1-3 month range still renders as a full,
  reasonably-proportioned chart rather than a tiny sliver.
- Plot internal padding: `{ top: 24, right: 24, bottom: 44, left: 16 }` —
  `left` shrinks from the old 64 to 16 since the value-axis text moved out
  to the separate fixed panel; the plot itself no longer needs to reserve
  room for it.

## Removed behavior

- The `maxLabels`/`labelIndices` thinning logic added in the prior fix is
  deleted entirely from `renderCashFlowChart()`. It solved the wrong layer
  of the problem (labels only) instead of the real one (not enough
  horizontal room for the data itself).

## Testing

- The existing "chart thins x-axis labels when many months are in view"
  test is replaced: with real scrolling, seeding 18 months should now
  assert **18** labels are present (`text[text-anchor="middle"]` count),
  not a capped 12.
- A new test asserts scrolling actually engages for wide ranges: seed
  enough months that the plot's rendered width exceeds a fixed viewport
  (e.g. 24 months → plot width ≈ 872px floor is exceeded), then check
  `#cash-flow-chart`'s `width` attribute/bounding box is wider than
  `#cash-flow-scroll`'s client width, and that after opening,
  `scrollLeft` is at (or very near) `scrollWidth - clientWidth` (i.e.
  scrolled to the right edge).
- Existing tests that assert circle/path counts and tooltip text
  (`tests/cash-flow-popup.spec.js`) are otherwise unaffected — the series
  data and tooltip format don't change, only the sizing/scroll mechanics.

## Out of scope

- No frozen/sticky behavior beyond the axis panel — no minimap, no
  zoom, no scroll-position memory across filter changes (every render
  snaps back to the right edge).
- No change to any other chart in the app (Year-End Outlook, sparklines).
