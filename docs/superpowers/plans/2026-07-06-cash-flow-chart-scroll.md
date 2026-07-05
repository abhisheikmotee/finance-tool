# Scrollable Cash Flow Trend Chart Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** The Cash Flow Trend popup's chart grows with the amount of data instead of squeezing everything into a fixed width — a fixed value-axis panel stays put on the left while the plot (lines, points, per-month labels) scrolls horizontally, giving every month real spacing.

**Architecture:** Split the single `#cash-flow-chart` SVG into two synced pieces sharing one y-scale: a small fixed `#cash-flow-axis` SVG (value labels + tick marks, never scrolls) and `#cash-flow-chart` itself (now the scrolling plot, sized at 64px per month with an 872px floor), sitting inside a new `overflow-x: auto` div (`#cash-flow-scroll`). Every render snaps the scroll position to the right edge (most recent months). The prior fix's 12-label cap is removed — every month gets its own label now that scrolling provides real room.

**Tech Stack:** Vanilla JS, vanilla CSS, static HTML — no new dependencies, no custom scrollbar (native browser scrollbar, already themed globally via this app's `*::-webkit-scrollbar` rules).

## Global Constraints

- No new npm dependencies.
- `#cash-flow-chart` keeps its existing id and role/aria-label — existing tests and the tooltip/circle/path structure inside it are unaffected, only its sizing mechanics change.
- Axis panel width: 64px. Plot per-month width: 64px. Plot minimum width: 872px. Plot internal padding: `{ top: 24, right: 24, bottom: 44, left: 16 }`.
- No frozen behavior beyond the axis panel — no minimap, no zoom, no scroll-position memory across renders (every render snaps to the right edge).
- No change to any other chart in the app (Year-End Outlook `#trendline-chart`, tile sparklines).

---

## File Structure

- **`index.html`** — replace the `.trendline-chart-wrap` block around `#cash-flow-chart` (currently lines 290-292) with a two-panel frame: fixed axis SVG + scrollable div wrapping the plot SVG.
- **`styles.css`** — replace the `.cash-flow-chart { min-height: 360px; }` rule (currently lines 1864-1866) with `.cash-flow-chart-frame`, `.cash-flow-axis`, `.cash-flow-scroll`, `.cash-flow-plot`.
- **`app.js`** — `cacheElements()` (currently ~line 232-234): cache the two new elements. `renderCashFlowChart()` (currently lines 1678-1746): full rewrite per the design — two-panel rendering, dynamic plot width, no label thinning, auto-scroll-to-end.
- **`tests/cash-flow-popup.spec.js`** — replace the "chart thins x-axis labels when many months are in view" test (added by the prior fix) with a test proving every month gets a label and the chart actually scrolls for wide ranges.

---

### Task 1: Scrollable two-panel chart

**Files:**
- Modify: `index.html:290-292`
- Modify: `styles.css:1864-1866`
- Modify: `app.js` (`cacheElements()` ~line 232-234, `renderCashFlowChart()` lines 1678-1746)
- Modify: `tests/cash-flow-popup.spec.js` (replace the label-thinning test)

**Interfaces:**
- Consumes: `computeMonthlyCashFlowSeries(transactions)` (unchanged, from the earlier Net Cash Flow popup work), `formatMonthShort`, `compactMoneyFormat`, `moneyFormat`, `escapeHtml` (all unchanged, existing).
- Produces: no new function names — `renderCashFlowChart()` keeps its existing signature and existing callers (`openCashFlowModal()`, `renderAll()`'s conditional call) are unaffected.

- [ ] **Step 1: Write the failing test**

In `tests/cash-flow-popup.spec.js`, replace this entire test (added by the prior label-thinning fix):

```js
  test("chart thins x-axis labels when many months are in view", async ({ page }) => {
    await freezeDate(page, "2026-07-06T12:00:00Z");
    await page.goto("/");

    const rows = [];
    for (let i = 0; i < 18; i += 1) {
      const year = 2024 + Math.floor(i / 12);
      const month = String((i % 12) + 1).padStart(2, "0");
      rows.push({ txnDate: `${year}-${month}-15`, accountNumber: "0001", debit: 1000, credit: 2000, balance: 1000 });
    }
    await seedTransactions(page, rows);
    await page.evaluate(() => {
      state.quickFilters.datePreset = "all";
      applyFilters();
      renderAll();
    });

    await page.locator('.metric-tile[data-tile="net-cash-flow"]').click();

    const chart = page.locator("#cash-flow-chart");
    await expect(chart.locator("circle")).toHaveCount(18 * 3);
    const labelCount = await chart.locator('text[text-anchor="middle"]').count();
    expect(labelCount).toBeLessThanOrEqual(12);
  });
```

with:

```js
  test("chart gives every month its own label via horizontal scroll instead of thinning them", async ({ page }) => {
    await freezeDate(page, "2026-07-06T12:00:00Z");
    await page.goto("/");

    const rows = [];
    for (let i = 0; i < 30; i += 1) {
      const year = 2024 + Math.floor(i / 12);
      const month = String((i % 12) + 1).padStart(2, "0");
      rows.push({ txnDate: `${year}-${month}-15`, accountNumber: "0001", debit: 1000, credit: 2000, balance: 1000 });
    }
    await seedTransactions(page, rows);
    await page.evaluate(() => {
      state.quickFilters.datePreset = "all";
      applyFilters();
      renderAll();
    });

    await page.locator('.metric-tile[data-tile="net-cash-flow"]').click();

    const chart = page.locator("#cash-flow-chart");
    await expect(chart.locator("circle")).toHaveCount(30 * 3);
    await expect(chart.locator('text[text-anchor="middle"]')).toHaveCount(30);

    const scrollState = await page.locator("#cash-flow-scroll").evaluate((el) => ({
      scrollLeft: el.scrollLeft,
      scrollWidth: el.scrollWidth,
      clientWidth: el.clientWidth,
    }));
    expect(scrollState.scrollWidth).toBeGreaterThan(scrollState.clientWidth);
    expect(scrollState.scrollLeft).toBeGreaterThanOrEqual(scrollState.scrollWidth - scrollState.clientWidth - 1);
  });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx playwright test tests/cash-flow-popup.spec.js -g "gives every month its own label"`
Expected: FAIL — `#cash-flow-scroll` doesn't exist yet (locator resolves to nothing / `evaluate` throws), and with the current fixed-width chart the label count is capped at 12, not 30.

- [ ] **Step 3: Update the markup**

In `index.html`, replace:

```html
      <div class="trendline-chart-wrap">
        <svg id="cash-flow-chart" class="trendline-chart cash-flow-chart" viewBox="0 0 960 360" role="img" aria-label="Monthly credit, debit, and net cash flow"></svg>
      </div>
```

with:

```html
      <div class="cash-flow-chart-frame">
        <svg id="cash-flow-axis" class="cash-flow-axis" width="64" height="360" role="presentation" aria-hidden="true"></svg>
        <div id="cash-flow-scroll" class="cash-flow-scroll">
          <svg id="cash-flow-chart" class="cash-flow-plot" height="360" role="img" aria-label="Monthly credit, debit, and net cash flow"></svg>
        </div>
      </div>
```

- [ ] **Step 4: Update the CSS**

In `styles.css`, replace:

```css
.cash-flow-chart {
  min-height: 360px;
}
```

with:

```css
.cash-flow-chart-frame {
  display: flex;
  border-radius: 20px;
  border: 1px solid var(--line);
  background: rgba(255, 255, 255, 0.82);
  overflow: hidden;
}

.cash-flow-axis {
  flex: 0 0 auto;
  display: block;
}

.cash-flow-scroll {
  flex: 1 1 auto;
  overflow-x: auto;
  overflow-y: hidden;
}

.cash-flow-plot {
  display: block;
}
```

- [ ] **Step 5: Cache the new elements**

In `app.js`, in `cacheElements()`, immediately after the line `els.cashFlowChart = document.getElementById("cash-flow-chart");`, add:

```js
  els.cashFlowAxis = document.getElementById("cash-flow-axis");
  els.cashFlowScroll = document.getElementById("cash-flow-scroll");
```

- [ ] **Step 6: Rewrite `renderCashFlowChart()`**

In `app.js`, replace the entire current function body:

```js
function renderCashFlowChart() {
  const series = computeMonthlyCashFlowSeries(state.filteredTransactions);
  if (!series.length) {
    els.cashFlowChart.innerHTML = `<foreignObject x="0" y="0" width="960" height="360"><div xmlns="http://www.w3.org/1999/xhtml" class="trend-empty">No monthly insights for the current filters. Try widening the date range or clearing bank filters.</div></foreignObject>`;
    return;
  }

  const width = 960;
  const height = 360;
  const pad = { top: 24, right: 24, bottom: 44, left: 64 };
  const points = series.map((row) => ({
    month: row.month,
    credit: row.credit,
    debit: row.debit,
    net: row.credit - row.debit,
  }));
  const allValues = points.flatMap((point) => [point.credit, point.debit, point.net]);
  allValues.push(0);
  let minY = Math.min(...allValues);
  let maxY = Math.max(...allValues);
  if (minY === maxY) {
    minY -= 1;
    maxY += 1;
  }
  const rangePadding = (maxY - minY) * 0.15;
  minY -= rangePadding;
  maxY += rangePadding;
  const xStep = points.length > 1 ? (width - pad.left - pad.right) / (points.length - 1) : 0;
  const xForIndex = (index) => pad.left + xStep * index;
  const yForValue = (value) => pad.top + ((maxY - value) / (maxY - minY)) * (height - pad.top - pad.bottom);
  const monthLabel = (month) => `${formatMonthShort(month)} ${month.slice(0, 4)}`;

  const seriesConfig = [
    { key: "credit", lineClass: "cash-flow-credit-line", pointClass: "cash-flow-point-credit", label: "Credit" },
    { key: "debit", lineClass: "cash-flow-debit-line", pointClass: "cash-flow-point-debit", label: "Debit" },
    { key: "net", lineClass: "cash-flow-net-line", pointClass: "cash-flow-point-net", label: "Net Cash Flow" },
  ];

  const pathFor = (key) => points
    .map((point, index) => `${index === 0 ? "M" : "L"} ${xForIndex(index).toFixed(2)} ${yForValue(point[key]).toFixed(2)}`)
    .join(" ");

  const gridValues = [minY, (minY + maxY) / 2, maxY];
  const maxLabels = 12;
  const lastIndex = points.length - 1;
  const labelCount = Math.min(maxLabels, points.length);
  const labelIndices = new Set(Array.from({ length: labelCount }, (_, i) => (
    labelCount > 1 ? Math.round((i * lastIndex) / (labelCount - 1)) : 0
  )));

  els.cashFlowChart.innerHTML = `
    ${gridValues.map((value) => `
      <g>
        <line class="trend-gridline ${Math.abs(value) < 0.0001 ? "trend-zero-line" : ""}" x1="${pad.left}" y1="${yForValue(value)}" x2="${width - pad.right}" y2="${yForValue(value)}"></line>
        <text class="trend-label" x="10" y="${yForValue(value) + 4}">${escapeHtml(compactMoneyFormat(value))}</text>
      </g>
    `).join("")}
    <line class="trend-axis" x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}"></line>
    ${seriesConfig.map((config) => `<path class="${config.lineClass}" d="${pathFor(config.key)}"></path>`).join("")}
    ${points.map((point, index) => `
      <g>
        ${seriesConfig.map((config) => `
          <circle class="trend-point ${config.pointClass}" cx="${xForIndex(index)}" cy="${yForValue(point[config.key])}" r="4.5">
            <title>${escapeHtml(`${monthLabel(point.month)} ${config.label}: ${moneyFormat(point[config.key])}`)}</title>
          </circle>
        `).join("")}
        ${labelIndices.has(index) ? `<text class="trend-label" x="${xForIndex(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(monthLabel(point.month))}</text>` : ""}
      </g>
    `).join("")}
  `;
}
```

with:

```js
function renderCashFlowChart() {
  const series = computeMonthlyCashFlowSeries(state.filteredTransactions);
  const height = 360;
  const axisWidth = 64;
  const minPlotWidth = 872;
  const pxPerMonth = 64;
  const plotPad = { top: 24, right: 24, bottom: 44, left: 16 };

  if (!series.length) {
    els.cashFlowAxis.innerHTML = "";
    els.cashFlowChart.setAttribute("width", String(minPlotWidth));
    els.cashFlowChart.innerHTML = `<foreignObject x="0" y="0" width="${minPlotWidth}" height="${height}"><div xmlns="http://www.w3.org/1999/xhtml" class="trend-empty">No monthly insights for the current filters. Try widening the date range or clearing bank filters.</div></foreignObject>`;
    return;
  }

  const points = series.map((row) => ({
    month: row.month,
    credit: row.credit,
    debit: row.debit,
    net: row.credit - row.debit,
  }));
  const allValues = points.flatMap((point) => [point.credit, point.debit, point.net]);
  allValues.push(0);
  let minY = Math.min(...allValues);
  let maxY = Math.max(...allValues);
  if (minY === maxY) {
    minY -= 1;
    maxY += 1;
  }
  const rangePadding = (maxY - minY) * 0.15;
  minY -= rangePadding;
  maxY += rangePadding;

  const plotWidth = Math.max(minPlotWidth, plotPad.left + plotPad.right + pxPerMonth * Math.max(points.length - 1, 1));
  const xStep = points.length > 1 ? (plotWidth - plotPad.left - plotPad.right) / (points.length - 1) : 0;
  const xForIndex = (index) => plotPad.left + xStep * index;
  const yForValue = (value) => plotPad.top + ((maxY - value) / (maxY - minY)) * (height - plotPad.top - plotPad.bottom);
  const monthLabel = (month) => `${formatMonthShort(month)} ${month.slice(0, 4)}`;

  const seriesConfig = [
    { key: "credit", lineClass: "cash-flow-credit-line", pointClass: "cash-flow-point-credit", label: "Credit" },
    { key: "debit", lineClass: "cash-flow-debit-line", pointClass: "cash-flow-point-debit", label: "Debit" },
    { key: "net", lineClass: "cash-flow-net-line", pointClass: "cash-flow-point-net", label: "Net Cash Flow" },
  ];

  const pathFor = (key) => points
    .map((point, index) => `${index === 0 ? "M" : "L"} ${xForIndex(index).toFixed(2)} ${yForValue(point[key]).toFixed(2)}`)
    .join(" ");

  const gridValues = [minY, (minY + maxY) / 2, maxY];

  els.cashFlowAxis.innerHTML = gridValues.map((value) => `
    <g>
      <text class="trend-label" x="10" y="${yForValue(value) + 4}">${escapeHtml(compactMoneyFormat(value))}</text>
      <line class="trend-gridline ${Math.abs(value) < 0.0001 ? "trend-zero-line" : ""}" x1="${axisWidth - 8}" y1="${yForValue(value)}" x2="${axisWidth}" y2="${yForValue(value)}"></line>
    </g>
  `).join("");

  els.cashFlowChart.setAttribute("width", String(plotWidth));
  els.cashFlowChart.innerHTML = `
    ${gridValues.map((value) => `<line class="trend-gridline ${Math.abs(value) < 0.0001 ? "trend-zero-line" : ""}" x1="0" y1="${yForValue(value)}" x2="${plotWidth}" y2="${yForValue(value)}"></line>`).join("")}
    <line class="trend-axis" x1="0" y1="${height - plotPad.bottom}" x2="${plotWidth}" y2="${height - plotPad.bottom}"></line>
    ${seriesConfig.map((config) => `<path class="${config.lineClass}" d="${pathFor(config.key)}"></path>`).join("")}
    ${points.map((point, index) => `
      <g>
        ${seriesConfig.map((config) => `
          <circle class="trend-point ${config.pointClass}" cx="${xForIndex(index)}" cy="${yForValue(point[config.key])}" r="4.5">
            <title>${escapeHtml(`${monthLabel(point.month)} ${config.label}: ${moneyFormat(point[config.key])}`)}</title>
          </circle>
        `).join("")}
        <text class="trend-label" x="${xForIndex(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(monthLabel(point.month))}</text>
      </g>
    `).join("")}
  `;

  els.cashFlowScroll.scrollLeft = els.cashFlowScroll.scrollWidth;
}
```

- [ ] **Step 7: Run the focused tests to verify they pass**

Run: `npx playwright test tests/cash-flow-popup.spec.js`
Expected: PASS (all 5 tests — the 4 pre-existing ones plus the new scroll/label test)

- [ ] **Step 8: Run the full suite to check for regressions**

Run: `npx playwright test`
Expected: PASS (no regressions in `tests/dashboard-smoke.spec.js`)

- [ ] **Step 9: Commit**

```bash
git add index.html styles.css app.js tests/cash-flow-popup.spec.js
git commit -m "$(cat <<'EOF'
Make Cash Flow Trend chart scroll instead of squeezing months

Splits the chart into a fixed value-axis panel and a horizontally
scrollable plot panel sized at 64px per month (872px floor), so wide
date ranges get real per-month spacing instead of being crushed into
a fixed width. Every month now gets its own x-axis label; the
previous fix's 12-label cap is removed since scrolling makes it
unnecessary. The plot auto-scrolls to the most recent months on every
render.
EOF
)"
```

---

## Self-Review

**Spec coverage:**
- Two-panel split (fixed axis + scrolling plot), synced y-scale → Step 6 (`renderCashFlowChart` rewrite computes `yForValue` once and uses it for both `els.cashFlowAxis` and `els.cashFlowChart`). ✓
- 64px/month, 872px floor, `{top:24,right:24,bottom:44,left:16}` padding, 64px axis width → all present verbatim in Step 6's constants. ✓
- Every month gets a label, no thinning → Step 6 renders one `<text text-anchor="middle">` per point unconditionally; Step 1's test asserts count equals point count (30), not a cap. ✓
- Auto-scroll to right edge on every render → `els.cashFlowScroll.scrollLeft = els.cashFlowScroll.scrollWidth;` at the end of Step 6; Step 1's test asserts `scrollLeft` is at the far edge. ✓
- Native scrollbar, no new dependency → `.cash-flow-scroll { overflow-x: auto; }` in Step 4, nothing else. ✓
- `#cash-flow-chart` keeps its id/role/aria-label → Step 3's markup preserves them exactly. ✓

**Placeholder scan:** No TBD/TODO; every step has complete code or an exact command with expected output.

**Type consistency:** `renderCashFlowChart()` keeps its zero-argument signature; no other function's name or signature changes, so `openCashFlowModal()` and `renderAll()`'s existing calls to it need no changes (not touched by this plan).

**Scope check:** Single cohesive change (one chart, one rendering function, its markup and CSS) — appropriately sized for one task.
