# Net Cash Flow Popup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Clicking the "Net Cash Flow" tile in the Monthly Cash Flow insights section opens a large popup showing Credit, Debit, and Net Cash Flow as three lines on one chart, scoped to the currently active filters, updating live if those filters change while the popup is open.

**Architecture:** `app.js` is a single dependency-free script (no build step, no framework) that mutates a module-level `state` object and re-renders DOM via template-string `innerHTML` assignment. The popup follows the app's existing `#import-panel` overlay pattern (hidden `.overlay-panel` + `.overlay-backdrop` + `.overlay-card`) and its existing hand-rolled SVG chart pattern (`renderTrendlineChart`, using native SVG `<title>` elements for tooltips instead of custom JS tooltip tracking).

**Tech Stack:** Vanilla JS, vanilla CSS, static HTML — no new dependencies. Tests use the existing Playwright suite (`tests/*.spec.js`), served via `python3 -m http.server 4173` per `playwright.config.js`.

## Global Constraints

- No new npm dependencies (project has zero `dependencies`, only `@playwright/test` as a devDependency).
- Follow the existing `.trend-*` CSS naming and visual conventions already used by `renderTrendlineChart` (gridlines, axis, point circles, labels) rather than inventing a new chart visual style.
- Reuse the existing `.overlay-panel` / `.overlay-backdrop` / `.overlay-card` / `.icon-btn.icon-btn-subtle` markup and CSS already used by `#import-panel`, rather than creating parallel modal infrastructure.
- Only the Net Cash Flow tile becomes clickable; Debit, Credit, and Current Balance tiles are unchanged.
- `state` and every top-level function in `app.js` (e.g. `applyFilters`, `renderAll`, `buildTransaction`) are reachable as bare identifiers from Playwright's `page.evaluate`, even though they are not properties of `window` (verified directly: classic `<script>` top-level `const`/`function` declarations are visible to code evaluated in the same page realm). Tests use this to seed `state.transactions` directly instead of driving CSV import through the UI.

---

## File Structure

- **`app.js`** — modify:
  - `state` initializer (~line 160-188): add `cashFlowModalOpen` flag.
  - New function `computeMonthlyCashFlowSeries(transactions)`, placed immediately before `renderMonthlySummary` (~line 1516): single source of truth for monthly Credit/Debit totals, replacing the ad-hoc `monthTotals`/`monthlySeries` computation currently inlined in `renderMonthlySummary`.
  - `renderMonthlySummary` (~line 1516-1630): use the new helper; tag the Net Cash Flow tile with `data-tile="net-cash-flow"` and (from Task 2) clickable affordances.
  - New functions `openCashFlowModal`, `closeCashFlowModal`, `renderCashFlowChart`, placed after `renderMonthlySummary`.
  - `renderAll` (~line 1194-1207): call `renderCashFlowChart()` when the popup is open, so every existing `renderAll()` call site (filter changes, imports, tax edits, etc.) keeps the popup live.
  - `cacheElements` (~line 203-249): cache the new modal's DOM elements.
  - `bindEvents` (~line 251-348): wire tile click/keydown to open, close button/backdrop to close.
- **`index.html`** — add a `#cash-flow-modal` overlay block after the existing `#import-panel` block (~line 273), modeled on it directly.
- **`styles.css`** — add legend and per-series line/point color classes, plus a clickable-tile affordance class.
- **`tests/cash-flow-popup.spec.js`** — new file, following the existing `tests/dashboard-smoke.spec.js` conventions (same `freezeDate` helper).

---

### Task 1: Shared monthly cash-flow series + sparkline bug fix

**Files:**
- Modify: `app.js:160-188` (state), `app.js:1516-1630` (`renderMonthlySummary`)
- Test: `tests/cash-flow-popup.spec.js` (create)

**Interfaces:**
- Produces: `computeMonthlyCashFlowSeries(transactions: Transaction[]) => Array<{ month: string, debit: number, credit: number }>`, sorted ascending by `month` (`YYYY-MM`). Task 2's `renderCashFlowChart` consumes this directly.
- Produces: `data-tile="net-cash-flow"` attribute on the Net Cash Flow tile's `<article>`. Task 2's click/keydown delegation and tests select tiles via `[data-tile="net-cash-flow"]`.

This task also fixes a pre-existing bug: the Net Cash Flow and Current Balance tile sparklines currently plot one point per **(month, account)** pair instead of one point per month, because they read from `monthlyRows` (per-account rows) instead of the per-month-aggregated series the Debit/Credit tiles already use correctly. This surfaced while building the shared helper the popup needs, and both tiles must show correct data since the Net Cash Flow tile is about to become the popup's entry point.

- [ ] **Step 1: Write the failing test**

Create `tests/cash-flow-popup.spec.js`:

```js
const { test, expect } = require("@playwright/test");

async function freezeDate(page, isoDate) {
  await page.addInitScript(({ now }) => {
    const fixedTime = new Date(now).getTime();
    const OriginalDate = Date;

    class MockDate extends OriginalDate {
      constructor(...args) {
        if (args.length === 0) {
          super(fixedTime);
          return;
        }
        super(...args);
      }

      static now() {
        return fixedTime;
      }

      static parse(value) {
        return OriginalDate.parse(value);
      }

      static UTC(...args) {
        return OriginalDate.UTC(...args);
      }
    }

    window.Date = MockDate;
  }, { now: isoDate });
}

async function seedTransactions(page, rows) {
  await page.evaluate((transactionRows) => {
    state.transactions = transactionRows.map((row, index) => buildTransaction({
      bankName: row.bankName || "SBM",
      accountNumber: row.accountNumber || "0001",
      currency: row.currency || "MUR",
      txnDate: row.txnDate,
      valueDate: row.txnDate,
      reference: "",
      description: "Test transaction",
      debit: row.debit || 0,
      credit: row.credit || 0,
      balance: row.balance || 0,
      sourceFile: "test-fixture.csv",
      statementOrder: index,
    }));
    applyFilters();
    renderAll();
  }, rows);
}

test.describe("Net Cash Flow popup", () => {
  test("Net Cash Flow and Current Balance sparklines plot one point per month", async ({ page }) => {
    await freezeDate(page, "2026-07-06T12:00:00Z");
    await page.goto("/");

    await seedTransactions(page, [
      { txnDate: "2026-01-15", accountNumber: "0001", debit: 1000, credit: 0, balance: 9000 },
      { txnDate: "2026-01-15", accountNumber: "0002", debit: 500, credit: 0, balance: 4500 },
      { txnDate: "2026-02-10", accountNumber: "0001", debit: 0, credit: 3000, balance: 12000 },
      { txnDate: "2026-02-10", accountNumber: "0002", debit: 0, credit: 1000, balance: 5500 },
    ]);

    const netCashFlowTile = page.locator('.metric-tile[data-tile="net-cash-flow"]');
    const netPoints = await netCashFlowTile.locator("polyline").getAttribute("points");
    expect(netPoints.trim().split(/\s+/)).toHaveLength(2);

    const currentBalanceTile = page.locator(".metric-tile", { hasText: "Current Balance" });
    const balancePoints = await currentBalanceTile.locator("polyline").getAttribute("points");
    expect(balancePoints.trim().split(/\s+/)).toHaveLength(2);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx playwright test tests/cash-flow-popup.spec.js`
Expected: FAIL — both assertions receive length 4 (this was verified directly: with 2 accounts × 2 months seeded, the current code renders 4 sparkline points instead of 2 for both the Net Cash Flow and Current Balance tiles). Also `[data-tile="net-cash-flow"]` matches nothing yet, since that attribute doesn't exist until Step 3.

- [ ] **Step 3: Add `computeMonthlyCashFlowSeries` and refactor `renderMonthlySummary`**

In `app.js`, insert this new function immediately before `function renderMonthlySummary() {` (~line 1516):

```js
function computeMonthlyCashFlowSeries(transactions) {
  const monthTotals = new Map();
  transactions.forEach((txn) => {
    const month = txn.txnDate.slice(0, 7);
    const current = monthTotals.get(month) || { month, debit: 0, credit: 0 };
    current.debit += toInsightAmount(txn.debit, txn.currency);
    current.credit += toInsightAmount(txn.credit, txn.currency);
    monthTotals.set(month, current);
  });
  return Array.from(monthTotals.values()).sort((a, b) => a.month.localeCompare(b.month));
}
```

Inside `renderMonthlySummary`, replace this block:

```js
  const monthlyRows = Array.from(summaryMap.values())
    .sort((a, b) => a.month.localeCompare(b.month) || a.accountLabel.localeCompare(b.accountLabel));
  const monthTotals = Array.from(summaryMap.values()).reduce((map, row) => {
    const current = map.get(row.month) || { debit: 0, credit: 0 };
    current.debit += row.totalDebit;
    current.credit += row.totalCredit;
    map.set(row.month, current);
    return map;
  }, new Map());
  const monthlySeries = Array.from(monthTotals.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  const monthlyDebitSeries = monthlySeries.map(([_, value]) => value.debit);
  const monthlyCreditSeries = monthlySeries.map(([_, value]) => value.credit);
  const avgMonthlyDebit = monthlyDebitSeries.length ? average(monthlyDebitSeries) : 0;
  const avgMonthlyCredit = monthlyCreditSeries.length ? average(monthlyCreditSeries) : 0;
  const currentBalance = Array.from(latestBalanceByAccount.values())
    .reduce((sum, txn) => sum + toInsightAmount(txn.balance, txn.currency), 0);
  const summaryTiles = [
    {
      label: "Debit",
      value: moneyFormat(totalDebit),
      subtext: `Avg ${moneyFormat(avgMonthlyDebit)} per month`,
      toneClass: "is-negative",
      sparkValues: monthlyDebitSeries,
    },
    {
      label: "Credit",
      value: moneyFormat(totalCredit),
      subtext: `Avg ${moneyFormat(avgMonthlyCredit)} per month`,
      toneClass: "is-positive",
      sparkValues: monthlyCreditSeries,
    },
    {
      label: "Net Cash Flow",
      value: moneyFormat(totalCredit - totalDebit),
      subtext: "Credit minus debit",
      toneClass: totalCredit - totalDebit >= 0 ? "is-positive" : "is-negative",
      sparkValues: monthlyRows.map((row) => row.totalCredit - row.totalDebit),
    },
    {
      label: "Current Balance",
      value: moneyFormat(currentBalance),
      subtext: "Latest visible balance by account",
      toneClass: currentBalance >= 0 ? "is-positive" : "is-negative",
      sparkValues: buildRunningSeries(monthlyRows.map((row) => row.totalCredit - row.totalDebit)),
    },
  ];

  els.monthlySummaryMetrics.innerHTML = summaryTiles.map((metric) => `
    <article class="metric-tile metric-tile-inline ${metric.toneClass}">
      <div class="metric-tile-body">
        <div class="metric-copy">
          <div class="metric-label">${escapeHtml(metric.label)}</div>
          <div class="metric-value">${escapeHtml(metric.value)}</div>
          <div class="metric-subtext">${escapeHtml(metric.subtext)}</div>
        </div>
        <div class="metric-spark-wrap">
          ${renderSparkline(metric.sparkValues, metric.toneClass)}
        </div>
      </div>
    </article>
  `).join("");
```

with:

```js
  const monthlyRows = Array.from(summaryMap.values())
    .sort((a, b) => a.month.localeCompare(b.month) || a.accountLabel.localeCompare(b.accountLabel));
  const monthlyCashFlowSeries = computeMonthlyCashFlowSeries(state.filteredTransactions);
  const monthlyDebitSeries = monthlyCashFlowSeries.map((row) => row.debit);
  const monthlyCreditSeries = monthlyCashFlowSeries.map((row) => row.credit);
  const monthlyNetSeries = monthlyCashFlowSeries.map((row) => row.credit - row.debit);
  const avgMonthlyDebit = monthlyDebitSeries.length ? average(monthlyDebitSeries) : 0;
  const avgMonthlyCredit = monthlyCreditSeries.length ? average(monthlyCreditSeries) : 0;
  const currentBalance = Array.from(latestBalanceByAccount.values())
    .reduce((sum, txn) => sum + toInsightAmount(txn.balance, txn.currency), 0);
  const summaryTiles = [
    {
      label: "Debit",
      value: moneyFormat(totalDebit),
      subtext: `Avg ${moneyFormat(avgMonthlyDebit)} per month`,
      toneClass: "is-negative",
      sparkValues: monthlyDebitSeries,
    },
    {
      label: "Credit",
      value: moneyFormat(totalCredit),
      subtext: `Avg ${moneyFormat(avgMonthlyCredit)} per month`,
      toneClass: "is-positive",
      sparkValues: monthlyCreditSeries,
    },
    {
      label: "Net Cash Flow",
      tileKey: "net-cash-flow",
      value: moneyFormat(totalCredit - totalDebit),
      subtext: "Credit minus debit",
      toneClass: totalCredit - totalDebit >= 0 ? "is-positive" : "is-negative",
      sparkValues: monthlyNetSeries,
    },
    {
      label: "Current Balance",
      value: moneyFormat(currentBalance),
      subtext: "Latest visible balance by account",
      toneClass: currentBalance >= 0 ? "is-positive" : "is-negative",
      sparkValues: buildRunningSeries(monthlyNetSeries),
    },
  ];

  els.monthlySummaryMetrics.innerHTML = summaryTiles.map((metric) => `
    <article class="metric-tile metric-tile-inline ${metric.toneClass}"${metric.tileKey ? ` data-tile="${metric.tileKey}"` : ""}>
      <div class="metric-tile-body">
        <div class="metric-copy">
          <div class="metric-label">${escapeHtml(metric.label)}</div>
          <div class="metric-value">${escapeHtml(metric.value)}</div>
          <div class="metric-subtext">${escapeHtml(metric.subtext)}</div>
        </div>
        <div class="metric-spark-wrap">
          ${renderSparkline(metric.sparkValues, metric.toneClass)}
        </div>
      </div>
    </article>
  `).join("");
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx playwright test tests/cash-flow-popup.spec.js`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add app.js tests/cash-flow-popup.spec.js
git commit -m "$(cat <<'EOF'
Fix Net Cash Flow/Current Balance sparklines to plot one point per month

Extracts computeMonthlyCashFlowSeries() as the single source of truth
for monthly credit/debit totals. The two tiles previously read from
monthlyRows (per account-month pairs), so with multiple accounts they
plotted duplicate points per month instead of one aggregated point.
EOF
)"
```

---

### Task 2: Net Cash Flow popup — modal, live chart, open/close, keyboard access

**Files:**
- Modify: `index.html` (insert after the `#import-panel` block, ~line 273)
- Modify: `styles.css` (add legend/line/point/clickable-tile classes)
- Modify: `app.js` — `state` (~line 160-188), `cacheElements` (~line 203-249), `bindEvents` (~line 251-348), `renderMonthlySummary`'s tile markup (from Task 1), `renderAll` (~line 1194-1207), plus new `openCashFlowModal`/`closeCashFlowModal`/`renderCashFlowChart` functions after `renderMonthlySummary`
- Test: `tests/cash-flow-popup.spec.js` (append)

**Interfaces:**
- Consumes: `computeMonthlyCashFlowSeries(transactions)` from Task 1, `formatMonthShort(monthKey)` (existing, `app.js:3538`), `compactMoneyFormat`/`moneyFormat`/`escapeHtml` (existing).
- Consumes: `data-tile="net-cash-flow"` attribute from Task 1 for click/keydown delegation.
- Produces: `state.cashFlowModalOpen: boolean`, `openCashFlowModal()`, `closeCashFlowModal()`, `renderCashFlowChart()` — no later task depends on these, but `renderAll()` calls `renderCashFlowChart()` conditionally.

- [ ] **Step 1: Write the failing tests**

Append to `tests/cash-flow-popup.spec.js`, inside the existing `test.describe("Net Cash Flow popup", () => { ... })` block (add after the Task 1 test, before the closing `});`):

```js
  test("clicking Net Cash Flow tile opens a 3-line chart popup with correct tooltips, and it closes via button and backdrop", async ({ page }) => {
    await freezeDate(page, "2026-07-06T12:00:00Z");
    await page.goto("/");

    await seedTransactions(page, [
      { txnDate: "2026-01-15", accountNumber: "0001", debit: 1000, credit: 3000, balance: 9000 },
      { txnDate: "2026-02-10", accountNumber: "0001", debit: 2000, credit: 5000, balance: 12000 },
    ]);

    const modal = page.locator("#cash-flow-modal");
    await expect(modal).toBeHidden();

    await page.locator('.metric-tile[data-tile="net-cash-flow"]').click();
    await expect(modal).toBeVisible();

    const chart = page.locator("#cash-flow-chart");
    await expect(chart.locator("path.cash-flow-credit-line")).toHaveCount(1);
    await expect(chart.locator("path.cash-flow-debit-line")).toHaveCount(1);
    await expect(chart.locator("path.cash-flow-net-line")).toHaveCount(1);
    await expect(chart.locator("circle")).toHaveCount(6);

    const titles = await chart.locator("circle title").allTextContents();
    expect(titles).toContain("Jan 2026 Credit: 3,000");
    expect(titles).toContain("Jan 2026 Debit: 1,000");
    expect(titles).toContain("Feb 2026 Net Cash Flow: 3,000");

    await page.locator("#close-cash-flow-modal").click();
    await expect(modal).toBeHidden();

    await page.locator('.metric-tile[data-tile="net-cash-flow"]').click();
    await expect(modal).toBeVisible();
    await page.locator("#cash-flow-modal .overlay-backdrop").click();
    await expect(modal).toBeHidden();
  });

  test("cash flow popup updates live when filters change while open", async ({ page }) => {
    await freezeDate(page, "2026-07-06T12:00:00Z");
    await page.goto("/");

    await seedTransactions(page, [
      { txnDate: "2026-01-15", accountNumber: "0001", debit: 1000, credit: 3000, balance: 9000 },
      { txnDate: "2026-02-10", accountNumber: "0002", debit: 2000, credit: 5000, balance: 12000 },
    ]);

    await page.locator('.metric-tile[data-tile="net-cash-flow"]').click();
    await expect(page.locator("#cash-flow-chart circle")).toHaveCount(6);

    await page.locator("#account-filter").selectOption({ label: "SBM 0001 (MUR)" });
    await expect(page.locator("#cash-flow-chart circle")).toHaveCount(3);
    const titles = await page.locator("#cash-flow-chart circle title").allTextContents();
    expect(titles).toContain("Jan 2026 Credit: 3,000");
  });

  test("Net Cash Flow tile can be opened with the keyboard", async ({ page }) => {
    await freezeDate(page, "2026-07-06T12:00:00Z");
    await page.goto("/");

    await seedTransactions(page, [
      { txnDate: "2026-01-15", accountNumber: "0001", debit: 1000, credit: 3000, balance: 9000 },
    ]);

    await page.locator('.metric-tile[data-tile="net-cash-flow"]').focus();
    await page.keyboard.press("Enter");
    await expect(page.locator("#cash-flow-modal")).toBeVisible();
  });
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx playwright test tests/cash-flow-popup.spec.js`
Expected: FAIL — `#cash-flow-modal` doesn't exist yet, so every new test times out locating it.

- [ ] **Step 3: Add the modal markup to `index.html`**

Insert immediately after the closing `</div>` of `#import-panel` (~line 273), before `<script src="./app.js"></script>`:

```html
  <div id="cash-flow-modal" class="overlay-panel" hidden>
    <div class="overlay-backdrop" data-close-cash-flow-modal></div>
    <section class="glass-card overlay-card" role="dialog" aria-modal="true" aria-labelledby="cash-flow-modal-title">
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
      <div class="trendline-chart-wrap">
        <svg id="cash-flow-chart" class="trendline-chart cash-flow-chart" viewBox="0 0 960 360" role="img" aria-label="Monthly credit, debit, and net cash flow"></svg>
      </div>
    </section>
  </div>
```

- [ ] **Step 4: Add CSS for the legend, chart lines/points, and clickable tile**

Append to `styles.css` (after the `.trend-empty` block, ~line 1863):

```css
.cash-flow-chart {
  min-height: 360px;
}

.cash-flow-legend {
  display: flex;
  align-items: center;
  gap: 18px;
  margin-bottom: 14px;
  color: var(--muted-soft);
  font-size: 0.86rem;
  font-weight: 600;
}

.legend-swatch {
  display: inline-block;
  width: 12px;
  height: 12px;
  border-radius: 999px;
  margin-right: 6px;
}

.legend-credit {
  background: var(--primary);
}

.legend-debit {
  background: var(--danger);
}

.legend-net {
  background: var(--gold);
}

.cash-flow-credit-line {
  fill: none;
  stroke: var(--primary);
  stroke-width: 3;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.cash-flow-debit-line {
  fill: none;
  stroke: var(--danger);
  stroke-width: 3;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.cash-flow-net-line {
  fill: none;
  stroke: var(--gold);
  stroke-width: 3;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.cash-flow-point-credit {
  stroke: var(--primary);
}

.cash-flow-point-debit {
  stroke: var(--danger);
}

.cash-flow-point-net {
  stroke: var(--gold);
}

.metric-tile-clickable {
  cursor: pointer;
}

.metric-tile-clickable:focus-visible {
  outline: 2px solid var(--primary);
  outline-offset: 2px;
}
```

- [ ] **Step 5: Make the Net Cash Flow tile clickable and keyboard-focusable**

In `app.js`, inside `renderMonthlySummary` (edited by Task 1), replace:

```js
  els.monthlySummaryMetrics.innerHTML = summaryTiles.map((metric) => `
    <article class="metric-tile metric-tile-inline ${metric.toneClass}"${metric.tileKey ? ` data-tile="${metric.tileKey}"` : ""}>
```

with:

```js
  els.monthlySummaryMetrics.innerHTML = summaryTiles.map((metric) => `
    <article class="metric-tile metric-tile-inline ${metric.toneClass}${metric.tileKey === "net-cash-flow" ? " metric-tile-clickable" : ""}"${metric.tileKey ? ` data-tile="${metric.tileKey}"` : ""}${metric.tileKey === "net-cash-flow" ? ' role="button" tabindex="0" aria-haspopup="dialog"' : ""}>
```

- [ ] **Step 6: Add `cashFlowModalOpen` to state**

In `app.js`, in the `state` initializer, replace:

```js
  shouldScrollNextTaxReceipt: true,
};
```

with:

```js
  shouldScrollNextTaxReceipt: true,
  cashFlowModalOpen: false,
};
```

- [ ] **Step 7: Cache the new DOM elements**

In `app.js`, in `cacheElements()`, immediately after the line `els.categoryBarList = document.getElementById("category-bar-list");`, add:

```js
  els.cashFlowModal = document.getElementById("cash-flow-modal");
  els.closeCashFlowModal = document.getElementById("close-cash-flow-modal");
  els.cashFlowChart = document.getElementById("cash-flow-chart");
```

- [ ] **Step 8: Add open/close functions and the chart renderer**

In `app.js`, immediately after `renderMonthlySummary`'s closing `}` (before `function renderTrendInsights() {`), add:

```js
function openCashFlowModal() {
  state.cashFlowModalOpen = true;
  els.cashFlowModal.hidden = false;
  document.body.classList.add("overlay-open");
  renderCashFlowChart();
}

function closeCashFlowModal() {
  state.cashFlowModalOpen = false;
  els.cashFlowModal.hidden = true;
  document.body.classList.remove("overlay-open");
}

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
        <text class="trend-label" x="${xForIndex(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(monthLabel(point.month))}</text>
      </g>
    `).join("")}
  `;
}
```

- [ ] **Step 9: Wire click/keydown to open, and button/backdrop to close**

In `app.js`, in `bindEvents()`, immediately before the function's closing `}` (after the `els.taxForecastYearSelect.addEventListener("change", ...)` block), add:

```js
  els.monthlySummaryMetrics.addEventListener("click", (event) => {
    const tile = event.target.closest('[data-tile="net-cash-flow"]');
    if (tile) {
      openCashFlowModal();
    }
  });
  els.monthlySummaryMetrics.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    const tile = event.target.closest('[data-tile="net-cash-flow"]');
    if (!tile) return;
    event.preventDefault();
    openCashFlowModal();
  });
  els.closeCashFlowModal.addEventListener("click", closeCashFlowModal);
  els.cashFlowModal.addEventListener("click", (event) => {
    if (event.target instanceof HTMLElement && event.target.hasAttribute("data-close-cash-flow-modal")) {
      closeCashFlowModal();
    }
  });
```

- [ ] **Step 10: Make `renderAll` keep the chart live while the popup is open**

In `app.js`, replace:

```js
function renderAll() {
  refreshAutoTaxExpectedExpenses();
  renderLedgerStatus();
  renderFilterOptions();
  renderQuickFilterChips();
  renderActiveFilterSummary();
  renderMetrics();
  renderTransactionsTable();
  renderMonthlySummary();
  renderTrendInsights();
  renderTaxActionQueue();
  renderTaxTable();
  renderTaxSummary();
}
```

with:

```js
function renderAll() {
  refreshAutoTaxExpectedExpenses();
  renderLedgerStatus();
  renderFilterOptions();
  renderQuickFilterChips();
  renderActiveFilterSummary();
  renderMetrics();
  renderTransactionsTable();
  renderMonthlySummary();
  renderTrendInsights();
  renderTaxActionQueue();
  renderTaxTable();
  renderTaxSummary();
  if (state.cashFlowModalOpen) {
    renderCashFlowChart();
  }
}
```

- [ ] **Step 11: Run tests to verify they pass**

Run: `npx playwright test tests/cash-flow-popup.spec.js`
Expected: PASS (all 4 tests, including Task 1's)

- [ ] **Step 12: Run the full existing suite to check for regressions**

Run: `npx playwright test`
Expected: PASS (no regressions in `tests/dashboard-smoke.spec.js`)

- [ ] **Step 13: Commit**

```bash
git add app.js index.html styles.css tests/cash-flow-popup.spec.js
git commit -m "$(cat <<'EOF'
Add Net Cash Flow popup with live-updating credit/debit/net chart

Clicking the Net Cash Flow tile (mouse or keyboard) opens a large
popup with a 3-line SVG chart of monthly Credit, Debit, and Net Cash
Flow, driven by the same filtered data as the tiles. Filter changes
made while the popup is open re-render it live via renderAll().
EOF
)"
```

---

## Self-Review

**Spec coverage:**
- Only Net Cash Flow tile clickable → Task 1 Step 3 (`tileKey` only on that tile), Task 2 Step 5 (clickable class/role only on that tile). ✓
- 3-line chart (credit/debit/net) with legend + hover tooltips → Task 2 Steps 3-4 (legend markup/CSS), Step 8 (`renderCashFlowChart`, native `<title>` tooltips). ✓
- Driven by current filter → both tasks read `state.filteredTransactions`. ✓
- Live update while open → Task 2 Step 10 (`renderAll` conditional call) + Step 1's live-update test. ✓
- Hand-rolled SVG, no new dependencies → confirmed in Global Constraints; no `package.json` changes anywhere in the plan. ✓
- Close via button/backdrop, no Escape handler (per spec's explicit "Out of scope") → Task 2 Step 9 only wires click handlers, matching `#import-panel`'s existing behavior. ✓

**Placeholder scan:** No TBD/TODO markers; every step has complete code or an exact command with expected output.

**Type consistency:** `computeMonthlyCashFlowSeries` return shape (`{ month, debit, credit }`) is used identically in Task 1 (`renderMonthlySummary`) and Task 2 (`renderCashFlowChart`). `data-tile="net-cash-flow"` string is consistent across the tile markup, click/keydown delegation, and every test selector.

**Scope check:** Single cohesive feature (one clickable tile, one popup), two tasks, both directly required — appropriately sized for one plan.
