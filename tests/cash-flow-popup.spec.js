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
    await page.locator("#cash-flow-modal .overlay-backdrop").click({ position: { x: 10, y: 10 } });
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
});
