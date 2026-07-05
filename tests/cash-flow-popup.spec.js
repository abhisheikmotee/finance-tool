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
