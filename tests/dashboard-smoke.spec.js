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

test("dashboard loads core sections", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByRole("heading", { name: "At a Glance", exact: true })).toBeVisible();
  await expect(page.locator(".metrics-card #tax-action-queue")).toBeVisible();
  await expect(page.getByRole("heading", { name: "Action Queue", exact: true })).toHaveCount(0);
  await expect(page.getByRole("heading", { name: "Options", exact: true })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Transactions Table", exact: true })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Monthly Cash Flow", exact: true })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Year-End Outlook", exact: true })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Tracker", exact: true })).toBeVisible();
});

test("tax forecast shows dynamic tax year wording", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByText(/Target saved amount for Jul \d{4} to Jun \d{4}/)).toBeVisible();
  await expect(page.getByText(/Auto uses expenses so far from Jul \d{4} to Jun \d{4}/)).toBeVisible();
});

test("add row prefills next invoice date, GBP rate, and salary amount", async ({ page }) => {
  await page.route("https://open.er-api.com/v6/latest/GBP", async (route) => {
    await route.fulfill({
      status: 200,
      contentType: "application/json",
      body: JSON.stringify({
        result: "success",
        rates: { MUR: 61.97 },
        time_last_update_utc: "Mon, 23 Mar 2026 00:00:01 +0000",
      }),
    });
  });

  await page.goto("/");

  const rows = page.locator("#tax-body tr");
  const initialCount = await rows.count();

  await page.getByRole("button", { name: "Add Row", exact: true }).click();

  await expect(rows).toHaveCount(initialCount + 1);

  const newRow = rows.nth(initialCount);
  await expect(newRow.locator('[data-field="invoicedDate"]')).toHaveValue("2026-07-31");
  await expect(newRow.locator('[data-field="dateReceived"]')).toHaveValue("");
  await expect(newRow.locator('[data-field="dateCsgPaid"]')).toHaveValue("");
  await expect(newRow.locator('[data-field="exchangeRate"]')).toHaveValue("61.97");
  await expect(newRow.locator('[data-field="amountReceivedGbp"]')).toHaveValue("8050");
});

test("next pending received date is highlighted and scrolled into view", async ({ page }) => {
  await page.goto("/");

  const nextReceiptInput = page.locator('[data-next-tax-receipt-input="true"]');
  await expect(nextReceiptInput).toHaveValue("");
  await expect(nextReceiptInput).toBeVisible();
  await expect(nextReceiptInput).toHaveAttribute("data-index", "8");

  const taxTableWrap = page.locator(".tax-table-wrap");
  await expect.poll(async () => taxTableWrap.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
});

test("tax action queue shows pending receipt follow-up on load", async ({ page }) => {
  await freezeDate(page, "2026-04-02T12:00:00Z");
  await page.goto("/");

  const taxActionQueue = page.locator("#tax-action-queue");
  await expect(taxActionQueue).toContainText("Confirm receipt date");
  await expect(taxActionQueue).toContainText("Expected receipt 27 Mar 2026");
});

test("tax action queue flags overdue CSG work after payment details are cleared", async ({ page }) => {
  await freezeDate(page, "2026-04-02T12:00:00Z");
  await page.goto("/");

  const taxActionQueue = page.locator("#tax-action-queue");
  await page.locator('[data-index="7"][data-field="dateCsgPaid"]').fill("");
  await page.locator('[data-index="7"][data-field="csgPaymentReference"]').fill("");

  await expect(taxActionQueue).toContainText("CSG payment overdue");
  await expect(taxActionQueue).toContainText("Due 31 Mar 2026");

  await taxActionQueue.locator('[data-tax-queue-jump="7"]').click();
  await expect(page.locator('[data-index="7"][data-field="dateCsgPaid"]')).toBeFocused();
});
