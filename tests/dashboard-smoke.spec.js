const { test, expect } = require("@playwright/test");

test("dashboard loads core sections", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByRole("heading", { name: "At a Glance", exact: true })).toBeVisible();
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
