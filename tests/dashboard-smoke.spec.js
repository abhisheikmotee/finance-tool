const { test, expect } = require("@playwright/test");

test("dashboard loads core sections", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByRole("heading", { name: "At a Glance" })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Transactions" })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Transactions Table" })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Monthly Cash Flow" })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Year-End Outlook" })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Tracker" })).toBeVisible();
});

test("tax forecast shows dynamic tax year wording", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByText(/Target saved amount for Jul \d{4} to Jun \d{4}/)).toBeVisible();
  await expect(page.getByText(/Auto uses expenses so far from Jul \d{4} to Jun \d{4}/)).toBeVisible();
});
