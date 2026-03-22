const { test, expect } = require("@playwright/test");

test("dashboard loads core sections", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByRole("heading", { name: "At a Glance", exact: true })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Transactions", exact: true })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Categorization Inbox", exact: true })).toBeVisible();
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
