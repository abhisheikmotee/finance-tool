const DB_NAME = "finance-ledger-browser-db";
const DB_VERSION = 2;
const SETTINGS_STORE = "settings";
const GBP_TO_MUR_RATE_KEY = "gbpToMurRate";
const GBP_TO_MUR_API_URL = "https://open.er-api.com/v6/latest/GBP";
const TAX_ENTRIES_KEY = "taxEntries";
const TAX_EXPECTED_EXPENSES_KEY = "taxExpectedExpenses";
const TAX_EXPECTED_EXPENSES_MODE_KEY = "taxExpectedExpensesMode";
const DEFAULT_EXPECTED_EXPENSES = 2500000;
const AUTO_TAX_EXPENSES_MODE = "auto";
const MANUAL_TAX_EXPENSES_MODE = "manual";
const DAILY_SALARY_RATE_GBP = 350;
const MAURITIUS_PUBLIC_HOLIDAYS = {
  2025: new Set([
    "2025-01-01",
    "2025-01-02",
    "2025-02-01",
    "2025-02-11",
    "2025-02-26",
    "2025-03-12",
    "2025-03-30",
    "2025-05-01",
    "2025-08-28",
    "2025-10-20",
    "2025-10-21",
    "2025-11-01",
    "2025-12-25",
  ]),
  2026: new Set([
    "2026-01-01",
    "2026-01-02",
    "2026-02-01",
    "2026-02-15",
    "2026-02-17",
    "2026-03-12",
    "2026-03-19",
    "2026-03-21",
    "2026-05-01",
    "2026-08-15",
    "2026-09-16",
    "2026-11-02",
    "2026-11-08",
    "2026-12-25",
  ]),
};
const FORECAST_SPEND_CATEGORIES = new Set([
  "Food Delivery",
  "Dining / Restaurants",
  "Shopping",
  "Insurance",
  "Transfers",
]);
const DEFAULT_TAX_ENTRIES = [
  {
    invoicedDate: "2025-06-30",
    dateReceived: "2025-07-29",
    dateCsgPaid: "2025-08-15",
    exchangeRate: "59.99",
    amountReceivedGbp: "7350",
    csgPaymentReference: "NIDM180785300014FFCSGP250780237",
  },
  {
    invoicedDate: "2025-07-31",
    dateReceived: "2025-08-27",
    dateCsgPaid: "2025-09-22",
    exchangeRate: "60.423",
    amountReceivedGbp: "8050",
    csgPaymentReference: "NIDM180785300014FFCSGP250892738",
  },
  {
    invoicedDate: "2025-08-29",
    dateReceived: "2025-09-26",
    dateCsgPaid: "2025-09-28",
    exchangeRate: "59.644",
    amountReceivedGbp: "7000",
    csgPaymentReference: "NIDM180785300014FFCSGP250981748",
  },
  {
    invoicedDate: "2025-09-30",
    dateReceived: "2025-10-29",
    dateCsgPaid: "2025-11-18",
    exchangeRate: "58.837",
    amountReceivedGbp: "7700",
    csgPaymentReference: "NIDM180785300014FFCSGP251066136",
  },
  {
    invoicedDate: "2025-10-31",
    dateReceived: "2025-11-26",
    dateCsgPaid: "2025-12-03",
    exchangeRate: "58.077",
    amountReceivedGbp: "8050",
    csgPaymentReference: "NIDM180785300014FFCSGP251186576",
  },
  {
    invoicedDate: "2025-11-28",
    dateReceived: "2025-12-29",
    dateCsgPaid: "2025-12-30",
    exchangeRate: "60.673",
    amountReceivedGbp: "7000",
    csgPaymentReference: "NIDM180785300014FFCSGP251225688",
  },
  {
    invoicedDate: "2025-12-31",
    dateReceived: "2026-01-28",
    dateCsgPaid: "2026-02-01",
    exchangeRate: "60.673",
    amountReceivedGbp: "5250",
    csgPaymentReference: "NIDM180785300014FFCSGP260149704",
  },
  {
    invoicedDate: "2026-01-30",
    dateReceived: "2026-02-25",
    dateCsgPaid: "2026-03-17",
    exchangeRate: "61.163",
    amountReceivedGbp: "7000",
    csgPaymentReference: "NIDM180785300014FFCSGP260220432",
  },
  {
    invoicedDate: "2026-02-27",
    dateReceived: "",
    dateCsgPaid: "",
    exchangeRate: "61.97",
    amountReceivedGbp: "6650",
    csgPaymentReference: "",
  },
  {
    invoicedDate: "2026-03-31",
    dateReceived: "",
    dateCsgPaid: "",
    exchangeRate: "61.97",
    amountReceivedGbp: "8050",
    csgPaymentReference: "",
  },
  {
    invoicedDate: "2026-04-30",
    dateReceived: "",
    dateCsgPaid: "",
    exchangeRate: "61.97",
    amountReceivedGbp: "7700",
    csgPaymentReference: "",
  },
  {
    invoicedDate: "2026-05-29",
    dateReceived: "",
    dateCsgPaid: "",
    exchangeRate: "61.97",
    amountReceivedGbp: "7000",
    csgPaymentReference: "",
  },
  {
    invoicedDate: "2026-06-30",
    dateReceived: "",
    dateCsgPaid: "",
    exchangeRate: "61.97",
    amountReceivedGbp: "7700",
    csgPaymentReference: "",
  },
];

const state = {
  transactions: [],
  filteredTransactions: [],
  sortedTransactions: [],
  imports: [],
  taxEntries: [],
  taxExpectedExpenses: DEFAULT_EXPECTED_EXPENSES,
  taxExpectedExpensesMode: AUTO_TAX_EXPENSES_MODE,
  currentPage: 1,
  pageSize: 10,
  ledgerHandle: null,
  ledgerName: "",
  saveMode: "download",
  gbpToMurRate: null,
  gbpToMurDate: "",
  quickFilters: {
    datePreset: "this-year",
    salaryOnly: false,
    tradingOnly: false,
    excludeOwnAccountTransfers: false,
    bankOnly: "all",
  },
  tableSort: {
    key: "txnDate",
    direction: "desc",
  },
  shouldScrollNextTaxReceipt: true,
};

const els = {};

document.addEventListener("DOMContentLoaded", async () => {
  cacheElements();
  bindEvents();
  await restoreLedgerHandle();
  await restoreExchangeRate();
  await restoreTaxEntries();
  await refreshExchangeRate();
  applyFilters();
  renderAll();
});

function cacheElements() {
  els.fileInput = document.getElementById("file-input");
  els.ledgerFileInput = document.getElementById("ledger-file-input");
  els.chooseLedgerBtn = document.getElementById("choose-ledger-btn");
  els.newLedgerBtn = document.getElementById("new-ledger-btn");
  els.saveLedgerBtn = document.getElementById("save-ledger-btn");
  els.dropZone = document.getElementById("drop-zone");
  els.importPanel = document.getElementById("import-panel");
  els.toggleImportPanel = document.getElementById("toggle-import-panel");
  els.closeImportPanel = document.getElementById("close-import-panel");
  els.ledgerStatus = document.getElementById("ledger-status");
  els.importLog = document.getElementById("import-log");
  els.metricsGrid = document.getElementById("metrics-grid");
  els.metricsLastImport = document.getElementById("metrics-last-import");
  els.trendMetricsGrid = document.getElementById("trend-metrics-grid");
  els.searchInput = document.getElementById("search-input");
  els.accountFilter = document.getElementById("account-filter");
  els.quickFilterChips = document.getElementById("quick-filter-chips");
  els.clearFiltersBtn = document.getElementById("clear-filters-btn");
  els.activeFilterSummary = document.getElementById("active-filter-summary");
  els.fromDate = document.getElementById("from-date");
  els.toDate = document.getElementById("to-date");
  els.pageSize = document.getElementById("page-size");
  els.transactionsBody = document.getElementById("transactions-body");
  els.monthlySummaryMetrics = document.getElementById("monthly-summary-metrics");
  els.monthlySummaryBody = document.getElementById("monthly-summary-body");
  els.categoryBarList = document.getElementById("category-bar-list");
  els.forecastSummaryBody = document.getElementById("forecast-summary-body");
  els.forecastSummaryCards = document.getElementById("forecast-summary-cards");
  els.trendlineChart = document.getElementById("trendline-chart");
  els.trendlineSummary = document.getElementById("trendline-summary");
  els.taxBody = document.getElementById("tax-body");
  els.taxSummaryBody = document.getElementById("tax-summary-body");
  els.addTaxRowBtn = document.getElementById("add-tax-row-btn");
  els.taxTableWrap = document.querySelector(".tax-table-wrap");
  els.tableExportBtn = document.getElementById("table-export-btn");
  els.tableCount = document.getElementById("table-count");
  els.prevPage = document.getElementById("prev-page");
  els.nextPage = document.getElementById("next-page");
  els.pageIndicator = document.getElementById("page-indicator");
  els.paginationBar = document.getElementById("pagination-bar");
  els.pageJumpInput = document.getElementById("page-jump-input");
  els.pageJumpBtn = document.getElementById("page-jump-btn");
}

function bindEvents() {
  els.chooseLedgerBtn.addEventListener("click", openExistingLedger);
  els.newLedgerBtn.addEventListener("click", createNewLedger);
  els.saveLedgerBtn.addEventListener("click", saveLedgerToDisk);
  els.ledgerFileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (file) {
      await loadLedgerFromFile(file);
    }
    event.target.value = "";
  });
  els.fileInput.addEventListener("change", async (event) => {
    await importFiles(Array.from(event.target.files || []));
    event.target.value = "";
  });
  els.toggleImportPanel.addEventListener("click", openImportPanel);
  els.closeImportPanel.addEventListener("click", closeImportPanel);
  els.importPanel.addEventListener("click", (event) => {
    if (event.target instanceof HTMLElement && event.target.hasAttribute("data-close-import-panel")) {
      closeImportPanel();
    }
  });

  ["dragenter", "dragover"].forEach((type) => {
    els.dropZone.addEventListener(type, (event) => {
      event.preventDefault();
      els.dropZone.classList.add("drag-over");
    });
  });

  ["dragleave", "drop"].forEach((type) => {
    els.dropZone.addEventListener(type, (event) => {
      event.preventDefault();
      els.dropZone.classList.remove("drag-over");
    });
  });

  els.dropZone.addEventListener("drop", async (event) => {
    const files = Array.from(event.dataTransfer?.files || []).filter((file) => /\.(csv|tsv)$/i.test(file.name));
    await importFiles(files);
  });

  [els.searchInput, els.accountFilter, els.fromDate, els.toDate].forEach((input) => {
    input.addEventListener("input", onFilterChange);
    input.addEventListener("change", onFilterChange);
  });

  els.quickFilterChips.addEventListener("click", handleQuickFilterChipClick);
  els.clearFiltersBtn.addEventListener("click", clearFilters);
  els.tableExportBtn.addEventListener("click", exportTransactionsCsv);
  document.querySelector(".table-card thead")?.addEventListener("click", handleTableSortClick);

  els.pageSize.addEventListener("change", () => {
    state.pageSize = Number(els.pageSize.value);
    state.currentPage = 1;
    renderTransactionsTable({ preservePaginationPosition: true });
  });

  els.prevPage.addEventListener("click", () => {
    if (state.currentPage > 1) {
      state.currentPage -= 1;
      renderTransactionsTable({ preservePaginationPosition: true });
    }
  });

  els.nextPage.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil(state.sortedTransactions.length / state.pageSize));
    if (state.currentPage < totalPages) {
      state.currentPage += 1;
      renderTransactionsTable({ preservePaginationPosition: true });
    }
  });
  els.pageJumpBtn.addEventListener("click", () => jumpToPage());
  els.pageJumpInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      jumpToPage();
    }
  });

  els.addTaxRowBtn.addEventListener("click", async () => {
    state.taxEntries.push(createPrefilledTaxEntry());
    state.shouldScrollNextTaxReceipt = true;
    renderTaxTable();
    renderTaxSummary();
    await persistTaxEntries();
  });

  els.taxBody.addEventListener("input", handleTaxTableInput);
  els.taxBody.addEventListener("click", handleTaxTableClick);
  els.taxSummaryBody.addEventListener("change", handleTaxSummaryInput);
}

function onFilterChange() {
  state.currentPage = 1;
  applyFilters();
  renderAll();
}

function clearFilters() {
  state.quickFilters = {
    datePreset: "this-year",
    salaryOnly: false,
    tradingOnly: false,
    excludeOwnAccountTransfers: false,
    bankOnly: "all",
  };
  els.searchInput.value = "";
  els.accountFilter.value = "all";
  els.fromDate.value = "";
  els.toDate.value = "";
  state.currentPage = 1;
  applyFilters();
  renderAll();
}

function handleQuickFilterChipClick(event) {
  const button = event.target.closest("[data-chip]");
  if (!button) return;

  const { chip } = button.dataset;
  if (chip === "this-month" || chip === "last-3-months" || chip === "this-year" || chip === "last-year" || chip === "current-tax-year") {
    state.quickFilters.datePreset = state.quickFilters.datePreset === chip ? "all" : chip;
  } else if (chip === "salary-only") {
    state.quickFilters.salaryOnly = !state.quickFilters.salaryOnly;
  } else if (chip === "trading-only") {
    state.quickFilters.tradingOnly = !state.quickFilters.tradingOnly;
  } else if (chip === "own-account-transfers-excluded") {
    state.quickFilters.excludeOwnAccountTransfers = !state.quickFilters.excludeOwnAccountTransfers;
  } else if (chip === "sbm-only" || chip === "mcb-only") {
    state.quickFilters.bankOnly = state.quickFilters.bankOnly === chip ? "all" : chip;
  }

  state.currentPage = 1;
  applyFilters();
  renderAll();
}

function handleTableSortClick(event) {
  const trigger = event.target.closest("[data-sort-key]");
  if (!trigger) return;

  const nextKey = trigger.getAttribute("data-sort-key");
  if (!nextKey) return;

  if (state.tableSort.key === nextKey) {
    state.tableSort.direction = state.tableSort.direction === "asc" ? "desc" : "asc";
  } else {
    state.tableSort.key = nextKey;
    state.tableSort.direction = nextKey === "txnDate" ? "desc" : "asc";
  }

  state.currentPage = 1;
  applyFilters();
  renderAll();
}

async function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(SETTINGS_STORE)) {
        db.createObjectStore(SETTINGS_STORE, { keyPath: "key" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getSetting(key) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SETTINGS_STORE, "readonly");
    const request = tx.objectStore(SETTINGS_STORE).get(key);
    request.onsuccess = () => resolve(request.result?.value ?? null);
    request.onerror = () => reject(request.error);
    tx.oncomplete = () => db.close();
  });
}

async function setSetting(key, value) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SETTINGS_STORE, "readwrite");
    tx.objectStore(SETTINGS_STORE).put({ key, value });
    tx.oncomplete = () => {
      db.close();
      resolve();
    };
    tx.onerror = () => {
      db.close();
      reject(tx.error);
    };
  });
}

async function deleteSetting(key) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SETTINGS_STORE, "readwrite");
    tx.objectStore(SETTINGS_STORE).delete(key);
    tx.oncomplete = () => {
      db.close();
      resolve();
    };
    tx.onerror = () => {
      db.close();
      reject(tx.error);
    };
  });
}

async function restoreLedgerHandle() {
  if (!supportsFileSystemAccess()) {
    state.saveMode = "download";
    updateLedgerStatus("Open and save a ledger JSON file from this page. No local server needed.");
    logMessage("Direct file mode active. Use Open to load a ledger JSON and Save to download updates.");
    return;
  }

  const storedHandle = await getSetting("ledgerHandle");
  if (!storedHandle) {
    updateLedgerStatus("Choose or create a ledger file to get started.");
    return;
  }

  try {
    const permission = await storedHandle.queryPermission({ mode: "readwrite" });
    if (permission !== "granted") {
      updateLedgerStatus("Ledger file remembered, but permission must be re-granted.");
      state.ledgerHandle = storedHandle;
      state.ledgerName = storedHandle.name || "";
      state.saveMode = "file-handle";
      return;
    }
    await loadLedgerFromHandle(storedHandle);
    logMessage(`Loaded ledger file: ${state.ledgerName}`);
  } catch (error) {
    updateLedgerStatus("Could not reopen the last ledger file. Choose it again.");
    logMessage(`Ledger reopen failed: ${error.message}`);
    await deleteSetting("ledgerHandle");
  }
}

async function restoreExchangeRate() {
  try {
    const savedRate = await getSetting(GBP_TO_MUR_RATE_KEY);
    if (savedRate && Number.isFinite(Number(savedRate.rate))) {
      state.gbpToMurRate = Number(savedRate.rate);
      state.gbpToMurDate = savedRate.date || "";
    }
  } catch (error) {
    logMessage(`Could not restore cached GBP/MUR rate: ${error.message}`);
  }
}

async function restoreTaxEntries() {
  try {
    const savedEntries = await getSetting(TAX_ENTRIES_KEY);
    if (Array.isArray(savedEntries)) {
      state.taxEntries = savedEntries.map(normalizeTaxEntry);
    } else {
      state.taxEntries = cloneDefaultTaxEntries();
      await persistTaxEntries();
    }
  } catch (error) {
    logMessage(`Could not restore tax entries: ${error.message}`);
  }

  state.shouldScrollNextTaxReceipt = true;

  try {
    const savedExpectedExpenses = await getSetting(TAX_EXPECTED_EXPENSES_KEY);
    const savedExpectedExpensesMode = await getSetting(TAX_EXPECTED_EXPENSES_MODE_KEY);
    state.taxExpectedExpensesMode = savedExpectedExpensesMode === MANUAL_TAX_EXPENSES_MODE
      ? MANUAL_TAX_EXPENSES_MODE
      : AUTO_TAX_EXPENSES_MODE;
    if (state.taxExpectedExpensesMode === MANUAL_TAX_EXPENSES_MODE && savedExpectedExpenses !== null && savedExpectedExpenses !== "") {
      state.taxExpectedExpenses = Number(savedExpectedExpenses) || DEFAULT_EXPECTED_EXPENSES;
    }
  } catch (error) {
    logMessage(`Could not restore tax expected expenses: ${error.message}`);
  }
}

async function refreshExchangeRate() {
  try {
    const response = await fetch(GBP_TO_MUR_API_URL, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const data = await response.json();
    if (data?.result !== "success") {
      throw new Error(data?.["error-type"] || "Unexpected exchange rate response");
    }

    const rate = Number(data?.rates?.MUR);
    if (!Number.isFinite(rate)) {
      throw new Error("GBP/MUR rate missing in API response");
    }

    state.gbpToMurRate = rate;
    state.gbpToMurDate = data.time_last_update_utc || "";
    await setSetting(GBP_TO_MUR_RATE_KEY, {
      rate,
      date: state.gbpToMurDate,
      fetchedAt: new Date().toISOString(),
    });
  } catch (error) {
    logMessage(`Could not refresh GBP/MUR rate: ${error.message}`);
  }
}

function supportsFileSystemAccess() {
  return typeof window.showOpenFilePicker === "function" && typeof window.showSaveFilePicker === "function";
}

async function createNewLedger() {
  state.transactions = [];
  state.imports = [];
  state.filteredTransactions = [];
  state.taxEntries = cloneDefaultTaxEntries();
  state.taxExpectedExpensesMode = AUTO_TAX_EXPENSES_MODE;
  state.taxExpectedExpenses = calculateSuggestedTaxExpectedExpenses(state.transactions);
  state.currentPage = 1;
  state.shouldScrollNextTaxReceipt = true;
  state.ledgerHandle = null;
  state.ledgerName = "finance-ledger.json";
  state.saveMode = supportsFileSystemAccess() ? "download" : "download";
  await persistTaxEntries();
  await persistTaxExpectedExpenses();
  applyFilters();
  renderAll();
  updateLedgerStatus(`Ledger file ready: ${state.ledgerName}. Click Save to download it.`);
  logMessage(`Started a new ledger: ${state.ledgerName}`);
}

async function openExistingLedger() {
  if (!supportsFileSystemAccess()) {
    els.ledgerFileInput.click();
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [{
        description: "Ledger JSON",
        accept: { "application/json": [".json"] },
      }],
    });
    await loadLedgerFromHandle(handle, true);
    const permissionGranted = await ensureReadWritePermission(handle, true);
    if (permissionGranted) {
      updateLedgerStatus(`Ledger file ready: ${state.ledgerName}`);
      logMessage(`Write access granted for ${state.ledgerName}. Save will overwrite this file.`);
    } else {
      updateLedgerStatus(`Ledger file ready: ${state.ledgerName}. Click Save to retry permission.`);
      logMessage(`Write access not granted yet for ${state.ledgerName}.`);
    }
    logMessage(`Opened ledger file: ${state.ledgerName}`);
  } catch (error) {
    if (error?.name !== "AbortError") {
      els.ledgerFileInput.click();
    }
  }
}

async function loadLedgerFromHandle(handle, storeHandle = false) {
  const file = await handle.getFile();
  const text = await file.text();
  const data = text.trim() ? JSON.parse(text) : emptyLedgerSnapshot();
  state.ledgerHandle = handle;
  state.ledgerName = handle.name || "";
  state.saveMode = "file-handle";
  state.transactions = normalizeTransactionRows(data.transactions || []);
  state.imports = (data.imports || []).sort((a, b) => a.importedAt.localeCompare(b.importedAt));
  state.taxEntries = Array.isArray(data.taxEntries) ? data.taxEntries.map(normalizeTaxEntry) : [];
  state.taxExpectedExpensesMode = data.taxExpectedExpensesMode === MANUAL_TAX_EXPENSES_MODE
    ? MANUAL_TAX_EXPENSES_MODE
    : AUTO_TAX_EXPENSES_MODE;
  state.taxExpectedExpenses = state.taxExpectedExpensesMode === MANUAL_TAX_EXPENSES_MODE
    ? (Number(data.taxExpectedExpenses) || DEFAULT_EXPECTED_EXPENSES)
    : calculateSuggestedTaxExpectedExpenses(state.transactions);
  state.shouldScrollNextTaxReceipt = true;
  await persistTaxEntries();
  await persistTaxExpectedExpenses();
  if (storeHandle) {
    await setSetting("ledgerHandle", handle);
  }
  updateLedgerStatus(`Ledger file ready: ${state.ledgerName}`);
  applyFilters();
  renderAll();
}

async function loadLedgerFromFile(file) {
  const text = await file.text();
  const data = text.trim() ? JSON.parse(text) : emptyLedgerSnapshot();
  state.ledgerHandle = null;
  state.ledgerName = file.name || "finance-ledger.json";
  state.saveMode = "download";
  state.transactions = normalizeTransactionRows(data.transactions || []);
  state.imports = (data.imports || []).sort((a, b) => a.importedAt.localeCompare(b.importedAt));
  state.taxEntries = Array.isArray(data.taxEntries) ? data.taxEntries.map(normalizeTaxEntry) : [];
  state.taxExpectedExpensesMode = data.taxExpectedExpensesMode === MANUAL_TAX_EXPENSES_MODE
    ? MANUAL_TAX_EXPENSES_MODE
    : AUTO_TAX_EXPENSES_MODE;
  state.taxExpectedExpenses = state.taxExpectedExpensesMode === MANUAL_TAX_EXPENSES_MODE
    ? (Number(data.taxExpectedExpenses) || DEFAULT_EXPECTED_EXPENSES)
    : calculateSuggestedTaxExpectedExpenses(state.transactions);
  state.shouldScrollNextTaxReceipt = true;
  await persistTaxEntries();
  await persistTaxExpectedExpenses();
  updateLedgerStatus(`Ledger file ready: ${state.ledgerName}. Save will download the updated file.`);
  applyFilters();
  renderAll();
  logMessage(`Opened ledger file: ${state.ledgerName}`);
}

async function saveLedgerToDisk() {
  const snapshot = JSON.stringify(buildLedgerSnapshot(), null, 2);

  if (state.ledgerHandle && state.saveMode === "file-handle") {
    try {
      const permission = await ensureReadWritePermission(state.ledgerHandle, true);
      if (!permission) {
        updateLedgerStatus("Ledger permission was not granted.");
        return;
      }

      const writable = await state.ledgerHandle.createWritable();
      await writable.write(snapshot);
      await writable.close();
      await setSetting("ledgerHandle", state.ledgerHandle);
      updateLedgerStatus(`Ledger file ready: ${state.ledgerName}. Changes saved.`);
      return;
    } catch (error) {
      updateLedgerStatus(`Could not save to ${state.ledgerName}.`);
      logMessage(`Direct save failed for ${state.ledgerName}: ${error.message}`);
      return;
    }
  }

  downloadLedgerSnapshot(snapshot, state.ledgerName || "finance-ledger.json");
  updateLedgerStatus(`Downloaded ${state.ledgerName || "finance-ledger.json"}. Replace your old ledger file with it.`);
}

async function ensureReadWritePermission(handle, requestIfNeeded = false) {
  if ((await handle.queryPermission({ mode: "readwrite" })) === "granted") return true;
  if (!requestIfNeeded) return false;
  return (await handle.requestPermission({ mode: "readwrite" })) === "granted";
}

function buildLedgerSnapshot() {
  return {
    version: 1,
    exportedAt: new Date().toISOString(),
    transactions: state.transactions,
    imports: state.imports,
    taxEntries: state.taxEntries,
    taxExpectedExpenses: state.taxExpectedExpenses,
    taxExpectedExpensesMode: state.taxExpectedExpensesMode,
  };
}

function emptyLedgerSnapshot() {
  return {
    version: 1,
    exportedAt: null,
    transactions: [],
    imports: [],
    taxEntries: cloneDefaultTaxEntries(),
    taxExpectedExpenses: DEFAULT_EXPECTED_EXPENSES,
    taxExpectedExpensesMode: AUTO_TAX_EXPENSES_MODE,
  };
}

function normalizeTransactionRows(rows) {
  return rows
    .map((row) => ({
      ...row,
      debit: Number(row.debit || 0),
      credit: Number(row.credit || 0),
      balance: Number(row.balance || 0),
    }))
    .sort((a, b) =>
      a.txnDate.localeCompare(b.txnDate) ||
      a.valueDate.localeCompare(b.valueDate) ||
      a.accountLabel.localeCompare(b.accountLabel) ||
      a.rowHash.localeCompare(b.rowHash)
    );
}

async function importFiles(files) {
  if (!files.length) {
    logMessage("No CSV files selected.");
    return;
  }

  if (!state.ledgerHandle && !state.ledgerName) {
    window.alert("Create a new ledger or open an existing ledger JSON file first.");
    return;
  }

  updateStatus("Importing...");
  const existingImportHashes = new Set(state.imports.map((item) => item.sourceHash));
  const existingRowHashes = new Set(state.transactions.map((item) => item.rowHash));
  const appendedTransactions = [];
  const appendedImports = [];

  for (const file of files) {
    try {
      const text = await file.text();
      const sourceHash = await stableHash(text);
      if (existingImportHashes.has(sourceHash)) {
        logMessage(`Skipped duplicate statement content: ${file.name}`);
        continue;
      }

      const parsed = await parseStatementFile(file.name, text);
      let addedForFile = 0;
      for (const txn of parsed.transactions) {
        if (existingRowHashes.has(txn.rowHash)) {
          continue;
        }
        existingRowHashes.add(txn.rowHash);
        appendedTransactions.push(txn);
        addedForFile += 1;
      }

      appendedImports.push({
        sourceHash,
        sourceFile: file.name,
        importedAt: new Date().toISOString(),
        detectedAccount: parsed.accountLabel,
        transactionsAdded: addedForFile,
      });
      existingImportHashes.add(sourceHash);
      logMessage(`Imported ${file.name} into ${parsed.accountLabel}. Added ${addedForFile} new transaction(s).`);
    } catch (error) {
      logMessage(`Failed to import ${file.name}: ${error.message}`);
    }
  }

  state.transactions = normalizeTransactionRows([...state.transactions, ...appendedTransactions]);
  state.imports = [...state.imports, ...appendedImports].sort((a, b) => a.importedAt.localeCompare(b.importedAt));
  applyFilters();
  renderAll();
  if (state.ledgerHandle && state.saveMode === "file-handle") {
    if (await ensureReadWritePermission(state.ledgerHandle)) {
      await saveLedgerToDisk();
    } else {
      updateLedgerStatus(`Imported changes are in memory for ${state.ledgerName}. Click Save to overwrite the ledger file.`);
      logMessage(`Imported data is pending save for ${state.ledgerName}. Manual Save is needed to write to disk.`);
    }
  } else {
    await saveLedgerToDisk();
  }
  updateStatus(`Imported ${appendedImports.length} file(s), added ${appendedTransactions.length} row(s).`);
}

async function parseStatementFile(fileName, text) {
  if (text.includes("Transaction Date,Value Date,Reference,Description,Money out,Money in,Balance")) {
    return parseMcbStatement(fileName, text);
  }
  if (text.includes("Instrument ID,Transaction Date,Value Date,Branch Code,Remarks,Debit Amount,Credit Amount,Balance")) {
    return parseSbmStatement(fileName, text);
  }
  if (text.includes("Txn Date\t Description - Instruction No.\t Debit\t Credit\t Running Balance")) {
    return parseLegacyMcbStatement(fileName, text);
  }
  if (text.includes("Txn Date\t Description - Instruction No.\t Debit\t Credit\t Balance  GBP")) {
    return parseLegacyMcbStatement(fileName, text);
  }
  if (text.includes("Txn Date\t Description / Instruction No.\t Debit\t Credit\t Running Balance")) {
    return parseLegacySbmStatement(fileName, text);
  }
  throw new Error("Unsupported CSV layout");
}

function parseMcbStatement(fileName, text) {
  const accountMatch = text.match(/Account Number\s+(\d+)/);
  const currencyMatch = text.match(/Account Currency\s+([A-Z]{3})/);
  if (!accountMatch || !currencyMatch) throw new Error("Could not detect MCB account details");

  const accountNumber = accountMatch[1];
  const currency = currencyMatch[1];
  const accountLabel = `MCB ${accountNumber} (${currency})`;
  const lines = text.split(/\r?\n/);
  const headerIndex = lines.findIndex((line) =>
    line.includes("Transaction Date,Value Date,Reference,Description,Money out,Money in,Balance")
  );
  const rows = parseCsvText(lines.slice(headerIndex).join("\n"));

  const transactions = rows.slice(1).filter((row) => row.some(Boolean)).map((row) => {
    return buildTransaction({
      bankName: "MCB",
      accountNumber,
      currency,
      txnDate: parseDate(row[0]),
      valueDate: parseDate(row[1]),
      reference: normalizeSpace(row[2]),
      description: normalizeSpace(row[3]),
      debit: parseAmount(row[4]),
      credit: parseAmount(row[5]),
      balance: parseAmount(row[6]),
      sourceFile: fileName,
    });
  });

  return { accountLabel, transactions };
}

function parseSbmStatement(fileName, text) {
  const accountMatch = text.match(/Account:,+\s*(\d+)\s*-\s*(\d+)/);
  if (!accountMatch) throw new Error("Could not detect SBM account details");

  const accountNumber = accountMatch[1];
  const currency = "MUR";
  const accountLabel = `SBM ${accountNumber} (${currency})`;
  const lines = text.split(/\r?\n/);
  const headerIndex = lines.findIndex((line) =>
    line.startsWith("Instrument ID,Transaction Date,Value Date,Branch Code,Remarks,Debit Amount,Credit Amount,Balance")
  );
  const rows = parseCsvText(lines.slice(headerIndex).join("\n"));

  const transactions = rows.slice(1).filter((row) => row.some(Boolean)).map((row) =>
    buildTransaction({
      bankName: "SBM",
      accountNumber,
      currency,
      txnDate: parseDate(row[1]),
      valueDate: parseDate(row[2]),
      reference: normalizeSpace(row[3]),
      description: normalizeSpace(row[4]),
      debit: parseAmount(row[5]),
      credit: parseAmount(row[6]),
      balance: parseAmount(row[7]),
      sourceFile: fileName,
    })
  );

  return { accountLabel, transactions };
}

function parseLegacyMcbStatement(fileName, text) {
  const metadata = detectLegacyAccountFromFileName(fileName);
  if (!metadata || metadata.bankName !== "MCB") {
    throw new Error("Could not detect legacy MCB account details from file name");
  }

  const rows = parseDelimitedText(text, "\t");
  const transactions = rows
    .slice(1)
    .filter((row) => row.some((cell) => (cell || "").trim()))
    .map((row) => buildTransaction({
      bankName: "MCB",
      accountNumber: metadata.accountNumber,
      currency: metadata.currency,
      txnDate: parseDate(row[0]),
      valueDate: parseDate(row[0]),
      reference: "",
      description: normalizeSpace(row[1]),
      debit: parseAmount(row[2]),
      credit: parseAmount(row[3]),
      balance: parseAmount(row[4]),
      sourceFile: fileName,
    }));

  return { accountLabel: `MCB ${metadata.accountNumber} (${metadata.currency})`, transactions };
}

function parseLegacySbmStatement(fileName, text) {
  const metadata = detectLegacyAccountFromFileName(fileName);
  if (!metadata || metadata.bankName !== "SBM") {
    throw new Error("Could not detect legacy SBM account details from file name");
  }

  const rows = parseDelimitedText(text, "\t");
  const transactions = rows
    .slice(1)
    .filter((row) => row.some((cell) => (cell || "").trim()))
    .map((row) => buildTransaction({
      bankName: "SBM",
      accountNumber: metadata.accountNumber,
      currency: metadata.currency,
      txnDate: parseDate(row[0]),
      valueDate: parseDate(row[0]),
      reference: "",
      description: normalizeSpace(row[1]),
      debit: parseAmount(row[2]),
      credit: parseAmount(row[3]),
      balance: parseAmount(row[4]),
      sourceFile: fileName,
    }));

  return { accountLabel: `SBM ${metadata.accountNumber} (${metadata.currency})`, transactions };
}

function buildTransaction(input) {
  const accountLabel = `${input.bankName} ${input.accountNumber} (${input.currency})`;
  const signature = [
    accountLabel,
    input.txnDate,
    input.valueDate,
    input.reference,
    input.description,
    input.debit.toFixed(2),
    input.credit.toFixed(2),
    input.balance.toFixed(2),
  ].join("||");

  return {
    bankName: input.bankName,
    accountNumber: input.accountNumber,
    currency: input.currency,
    accountLabel,
    txnDate: input.txnDate,
    valueDate: input.valueDate,
    reference: input.reference,
    description: input.description,
    debit: input.debit,
    credit: input.credit,
    balance: input.balance,
    sourceFile: input.sourceFile,
    rowHash: simpleHash(signature),
  };
}

function parseCsvText(text) {
  return parseDelimitedText(text, ",");
}

function parseDelimitedText(text, delimiter) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        cell += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

function parseAmount(raw) {
  const value = (raw || "").trim().replace(/,/g, "");
  return value ? Number(value) : 0;
}

function normalizeSpace(raw) {
  return (raw || "").replace(/\s+/g, " ").trim();
}

function parseDate(raw) {
  const value = (raw || "").trim();
  if (/^\d{2}-[A-Za-z]{3}-\d{4}$/.test(value)) {
    const [day, mon, year] = value.split("-");
    const month = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].indexOf(mon);
    return `${year}-${String(month + 1).padStart(2, "0")}-${day}`;
  }
  if (/^\d{1,2}-[A-Za-z]{3}-\d{2}$/.test(value)) {
    const [day, mon, year] = value.split("-");
    const month = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].indexOf(mon);
    const fullYear = Number(year) >= 70 ? `19${year}` : `20${year}`;
    return `${fullYear}-${String(month + 1).padStart(2, "0")}-${String(Number(day)).padStart(2, "0")}`;
  }
  if (/^\d{2}-\d{2}-\d{4}$/.test(value)) {
    const [day, month, year] = value.split("-");
    return `${year}-${month}-${day}`;
  }
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(value)) {
    const [day, month, year] = value.split("/");
    return `${year}-${month}-${day}`;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  throw new Error(`Unsupported date format: ${value}`);
}

function detectLegacyAccountFromFileName(fileName) {
  const normalized = (fileName || "").toUpperCase();
  if (normalized.includes("MCB_.000444921885")) {
    return { bankName: "MCB", accountNumber: "000444921885", currency: "MUR" };
  }
  if (normalized.includes("MCB_.000450553604")) {
    return { bankName: "MCB", accountNumber: "000450553604", currency: "GBP" };
  }
  if (normalized.includes("SBM_.01810100161110")) {
    return { bankName: "SBM", accountNumber: "01810100161110", currency: "MUR" };
  }
  return null;
}

function simpleHash(input) {
  let hash = 0;
  for (let i = 0; i < input.length; i += 1) {
    hash = (hash << 5) - hash + input.charCodeAt(i);
    hash |= 0;
  }
  return `row_${Math.abs(hash)}_${input.length}`;
}

async function stableHash(input) {
  if (window.crypto?.subtle) {
    const bytes = new TextEncoder().encode(input);
    const digest = await crypto.subtle.digest("SHA-256", bytes);
    return Array.from(new Uint8Array(digest)).map((item) => item.toString(16).padStart(2, "0")).join("");
  }
  return simpleHash(input);
}

function applyFilters() {
  const query = els.searchInput.value.trim().toLowerCase();
  const account = els.accountFilter.value || "all";
  const { fromDate, toDate } = getEffectiveDateRange();
  const { salaryOnly, tradingOnly, excludeOwnAccountTransfers, bankOnly } = state.quickFilters;

  state.filteredTransactions = state.transactions.filter((txn) => {
    const matchesQuery = !query || [txn.description, txn.reference, txn.sourceFile, txn.accountLabel]
      .some((value) => value.toLowerCase().includes(query));
    const matchesAccount = account === "all" || txn.accountLabel === account;
    const matchesFromDate = !fromDate || txn.txnDate >= fromDate;
    const matchesToDate = !toDate || txn.txnDate <= toDate;
    const matchesSalary = !salaryOnly || isSalaryTransaction(txn);
    const matchesTrading = !tradingOnly || isTradingTransaction(txn);
    const matchesOwnAccountTransfer = !excludeOwnAccountTransfers || !isOwnAccountTransferTransaction(txn);
    const matchesBank = bankOnly === "all"
      || (bankOnly === "sbm-only" && txn.bankName === "SBM")
      || (bankOnly === "mcb-only" && txn.bankName === "MCB");
    return matchesQuery
      && matchesAccount
      && matchesFromDate
      && matchesToDate
      && matchesSalary
      && matchesTrading
      && matchesOwnAccountTransfer
      && matchesBank;
  });
  state.sortedTransactions = sortTransactions(state.filteredTransactions);
}

function openImportPanel() {
  els.importPanel.hidden = false;
  document.body.classList.add("overlay-open");
}

function closeImportPanel() {
  els.importPanel.hidden = true;
  document.body.classList.remove("overlay-open");
}

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
  renderTaxTable();
  renderTaxSummary();
}

function renderQuickFilterChips() {
  const usingManualDates = Boolean(els.fromDate.value || els.toDate.value);
  const activeMap = {
    "this-month": !usingManualDates && state.quickFilters.datePreset === "this-month",
    "last-3-months": !usingManualDates && state.quickFilters.datePreset === "last-3-months",
    "this-year": !usingManualDates && state.quickFilters.datePreset === "this-year",
    "current-tax-year": !usingManualDates && state.quickFilters.datePreset === "current-tax-year",
    "last-year": !usingManualDates && state.quickFilters.datePreset === "last-year",
    "salary-only": state.quickFilters.salaryOnly,
    "trading-only": state.quickFilters.tradingOnly,
    "own-account-transfers-excluded": state.quickFilters.excludeOwnAccountTransfers,
    "sbm-only": state.quickFilters.bankOnly === "sbm-only",
    "mcb-only": state.quickFilters.bankOnly === "mcb-only",
  };

  els.quickFilterChips.querySelectorAll("[data-chip]").forEach((button) => {
    const isActive = Boolean(activeMap[button.dataset.chip]);
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });
}

function renderLedgerStatus() {
  if (state.ledgerName) {
    updateLedgerStatus(`Ledger file ready: ${state.ledgerName}`);
  } else if (supportsFileSystemAccess()) {
    updateLedgerStatus("Choose or create a ledger file to get started.");
  }
}

function renderFilterOptions() {
  const currentAccount = els.accountFilter.value || "all";
  const accounts = Array.from(new Set(state.transactions.map((txn) => txn.accountLabel))).sort();
  populateSelect(els.accountFilter, ["all", ...accounts], "All Accounts", currentAccount);
}

function populateSelect(select, values, allLabel, currentValue) {
  select.innerHTML = "";
  values.forEach((value, index) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = index === 0 ? allLabel : value;
    select.appendChild(option);
  });
  if (values.includes(currentValue)) {
    select.value = currentValue;
  }
}

function renderActiveFilterSummary() {
  const pills = [];
  const query = els.searchInput.value.trim();
  const account = els.accountFilter.value || "all";
  const { fromDate, toDate } = getEffectiveDateRange();

  if (query) pills.push({ label: "Search", value: query });
  if (account !== "all") pills.push({ label: "Account", value: account });
  if (fromDate || toDate) {
    pills.push({
      label: "Window",
      value: `${fromDate || "Start"} - ${toDate || "Now"}`,
    });
  }
  if (!els.fromDate.value && !els.toDate.value && state.quickFilters.datePreset !== "all") {
    pills.push({ label: "Preset", value: prettifyChipLabel(state.quickFilters.datePreset) });
  }
  if (state.quickFilters.salaryOnly) pills.push({ label: "Quick", value: "Salary only" });
  if (state.quickFilters.tradingOnly) pills.push({ label: "Quick", value: "Trading only" });
  if (state.quickFilters.excludeOwnAccountTransfers) pills.push({ label: "Quick", value: "Own account transfers excluded" });
  if (state.quickFilters.bankOnly === "sbm-only") pills.push({ label: "Bank", value: "SBM only" });
  if (state.quickFilters.bankOnly === "mcb-only") pills.push({ label: "Bank", value: "MCB only" });

  els.activeFilterSummary.hidden = pills.length === 0;
  els.activeFilterSummary.innerHTML = pills.map((pill) => `
    <div class="active-filter-pill"><strong>${escapeHtml(pill.label)}:</strong> ${escapeHtml(pill.value)}</div>
  `).join("");
}

function prettifyChipLabel(value) {
  return value
    .replace(/-/g, " ")
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function sortTransactions(rows) {
  const direction = state.tableSort.direction === "asc" ? 1 : -1;
  const picker = {
    txnDate: (txn) => txn.txnDate,
    accountLabel: (txn) => txn.accountLabel,
    description: (txn) => txn.description,
    debit: (txn) => txn.debit,
    credit: (txn) => txn.credit,
    balance: (txn) => txn.balance,
  }[state.tableSort.key] || ((txn) => txn.txnDate);

  return [...rows].sort((a, b) => {
    const aValue = picker(a);
    const bValue = picker(b);
    if (typeof aValue === "number" && typeof bValue === "number") {
      return direction * (aValue - bValue || a.txnDate.localeCompare(b.txnDate));
    }
    return direction * (
      String(aValue).localeCompare(String(bValue))
      || a.txnDate.localeCompare(b.txnDate)
      || a.accountLabel.localeCompare(b.accountLabel)
    );
  });
}

function renderMetrics() {
  const latestImport = [...state.imports].sort((a, b) => b.importedAt.localeCompare(a.importedAt))[0];
  const balances = summarizeAccountBalances(state.transactions);
  const gbpRate = getGbpToMurRate();
  const metrics = [
    {
      label: "SBM",
      value: moneyFormat(balances.sbmMur),
      subtext: "MUR balance",
    },
    {
      label: "MCB MUR",
      value: moneyFormat(balances.mcbMur),
      subtext: "MUR balance",
    },
    {
      label: "MCB GBP",
      value: moneyFormat(balances.mcbGbpInMur),
      subtext: `${numberFormat(balances.mcbGbp)} GBP (rate ${gbpRate.toFixed(2)})`,
    },
    {
      label: "Sum Of All",
      value: moneyFormat(balances.totalMur),
      subtext: "Total in MUR",
    },
  ];

  els.metricsGrid.innerHTML = metrics.map((metric) => `
    <article class="metric-tile">
      <div class="metric-label">${escapeHtml(metric.label)}</div>
      <div class="metric-value">${escapeHtml(metric.value)}</div>
      <div class="metric-subtext">${escapeHtml(metric.subtext)}</div>
    </article>
  `).join("");

  els.metricsLastImport.textContent = latestImport ? `Last import: ${formatDateTime(latestImport.importedAt)}` : "Last import: none";
}

function summarizeTransactions(transactions) {
  const accountSet = new Set();
  let totalDebit = 0;
  let totalCredit = 0;
  transactions.forEach((txn) => {
    accountSet.add(txn.accountLabel);
    totalDebit += toInsightAmount(txn.debit, txn.currency);
    totalCredit += toInsightAmount(txn.credit, txn.currency);
  });
  return {
    transactionCount: transactions.length,
    accountCount: accountSet.size,
    totalDebit,
    totalCredit,
  };
}

function summarizeAccountBalances(transactions) {
  const latestByAccount = new Map();
  transactions.forEach((txn) => {
    const existing = latestByAccount.get(txn.accountLabel);
    if (!existing || txn.txnDate >= existing.txnDate) {
      latestByAccount.set(txn.accountLabel, txn);
    }
  });

  let sbmMur = 0;
  let mcbMur = 0;
  let mcbGbp = 0;

  latestByAccount.forEach((txn) => {
    if (txn.bankName === "SBM" && txn.currency === "MUR") {
      sbmMur += txn.balance;
    } else if (txn.bankName === "MCB" && txn.currency === "MUR") {
      mcbMur += txn.balance;
    } else if (txn.bankName === "MCB" && txn.currency === "GBP") {
      mcbGbp += txn.balance;
    }
  });

  const mcbGbpInMur = toInsightAmount(mcbGbp, "GBP");
  return {
    sbmMur,
    mcbMur,
    mcbGbp,
    mcbGbpInMur,
    totalMur: sbmMur + mcbMur + mcbGbpInMur,
  };
}

function renderTransactionsTable(options = {}) {
  const { preservePaginationPosition = false } = options;
  const previousPaginationTop = preservePaginationPosition && els.paginationBar
    ? els.paginationBar.getBoundingClientRect().top
    : null;
  const totalRows = state.sortedTransactions.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / state.pageSize));
  if (state.currentPage > totalPages) state.currentPage = totalPages;

  const start = (state.currentPage - 1) * state.pageSize;
  const end = start + state.pageSize;
  const pageRows = state.sortedTransactions.slice(start, end);
  const totalLedgerRows = state.transactions.length;

  els.tableCount.textContent = `Showing ${numberFormat(pageRows.length)} of ${numberFormat(totalRows)} row${totalRows === 1 ? "" : "s"}`;
  els.tableCount.title = `${numberFormat(totalRows)} filtered row${totalRows === 1 ? "" : "s"} from ${numberFormat(totalLedgerRows)} total ledger row${totalLedgerRows === 1 ? "" : "s"}`;
  els.pageIndicator.textContent = `Page ${state.currentPage} of ${totalPages}`;
  els.prevPage.disabled = state.currentPage === 1;
  els.nextPage.disabled = state.currentPage === totalPages;
  els.pageJumpInput.max = String(totalPages);
  els.pageJumpInput.value = String(state.currentPage);
  renderTableSortState();

  if (!pageRows.length) {
    els.transactionsBody.innerHTML = `<tr><td colspan="6" class="empty-state">No transactions match the current filters. Try widening the date range or clearing bank and salary filters.</td></tr>`;
    restorePaginationPosition(previousPaginationTop);
    return;
  }

  els.transactionsBody.innerHTML = pageRows.map((txn) => `
    <tr>
      <td>${escapeHtml(txn.txnDate)}</td>
      <td>${renderAccountBadge(txn.accountLabel)}</td>
      <td>${escapeHtml(txn.description)}</td>
      <td class="${txn.debit > 0 ? "amount-negative" : ""}">${moneyFormat(txn.debit)}</td>
      <td class="${txn.credit > 0 ? "amount-positive" : ""}">${moneyFormat(txn.credit)}</td>
      <td>${moneyFormat(txn.balance)}</td>
    </tr>
  `).join("");

  restorePaginationPosition(previousPaginationTop);
}

function renderTableSortState() {
  document.querySelectorAll("[data-sort-key]").forEach((button) => {
    button.classList.remove("sort-asc", "sort-desc");
    if (button.getAttribute("data-sort-key") === state.tableSort.key) {
      button.classList.add(state.tableSort.direction === "asc" ? "sort-asc" : "sort-desc");
    }
  });
}

function jumpToPage() {
  const totalPages = Math.max(1, Math.ceil(state.sortedTransactions.length / state.pageSize));
  const requestedPage = Number(els.pageJumpInput.value || 1);
  if (!Number.isFinite(requestedPage)) return;
  state.currentPage = Math.min(totalPages, Math.max(1, Math.round(requestedPage)));
  renderTransactionsTable({ preservePaginationPosition: true });
}

function restorePaginationPosition(previousPaginationTop) {
  if (previousPaginationTop === null || !els.paginationBar) return;

  requestAnimationFrame(() => {
    const currentPaginationTop = els.paginationBar.getBoundingClientRect().top;
    window.scrollBy({
      top: currentPaginationTop - previousPaginationTop,
      left: 0,
      behavior: "auto",
    });
  });
}

function buildRunningSeries(values) {
  let running = 0;
  return values.map((value) => {
    running += value;
    return running;
  });
}

function renderSparkline(values, toneClass = "") {
  const safeValues = values.length ? values : [0];
  const width = 132;
  const height = 34;
  const min = Math.min(...safeValues);
  const max = Math.max(...safeValues);
  const span = max - min || 1;
  const step = safeValues.length > 1 ? width / (safeValues.length - 1) : width;
  const points = safeValues.map((value, index) => {
    const x = safeValues.length > 1 ? index * step : width / 2;
    const y = height - (((value - min) / span) * (height - 6) + 3);
    return `${x.toFixed(2)},${y.toFixed(2)}`;
  }).join(" ");
  const sparkClass = toneClass === "is-negative" ? "negative" : "positive";

  return `
    <svg class="metric-sparkline ${sparkClass}" viewBox="0 0 ${width} ${height}" aria-hidden="true" focusable="false">
      <polyline points="${points}"></polyline>
    </svg>
  `;
}

function renderMonthlySummary() {
  const summaryMap = new Map();
  const latestBalanceByAccount = new Map();
  let totalDebit = 0;
  let totalCredit = 0;

  state.filteredTransactions.forEach((txn) => {
    const key = `${txn.txnDate.slice(0, 7)}||${txn.accountLabel}`;
    const current = summaryMap.get(key) || {
      month: txn.txnDate.slice(0, 7),
      accountLabel: txn.accountLabel,
      totalDebit: 0,
      totalCredit: 0,
    };
    const debitAmount = toInsightAmount(txn.debit, txn.currency);
    const creditAmount = toInsightAmount(txn.credit, txn.currency);
    current.totalDebit += debitAmount;
    current.totalCredit += creditAmount;
    summaryMap.set(key, current);

    totalDebit += debitAmount;
    totalCredit += creditAmount;

    const existingBalance = latestBalanceByAccount.get(txn.accountLabel);
    if (!existingBalance || txn.txnDate >= existingBalance.txnDate) {
      latestBalanceByAccount.set(txn.accountLabel, txn);
    }
  });

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

  if (!monthlyRows.length) {
    els.monthlySummaryBody.innerHTML = `<tr><td colspan="5" class="empty-state">No monthly insights for the current filters. Try widening the date range or clearing bank filters.</td></tr>`;
    return;
  }

  let previousMonth = "";
  let monthBandIndex = -1;
  els.monthlySummaryBody.innerHTML = monthlyRows.map((row) => {
    if (row.month !== previousMonth) {
      monthBandIndex += 1;
      previousMonth = row.month;
    }
    const monthBandClass = monthBandIndex % 2 === 0 ? "month-band-even" : "month-band-odd";
    return `
    <tr class="${monthBandClass}">
      <td>${escapeHtml(row.month)}</td>
      <td>${renderAccountBadge(row.accountLabel)}</td>
      <td class="amount-negative">${moneyFormat(row.totalDebit)}</td>
      <td class="amount-positive">${moneyFormat(row.totalCredit)}</td>
      <td class="${row.totalCredit - row.totalDebit >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(row.totalCredit - row.totalDebit)}</td>
    </tr>
  `;
  }).join("");
}

function renderTrendInsights() {
  const filtered = getChronologicalTransactions(state.filteredTransactions);
  const planningTransactions = getChronologicalTransactions(state.transactions);
  const planningYear = getPlanningYear();
  const categoryMap = new Map();
  const forecastMap = new Map();

  planningTransactions.forEach((txn) => {
    const forecastEntry = forecastMap.get(txn.accountLabel) || {
      accountLabel: txn.accountLabel,
      bankName: txn.bankName,
      currency: txn.currency,
      lastBalance: 0,
      lastDate: "",
      net30: 0,
    };
    if (txn.txnDate >= forecastEntry.lastDate) {
      forecastEntry.lastDate = txn.txnDate;
      forecastEntry.lastBalance = toInsightAmount(txn.balance, txn.currency);
    }
    forecastMap.set(txn.accountLabel, forecastEntry);
  });

  const lastTxnDate = planningTransactions.length ? planningTransactions[planningTransactions.length - 1].txnDate : "";
  const rollingStart = lastTxnDate ? shiftDate(lastTxnDate, -30) : "";
  planningTransactions.forEach((txn) => {
    if (rollingStart && txn.txnDate >= rollingStart) {
      const forecastEntry = forecastMap.get(txn.accountLabel);
      forecastEntry.net30 += toInsightNet(txn);
    }
  });

  filtered.forEach((txn) => {
    if (toInsightAmount(txn.debit, txn.currency) > 0) {
      const category = categorizeTransaction(txn);
      const categoryEntry = categoryMap.get(category) || { category, count: 0, total: 0 };
      categoryEntry.count += 1;
      categoryEntry.total += toInsightAmount(txn.debit, txn.currency);
      categoryMap.set(category, categoryEntry);
    }
  });

  const baseForecastRows = Array.from(forecastMap.values())
    .sort((a, b) => a.accountLabel.localeCompare(b.accountLabel))
    .map((row) => ({
      ...row,
      endOfYear: row.lastBalance,
    }));
  const forecastPlan = buildForecastPlan(planningTransactions, baseForecastRows, planningYear);
  const forecastRows = baseForecastRows.map((row) => {
      const endOfYear = forecastPlan.accountEndBalances.get(row.accountLabel)
        ?? projectAccountYearEnd(row, planningTransactions, planningYear);
      return {
        ...row,
        endOfYear,
      };
    });
  const totalCurrentBalance = forecastRows.reduce((sum, row) => sum + row.lastBalance, 0);
  const projectedClosing = forecastRows.length
    ? forecastRows.reduce((sum, row) => sum + row.endOfYear, 0)
    : totalCurrentBalance;
  const yearlyProjection = buildYearlyProjection(planningTransactions, totalCurrentBalance, projectedClosing, forecastPlan, planningYear);
  const knownFutureOutflows = forecastPlan.projectedCsgTotal
    + forecastPlan.projectedIncomeTaxTotal;

  const trendTiles = [
    {
      label: "Projected Year-End Balance",
      value: moneyFormat(projectedClosing),
      subtext: `Expected total balance by Dec ${planningYear}`,
      tooltip: `All tracked accounts combined. Based on current balances, confirmed Tax Tracker items, and estimated non-tax spend through Dec ${planningYear}.`,
    },
    {
      label: "Forecast Change",
      value: moneyFormat(projectedClosing - totalCurrentBalance),
      subtext: "Projected movement from today to year end",
      tooltip: "Projected year-end balance minus current all-account balance.",
      toneClass: projectedClosing - totalCurrentBalance >= 0 ? "is-positive" : "is-negative",
    },
    {
      label: "Confirmed Future Income",
      value: moneyFormat(forecastPlan.projectedIncomeTotal),
      subtext: "Scheduled Tax Tracker receipts not yet received",
      tooltip: `Confirmed future Tax Tracker income for ${planningYear} only, based on pending receipt dates after the latest imported transaction date.`,
      toneClass: "is-positive",
    },
    {
      label: "Known Future Outflows",
      value: moneyFormat(knownFutureOutflows),
      subtext: "Planned income tax and Tax Tracker CSG still ahead",
      detail: `+ ${moneyFormat(forecastPlan.projectedRecurringSpendTotal)} estimated non-tax spend`,
      tooltip: `Known outflows include ${moneyFormat(forecastPlan.projectedIncomeTaxTotal)} income tax and ${moneyFormat(forecastPlan.projectedCsgTotal)} Tax Tracker CSG for ${planningYear}. Estimated non-tax spend is shown separately below.`,
      toneClass: "is-negative",
    },
  ];

  els.trendMetricsGrid.innerHTML = trendTiles.map((metric) => `
    <article class="metric-tile ${metric.toneClass || ""}" ${metric.tooltip ? `title="${escapeHtml(metric.tooltip)}"` : ""}>
      <div class="metric-label">${escapeHtml(metric.label)}</div>
      <div class="metric-value">${escapeHtml(metric.value)}</div>
      <div class="metric-subtext">${escapeHtml(metric.subtext)}</div>
      ${metric.detail ? `<div class="metric-detail">${escapeHtml(metric.detail)}</div>` : ""}
    </article>
  `).join("");

  els.trendlineSummary.textContent = yearlyProjection.summaryText;
  renderTrendlineChart(yearlyProjection);

  const allCategoryRows = Array.from(categoryMap.values()).sort((a, b) => b.total - a.total);
  renderCategoryBars(allCategoryRows);

  els.forecastSummaryCards.innerHTML = forecastRows.length ? forecastRows.map((row) => `
    <article class="forecast-card">
      <div class="forecast-card-title">${renderAccountBadge(row.accountLabel)}</div>
      <div class="forecast-card-grid">
        <div>
          <div class="forecast-card-label">Last balance</div>
          <div class="forecast-card-value">${moneyFormat(row.lastBalance)}</div>
        </div>
        <div>
          <div class="forecast-card-label">30d flow</div>
          <div class="forecast-card-value ${row.net30 >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(row.net30)}</div>
        </div>
        <div>
          <div class="forecast-card-label">Year-end</div>
          <div class="forecast-card-value ${row.endOfYear >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(row.endOfYear)}</div>
        </div>
      </div>
    </article>
  `).join("") : `<div class="empty-state">No account outlook is available yet for the imported ledger.</div>`;
}

function renderCategoryBars(rows) {
  if (!rows.length) {
    els.categoryBarList.innerHTML = `<div class="empty-state">No visible spend categories to compare. Try widening the date range or clearing bank filters.</div>`;
    return;
  }

  const maxTotal = Math.max(...rows.map((row) => row.total), 1);
  els.categoryBarList.innerHTML = rows.map((row) => `
    <div class="category-bar-item" title="${escapeHtml(`${row.category}: ${moneyFormat(row.total)}`)}">
      <div class="category-bar-label">${escapeHtml(row.category)}</div>
      <div class="category-bar-track">
        <div class="category-bar-fill" style="--fill:${((row.total / maxTotal) * 100).toFixed(2)}%"></div>
      </div>
      <div class="category-bar-value">${escapeHtml(compactMoneyFormat(row.total))}</div>
    </div>
  `).join("");
}

function buildYearlyProjection(transactions, totalCurrentBalance = 0, projectedClosingBalance = 0, forecastPlan = null, targetYear = getPlanningYear()) {
  const yearTransactions = transactions.filter((txn) => txn.txnDate.startsWith(`${targetYear}-`));
  const latestDate = getLatestTransactionDate(yearTransactions);
  const lastMonthIndex = latestDate ? Number(latestDate.slice(5, 7)) : 0;
  const todayMonthIndex = Number(getTodayDateString().slice(5, 7));
  const monthsElapsed = Math.max(lastMonthIndex, 0);
  const monthlyTotals = new Map();
  const activeMonths = new Set();

  yearTransactions.forEach((txn) => {
    const month = txn.txnDate.slice(0, 7);
    monthlyTotals.set(month, (monthlyTotals.get(month) || 0) + toInsightNet(txn));
    if (toInsightNet(txn) !== 0) {
      activeMonths.add(month);
    }
  });

  const actualNetSeries = buildMonthlySeries(targetYear, monthsElapsed, monthlyTotals);
  const actualNetTotal = actualNetSeries.reduce((sum, point) => sum + point.y, 0);
  const futureMonthlyNetMap = forecastPlan?.futureMonthlyNetMap || new Map();
  const effectiveMonthIndex = actualNetSeries.length ? monthsElapsed : todayMonthIndex;
  const monthsRemaining = Math.max(0, 12 - effectiveMonthIndex);
  const fallbackProjectedNetValues = actualNetSeries.length
    ? projectFutureMonthlyValues(actualNetSeries, monthsElapsed, monthsRemaining)
    : Array.from({ length: monthsRemaining }, () => 0);
  const projectedNetValues = Array.from({ length: monthsRemaining }, (_, index) => {
    const monthNumber = effectiveMonthIndex + index + 1;
    const monthKey = `${targetYear}-${String(monthNumber).padStart(2, "0")}`;
    if (futureMonthlyNetMap.has(monthKey)) {
      return futureMonthlyNetMap.get(monthKey) || 0;
    }
    return fallbackProjectedNetValues[index] || 0;
  });
  const impliedStartingBalance = totalCurrentBalance - actualNetTotal;
  let runningBalance = impliedStartingBalance;
  const actualSeries = actualNetSeries.map((point) => {
    runningBalance += point.y;
    return {
      ...point,
      y: runningBalance,
    };
  });

  const projectionSeries = [];
  for (let i = 1; i <= monthsRemaining; i += 1) {
    const absoluteMonthIndex = effectiveMonthIndex + i;
    const month = `${targetYear}-${String(absoluteMonthIndex).padStart(2, "0")}`;
    runningBalance += projectedNetValues[i - 1] ?? 0;
    projectionSeries.push({
      label: formatMonthShort(month),
      month,
      x: absoluteMonthIndex,
      monthNumber: absoluteMonthIndex,
      y: runningBalance,
      forecast: true,
    });
  }

  if (!actualSeries.length && Number.isFinite(totalCurrentBalance)) {
    const anchorMonthNumber = Math.max(1, Math.min(12, effectiveMonthIndex));
    const anchorMonth = `${targetYear}-${String(anchorMonthNumber).padStart(2, "0")}`;
    actualSeries.push({
      label: formatMonthShort(anchorMonth),
      month: anchorMonth,
      x: anchorMonthNumber,
      monthNumber: anchorMonthNumber,
      y: totalCurrentBalance,
      forecast: false,
    });
    runningBalance = totalCurrentBalance;
  }

  const visibleMonths = actualSeries.length;
  const summarySubtext = visibleMonths
    ? `Balance path from Jan to ${formatMonthShort(`${targetYear}-${String(monthsElapsed || 1).padStart(2, "0")}`)} ${targetYear}, including quiet months as zero`
    : `Waiting for ${targetYear} transactions to estimate balance growth`;
  const balanceSubtext = yearTransactions.length
    ? `Projected total balance across all tracked accounts through Dec ${targetYear}`
    : "Import current-year statements to estimate a closing balance";
  const summaryText = forecastPlan
    ? `Trendline blends confirmed Tax Tracker income and tax outflows with trailing-average day-to-day spending to target ${moneyFormat(projectedClosingBalance)} by Dec ${targetYear}.`
    : visibleMonths
      ? `Trendline shows expected balance growth toward ${moneyFormat(projectedClosingBalance)} by Dec ${targetYear}.${activeMonths.size < 4 ? " Low confidence: fewer than 4 active months are visible, so this uses an average-month fallback." : ""}`
      : `Import transactions dated in ${targetYear} to generate the balance-growth forecast.`;

  return {
    year: targetYear,
    actualSeries,
    projectionSeries,
    monthsRemaining,
    summarySubtext,
    balanceSubtext,
    summaryText,
  };
}

function renderTrendlineChart(projection) {
  if (!projection.actualSeries.length) {
    els.trendlineChart.innerHTML = `<foreignObject x="0" y="0" width="760" height="280"><div xmlns="http://www.w3.org/1999/xhtml" class="trend-empty">No current-year monthly data is visible yet.</div></foreignObject>`;
    return;
  }

  const width = 760;
  const height = 280;
  const pad = { top: 24, right: 24, bottom: 44, left: 56 };
  const allPoints = [...projection.actualSeries, ...projection.projectionSeries];
  const allValues = allPoints.map((point) => point.y);
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
  const xStep = allPoints.length > 1 ? (width - pad.left - pad.right) / (allPoints.length - 1) : 0;
  const xForIndex = (index) => pad.left + xStep * index;
  const yForValue = (value) => pad.top + ((maxY - value) / (maxY - minY)) * (height - pad.top - pad.bottom);
  const actualPath = projection.actualSeries.map((point, index) => `${index === 0 ? "M" : "L"} ${xForIndex(index).toFixed(2)} ${yForValue(point.y).toFixed(2)}`).join(" ");
  const projectionStartIndex = Math.max(0, projection.actualSeries.length - 1);
  const projectionPathPoints = [allPoints[projectionStartIndex], ...projection.projectionSeries];
  const projectionPath = projectionPathPoints
    .map((point, index) => {
      const absoluteIndex = projectionStartIndex + index;
      return `${index === 0 ? "M" : "L"} ${xForIndex(absoluteIndex).toFixed(2)} ${yForValue(point.y).toFixed(2)}`;
    })
    .join(" ");
  const zeroY = yForValue(0);
  const gridValues = [minY, (minY + maxY) / 2, maxY];
  const todayMonthNumber = projection.year === getPlanningYear() ? Number(getTodayDateString().slice(5, 7)) : 0;
  const hasTodayMarker = todayMonthNumber >= 1 && todayMonthNumber <= allPoints.length;
  const todayMarkerX = hasTodayMarker ? xForIndex(todayMonthNumber - 1) : 0;

  els.trendlineChart.innerHTML = `
    ${gridValues.map((value) => `
      <g>
        <line class="trend-gridline ${Math.abs(value) < 0.0001 ? "trend-zero-line" : ""}" x1="${pad.left}" y1="${yForValue(value)}" x2="${width - pad.right}" y2="${yForValue(value)}"></line>
        <text class="trend-label" x="10" y="${yForValue(value) + 4}">${escapeHtml(compactMoneyFormat(value))}</text>
      </g>
    `).join("")}
    <line class="trend-axis" x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}"></line>
    ${hasTodayMarker ? `
      <line class="trend-today-marker" x1="${todayMarkerX}" y1="${pad.top}" x2="${todayMarkerX}" y2="${height - pad.bottom}"></line>
      <text class="trend-today-label" x="${todayMarkerX}" y="${pad.top - 6}" text-anchor="middle">Today</text>
    ` : ""}
    ${actualPath ? `<path class="trend-actual" d="${actualPath}"></path>` : ""}
    ${projection.projectionSeries.length ? `<path class="trend-projection" d="${projectionPath}"></path>` : ""}
    ${allPoints.map((point, index) => `
      <g>
        <circle class="trend-point ${point.forecast ? "forecast" : ""}" cx="${xForIndex(index)}" cy="${yForValue(point.y)}" r="4.5">
          <title>${escapeHtml(`${point.label} ${projection.year}: ${moneyFormat(point.y)}`)}</title>
        </circle>
        <text class="trend-label" x="${xForIndex(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(point.label)}</text>
      </g>
    `).join("")}
    <text class="trend-value-label" x="${width - pad.right}" y="${yForValue(allPoints[allPoints.length - 1].y) - 10}" text-anchor="end">
      ${escapeHtml(`Year-end trend ${compactMoneyFormat(allPoints[allPoints.length - 1].y)}`)}
    </text>
  `;
}

function detectProjectionYear(transactions) {
  const currentYear = new Date().getFullYear();
  if (!transactions.length) {
    return currentYear;
  }
  return Math.max(currentYear, Number(transactions[transactions.length - 1].txnDate.slice(0, 4)));
}

function linearRegression(values) {
  if (!values.length) {
    return { slope: 0, intercept: 0 };
  }
  const normalized = values.map((value, index) => (
    typeof value === "number"
      ? { x: index + 1, y: value }
      : { x: Number(value.x), y: Number(value.y) }
  ));
  if (normalized.length === 1) {
    return { slope: 0, intercept: normalized[0].y };
  }

  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumXX = 0;
  normalized.forEach((point) => {
    sumX += point.x;
    sumY += point.y;
    sumXY += point.x * point.y;
    sumXX += point.x * point.x;
  });

  const count = normalized.length;
  const denominator = count * sumXX - sumX * sumX;
  const slope = denominator ? (count * sumXY - sumX * sumY) / denominator : 0;
  const intercept = (sumY - slope * sumX) / count;
  return { slope, intercept };
}

function sumProjectedMonths(slope, intercept, startIndex, count) {
  let total = 0;
  for (let i = 0; i < count; i += 1) {
    total += intercept + slope * (startIndex + i);
  }
  return total;
}

function projectAccountYearEnd(accountForecast, transactions, targetYear) {
  const accountTransactions = transactions.filter((txn) =>
    txn.accountLabel === accountForecast.accountLabel && txn.txnDate.startsWith(`${targetYear}-`)
  );

  if (!accountTransactions.length || !accountForecast.lastDate) {
    return accountForecast.lastBalance;
  }

  const monthlyNetByAccount = new Map();
  accountTransactions.forEach((txn) => {
    const month = txn.txnDate.slice(0, 7);
    monthlyNetByAccount.set(month, (monthlyNetByAccount.get(month) || 0) + toInsightNet(txn));
  });

  const lastMonthNumber = Number(accountForecast.lastDate.slice(5, 7));
  const monthSeries = buildMonthlySeries(targetYear, lastMonthNumber, monthlyNetByAccount, true);
  const monthValues = monthSeries.map((point) => point.y);
  const monthsRemaining = Math.max(0, 12 - lastMonthNumber);

  if (monthsRemaining === 0) {
    return accountForecast.lastBalance;
  }

  const projectedNet = projectFutureMonthlyValues(monthSeries, lastMonthNumber, monthsRemaining)
    .reduce((sum, value) => sum + value, 0);

  return accountForecast.lastBalance + projectedNet;
}

function buildForecastPlan(transactions, forecastRows, targetYear) {
  if (!forecastRows.length) {
    return {
      futureMonthlyNetMap: new Map(),
      accountEndBalances: new Map(),
      projectedIncomeTotal: 0,
      projectedCsgTotal: 0,
      projectedIncomeTaxTotal: 0,
      projectedRecurringSpendTotal: 0,
    };
  }

  const yearTransactions = transactions.filter((txn) => txn.txnDate.startsWith(`${targetYear}-`));
  const latestYearDate = getLatestTransactionDate(yearTransactions);
  const currentMonth = Number(getTodayDateString().slice(5, 7));
  const forecastStartMonth = latestYearDate ? Number(latestYearDate.slice(5, 7)) + 1 : currentMonth + 1;
  const futureMonthlyNetMap = new Map();
  const accountEndBalances = new Map(forecastRows.map((row) => [row.accountLabel, row.lastBalance]));
  const recurringMonthlySpend = estimateRecurringMonthlySpend(transactions);
  const spendShares = estimateSpendingAccountShares(transactions, forecastRows);
  const operatingBuffers = estimateOperatingBuffers(transactions, forecastRows, spendShares, recurringMonthlySpend);
  const incomeAccountLabel = pickForecastAccount(forecastRows, (row) => row.bankName === "MCB" && row.currency === "GBP");
  const sbmMurAccountLabel = pickForecastAccount(forecastRows, (row) => row.bankName === "SBM" && row.currency === "MUR");
  const mcbMurAccountLabel = pickForecastAccount(forecastRows, (row) => row.bankName === "MCB" && row.currency === "MUR");
  const taxSummary = computeTaxSummary();
  const taxYearRange = getCurrentTaxYearRange();
  const projectedIncomeByMonth = new Map();
  const projectedCsgByMonth = new Map();
  const projectedMonthlySpendByAccount = new Map();

  spendShares.forEach((share, accountLabel) => {
    projectedMonthlySpendByAccount.set(accountLabel, (recurringMonthlySpend || 0) * share);
  });
  if (!projectedMonthlySpendByAccount.size && recurringMonthlySpend > 0) {
    const fallbackAccount = sbmMurAccountLabel || mcbMurAccountLabel;
    if (fallbackAccount) {
      projectedMonthlySpendByAccount.set(fallbackAccount, recurringMonthlySpend);
    }
  }

  state.taxEntries.forEach((entry) => {
    const computed = computeTaxEntry(entry);
    const incomeDate = getEffectiveTaxReceiptDate(entry);
    if (incomeDate && incomeDate > latestYearDate && incomeDate.startsWith(`${targetYear}-`)) {
      const monthKey = incomeDate.slice(0, 7);
      projectedIncomeByMonth.set(monthKey, (projectedIncomeByMonth.get(monthKey) || 0) + computed.amountReceivedMur);
    }

    const csgDate = getEffectiveTaxCsgDate(entry);
    if (csgDate && csgDate > latestYearDate && csgDate.startsWith(`${targetYear}-`)) {
      const monthKey = csgDate.slice(0, 7);
      projectedCsgByMonth.set(monthKey, (projectedCsgByMonth.get(monthKey) || 0) + computed.csgAmountPaidMur);
    }
  });

  const incomeTaxByMonth = new Map();
  const incomeTaxPaymentDate = `${targetYear}-09-30`;
  if (
    incomeAccountLabel &&
    taxYearRange.endYear === targetYear &&
    incomeTaxPaymentDate > latestYearDate
  ) {
    incomeTaxByMonth.set(`${targetYear}-09`, taxSummary.expectedIncomeTax);
  }

  const projectedIncomeTotal = Array.from(projectedIncomeByMonth.values()).reduce((sum, value) => sum + value, 0);
  const projectedCsgTotal = Array.from(projectedCsgByMonth.values()).reduce((sum, value) => sum + value, 0);
  const projectedIncomeTaxTotal = Array.from(incomeTaxByMonth.values()).reduce((sum, value) => sum + value, 0);
  const projectedRecurringSpendTotal = Array.from(projectedMonthlySpendByAccount.values())
    .reduce((sum, value) => sum + value, 0) * Math.max(0, 13 - forecastStartMonth);

  for (let monthNumber = forecastStartMonth; monthNumber <= 12; monthNumber += 1) {
    const monthKey = `${targetYear}-${String(monthNumber).padStart(2, "0")}`;
    let monthNet = 0;

    const projectedIncome = projectedIncomeByMonth.get(monthKey) || 0;
    if (projectedIncome && incomeAccountLabel) {
      accountEndBalances.set(incomeAccountLabel, (accountEndBalances.get(incomeAccountLabel) || 0) + projectedIncome);
    }
    monthNet += projectedIncome;

    projectedMonthlySpendByAccount.forEach((monthlySpend, accountLabel) => {
      accountEndBalances.set(accountLabel, (accountEndBalances.get(accountLabel) || 0) - monthlySpend);
      monthNet -= monthlySpend;
    });

    const projectedCsg = projectedCsgByMonth.get(monthKey) || 0;
    if (projectedCsg && sbmMurAccountLabel) {
      accountEndBalances.set(sbmMurAccountLabel, (accountEndBalances.get(sbmMurAccountLabel) || 0) - projectedCsg);
    }
    monthNet -= projectedCsg;

    const projectedIncomeTax = incomeTaxByMonth.get(monthKey) || 0;
    if (projectedIncomeTax && incomeAccountLabel) {
      accountEndBalances.set(incomeAccountLabel, (accountEndBalances.get(incomeAccountLabel) || 0) - projectedIncomeTax);
    }
    monthNet -= projectedIncomeTax;

    if (incomeAccountLabel) {
      projectedMonthlySpendByAccount.forEach((_, accountLabel) => {
        const buffer = operatingBuffers.get(accountLabel);
        if (!buffer) return;
        const currentBalance = accountEndBalances.get(accountLabel) || 0;
        if (currentBalance >= buffer.minimum) return;

        const sourceBalance = accountEndBalances.get(incomeAccountLabel) || 0;
        const desiredTopUp = Math.max(0, buffer.target - currentBalance);
        const topUp = Math.min(desiredTopUp, Math.max(0, sourceBalance));
        if (topUp <= 0) return;

        accountEndBalances.set(accountLabel, currentBalance + topUp);
        accountEndBalances.set(incomeAccountLabel, sourceBalance - topUp);
      });

      if (sbmMurAccountLabel) {
        const sbmBuffer = operatingBuffers.get(sbmMurAccountLabel);
        const currentBalance = accountEndBalances.get(sbmMurAccountLabel) || 0;
        if (sbmBuffer && currentBalance < sbmBuffer.minimum) {
          const sourceBalance = accountEndBalances.get(incomeAccountLabel) || 0;
          const desiredTopUp = Math.max(0, sbmBuffer.target - currentBalance);
          const topUp = Math.min(desiredTopUp, Math.max(0, sourceBalance));
          if (topUp > 0) {
            accountEndBalances.set(sbmMurAccountLabel, currentBalance + topUp);
            accountEndBalances.set(incomeAccountLabel, sourceBalance - topUp);
          }
        }
      }
    }

    futureMonthlyNetMap.set(monthKey, monthNet);
  }

  return {
    futureMonthlyNetMap,
    accountEndBalances,
    projectedIncomeTotal,
    projectedCsgTotal,
    projectedIncomeTaxTotal,
    projectedRecurringSpendTotal,
  };
}

function getChronologicalTransactions(transactions) {
  return [...transactions].sort((a, b) => (
    a.txnDate.localeCompare(b.txnDate)
    || a.accountLabel.localeCompare(b.accountLabel)
    || a.description.localeCompare(b.description)
  ));
}

function getLatestTransactionDate(transactions) {
  return transactions.reduce((latestDate, txn) => (
    txn.txnDate > latestDate ? txn.txnDate : latestDate
  ), "");
}

function getPlanningYear() {
  return Number(getTodayDateString().slice(0, 4));
}

function getCurrentTaxYearRange(today = getTodayDateString()) {
  const year = Number(today.slice(0, 4));
  const month = Number(today.slice(5, 7));
  const startYear = month >= 7 ? year : year - 1;
  const endYear = startYear + 1;
  return {
    start: `${startYear}-07-01`,
    end: `${endYear}-06-30`,
    label: `Jul ${startYear} to Jun ${endYear}`,
    startYear,
    endYear,
  };
}

function estimateRecurringMonthlySpend(transactions) {
  if (!transactions.length) return 0;
  const latestDate = getLatestTransactionDate(transactions);
  const monthTotals = new Map();

  transactions.forEach((txn) => {
    if (toInsightAmount(txn.debit, txn.currency) <= 0) return;
    const category = categorizeTransaction(txn);
    if (!FORECAST_SPEND_CATEGORIES.has(category)) return;
    const monthKey = txn.txnDate.slice(0, 7);
    monthTotals.set(monthKey, (monthTotals.get(monthKey) || 0) + toInsightAmount(txn.debit, txn.currency));
  });

  const historyMonths = [];
  let cursor = new Date(`${latestDate.slice(0, 7)}-01T00:00:00`);
  for (let i = 0; i < 36; i += 1) {
    historyMonths.unshift(`${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`);
    cursor.setMonth(cursor.getMonth() - 1);
  }

  const historySeries = historyMonths.map((monthKey) => monthTotals.get(monthKey) || 0);
  const observedValues = historySeries.filter((value) => value > 0);
  if (!observedValues.length) {
    return 0;
  }
  if (observedValues.length < 6) {
    return average(observedValues);
  }

  const longTermBaseline = winsorizedAverage(historySeries, 0.1, 0.9);
  const recentSeries = historySeries.slice(-12);
  const recentMedian = median(recentSeries);
  const trendProjection = projectSpendRunRate(historySeries);

  return Math.max(0, 0.5 * longTermBaseline + 0.3 * recentMedian + 0.2 * trendProjection);
}

function estimateSpendingAccountShares(transactions, forecastRows) {
  const spendByAccount = new Map();
  const spendAccounts = forecastRows.filter((row) => row.currency === "MUR" && (row.bankName === "MCB" || row.bankName === "SBM"));

  transactions.forEach((txn) => {
    if (toInsightAmount(txn.debit, txn.currency) <= 0) return;
    const category = categorizeTransaction(txn);
    if (!FORECAST_SPEND_CATEGORIES.has(category)) return;
    if (!spendAccounts.some((row) => row.accountLabel === txn.accountLabel)) return;
    spendByAccount.set(txn.accountLabel, (spendByAccount.get(txn.accountLabel) || 0) + Number(txn.debit || 0));
  });

  const total = Array.from(spendByAccount.values()).reduce((sum, value) => sum + value, 0);
  if (!total) {
    const fallback = new Map();
    if (spendAccounts.length === 1) {
      fallback.set(spendAccounts[0].accountLabel, 1);
      return fallback;
    }
    spendAccounts.forEach((row) => fallback.set(row.accountLabel, 1 / Math.max(1, spendAccounts.length)));
    return fallback;
  }

  const shares = new Map();
  spendByAccount.forEach((value, accountLabel) => {
    shares.set(accountLabel, value / total);
  });
  return shares;
}

function pickForecastAccount(forecastRows, predicate) {
  const match = forecastRows.find(predicate);
  return match ? match.accountLabel : "";
}

function estimateOperatingBuffers(transactions, forecastRows, spendShares, recurringMonthlySpend) {
  const buffers = new Map();
  const currentYear = getPlanningYear();
  const spendAccounts = forecastRows.filter((row) => row.currency === "MUR" && (row.bankName === "MCB" || row.bankName === "SBM"));

  spendAccounts.forEach((row) => {
    const monthlyBalances = new Map();
    transactions.forEach((txn) => {
      if (txn.accountLabel !== row.accountLabel) return;
      const year = Number(txn.txnDate.slice(0, 4));
      if (year < currentYear - 1) return;
      monthlyBalances.set(txn.txnDate.slice(0, 7), Number(txn.balance || 0));
    });

    const recentBalances = Array.from(monthlyBalances.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .slice(-6)
      .map(([, balance]) => balance)
      .filter((balance) => Number.isFinite(balance));
    const share = spendShares.get(row.accountLabel) || 0;
    const monthlySpendShare = recurringMonthlySpend * share;
    const fallbackMinimum = monthlySpendShare * 0.4;
    const fallbackTarget = monthlySpendShare * 0.9;

    if (!recentBalances.length) {
      buffers.set(row.accountLabel, {
        minimum: fallbackMinimum,
        target: Math.max(fallbackTarget, fallbackMinimum),
      });
      return;
    }

    const sorted = [...recentBalances].sort((a, b) => a - b);
    const minimum = Math.max(percentile(sorted, 0.35), fallbackMinimum);
    const target = Math.max(percentile(sorted, 0.65), minimum, fallbackTarget);
    buffers.set(row.accountLabel, { minimum, target });
  });

  return buffers;
}

function getEffectiveTaxCsgDate(entry) {
  if (entry.dateCsgPaid) {
    return entry.dateCsgPaid;
  }
  const baseDate = getEffectiveTaxReceiptDate(entry);
  if (!baseDate) {
    return "";
  }
  const nextMonth = new Date(`${baseDate}T00:00:00`);
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  return nextMonth.toISOString().slice(0, 10);
}

function buildMonthlySeries(year, lastMonthNumber, totalsMap, clampOutliers = false) {
  const safeLastMonth = Math.max(0, Number(lastMonthNumber) || 0);
  const series = [];

  for (let monthNumber = 1; monthNumber <= safeLastMonth; monthNumber += 1) {
    const monthKey = `${year}-${String(monthNumber).padStart(2, "0")}`;
    series.push({
      label: formatMonthShort(monthKey),
      month: monthKey,
      monthNumber,
      x: monthNumber,
      y: totalsMap.get(monthKey) || 0,
      forecast: false,
    });
  }

  if (!clampOutliers || series.length < 4) {
    return series;
  }

  const cappedValues = capSeriesOutliers(series.map((point) => point.y));
  return series.map((point, index) => ({ ...point, y: cappedValues[index] }));
}

function capSeriesOutliers(values) {
  if (values.length < 4) {
    return values;
  }

  const sorted = [...values].sort((a, b) => a - b);
  const low = percentile(sorted, 0.15);
  const high = percentile(sorted, 0.85);
  return values.map((value) => Math.min(high, Math.max(low, value)));
}

function percentile(sortedValues, ratio) {
  if (!sortedValues.length) {
    return 0;
  }

  const index = (sortedValues.length - 1) * ratio;
  const lower = Math.floor(index);
  const upper = Math.ceil(index);

  if (lower === upper) {
    return sortedValues[lower];
  }

  const weight = index - lower;
  return sortedValues[lower] * (1 - weight) + sortedValues[upper] * weight;
}

function average(values) {
  if (!values.length) {
    return 0;
  }
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function median(values) {
  if (!values.length) {
    return 0;
  }
  const sorted = [...values].sort((a, b) => a - b);
  return percentile(sorted, 0.5);
}

function winsorizedAverage(values, lowerRatio = 0.1, upperRatio = 0.9) {
  if (!values.length) {
    return 0;
  }
  const sorted = [...values].sort((a, b) => a - b);
  const lowerBound = percentile(sorted, lowerRatio);
  const upperBound = percentile(sorted, upperRatio);
  const cappedValues = values.map((value) => Math.min(upperBound, Math.max(lowerBound, value)));
  return average(cappedValues);
}

function projectSpendRunRate(historySeries) {
  if (!historySeries.length) {
    return 0;
  }

  const modeledSeries = historySeries.slice(-18);
  if (modeledSeries.filter((value) => value > 0).length < 4) {
    return winsorizedAverage(historySeries, 0.15, 0.85);
  }

  const smoothedSeries = modeledSeries.map((_, index) => {
    const start = Math.max(0, index - 1);
    const end = Math.min(modeledSeries.length, index + 2);
    return winsorizedAverage(modeledSeries.slice(start, end), 0.15, 0.85);
  });
  const regression = linearRegression(smoothedSeries.map((value, index) => ({
    x: index + 1,
    y: value,
  })));
  const projectedValues = Array.from({ length: 3 }, (_, index) => (
    regression.intercept + regression.slope * (smoothedSeries.length + index + 1)
  ));
  const sorted = [...historySeries].sort((a, b) => a - b);
  const floor = Math.max(0, percentile(sorted, 0.2) * 0.85);
  const ceiling = percentile(sorted, 0.85) * 1.15 + 25000;
  return average(projectedValues.map((value) => Math.min(ceiling, Math.max(floor, value))));
}

function projectFutureMonthlyValues(monthSeries, lastMonthNumber, monthsRemaining) {
  const monthValues = monthSeries.map((point) => point.y);
  const historyCount = monthValues.filter((value) => value !== 0).length;
  const averageMonthlyNet = average(monthValues);

  if (historyCount < 4) {
    return Array.from({ length: monthsRemaining }, () => averageMonthlyNet);
  }

  const regression = linearRegression(monthSeries.map((point) => ({
    x: point.monthNumber,
    y: point.y,
  })));
  const cappedFloor = averageMonthlyNet - Math.abs(averageMonthlyNet) * 1.5 - 50000;
  const cappedCeiling = averageMonthlyNet + Math.abs(averageMonthlyNet) * 1.5 + 50000;

  return Array.from({ length: monthsRemaining }, (_, index) => {
    const monthNumber = lastMonthNumber + index + 1;
    const projectedValue = regression.intercept + regression.slope * monthNumber;
    return Math.min(cappedCeiling, Math.max(cappedFloor, projectedValue));
  });
}

function exportTransactionsCsv() {
  if (!state.filteredTransactions.length) {
    logMessage("Nothing to export yet.");
    return;
  }

  const headers = ["Txn Date", "Value Date", "Account", "Currency", "Description", "Debit", "Credit", "Balance", "Reference", "Source File"];
  const rows = state.sortedTransactions.map((txn) => [
    txn.txnDate,
    txn.valueDate,
    txn.accountLabel,
    txn.currency,
    txn.description,
    txn.debit.toFixed(2),
    txn.credit.toFixed(2),
    txn.balance.toFixed(2),
    txn.reference,
    txn.sourceFile,
  ]);

  const csv = [headers, ...rows].map((row) => row.map(csvEscape).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "finance_transactions_filtered_export.csv";
  link.click();
  URL.revokeObjectURL(url);
  logMessage(`Exported ${rows.length} filtered transaction(s) to CSV.`);
}

async function resetLedgerData() {
  if (!state.ledgerHandle) {
    logMessage("No ledger file selected.");
    return;
  }

  const confirmed = window.confirm("Clear all transactions from the current ledger file?");
  if (!confirmed) return;

  state.transactions = [];
  state.imports = [];
  state.filteredTransactions = [];
  state.sortedTransactions = [];
  state.taxEntries = [];
  state.taxExpectedExpensesMode = AUTO_TAX_EXPENSES_MODE;
  state.taxExpectedExpenses = calculateSuggestedTaxExpectedExpenses(state.transactions);
  state.currentPage = 1;
  state.shouldScrollNextTaxReceipt = true;
  await persistTaxEntries();
  await persistTaxExpectedExpenses();
  await saveLedgerToDisk();
  renderAll();
  updateStatus("Ledger cleared.");
  logMessage("Current ledger data cleared and saved.");
}

function renderTaxTable() {
  if (!state.taxEntries.length) {
    els.taxBody.innerHTML = `<tr><td colspan="10" class="empty-state">No tax rows yet. Click Add Row to start tracking invoices and CSG.</td></tr>`;
    state.shouldScrollNextTaxReceipt = false;
    return;
  }

  els.taxBody.innerHTML = state.taxEntries.map((entry, index) => {
    const computed = computeTaxEntry(entry);
    return `
      <tr data-tax-row="${index}">
        <td data-label="Invoiced Date"><input class="tax-input" type="date" data-index="${index}" data-field="invoicedDate" value="${escapeHtml(entry.invoicedDate)}"></td>
        <td data-label="Date Received"><input class="tax-input" type="date" data-index="${index}" data-field="dateReceived" value="${escapeHtml(entry.dateReceived)}"></td>
        <td data-label="Date CSG Paid"><input class="tax-input" type="date" data-index="${index}" data-field="dateCsgPaid" value="${escapeHtml(entry.dateCsgPaid)}"></td>
        <td data-label="Exchange Rate Buying Notes at Date Received"><input class="tax-input" type="number" step="0.001" min="0" data-index="${index}" data-field="exchangeRate" value="${escapeHtml(entry.exchangeRate)}"></td>
        <td data-label="Amount Received (GBP)"><input class="tax-input" type="number" step="0.01" min="0" data-index="${index}" data-field="amountReceivedGbp" value="${escapeHtml(entry.amountReceivedGbp)}"></td>
        <td data-label="Amount Received (MUR)"><div class="tax-readonly" data-tax-computed="amountReceivedMur">${moneyFormat(computed.amountReceivedMur)}</div></td>
        <td data-label="CSG Amount Paid (MUR)"><div class="tax-readonly" data-tax-computed="csgAmountPaidMur">${moneyFormat(computed.csgAmountPaidMur)}</div></td>
        <td data-label="CSG Payment Reference"><input class="tax-input" type="text" data-index="${index}" data-field="csgPaymentReference" value="${escapeHtml(entry.csgPaymentReference)}" placeholder="Payment reference"></td>
        <td data-label="Action"><button class="btn btn-ghost tax-action-btn" type="button" aria-label="Delete row" title="Delete row" data-tax-delete="${index}">&times;</button></td>
      </tr>
    `;
  }).join("");

  syncNextTaxReceiptTarget({ scroll: state.shouldScrollNextTaxReceipt });
  state.shouldScrollNextTaxReceipt = false;
}

function renderTaxSummary() {
  const summary = computeTaxSummary();
  const taxYearRange = getCurrentTaxYearRange();
  const expenseTargetModeLabel = state.taxExpectedExpensesMode === MANUAL_TAX_EXPENSES_MODE ? "Manual" : "Auto";
  const comparisons = [
    {
      label: "Income",
      expected: summary.expectedIncome,
      current: summary.soFarIncome,
      progressLabel: "received",
    },
    {
      label: "Income Tax",
      expected: summary.expectedIncomeTax,
      current: summary.soFarIncomeTax,
      progressLabel: "accrued",
    },
    {
      label: "CSG",
      expected: summary.expectedCsg,
      current: summary.soFarCsg,
      progressLabel: "paid",
    },
    {
      label: "Expenses",
      expected: state.taxExpectedExpenses,
      current: summary.soFarExpenses,
      progressLabel: "spent",
      inverseTone: true,
    },
    {
      label: "Saved",
      expected: summary.expectedSaved,
      current: summary.soFarSaved,
      progressLabel: "retained",
      emphasize: true,
    },
  ];

  els.taxSummaryBody.innerHTML = `
    <div class="tax-summary-cards">
      <article class="tax-summary-card expected">
        <p class="tax-summary-kicker">Expected by Year End</p>
        <div class="tax-summary-card-value ${summary.expectedSaved >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(summary.expectedSaved)}</div>
        <div class="tax-summary-card-copy">Target saved amount for ${escapeHtml(taxYearRange.label)} after tax, CSG, and planned expenses.</div>
      </article>
      <article class="tax-summary-card current">
        <p class="tax-summary-kicker">Current / So Far</p>
        <div class="tax-summary-card-value ${summary.soFarSaved >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(summary.soFarSaved)}</div>
        <div class="tax-summary-card-copy">${escapeHtml(formatGapLabel(summary.soFarSaved - summary.expectedSaved, "vs expected saved amount"))}</div>
      </article>
    </div>
    <section class="tax-expense-editor">
      <div>
        <div class="tax-expense-heading">
          <div class="tax-expense-title">Expected Expenses</div>
          <span class="mode-badge ${state.taxExpectedExpensesMode === MANUAL_TAX_EXPENSES_MODE ? "manual" : "auto"}">${escapeHtml(expenseTargetModeLabel)}</span>
        </div>
        <div class="tax-expense-copy">Set the full-year expense target here. The comparison below shows whether current spend is still within plan.</div>
        <div class="tax-expense-note">Auto uses expenses so far from ${escapeHtml(taxYearRange.label)} plus estimated remaining months from ledger history. Manual keeps your typed target.</div>
      </div>
      <div class="tax-expense-input-wrap">
        <label for="tax-expected-expenses-input">Expense Target</label>
        <input id="tax-expected-expenses-input" class="tax-summary-input" type="number" min="0" step="0.01" value="${escapeHtml(String(state.taxExpectedExpenses))}">
      </div>
    </section>
    <div class="tax-comparison-wrap">
      <table class="tax-comparison-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>Expected</th>
            <th>Current</th>
            <th>Gap</th>
            <th>Progress</th>
          </tr>
        </thead>
        <tbody>
          ${comparisons.map((row) => renderTaxComparisonRow(row)).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderTaxComparisonRow(row) {
  const gap = row.current - row.expected;
  const normalizedExpected = Math.abs(row.expected);
  const progressRaw = normalizedExpected > 0 ? (Math.abs(row.current) / normalizedExpected) * 100 : 0;
  const progress = Math.max(0, Math.min(progressRaw, 100));
  const isOver = progressRaw > 100;
  const gapToneClass = row.inverseTone
    ? (gap <= 0 ? "amount-positive" : "amount-negative")
    : (gap >= 0 ? "amount-positive" : "amount-negative");

  return `
    <tr class="${row.emphasize ? "tax-row-strong" : ""}">
      <td>${escapeHtml(row.label)}</td>
      <td class="tax-summary-value">${moneyFormat(row.expected)}</td>
      <td class="tax-summary-value ${row.emphasize ? (row.current >= 0 ? "amount-positive" : "amount-negative") : ""}">${moneyFormat(row.current)}</td>
      <td class="tax-summary-value ${gapToneClass}">${formatSignedMoney(gap)}</td>
      <td class="tax-progress-cell">
        <div class="tax-progress-stack">
          <div class="tax-progress-copy">
            <span>${escapeHtml(row.progressLabel)}</span>
            <strong>${escapeHtml(formatPercent(progressRaw))}</strong>
          </div>
          <div class="tax-progress-track">
            <div class="tax-progress-fill ${isOver ? "is-over" : ""}" style="--progress:${progress.toFixed(2)}%"></div>
          </div>
        </div>
      </td>
    </tr>
  `;
}

function handleTaxTableInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }

  const index = Number(target.dataset.index);
  const field = target.dataset.field;
  if (!Number.isInteger(index) || !field || !state.taxEntries[index]) {
    return;
  }

  const shouldScrollNextReceipt = field === "dateReceived"
    && target.hasAttribute("data-next-tax-receipt-input")
    && Boolean(target.value);
  state.taxEntries[index][field] = target.value;
  updateTaxComputedRow(index);
  if (field === "dateReceived") {
    syncNextTaxReceiptTarget({ scroll: shouldScrollNextReceipt });
  }
  renderTaxSummary();
  persistTaxEntries();
}

function handleTaxTableClick(event) {
  const trigger = event.target.closest("[data-tax-delete]");
  if (!trigger) {
    return;
  }

  const index = Number(trigger.getAttribute("data-tax-delete"));
  if (!Number.isInteger(index)) {
    return;
  }

  state.taxEntries.splice(index, 1);
  renderTaxTable();
  renderTaxSummary();
  persistTaxEntries();
}

function getNextTaxReceiptEntryIndex() {
  return state.taxEntries.findIndex((entry) => !String(entry?.dateReceived || "").trim());
}

function syncNextTaxReceiptTarget({ scroll = false } = {}) {
  if (!els.taxBody) {
    return;
  }

  els.taxBody.querySelectorAll("[data-next-tax-receipt]").forEach((row) => {
    row.classList.remove("tax-row-next-receipt");
    row.removeAttribute("data-next-tax-receipt");
  });

  els.taxBody.querySelectorAll("[data-next-tax-receipt-input]").forEach((input) => {
    input.classList.remove("tax-input-next-receipt");
    input.removeAttribute("data-next-tax-receipt-input");
  });

  const nextIndex = getNextTaxReceiptEntryIndex();
  if (nextIndex < 0) {
    return;
  }

  const row = els.taxBody.querySelector(`[data-tax-row="${nextIndex}"]`);
  const input = els.taxBody.querySelector(`[data-index="${nextIndex}"][data-field="dateReceived"]`);
  if (!(row instanceof HTMLTableRowElement) || !(input instanceof HTMLInputElement)) {
    return;
  }

  row.classList.add("tax-row-next-receipt");
  row.setAttribute("data-next-tax-receipt", "true");
  input.classList.add("tax-input-next-receipt");
  input.setAttribute("data-next-tax-receipt-input", "true");

  if (scroll) {
    scrollTaxReceiptTargetIntoView(row);
  }
}

function scrollTaxReceiptTargetIntoView(row) {
  if (!(els.taxTableWrap instanceof HTMLElement)) {
    return;
  }

  requestAnimationFrame(() => {
    if (!row.isConnected) {
      return;
    }

    const containerRect = els.taxTableWrap.getBoundingClientRect();
    const rowRect = row.getBoundingClientRect();
    const nextTop = els.taxTableWrap.scrollTop
      + (rowRect.top - containerRect.top)
      - (els.taxTableWrap.clientHeight / 2)
      + (rowRect.height / 2);

    els.taxTableWrap.scrollTo({
      top: Math.max(0, nextTop),
      left: els.taxTableWrap.scrollLeft,
      behavior: "smooth",
    });
  });
}

function handleTaxSummaryInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || target.id !== "tax-expected-expenses-input") {
    return;
  }

  state.taxExpectedExpensesMode = MANUAL_TAX_EXPENSES_MODE;
  state.taxExpectedExpenses = Number(target.value || 0);
  renderTaxSummary();
  persistTaxExpectedExpenses();
}

function updateTaxComputedRow(index) {
  const row = els.taxBody.querySelector(`[data-tax-row="${index}"]`);
  if (!row || !state.taxEntries[index]) {
    return;
  }

  const computed = computeTaxEntry(state.taxEntries[index]);
  const computedFields = {
    amountReceivedMur: moneyFormat(computed.amountReceivedMur),
    csgAmountPaidMur: moneyFormat(computed.csgAmountPaidMur),
  };

  Object.entries(computedFields).forEach(([field, value]) => {
    const cell = row.querySelector(`[data-tax-computed="${field}"]`);
    if (cell) {
      cell.textContent = value;
    }
  });
}

function createEmptyTaxEntry() {
  return {
    invoicedDate: "",
    dateReceived: "",
    dateCsgPaid: "",
    exchangeRate: "",
    amountReceivedGbp: "",
    csgPaymentReference: "",
  };
}

function createPrefilledTaxEntry() {
  const previousEntry = getLastTaxEntry();
  if (!previousEntry?.invoicedDate) {
    return createEmptyTaxEntry();
  }

  const invoicedDate = getNextInvoiceWorkingDate(previousEntry.invoicedDate);
  const nextRate = Number.isFinite(state.gbpToMurRate) ? formatTaxRate(state.gbpToMurRate) : "";
  const amountReceivedGbp = invoicedDate
    ? String(countWorkingDaysBetweenInvoices(previousEntry.invoicedDate, invoicedDate) * DAILY_SALARY_RATE_GBP)
    : "";

  return {
    invoicedDate,
    dateReceived: "",
    dateCsgPaid: "",
    exchangeRate: nextRate,
    amountReceivedGbp,
    csgPaymentReference: "",
  };
}

function cloneDefaultTaxEntries() {
  return DEFAULT_TAX_ENTRIES.map((entry) => ({ ...entry }));
}

function normalizeTaxEntry(entry) {
  return {
    invoicedDate: String(entry?.invoicedDate || ""),
    dateReceived: String(entry?.dateReceived || ""),
    dateCsgPaid: String(entry?.dateCsgPaid || ""),
    exchangeRate: String(entry?.exchangeRate || ""),
    amountReceivedGbp: String(entry?.amountReceivedGbp || ""),
    csgPaymentReference: String(entry?.csgPaymentReference || ""),
  };
}

function computeTaxEntry(entry) {
  const exchangeRate = Number(entry.exchangeRate || 0);
  const amountReceivedGbp = Number(entry.amountReceivedGbp || 0);
  const amountReceivedMur = roundMur(exchangeRate * amountReceivedGbp);
  const csgAmountPaidGbp = amountReceivedGbp * 0.9 * 0.03;
  const csgAmountPaidMur = roundMur(exchangeRate * csgAmountPaidGbp);

  return {
    amountReceivedMur,
    csgAmountPaidGbp,
    csgAmountPaidMur,
  };
}

function computeTaxSummary() {
  const taxYearRange = getCurrentTaxYearRange();
  const entriesInWindow = state.taxEntries.filter((entry) => isWithinTaxYear(getEffectiveTaxReceiptDate(entry), taxYearRange));
  const expectedIncome = entriesInWindow.reduce((sum, entry) => sum + computeTaxEntry(entry).amountReceivedMur, 0);
  const soFarIncome = entriesInWindow
    .filter((entry) => entry.dateReceived)
    .reduce((sum, entry) => sum + computeTaxEntry(entry).amountReceivedMur, 0);
  const expectedIncomeTax = calculateIncomeTax(expectedIncome);
  const soFarIncomeTax = calculateIncomeTax(soFarIncome);
  const expectedCsg = 0.03 * 0.9 * expectedIncome;
  const soFarCsg = entriesInWindow
    .filter((entry) => entry.csgPaymentReference.trim())
    .reduce((sum, entry) => sum + computeTaxEntry(entry).csgAmountPaidMur, 0);
  const soFarExpenses = getTaxYearExpensesSoFar(state.transactions, getTodayDateString(), taxYearRange);
  const expectedSaved = expectedIncome - expectedIncomeTax - expectedCsg - state.taxExpectedExpenses;
  const soFarSaved = soFarIncome - soFarIncomeTax - soFarCsg - soFarExpenses;

  return {
    expectedIncome,
    expectedIncomeTax,
    expectedCsg,
    soFarIncome,
    soFarIncomeTax,
    soFarCsg,
    soFarExpenses,
    expectedSaved,
    soFarSaved,
  };
}

function refreshAutoTaxExpectedExpenses() {
  if (state.taxExpectedExpensesMode !== AUTO_TAX_EXPENSES_MODE) {
    return;
  }
  state.taxExpectedExpenses = calculateSuggestedTaxExpectedExpenses(state.transactions);
}

function calculateSuggestedTaxExpectedExpenses(transactions) {
  const today = getTodayDateString();
  const taxYearRange = getCurrentTaxYearRange(today);
  const soFarExpenses = getTaxYearExpensesSoFar(transactions, today, taxYearRange);
  const remainingMonths = getRemainingTaxYearMonths(today, taxYearRange);
  const estimatedMonthlyExpenses = estimateMonthlyExpenseRunRate(transactions, today);
  const suggested = soFarExpenses + estimatedMonthlyExpenses * remainingMonths;
  return Math.max(0, roundMur(suggested));
}

function getTaxYearExpensesSoFar(transactions, today = getTodayDateString(), taxYearRange = getCurrentTaxYearRange(today)) {
  return transactions
    .filter((txn) =>
      txn.txnDate >= taxYearRange.start
      && txn.txnDate <= today
      && txn.currency === "MUR"
      && (txn.bankName === "MCB" || txn.bankName === "SBM")
    )
    .reduce((sum, txn) => sum + Number(txn.debit || 0), 0);
}

function getRemainingTaxYearMonths(today = getTodayDateString(), taxYearRange = getCurrentTaxYearRange(today)) {
  if (today >= taxYearRange.end) {
    return 0;
  }
  const currentMonthStart = new Date(`${today.slice(0, 7)}-01T00:00:00`);
  const taxYearEndMonthStart = new Date(`${taxYearRange.end.slice(0, 7)}-01T00:00:00`);
  const monthDiff = (taxYearEndMonthStart.getFullYear() - currentMonthStart.getFullYear()) * 12
    + (taxYearEndMonthStart.getMonth() - currentMonthStart.getMonth());
  return Math.max(0, monthDiff);
}

function estimateMonthlyExpenseRunRate(transactions, today = getTodayDateString()) {
  const currentMonthStart = `${today.slice(0, 7)}-01`;
  const monthlyTotals = new Map();

  transactions.forEach((txn) => {
    if (txn.txnDate >= currentMonthStart) return;
    if (txn.currency !== "MUR") return;
    if (txn.bankName !== "MCB" && txn.bankName !== "SBM") return;
    monthlyTotals.set(txn.txnDate.slice(0, 7), (monthlyTotals.get(txn.txnDate.slice(0, 7)) || 0) + Number(txn.debit || 0));
  });

  const historyMonthKeys = [];
  const cursor = new Date(`${currentMonthStart}T00:00:00`);
  cursor.setMonth(cursor.getMonth() - 1);
  for (let index = 0; index < 36; index += 1) {
    historyMonthKeys.unshift(`${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`);
    cursor.setMonth(cursor.getMonth() - 1);
  }

  const historySeries = historyMonthKeys.map((monthKey) => monthlyTotals.get(monthKey) || 0);
  const observedValues = historySeries.filter((value) => value > 0);
  if (!observedValues.length) {
    return DEFAULT_EXPECTED_EXPENSES / 12;
  }
  if (observedValues.length < 6) {
    return average(observedValues);
  }

  const longTermBaseline = winsorizedAverage(historySeries, 0.1, 0.9);
  const recentMedian = median(historySeries.slice(-12));
  const trendProjection = projectSpendRunRate(historySeries);

  return Math.max(0, 0.5 * longTermBaseline + 0.3 * recentMedian + 0.2 * trendProjection);
}

function calculateIncomeTax(amount) {
  if (amount <= 500000) {
    return 0;
  }
  if (amount <= 1000000) {
    return (amount - 500000) * 0.1;
  }
  return 500000 * 0.1 + (amount - 1000000) * 0.2;
}

function isWithinTaxYear(dateString, taxYearRange = getCurrentTaxYearRange()) {
  return Boolean(dateString) && dateString >= taxYearRange.start && dateString <= taxYearRange.end;
}

function getEffectiveTaxReceiptDate(entry) {
  if (entry.dateReceived) {
    return entry.dateReceived;
  }
  if (!entry.invoicedDate) {
    return "";
  }

  const invoiceDate = new Date(`${entry.invoicedDate}T00:00:00`);
  invoiceDate.setMonth(invoiceDate.getMonth() + 1);
  return invoiceDate.toISOString().slice(0, 10);
}

function getLastTaxEntry() {
  return state.taxEntries.at(-1) || null;
}

function getNextInvoiceWorkingDate(previousInvoiceDate) {
  const previousDate = parseDateString(previousInvoiceDate);
  if (!previousDate) {
    return "";
  }

  const year = previousDate.getFullYear();
  const month = previousDate.getMonth() + 1;
  const targetMonthDate = new Date(year, month + 1, 0);

  while (!isMauritiusWorkingDay(targetMonthDate)) {
    targetMonthDate.setDate(targetMonthDate.getDate() - 1);
  }

  return formatDateInputValue(targetMonthDate);
}

function countWorkingDaysBetweenInvoices(previousInvoiceDate, nextInvoiceDate) {
  const startDate = parseDateString(previousInvoiceDate);
  const endDate = parseDateString(nextInvoiceDate);
  if (!startDate || !endDate || endDate <= startDate) {
    return 0;
  }

  const cursor = new Date(startDate);
  cursor.setDate(cursor.getDate() + 1);
  let workingDays = 0;
  while (cursor <= endDate) {
    if (isMauritiusWorkingDay(cursor)) {
      workingDays += 1;
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return workingDays;
}

function isMauritiusWorkingDay(date) {
  return !isWeekend(date) && !isMauritiusPublicHoliday(date);
}

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function isMauritiusPublicHoliday(date) {
  const holidays = MAURITIUS_PUBLIC_HOLIDAYS[date.getFullYear()];
  return holidays ? holidays.has(formatDateInputValue(date)) : false;
}

function parseDateString(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) {
    return null;
  }
  const parsed = new Date(`${value}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatTaxRate(value) {
  return Number(value).toFixed(3).replace(/\.?0+$/, "");
}

function getTodayDateString() {
  return new Date().toISOString().slice(0, 10);
}

function roundMur(value) {
  return Math.round(Number(value || 0));
}

async function persistTaxEntries() {
  try {
    await setSetting(TAX_ENTRIES_KEY, state.taxEntries);
  } catch (error) {
    logMessage(`Could not save tax entries: ${error.message}`);
  }
}

async function persistTaxExpectedExpenses() {
  try {
    await setSetting(TAX_EXPECTED_EXPENSES_KEY, state.taxExpectedExpenses);
    await setSetting(TAX_EXPECTED_EXPENSES_MODE_KEY, state.taxExpectedExpensesMode);
  } catch (error) {
    logMessage(`Could not save tax expected expenses: ${error.message}`);
  }
}

function updateStatus(text) {
  if (!els.ledgerStatus) return;
  if (state.ledgerName) {
    els.ledgerStatus.textContent = `Ledger file ready: ${state.ledgerName}. ${text}`;
    return;
  }
  els.ledgerStatus.textContent = text;
}

function updateLedgerStatus(text) {
  els.ledgerStatus.textContent = text;
}

function downloadLedgerSnapshot(text, fileName) {
  const blob = new Blob([text], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function logMessage(text) {
  const timestamp = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  const line = document.createElement("div");
  line.textContent = `[${timestamp}] ${text}`;
  els.importLog.prepend(line);
}

function moneyFormat(value) {
  return Number(value || 0).toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  });
}

function toInsightAmount(amount, currency) {
  const numericAmount = Number(amount || 0);
  if (currency === "GBP") {
    return numericAmount * getGbpToMurRate();
  }
  return numericAmount;
}

function toInsightNet(txn) {
  return toInsightAmount(txn.credit, txn.currency) - toInsightAmount(txn.debit, txn.currency);
}

function buildInsightConversionNote(transactions) {
  const hasGbp = transactions.some((txn) => txn.currency === "GBP");
  if (!hasGbp) return "";
  if (Number.isFinite(state.gbpToMurRate)) {
    const rateDate = state.gbpToMurDate ? ` (${formatRateTimestamp(state.gbpToMurDate)})` : "";
    return `GBP converted to MUR at ${moneyFormat(state.gbpToMurRate)}${rateDate} for insights`;
  }
  return "GBP to MUR rate unavailable, using raw GBP amounts until rate loads";
}

function getGbpToMurRate() {
  if (Number.isFinite(state.gbpToMurRate)) {
    return state.gbpToMurRate;
  }
  return 1;
}

function numberFormat(value) {
  return Number(value || 0).toLocaleString();
}

function compactMoneyFormat(value) {
  return Number(value || 0).toLocaleString(undefined, {
    notation: "compact",
    maximumFractionDigits: 1,
  });
}

function formatSignedMoney(value) {
  const numeric = Number(value || 0);
  const absolute = moneyFormat(Math.abs(numeric));
  return `${numeric >= 0 ? "+" : "-"}${absolute}`;
}

function formatPercent(value) {
  if (!Number.isFinite(value)) return "0%";
  return `${Math.round(value)}%`;
}

function formatGapLabel(value, suffix = "") {
  const numeric = Number(value || 0);
  const prefix = numeric >= 0 ? `${formatSignedMoney(numeric)} ahead` : `${formatSignedMoney(numeric)} behind`;
  return suffix ? `${prefix} ${suffix}` : prefix;
}

function formatDateTime(value) {
  return new Date(value).toLocaleString();
}

function formatRateTimestamp(value) {
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? value : parsed.toLocaleString();
}

function shiftDate(dateString, days) {
  const date = new Date(`${dateString}T00:00:00`);
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function formatMonthShort(value) {
  const [year, month] = String(value).split("-");
  return new Date(Number(year), Number(month) - 1, 1).toLocaleString(undefined, {
    month: "short",
  });
}

function categorizeTransaction(txn) {
  const description = txn?.description || "";
  const reference = txn?.reference || "";
  const text = `${description} ${reference}`.toUpperCase();
  if (isSqliSalaryTransfer(txn)) {
    return "Salary";
  }
  const patterns = [
    ["DELIVOO", "Food Delivery"],
    ["KOON PO YUEN", "Dining / Restaurants"],
    ["HOT & POT", "Dining / Restaurants"],
    ["KENTUCKY FRIED CHICKEN", "Dining / Restaurants"],
    ["K.F.C", "Dining / Restaurants"],
    ["BLUE JUICE", "Dining / Restaurants"],
    ["NINE DRAGONS", "Dining / Restaurants"],
    ["STEERS", "Dining / Restaurants"],
    ["LITTLE ITA", "Dining / Restaurants"],
    ["LITTLE KING", "Dining / Restaurants"],
    ["MOVIE STORE", "Dining / Restaurants"],
    ["JANGO S CAFE", "Dining / Restaurants"],
    ["CHILLAX", "Dining / Restaurants"],
    ["SEN & KEN", "Dining / Restaurants"],
    ["DEBONNAIRS", "Dining / Restaurants"],
    ["RATATOUILLE", "Dining / Restaurants"],
    ["RESTOWAY", "Dining / Restaurants"],
    ["ARTISAN COFFEE", "Dining / Restaurants"],
    ["EL MONDO", "Dining / Restaurants"],
    ["NANDO'S", "Dining / Restaurants"],
    ["SCOTT", "Shopping"],
    ["WINNERS", "Shopping"],
    ["PRICE GURU", "Shopping"],
    ["ONEOONE MULTIMEDIA", "Shopping"],
    ["PIONEER MARKETING&TRADING", "Shopping"],
    ["INTERMART", "Groceries"],
    ["EBENE PHARMACY", "Pharmacy / Medical"],
    ["MEDACTIV", "Pharmacy / Medical"],
    ["PHARMACIE ST JEAN", "Pharmacy / Medical"],
    ["GALIEN PHARMACY", "Pharmacy / Medical"],
    ["SMS TOPUP", "Mobile / Telecom"],
    ["SMS TOP UP", "Mobile / Telecom"],
    ["EBANKING TOPUP", "Mobile / Telecom"],
    ["TELECOM", "Mobile / Telecom"],
    ["ICMARKET", "Trading / Broker"],
    ["MQL5", "Trading / Broker"],
    ["ICM PAY", "Trading / Broker"],
    ["IC MARKET", "Trading / Broker"],
    ["ZULU REPAY", "Trading / Broker"],
    ["MT4 REPAY", "Trading / Broker"],
    ["MISSING MT4", "Trading / Broker"],
    ["MRA", "Taxes / Government"],
    ["CSG", "Taxes / Government"],
    ["STATEINSURANCE", "Insurance"],
    ["LAPRUDENCE", "Insurance"],
    ["CREDIT INTEREST", "Interest"],
    ["VISA CARD PAYMENT", "Card Repayment"],
    ["ATM CASH WITHDRAWAL", "Cash Withdrawal"],
    ["ATM WITHDRAWAL", "Cash Withdrawal"],
    ["CASH WITHDRAWAL", "Cash Withdrawal"],
    ["SI EXECUTION CHARGE", "Bank Fees"],
    ["SI SC DR", "Bank Fees"],
    ["CHARGE AC-", "Bank Fees"],
    ["PAYPAL CHARGES", "Bank Fees"],
    ["INWARD TRANSFER", "Transfers In"],
    ["GBP TO SBM", "Transfers Between Own Accounts"],
    ["MCB TO SBM", "Transfers Between Own Accounts"],
    ["JUICE ACCOUNT TRANSFER", "Transfers"],
    ["JUICE TRANSFER", "Transfers"],
    ["BILL REFUND TO -", "Transfers"],
    ["BILL REFUND TO /", "Transfers"],
    ["LENDING TO -", "Transfers"],
    ["REFUND/", "Transfers"],
  ];
  for (const [needle, label] of patterns) {
    if (text.includes(needle)) return label;
  }
  return "Other";
}

function isSalaryTransaction(txn) {
  const text = `${txn.description || ""} ${txn.reference || ""}`.toUpperCase();
  const hasSalaryKeyword = /(SALARY|PAYROLL|WAGE|REMUNERATION)/.test(text);
  return txn.credit > 0 && (hasSalaryKeyword || isSqliSalaryTransfer(txn));
}

function isTradingTransaction(txn) {
  return categorizeTransaction(txn) === "Trading / Broker";
}

function isTransferTransaction(txn) {
  if (isSqliSalaryTransfer(txn)) return false;
  const category = categorizeTransaction(txn);
  if (category.startsWith("Transfers")) return true;
  const text = `${txn.description || ""} ${txn.reference || ""}`.toUpperCase();
  return text.includes("TRANSFER");
}

function isOwnAccountTransferTransaction(txn) {
  return categorizeTransaction(txn) === "Transfers Between Own Accounts";
}

function isSqliSalaryTransfer(txn) {
  const text = `${txn?.description || ""} ${txn?.reference || ""}`.toUpperCase();
  return txn?.credit > 0
    && txn?.bankName === "MCB"
    && txn?.currency === "GBP"
    && text.includes("INWARD TRANSFER")
    && text.includes("SQLI UK LIMIT");
}

function getEffectiveDateRange() {
  if (els.fromDate.value || els.toDate.value) {
    return {
      fromDate: els.fromDate.value,
      toDate: els.toDate.value,
    };
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const preset = state.quickFilters.datePreset;
  if (preset === "this-month") {
    return {
      fromDate: formatDateInputValue(new Date(today.getFullYear(), today.getMonth(), 1)),
      toDate: formatDateInputValue(today),
    };
  }
  if (preset === "last-3-months") {
    return {
      fromDate: formatDateInputValue(new Date(today.getFullYear(), today.getMonth() - 2, 1)),
      toDate: formatDateInputValue(today),
    };
  }
  if (preset === "this-year") {
    return {
      fromDate: formatDateInputValue(new Date(today.getFullYear(), 0, 1)),
      toDate: formatDateInputValue(today),
    };
  }
  if (preset === "current-tax-year") {
    const taxYearRange = getCurrentTaxYearRange(formatDateInputValue(today));
    return {
      fromDate: taxYearRange.start,
      toDate: taxYearRange.end,
    };
  }
  if (preset === "last-year") {
    return {
      fromDate: formatDateInputValue(new Date(today.getFullYear() - 1, 0, 1)),
      toDate: formatDateInputValue(new Date(today.getFullYear() - 1, 11, 31)),
    };
  }
  return { fromDate: "", toDate: "" };
}

function formatDateInputValue(date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
}

function renderAccountBadge(accountLabel) {
  const match = String(accountLabel || "").match(/^([A-Z]+)\s+(.+)\s+\(([^)]+)\)$/);
  if (!match) {
    return escapeHtml(accountLabel || "");
  }
  const [, bankName, accountNumber, currency] = match;
  return `
    <span class="account-badge">
      <span class="account-bank account-bank-${escapeHtml(bankName.toLowerCase())}">${escapeHtml(bankName)}</span>
      <span class="account-number">${escapeHtml(accountNumber)}</span>
      <span class="currency-badge currency-${escapeHtml(currency.toLowerCase())}">${escapeHtml(currency)}</span>
    </span>
  `;
}

function csvEscape(value) {
  const stringValue = String(value ?? "");
  if (/[",\n]/.test(stringValue)) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }
  return stringValue;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
