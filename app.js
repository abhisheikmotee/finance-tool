const DB_NAME = "finance-ledger-browser-db";
const DB_VERSION = 2;
const SETTINGS_STORE = "settings";
const GBP_TO_MUR_RATE_KEY = "gbpToMurRate";
const GBP_TO_MUR_API_URL = "https://open.er-api.com/v6/latest/GBP";
const TAX_ENTRIES_KEY = "taxEntries";
const TAX_EXPECTED_EXPENSES_KEY = "taxExpectedExpenses";
const TAX_YEAR_START = "2025-07-01";
const TAX_YEAR_END = "2026-06-30";
const DEFAULT_EXPECTED_EXPENSES = 2500000;
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
  imports: [],
  taxEntries: [],
  taxExpectedExpenses: DEFAULT_EXPECTED_EXPENSES,
  currentPage: 1,
  pageSize: 25,
  ledgerHandle: null,
  ledgerName: "",
  saveMode: "download",
  gbpToMurRate: null,
  gbpToMurDate: "",
  quickFilters: {
    datePreset: "all",
    salaryOnly: false,
    excludeTransfers: false,
    bankOnly: "all",
  },
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
  els.exportBtn = document.getElementById("export-btn");
  els.resetBtn = document.getElementById("reset-btn");
  els.chooseLedgerBtn = document.getElementById("choose-ledger-btn");
  els.newLedgerBtn = document.getElementById("new-ledger-btn");
  els.saveLedgerBtn = document.getElementById("save-ledger-btn");
  els.dropZone = document.getElementById("drop-zone");
  els.importStatus = document.getElementById("import-status");
  els.ledgerStatus = document.getElementById("ledger-status");
  els.importLog = document.getElementById("import-log");
  els.metricsGrid = document.getElementById("metrics-grid");
  els.metricsLastImport = document.getElementById("metrics-last-import");
  els.trendMetricsGrid = document.getElementById("trend-metrics-grid");
  els.searchInput = document.getElementById("search-input");
  els.accountFilter = document.getElementById("account-filter");
  els.currencyFilter = document.getElementById("currency-filter");
  els.quickFilterChips = document.getElementById("quick-filter-chips");
  els.fromDate = document.getElementById("from-date");
  els.toDate = document.getElementById("to-date");
  els.pageSize = document.getElementById("page-size");
  els.transactionsBody = document.getElementById("transactions-body");
  els.monthlySummaryMetrics = document.getElementById("monthly-summary-metrics");
  els.monthlySummaryBody = document.getElementById("monthly-summary-body");
  els.categorySummaryBody = document.getElementById("category-summary-body");
  els.forecastSummaryBody = document.getElementById("forecast-summary-body");
  els.trendlineChart = document.getElementById("trendline-chart");
  els.trendlineSummary = document.getElementById("trendline-summary");
  els.taxBody = document.getElementById("tax-body");
  els.taxSummaryBody = document.getElementById("tax-summary-body");
  els.addTaxRowBtn = document.getElementById("add-tax-row-btn");
  els.tableCount = document.getElementById("table-count");
  els.prevPage = document.getElementById("prev-page");
  els.nextPage = document.getElementById("next-page");
  els.pageIndicator = document.getElementById("page-indicator");
  els.paginationBar = document.getElementById("pagination-bar");
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
  els.exportBtn.addEventListener("click", exportTransactionsCsv);
  els.resetBtn.addEventListener("click", resetLedgerData);

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

  [els.searchInput, els.accountFilter, els.currencyFilter, els.fromDate, els.toDate].forEach((input) => {
    input.addEventListener("input", onFilterChange);
    input.addEventListener("change", onFilterChange);
  });

  els.quickFilterChips.addEventListener("click", handleQuickFilterChipClick);

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
    const totalPages = Math.max(1, Math.ceil(state.filteredTransactions.length / state.pageSize));
    if (state.currentPage < totalPages) {
      state.currentPage += 1;
      renderTransactionsTable({ preservePaginationPosition: true });
    }
  });

  els.addTaxRowBtn.addEventListener("click", async () => {
    state.taxEntries.push(createEmptyTaxEntry());
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

function handleQuickFilterChipClick(event) {
  const button = event.target.closest("[data-chip]");
  if (!button) return;

  const { chip } = button.dataset;
  if (chip === "this-month" || chip === "last-3-months" || chip === "this-year") {
    state.quickFilters.datePreset = state.quickFilters.datePreset === chip ? "all" : chip;
  } else if (chip === "salary-only") {
    state.quickFilters.salaryOnly = !state.quickFilters.salaryOnly;
  } else if (chip === "transfers-excluded") {
    state.quickFilters.excludeTransfers = !state.quickFilters.excludeTransfers;
  } else if (chip === "sbm-only" || chip === "mcb-only") {
    state.quickFilters.bankOnly = state.quickFilters.bankOnly === chip ? "all" : chip;
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
    updateLedgerStatus("Choose or create a ledger file to persist data on disk.");
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

  try {
    const savedExpectedExpenses = await getSetting(TAX_EXPECTED_EXPENSES_KEY);
    if (savedExpectedExpenses !== null && savedExpectedExpenses !== "") {
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
  state.taxExpectedExpenses = DEFAULT_EXPECTED_EXPENSES;
  state.currentPage = 1;
  state.ledgerHandle = null;
  state.ledgerName = "finance-ledger.json";
  state.saveMode = supportsFileSystemAccess() ? "download" : "download";
  await persistTaxEntries();
  await persistTaxExpectedExpenses();
  applyFilters();
  renderAll();
  updateLedgerStatus(`New ledger ready: ${state.ledgerName}. Click Save to download it.`);
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
      updateLedgerStatus(`Ledger file: ${state.ledgerName}`);
      logMessage(`Write access granted for ${state.ledgerName}. Save will overwrite this file.`);
    } else {
      updateLedgerStatus(`Ledger file: ${state.ledgerName}. Click Save to retry permission.`);
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
  state.taxExpectedExpenses = Number(data.taxExpectedExpenses) || DEFAULT_EXPECTED_EXPENSES;
  await persistTaxEntries();
  await persistTaxExpectedExpenses();
  if (storeHandle) {
    await setSetting("ledgerHandle", handle);
  }
  updateLedgerStatus(`Ledger file: ${state.ledgerName}`);
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
  state.taxExpectedExpenses = Number(data.taxExpectedExpenses) || DEFAULT_EXPECTED_EXPENSES;
  await persistTaxEntries();
  await persistTaxExpectedExpenses();
  updateLedgerStatus(`Opened ledger: ${state.ledgerName}. Save will download the updated file.`);
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
      updateLedgerStatus(`Saved to ${state.ledgerName}`);
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
  const currency = els.currencyFilter.value || "all";
  const { fromDate, toDate } = getEffectiveDateRange();
  const { salaryOnly, excludeTransfers, bankOnly } = state.quickFilters;

  state.filteredTransactions = state.transactions.filter((txn) => {
    const matchesQuery = !query || [txn.description, txn.reference, txn.sourceFile, txn.accountLabel]
      .some((value) => value.toLowerCase().includes(query));
    const matchesAccount = account === "all" || txn.accountLabel === account;
    const matchesCurrency = currency === "all" || txn.currency === currency;
    const matchesFromDate = !fromDate || txn.txnDate >= fromDate;
    const matchesToDate = !toDate || txn.txnDate <= toDate;
    const matchesSalary = !salaryOnly || isSalaryTransaction(txn);
    const matchesTransfer = !excludeTransfers || !isTransferTransaction(txn);
    const matchesBank = bankOnly === "all"
      || (bankOnly === "sbm-only" && txn.bankName === "SBM")
      || (bankOnly === "mcb-only" && txn.bankName === "MCB");
    return matchesQuery
      && matchesAccount
      && matchesCurrency
      && matchesFromDate
      && matchesToDate
      && matchesSalary
      && matchesTransfer
      && matchesBank;
  });
}

function renderAll() {
  renderLedgerStatus();
  renderFilterOptions();
  renderQuickFilterChips();
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
    "salary-only": state.quickFilters.salaryOnly,
    "transfers-excluded": state.quickFilters.excludeTransfers,
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
    updateLedgerStatus(`Ledger file: ${state.ledgerName}`);
  } else if (supportsFileSystemAccess()) {
    updateLedgerStatus("Choose or create a ledger file. Data is saved outside browser cache.");
  }
}

function renderFilterOptions() {
  const currentAccount = els.accountFilter.value || "all";
  const currentCurrency = els.currencyFilter.value || "all";
  const accounts = Array.from(new Set(state.transactions.map((txn) => txn.accountLabel))).sort();
  const currencies = Array.from(new Set(state.transactions.map((txn) => txn.currency))).sort();
  populateSelect(els.accountFilter, ["all", ...accounts], "All Accounts", currentAccount);
  populateSelect(els.currencyFilter, ["all", ...currencies], "All Currencies", currentCurrency);
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

function renderMetrics() {
  const latestImport = [...state.imports].sort((a, b) => b.importedAt.localeCompare(a.importedAt))[0];
  const balances = summarizeAccountBalances(state.transactions);
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
      value: `${moneyFormat(balances.mcbGbp)} GBP`,
      subtext: `${moneyFormat(balances.mcbGbpInMur)} MUR equivalent`,
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
  const totalRows = state.filteredTransactions.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / state.pageSize));
  if (state.currentPage > totalPages) state.currentPage = totalPages;

  const start = (state.currentPage - 1) * state.pageSize;
  const end = start + state.pageSize;
  const pageRows = state.filteredTransactions.slice(start, end);

  els.tableCount.textContent = `${numberFormat(totalRows)} row${totalRows === 1 ? "" : "s"}`;
  els.pageIndicator.textContent = `Page ${state.currentPage} of ${totalPages}`;
  els.prevPage.disabled = state.currentPage === 1;
  els.nextPage.disabled = state.currentPage === totalPages;

  if (!pageRows.length) {
    els.transactionsBody.innerHTML = `<tr><td colspan="8" class="empty-state">No transactions match the current filters.</td></tr>`;
    restorePaginationPosition(previousPaginationTop);
    return;
  }

  els.transactionsBody.innerHTML = pageRows.map((txn) => `
    <tr>
      <td>${escapeHtml(txn.txnDate)}</td>
      <td>${escapeHtml(txn.accountLabel)}</td>
      <td>${escapeHtml(txn.description)}</td>
      <td class="${txn.debit > 0 ? "amount-negative" : ""}">${moneyFormat(txn.debit)}</td>
      <td class="${txn.credit > 0 ? "amount-positive" : ""}">${moneyFormat(txn.credit)}</td>
      <td>${moneyFormat(txn.balance)}</td>
    </tr>
  `).join("");

  restorePaginationPosition(previousPaginationTop);
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

  const currentBalance = Array.from(latestBalanceByAccount.values())
    .reduce((sum, txn) => sum + toInsightAmount(txn.balance, txn.currency), 0);
  const summaryTiles = [
    {
      label: "Sum Of All Debit",
      value: moneyFormat(totalDebit),
      subtext: "Visible outgoing total",
      toneClass: "is-negative",
    },
    {
      label: "Sum Of All Credit",
      value: moneyFormat(totalCredit),
      subtext: "Visible incoming total",
      toneClass: "is-positive",
    },
    {
      label: "Net Cash Flow",
      value: moneyFormat(totalCredit - totalDebit),
      subtext: "Credit minus debit",
      toneClass: totalCredit - totalDebit >= 0 ? "is-positive" : "is-negative",
    },
    {
      label: "Current Balance",
      value: moneyFormat(currentBalance),
      subtext: "Latest visible balance by account",
      toneClass: currentBalance >= 0 ? "is-positive" : "is-negative",
    },
  ];

  els.monthlySummaryMetrics.innerHTML = summaryTiles.map((metric) => `
    <article class="metric-tile metric-tile-inline ${metric.toneClass}">
      <div class="metric-label">${escapeHtml(metric.label)}</div>
      <div class="metric-value">${escapeHtml(metric.value)}</div>
      <div class="metric-subtext">${escapeHtml(metric.subtext)}</div>
    </article>
  `).join("");

  const rows = Array.from(summaryMap.values())
    .sort((a, b) => a.month.localeCompare(b.month) || a.accountLabel.localeCompare(b.accountLabel));

  if (!rows.length) {
    els.monthlySummaryBody.innerHTML = `<tr><td colspan="5" class="empty-state">No monthly insights for the current filters.</td></tr>`;
    return;
  }

  els.monthlySummaryBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.month)}</td>
      <td>${escapeHtml(row.accountLabel)}</td>
      <td class="amount-negative">${moneyFormat(row.totalDebit)}</td>
      <td class="amount-positive">${moneyFormat(row.totalCredit)}</td>
      <td class="${row.totalCredit - row.totalDebit >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(row.totalCredit - row.totalDebit)}</td>
    </tr>
  `).join("");
}

function renderTrendInsights() {
  const filtered = state.filteredTransactions;
  const debits = filtered.filter((txn) => toInsightAmount(txn.debit, txn.currency) > 0);
  const monthlyMap = new Map();
  const categoryMap = new Map();
  const forecastMap = new Map();
  const conversionNote = buildInsightConversionNote(filtered);

  filtered.forEach((txn) => {
    const month = txn.txnDate.slice(0, 7);
    const monthEntry = monthlyMap.get(month) || { debit: 0, credit: 0 };
    monthEntry.debit += toInsightAmount(txn.debit, txn.currency);
    monthEntry.credit += toInsightAmount(txn.credit, txn.currency);
    monthlyMap.set(month, monthEntry);

    const forecastEntry = forecastMap.get(txn.accountLabel) || {
      accountLabel: txn.accountLabel,
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

  const lastTxnDate = filtered.length ? filtered[filtered.length - 1].txnDate : "";
  const rollingStart = lastTxnDate ? shiftDate(lastTxnDate, -30) : "";
  filtered.forEach((txn) => {
    if (rollingStart && txn.txnDate >= rollingStart) {
      const forecastEntry = forecastMap.get(txn.accountLabel);
      forecastEntry.net30 += toInsightNet(txn);
    }
    if (toInsightAmount(txn.debit, txn.currency) > 0) {
      const category = categorizeDescription(txn.description);
      const categoryEntry = categoryMap.get(category) || { category, count: 0, total: 0 };
      categoryEntry.count += 1;
      categoryEntry.total += toInsightAmount(txn.debit, txn.currency);
      categoryMap.set(category, categoryEntry);
    }
  });

  const monthlyRows = Array.from(monthlyMap.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  const monthlyDebitSeries = monthlyRows.map(([_, value]) => value.debit);
  const monthlyCreditSeries = monthlyRows.map(([_, value]) => value.credit);
  const netSeries = monthlyRows.map(([_, value]) => value.credit - value.debit);
  const avgMonthlyDebit = monthlyDebitSeries.length ? monthlyDebitSeries.reduce((sum, v) => sum + v, 0) / monthlyDebitSeries.length : 0;
  const avgMonthlyCredit = monthlyCreditSeries.length ? monthlyCreditSeries.reduce((sum, v) => sum + v, 0) / monthlyCreditSeries.length : 0;
  const avgMonthlyNet = netSeries.length ? netSeries.reduce((sum, v) => sum + v, 0) / netSeries.length : 0;
  const biggestDebit = debits.length ? Math.max(...debits.map((txn) => toInsightAmount(txn.debit, txn.currency))) : 0;
  const projectionYear = detectProjectionYear(filtered);
  const forecastRows = Array.from(forecastMap.values())
    .sort((a, b) => a.accountLabel.localeCompare(b.accountLabel))
    .map((row) => {
      const endOfYear = projectAccountYearEnd(row, filtered, projectionYear);
      return {
        ...row,
        endOfYear,
      };
    });
  const totalCurrentBalance = forecastRows.reduce((sum, row) => sum + row.lastBalance, 0);
  const projectedClosing = forecastRows.length
    ? forecastRows.reduce((sum, row) => sum + row.endOfYear, 0)
    : totalCurrentBalance;
  const yearlyProjection = buildYearlyProjection(filtered, totalCurrentBalance, projectedClosing);

  const trendTiles = [
    {
      label: "Average Monthly Debit",
      value: moneyFormat(avgMonthlyDebit),
      subtext: conversionNote || "Average outgoing total across visible months",
    },
    {
      label: "Average Monthly Credit",
      value: moneyFormat(avgMonthlyCredit),
      subtext: conversionNote || "Average incoming total across visible months",
    },
    {
      label: "Average Monthly Net Flow",
      value: moneyFormat(avgMonthlyNet),
      subtext: conversionNote || "Average credit minus debit across visible months",
    },
    {
      label: "Largest Debit Seen",
      value: moneyFormat(biggestDebit),
      subtext: conversionNote || "Highest single outgoing transaction in current view",
    },
  ];

  els.trendMetricsGrid.innerHTML = trendTiles.map((metric) => `
    <article class="metric-tile">
      <div class="metric-label">${escapeHtml(metric.label)}</div>
      <div class="metric-value">${escapeHtml(metric.value)}</div>
      <div class="metric-subtext">${escapeHtml(metric.subtext)}</div>
    </article>
  `).join("");

  els.trendlineSummary.textContent = yearlyProjection.summaryText;
  renderTrendlineChart(yearlyProjection);

  const categoryRows = Array.from(categoryMap.values()).sort((a, b) => b.total - a.total).slice(0, 8);
  els.categorySummaryBody.innerHTML = categoryRows.length ? categoryRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.category)}</td>
      <td>${numberFormat(row.count)}</td>
      <td class="amount-negative">${moneyFormat(row.total)}</td>
      <td>${moneyFormat(row.total / Math.max(1, row.count))}</td>
    </tr>
  `).join("") : `<tr><td colspan="4" class="empty-state">No spending categories for the current filters.</td></tr>`;

  els.forecastSummaryBody.innerHTML = forecastRows.length ? forecastRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.accountLabel)}</td>
      <td>${moneyFormat(row.lastBalance)}</td>
      <td class="${row.net30 >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(row.net30)}</td>
      <td class="${row.endOfYear >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(row.endOfYear)}</td>
    </tr>
  `).join("") : `<tr><td colspan="4" class="empty-state">No forecast rows for the current filters.</td></tr>`;
}

function buildYearlyProjection(transactions, totalCurrentBalance = 0, projectedClosingBalance = 0) {
  const targetYear = detectProjectionYear(transactions);
  const yearTransactions = transactions.filter((txn) => txn.txnDate.startsWith(`${targetYear}-`));
  const latestDate = yearTransactions.length ? yearTransactions[yearTransactions.length - 1].txnDate : "";
  const lastMonthIndex = latestDate ? Number(latestDate.slice(5, 7)) : 0;
  const monthsElapsed = Math.max(lastMonthIndex, 0);
  const monthsRemaining = Math.max(0, 12 - monthsElapsed);
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
  const projectedNetValues = projectFutureMonthlyValues(actualNetSeries, monthsElapsed, monthsRemaining);
  const actualNetTotal = actualNetSeries.reduce((sum, point) => sum + point.y, 0);
  const projectedNetTotal = projectedNetValues.reduce((sum, value) => sum + value, 0);
  const impliedStartingBalance = totalCurrentBalance - actualNetTotal;
  const desiredProjectedNet = projectedClosingBalance - totalCurrentBalance;
  const scalingFactor = projectedNetTotal && Number.isFinite(desiredProjectedNet)
    ? desiredProjectedNet / projectedNetTotal
    : 1;
  const scaledProjectedNetValues = projectedNetValues.map((value) => value * scalingFactor);
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
    const absoluteMonthIndex = monthsElapsed + i;
    const month = `${targetYear}-${String(absoluteMonthIndex).padStart(2, "0")}`;
    runningBalance += scaledProjectedNetValues[i - 1] ?? 0;
    projectionSeries.push({
      label: formatMonthShort(month),
      month,
      x: absoluteMonthIndex,
      monthNumber: absoluteMonthIndex,
      y: runningBalance,
      forecast: true,
    });
  }

  const visibleMonths = actualSeries.length;
  const summarySubtext = visibleMonths
    ? `Balance path from Jan to ${formatMonthShort(`${targetYear}-${String(monthsElapsed || 1).padStart(2, "0")}`)} ${targetYear}, including quiet months as zero`
    : `Waiting for ${targetYear} transactions to estimate balance growth`;
  const balanceSubtext = yearTransactions.length
    ? `Projected total balance across visible accounts through Dec ${targetYear}`
    : "Import current-year statements to estimate a closing balance";
  const summaryText = visibleMonths
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

  els.trendlineChart.innerHTML = `
    ${gridValues.map((value) => `
      <g>
        <line class="trend-gridline ${Math.abs(value) < 0.0001 ? "trend-zero-line" : ""}" x1="${pad.left}" y1="${yForValue(value)}" x2="${width - pad.right}" y2="${yForValue(value)}"></line>
        <text class="trend-label" x="10" y="${yForValue(value) + 4}">${escapeHtml(compactMoneyFormat(value))}</text>
      </g>
    `).join("")}
    <line class="trend-axis" x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}"></line>
    ${actualPath ? `<path class="trend-actual" d="${actualPath}"></path>` : ""}
    ${projection.projectionSeries.length ? `<path class="trend-projection" d="${projectionPath}"></path>` : ""}
    ${allPoints.map((point, index) => `
      <g>
        <circle class="trend-point ${point.forecast ? "forecast" : ""}" cx="${xForIndex(index)}" cy="${yForValue(point.y)}" r="4.5"></circle>
        <text class="trend-label" x="${xForIndex(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(point.label)}</text>
      </g>
    `).join("")}
    <text class="trend-value-label" x="${width - pad.right}" y="${yForValue(allPoints[allPoints.length - 1].y) - 10}" text-anchor="end">
      ${escapeHtml(`Year-end trend ${compactMoneyFormat(allPoints[allPoints.length - 1].y)}`)}
    </text>
  `;
}

function detectProjectionYear(transactions) {
  if (!transactions.length) {
    return new Date().getFullYear();
  }
  return Number(transactions[transactions.length - 1].txnDate.slice(0, 4));
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
  if (!state.transactions.length) {
    logMessage("Nothing to export yet.");
    return;
  }

  const headers = ["Txn Date", "Value Date", "Account", "Currency", "Description", "Debit", "Credit", "Balance", "Reference", "Source File"];
  const rows = state.transactions.map((txn) => [
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
  link.download = "finance_transactions_export.csv";
  link.click();
  URL.revokeObjectURL(url);
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
  state.taxEntries = [];
  state.taxExpectedExpenses = DEFAULT_EXPECTED_EXPENSES;
  state.currentPage = 1;
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
}

function renderTaxSummary() {
  const summary = computeTaxSummary();
  els.taxSummaryBody.innerHTML = `
    <tr>
      <td>Expected Income Year Jul 2025-Jun 2026</td>
      <td class="tax-summary-value">${moneyFormat(summary.expectedIncome)}</td>
      <td>So far Income Year Jul 2025-Jun 2026</td>
      <td class="tax-summary-value">${moneyFormat(summary.soFarIncome)}</td>
    </tr>
    <tr>
      <td>Expected Income Tax</td>
      <td class="tax-summary-value">${moneyFormat(summary.expectedIncomeTax)}</td>
      <td>So far Income Tax</td>
      <td class="tax-summary-value">${moneyFormat(summary.soFarIncomeTax)}</td>
    </tr>
    <tr>
      <td>Expected CSG</td>
      <td class="tax-summary-value">${moneyFormat(summary.expectedCsg)}</td>
      <td>So far CSG</td>
      <td class="tax-summary-value">${moneyFormat(summary.soFarCsg)}</td>
    </tr>
    <tr>
      <td>Expected Expenses</td>
      <td><input id="tax-expected-expenses-input" class="tax-summary-input" type="number" min="0" step="0.01" value="${escapeHtml(String(state.taxExpectedExpenses))}"></td>
      <td>So far Expenses</td>
      <td class="tax-summary-value">${moneyFormat(summary.soFarExpenses)}</td>
    </tr>
    <tr>
      <td class="tax-summary-label-strong">Total Amount Saved</td>
      <td class="tax-summary-value tax-summary-value-strong ${summary.expectedSaved >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(summary.expectedSaved)}</td>
      <td class="tax-summary-label-strong">Total Amount Saved</td>
      <td class="tax-summary-value tax-summary-value-strong ${summary.soFarSaved >= 0 ? "amount-positive" : "amount-negative"}">${moneyFormat(summary.soFarSaved)}</td>
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

  state.taxEntries[index][field] = target.value;
  updateTaxComputedRow(index);
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

function handleTaxSummaryInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || target.id !== "tax-expected-expenses-input") {
    return;
  }

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
  const entriesInWindow = state.taxEntries.filter((entry) => isWithinTaxYear(getEffectiveTaxReceiptDate(entry)));
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
  const soFarExpenses = state.transactions
    .filter((txn) =>
      txn.txnDate >= TAX_YEAR_START &&
      txn.txnDate <= getTodayDateString() &&
      txn.currency === "MUR" &&
      (txn.bankName === "MCB" || txn.bankName === "SBM")
    )
    .reduce((sum, txn) => sum + Number(txn.debit || 0), 0);
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

function calculateIncomeTax(amount) {
  if (amount <= 500000) {
    return 0;
  }
  if (amount <= 1000000) {
    return (amount - 500000) * 0.1;
  }
  return 500000 * 0.1 + (amount - 1000000) * 0.2;
}

function isWithinTaxYear(dateString) {
  return Boolean(dateString) && dateString >= TAX_YEAR_START && dateString <= TAX_YEAR_END;
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
  } catch (error) {
    logMessage(`Could not save tax expected expenses: ${error.message}`);
  }
}

function updateStatus(text) {
  els.importStatus.textContent = text;
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

function categorizeDescription(description) {
  const text = String(description || "").toUpperCase();
  const patterns = [
    ["DELIVOO", "Food Delivery"],
    ["SCOTT", "Shopping"],
    ["WINNERS", "Shopping"],
    ["ICMARKET", "Trading / Broker"],
    ["MRA", "Taxes / Government"],
    ["STATEINSURANCE", "Insurance"],
    ["LAPRUDENCE", "Insurance"],
    ["CREDIT INTEREST", "Interest"],
    ["INWARD TRANSFER", "Transfers In"],
    ["GBP TO SBM", "Transfers Between Own Accounts"],
    ["MCB TO SBM", "Transfers Between Own Accounts"],
    ["JUICE ACCOUNT TRANSFER", "Transfers"],
    ["JUICE TRANSFER", "Transfers"],
  ];
  for (const [needle, label] of patterns) {
    if (text.includes(needle)) return label;
  }
  return "Other";
}

function isSalaryTransaction(txn) {
  const text = `${txn.description || ""} ${txn.reference || ""}`.toUpperCase();
  const hasSalaryKeyword = /(SALARY|PAYROLL|WAGE|REMUNERATION)/.test(text);
  const isInvoiceFromSqli = text.includes("INWARD TRANSFER")
    && text.includes("INVOICE")
    && text.includes("SQLI UK LIMITED");
  return txn.credit > 0 && (hasSalaryKeyword || isInvoiceFromSqli);
}

function isTransferTransaction(txn) {
  const category = categorizeDescription(txn.description);
  if (category.startsWith("Transfers")) return true;
  const text = `${txn.description || ""} ${txn.reference || ""}`.toUpperCase();
  return text.includes("TRANSFER");
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
  return { fromDate: "", toDate: "" };
}

function formatDateInputValue(date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
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
