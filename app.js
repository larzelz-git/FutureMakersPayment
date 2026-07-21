const currencyFormatter = new Intl.NumberFormat("th-TH", {
  style: "currency",
  currency: "THB",
  maximumFractionDigits: 2,
});

const numberFormatter = new Intl.NumberFormat("th-TH", {
  maximumFractionDigits: 2,
});

let dashboardState = null;
let detailTableState = {
  sortKey: "paidDate",
  sortDirection: "desc",
  filters: {},
  globalQuery: "",
  selectedRows: new Set(),
};
const SOURCE_STORAGE_KEY = "summary-dashboard-source-config";
const INCOME_STORAGE_KEY = "summary-dashboard-income-rows";
const SOURCE_URL_KEYS = [
  "spreadsheetUrl",
  "spreadsheetId",
  "sheetName",
  "incomeSheetName",
  "pendingSheetName",
  "paidSheetName",
  "range",
  "incomeRange",
  "pendingRange",
  "paidDetailRange",
  "liveJsonUrl",
];
const DETAIL_COLUMNS = [
  { key: "category", label: "หมวดหลัก", type: "text" },
  { key: "subcategory", label: "หมวดย่อย", type: "text" },
  { key: "description", label: "Description", type: "text" },
  { key: "amount", label: "Final Amount", type: "number" },
  { key: "paidDate", label: "วันที่ชำระ", type: "text" },
];

let incomeRows = [];

function parseAmount(rawValue) {
  if (rawValue == null || rawValue === "-") {
    return 0;
  }

  return Number(String(rawValue).replaceAll(",", "")) || 0;
}

function formatCurrency(value) {
  return currencyFormatter.format(value);
}

function formatPercent(value) {
  return `${numberFormatter.format(value)}%`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeRowsWithCarryForward(rawRows) {
  let previousCategory = "";

  return rawRows
    .map((row) => {
      const safeRow = Array.isArray(row) ? row : [];
      const category = safeRow[0] || previousCategory;
      const subcategory = safeRow[1] || "";

      if (category) {
        previousCategory = category;
      }

      return [
        category,
        subcategory,
        safeRow[2] || "",
        safeRow[3] || "",
        safeRow[4] || "",
        safeRow[5] || "",
        ...safeRow.slice(6),
      ];
    })
    .filter((row) => row[0] || row[1]);
}

function pickPendingAmount(row) {
  const preferredAmount = parseAmount(row[4]) || parseAmount(row[2]);

  if (preferredAmount > 0) {
    return preferredAmount;
  }

  for (let index = row.length - 1; index >= 2; index -= 1) {
    const amount = parseAmount(row[index]);

    if (amount > 0) {
      return amount;
    }
  }

  return 0;
}

function preparePendingRows(rawRows) {
  return normalizeRowsWithCarryForward(rawRows || [])
    .map((row) => ({
      category: row[0],
      subcategory: row[1],
      amount: pickPendingAmount(row),
    }))
    .filter((row) => row.category && row.category !== "Grand Total" && row.amount > 0);
}

function preparePaidDetailRows(rawRows) {
  return (rawRows || [])
    .map((row, index) => {
      const safeRow = Array.isArray(row) ? row : [];
      const category = safeRow[0] || "";
      const subcategory = safeRow[1] || "";
      const description = safeRow[2] || "";
      const amount = parseAmount(safeRow[3]);
      const paidDate = safeRow[4] || "";

      return {
        id: [category, subcategory, description, amount, paidDate, index].join("|"),
        category,
        subcategory,
        description,
        amount,
        paidDate,
      };
    })
    .filter((row) => row.category && row.subcategory);
}

function prepareData(sourceData) {
  const grandTotal = parseAmount(sourceData.grandTotal);
  const actualGrandTotal = parseAmount(sourceData.actualGrandTotal);
  const normalizedRows = normalizeRowsWithCarryForward(sourceData.rows);
  const pendingRows = preparePendingRows(sourceData.pendingRows);
  const paidDetails = preparePaidDetailRows(sourceData.paidDetailRows);
  const paidDetailColumns = DETAIL_COLUMNS.map((column, index) => ({
    ...column,
    label: sourceData.paidDetailHeader?.[index] || column.label,
  }));
  const pendingGrandTotal = pendingRows.reduce((sum, row) => sum + row.amount, 0);

  const rows = normalizedRows.map(
    ([category, subcategory, total, categoryTotal, actualSubTotal, actualCategoryTotal]) => ({
      category,
      subcategory,
      budget: parseAmount(total),
      budgetCategoryTotal: parseAmount(categoryTotal),
      actual: parseAmount(actualSubTotal),
      actualCategoryTotal: parseAmount(actualCategoryTotal),
    })
  );

  const categoryMap = new Map();

  rows.forEach((row) => {
    if (!categoryMap.has(row.category)) {
      categoryMap.set(row.category, {
        category: row.category,
        budget: row.budgetCategoryTotal || 0,
        actual: row.actualCategoryTotal || 0,
        pending: 0,
        budgetFromChildren: 0,
        actualFromChildren: 0,
        pendingFromChildren: 0,
        entries: [],
      });
    }

    const categoryEntry = categoryMap.get(row.category);
    categoryEntry.entries.push(row);
    categoryEntry.budgetFromChildren += row.budget;
    categoryEntry.actualFromChildren += row.actual;
  });

  pendingRows.forEach((row) => {
    if (!categoryMap.has(row.category)) {
      categoryMap.set(row.category, {
        category: row.category,
        budget: 0,
        actual: 0,
        pending: 0,
        budgetFromChildren: 0,
        actualFromChildren: 0,
        pendingFromChildren: 0,
        entries: [],
      });
    }

    const categoryEntry = categoryMap.get(row.category);
    categoryEntry.pending += row.amount;
    categoryEntry.pendingFromChildren += row.amount;
  });

  categoryMap.forEach((categoryEntry) => {
    if (categoryEntry.budget === 0) {
      categoryEntry.budget = categoryEntry.budgetFromChildren;
    }

    if (categoryEntry.actual === 0) {
      categoryEntry.actual = categoryEntry.actualFromChildren;
    }
  });

  const categories = Array.from(categoryMap.values()).sort(
    (left, right) =>
      Math.max(right.budget, right.actual + right.pending) -
      Math.max(left.budget, left.actual + left.pending)
  );
  const activeRows = rows
    .filter((row) => row.budget > 0 || row.actual > 0)
    .sort((left, right) => Math.max(right.budget, right.actual) - Math.max(left.budget, left.actual));

  return {
    source: sourceData.source,
    grandTotal,
    actualGrandTotal,
    pendingGrandTotal,
    incomeSummary: sourceData.incomeSummary || normalizeIncomeSummaryMatrix([]),
    rows,
    pendingRows,
    paidDetails,
    paidDetailColumns,
    activeRows,
    categories,
  };
}

function setLiveStatus(message) {
  const target = document.getElementById("live-status");
  if (target) {
    target.textContent = message;
  }
}

function setRefreshDisabled(disabled) {
  const button = document.getElementById("refresh-button");
  if (button) {
    button.disabled = disabled;
  }
}

function getUrlSourceConfig() {
  const searchParams = new URLSearchParams(window.location.search);
  const spreadsheetUrl = searchParams.get("spreadsheetUrl") || searchParams.get("sourceUrl") || searchParams.get("url");
  const spreadsheetId = searchParams.get("spreadsheetId") || searchParams.get("sourceId") || searchParams.get("id");
  const sourceConfig = {};

  if (spreadsheetUrl) {
    sourceConfig.spreadsheetUrl = spreadsheetUrl;
    sourceConfig.spreadsheetId = spreadsheetId || extractSpreadsheetId(spreadsheetUrl);
  } else if (spreadsheetId) {
    sourceConfig.spreadsheetId = spreadsheetId;
  }

  SOURCE_URL_KEYS.forEach((key) => {
    const value = searchParams.get(key);

    if (value) {
      sourceConfig[key] = value;
    }
  });

  if (Object.keys(sourceConfig).length) {
    sourceConfig.fromUrl = true;
  }

  return sourceConfig;
}

function getSourceConfig() {
  const defaultSource = window.SUMMARY_DASHBOARD_DATA.source;
  const urlSource = getUrlSourceConfig();

  try {
    const saved = JSON.parse(window.localStorage.getItem(SOURCE_STORAGE_KEY) || "{}");
    const sourceConfig = {
      ...defaultSource,
      ...saved,
      ...urlSource,
    };

    if (sourceConfig.sheetName === "Summary") {
      sourceConfig.sheetName = "Summary รายจ่าย";
    }

    if (!sourceConfig.incomeSheetName) {
      sourceConfig.incomeSheetName = "Summary รายรับ";
    }

    if (!sourceConfig.incomeRange) {
      sourceConfig.incomeRange = "A:Z";
    }

    return sourceConfig;
  } catch (error) {
    return { ...defaultSource, ...urlSource };
  }
}

function saveLocalSourceConfig(sourceConfig) {
  window.localStorage.setItem(
    SOURCE_STORAGE_KEY,
    JSON.stringify({
      spreadsheetUrl: sourceConfig.spreadsheetUrl,
      spreadsheetId: sourceConfig.spreadsheetId,
      sheetName: sourceConfig.sheetName,
      incomeSheetName: sourceConfig.incomeSheetName,
      range: sourceConfig.range,
      incomeRange: sourceConfig.incomeRange,
      pendingSheetName: sourceConfig.pendingSheetName,
      pendingRange: sourceConfig.pendingRange,
      paidSheetName: sourceConfig.paidSheetName,
      paidDetailRange: sourceConfig.paidDetailRange,
      liveJsonUrl: sourceConfig.liveJsonUrl,
    })
  );
}

function updateGlobalSource(sourceConfig) {
  window.SUMMARY_DASHBOARD_DATA.source = {
    ...window.SUMMARY_DASHBOARD_DATA.source,
    ...sourceConfig,
  };
}

function buildSourceShareUrl(sourceConfig) {
  const url = new URL(window.location.href);

  SOURCE_URL_KEYS.forEach((key) => {
    url.searchParams.delete(key);
  });
  ["sourceUrl", "url", "sourceId", "id", "fromUrl"].forEach((key) => {
    url.searchParams.delete(key);
  });

  SOURCE_URL_KEYS.forEach((key) => {
    const value = sourceConfig[key];

    if (value) {
      url.searchParams.set(key, value);
    }
  });

  return url.toString();
}

function requestJsonp(url) {
  const callbackName = `sheetSaveCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;

  url.searchParams.set("callback", callbackName);

  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    const timeoutId = window.setTimeout(() => {
      cleanup();
      reject(new Error("Source config save timed out"));
    }, 12000);

    function cleanup() {
      window.clearTimeout(timeoutId);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (payload) => {
      cleanup();
      resolve(payload);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Failed to save source config"));
    };

    script.src = url.toString();
    document.head.appendChild(script);
  });
}

async function saveSourceConfig(sourceConfig) {
  const liveJsonUrl = window.SUMMARY_DASHBOARD_DATA.source.liveJsonUrl;

  if (!liveJsonUrl) {
    saveLocalSourceConfig(sourceConfig);
    return sourceConfig;
  }

  const url = new URL(liveJsonUrl);

  url.searchParams.set("action", "saveSource");
  url.searchParams.set("spreadsheetUrl", sourceConfig.spreadsheetUrl || "");
  url.searchParams.set("spreadsheetId", sourceConfig.spreadsheetId || "");
  url.searchParams.set("sheetName", sourceConfig.sheetName || "Summary รายจ่าย");
  url.searchParams.set("incomeSheetName", sourceConfig.incomeSheetName || "Summary รายรับ");
  url.searchParams.set("pendingSheetName", sourceConfig.pendingSheetName || "เตรียมจ่าย");
  url.searchParams.set("paidSheetName", sourceConfig.paidSheetName || "จ่ายแล้ว");
  url.searchParams.set("range", sourceConfig.range || "A:F");
  url.searchParams.set("incomeRange", sourceConfig.incomeRange || "A:Z");
  url.searchParams.set("pendingRange", sourceConfig.pendingRange || "A:G");
  url.searchParams.set("paidDetailRange", sourceConfig.paidDetailRange || "G:V");

  const payload = await requestJsonp(url);

  if (!payload.ok) {
    throw new Error(payload.error || "Source config save failed");
  }

  saveLocalSourceConfig(payload.source);
  return payload.source;
}

function extractSpreadsheetId(url) {
  const match = url.match(/\/spreadsheets\/d\/([^/]+)/);
  return match ? match[1] : "";
}

function parseGvizResponse(rawText) {
  const start = rawText.indexOf("{");
  const end = rawText.lastIndexOf("}");
  if (start === -1 || end === -1) {
    throw new Error("Invalid Google Visualization response");
  }

  return JSON.parse(rawText.slice(start, end + 1));
}

function getGvizCellValue(cell) {
  if (!cell) {
    return "";
  }

  if (cell.f != null) {
    return cell.f;
  }

  if (cell.v == null) {
    return "";
  }

  return String(cell.v);
}

function normalizeSheetMatrix(matrix, fallback) {
  if (!matrix.length) {
    throw new Error("Sheet returned no usable rows");
  }

  const snapshotDate = matrix[0]?.[6] || fallback.source.snapshotDate;
  const hasHeaderRow = matrix[0]?.[0] === "หมวดหลัก";
  const dataMatrix = hasHeaderRow ? matrix.slice(1) : matrix;
  const normalizedMatrix = normalizeRowsWithCarryForward(dataMatrix);
  const rows = normalizedMatrix.filter((row) => row[0] !== "Grand Total");
  const grandTotalRow = normalizedMatrix.find((row) => row[0] === "Grand Total") || [];

  return {
    ...fallback,
    source: {
      ...fallback.source,
      snapshotDate,
    },
    rows: rows.map((row) => [
      row[0] || "",
      row[1] || "",
      row[2] || "",
      row[3] || "",
      row[4] || "",
      row[5] || "",
    ]),
    pendingRows: fallback.pendingRows || [],
    grandTotal: grandTotalRow[2] || fallback.grandTotal,
    actualGrandTotal: grandTotalRow[4] || fallback.actualGrandTotal,
  };
}

function normalizePendingMatrix(matrix) {
  if (!matrix.length) {
    return [];
  }

  const hasHeaderRow = matrix[0]?.some((value) => String(value).includes("หมวด"));
  const dataMatrix = hasHeaderRow ? matrix.slice(1) : matrix;

  return normalizeRowsWithCarryForward(dataMatrix).filter((row) => row[0] !== "Grand Total");
}

function findAmountNearLabel(row, labelIndex) {
  for (let index = labelIndex + 1; index < row.length; index += 1) {
    const amount = parseAmount(row[index]);

    if (amount !== 0) {
      return amount;
    }
  }

  for (let index = 0; index < row.length; index += 1) {
    if (index !== labelIndex) {
      const amount = parseAmount(row[index]);

      if (amount !== 0) {
        return amount;
      }
    }
  }

  return 0;
}

function normalizeIncomeSummaryMatrix(matrix) {
  const summary = {
    salesMongo: 0,
    receivedTotal: 0,
    fee: 0,
    receivedWithFee: 0,
    channels: [],
    bankSummary: {
      latestDate: "",
      accounts: [],
      total: 0,
    },
  };
  const channelRows = [];
  const bankAccounts = [];
  let totalRow = null;
  let inBankSection = false;

  (matrix || []).forEach((row) => {
    const firstCell = String(row[0] || "").trim();

    if (firstCell === "ข้อมูลธนาคารล่าสุด") {
      summary.bankSummary.latestDate = row[1] || "";
    }

    if (firstCell === "รวมทั้งหมด") {
      totalRow = row;
    } else if (firstCell === "สรุปยอดคงเหลือธนาคาร") {
      inBankSection = true;
    } else if (inBankSection && firstCell === "ธนาคาร / บัญชี") {
      return;
    } else if (inBankSection && firstCell) {
      const bankBalance = parseAmount(row[3] || row[2] || row[1]);

      if (firstCell.includes("รวม")) {
        summary.bankSummary.total = bankBalance;
        inBankSection = false;
      } else {
        bankAccounts.push({
          account: firstCell,
          balance: bankBalance,
        });
      }
    } else if (
      firstCell &&
      !["ข้อมูลธนาคารล่าสุด", "ช่องทาง", "ยอดขายรวม", "เงินเข้าจริงสุทธิ"].includes(firstCell) &&
      !firstCell.startsWith("หมายเหตุ")
    ) {
      const salesMongo = parseAmount(row[1]);
      const fee = parseAmount(row[3]);
      const receivedWithFee = parseAmount(row[4]);

      if (salesMongo > 0 || receivedWithFee > 0) {
        channelRows.push({
          channel: firstCell,
          salesMongo,
          fee,
          receivedWithFee,
          netReceived: parseAmount(row[8]),
          gap: parseAmount(row[9]),
          receivedPercent: parseAmount(row[10]),
        });
      }
    }

    row.forEach((cell, cellIndex) => {
      const label = String(cell || "").replace(/\s+/g, " ").trim();
      const normalizedLabel = label.toLowerCase();
      const amount = findAmountNearLabel(row, cellIndex);

      if (!amount) {
        return;
      }

      if (label.includes("ยอดขาย") && normalizedLabel.includes("mongo")) {
        summary.salesMongo = amount;
      } else if (label.includes("รับจริงรวม") && label.includes("Fee")) {
        summary.receivedWithFee = amount;
      } else if (label.includes("รับจริงรวม") || label.includes("รับจริง")) {
        summary.receivedTotal = amount;
      } else if (label === "Fee" || label.includes("ค่า Fee") || label.includes("ค่าธรรมเนียม")) {
        summary.fee = amount;
      }
    });
  });

  if (totalRow) {
    summary.salesMongo = parseAmount(totalRow[1]) || summary.salesMongo;
    summary.fee = parseAmount(totalRow[3]) || summary.fee;
    summary.receivedWithFee = parseAmount(totalRow[4]) || summary.receivedWithFee;
    summary.receivedTotal = parseAmount(totalRow[8]) || summary.receivedTotal;
  }

  if (!summary.receivedWithFee) {
    summary.receivedWithFee = summary.receivedTotal + summary.fee;
  }

  summary.channels = channelRows;
  summary.bankSummary.accounts = bankAccounts;

  if (!summary.bankSummary.total) {
    summary.bankSummary.total = bankAccounts.reduce((sum, item) => sum + item.balance, 0);
  }

  return summary;
}

function hasIncomeSummaryValues(incomeSummary) {
  return Boolean(
    incomeSummary &&
      (parseAmount(incomeSummary.salesMongo) ||
        parseAmount(incomeSummary.receivedWithFee) ||
        (incomeSummary.channels || []).length)
  );
}

async function fetchIncomeSummaryForSource(source) {
  const sheetName = source.incomeSheetName || "Summary รายรับ";
  const range = source.incomeRange || "A:Z";
  const matrix = await fetchSheetMatrixViaScript(sheetName, range).catch(() => fetchSheetMatrix(sheetName, range).catch(() => []));

  return normalizeIncomeSummaryMatrix(matrix);
}

function normalizePaidDetailPayload(matrix) {
  if (!matrix.length) {
    return {
      paidDetailHeader: [],
      paidDetailRows: [],
    };
  }

  const mapRow = (row) => [
    row[14] || "",
    row[15] || "",
    row[0] || "",
    row[1] || "",
    row[2] || "",
  ];

  return {
    paidDetailHeader: mapRow(matrix[1] || []),
    paidDetailRows: matrix.slice(2).map(mapRow).filter((row) => row.some((value) => value !== "")),
  };
}

function fetchSheetMatrixViaScript(sheetName, range) {
  const fallback = window.SUMMARY_DASHBOARD_DATA;
  const source = fallback.source;
  const spreadsheetId = source.spreadsheetId || extractSpreadsheetId(source.spreadsheetUrl);
  const callbackName = `sheetCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
  const url = new URL(`https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq`);

  url.searchParams.set("sheet", sheetName);
  url.searchParams.set("range", range);
  url.searchParams.set("tqx", `responseHandler:${callbackName}`);

  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    const timeoutId = window.setTimeout(() => {
      cleanup();
      reject(new Error("Google Sheet request timed out"));
    }, 12000);

    function cleanup() {
      window.clearTimeout(timeoutId);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (payload) => {
      try {
        const matrix = (payload.table?.rows || [])
          .map((row) => (row.c || []).map(getGvizCellValue))
          .filter((row) => row.some((value) => value !== ""));
        cleanup();
        resolve(matrix);
      } catch (error) {
        cleanup();
        reject(error);
      }
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Failed to load Google Sheet script"));
    };

    script.src = url.toString();
    document.head.appendChild(script);
  });
}

async function fetchLiveSummaryDataViaScript() {
  const fallback = window.SUMMARY_DASHBOARD_DATA;
  const source = fallback.source;
  const summaryMatrix = await fetchSheetMatrixViaScript(source.sheetName, source.range);
  const pendingMatrix = await fetchSheetMatrixViaScript(
    source.pendingSheetName || "เตรียมจ่าย",
    source.pendingRange || "A:G"
  ).catch(() => []);
  const paidDetailMatrix = await fetchSheetMatrixViaScript(
    source.paidSheetName || "จ่ายแล้ว",
    source.paidDetailRange || "G:V"
  ).catch(() => []);
  const incomeMatrix = await fetchSheetMatrixViaScript(
    source.incomeSheetName || "Summary รายรับ",
    source.incomeRange || "A:Z"
  ).catch(() => []);
  const paidDetailPayload = normalizePaidDetailPayload(paidDetailMatrix);

  return {
    ...normalizeSheetMatrix(summaryMatrix, fallback),
    pendingRows: normalizePendingMatrix(pendingMatrix),
    incomeSummary: normalizeIncomeSummaryMatrix(incomeMatrix),
    ...paidDetailPayload,
  };
}

async function fetchSheetMatrix(sheetName, range) {
  const fallback = window.SUMMARY_DASHBOARD_DATA;
  const source = fallback.source;
  const spreadsheetId = source.spreadsheetId || extractSpreadsheetId(source.spreadsheetUrl);
  const url = new URL(`https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq`);

  url.searchParams.set("sheet", sheetName);
  url.searchParams.set("range", range);
  url.searchParams.set("tqx", "out:json");

  const response = await fetch(url.toString(), { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Google Sheet fetch failed: ${response.status}`);
  }

  const payload = parseGvizResponse(await response.text());

  return (payload.table?.rows || [])
    .map((row) => (row.c || []).map(getGvizCellValue))
    .filter((row) => row.some((value) => value !== ""));
}

async function fetchLiveSummaryData() {
  const fallback = window.SUMMARY_DASHBOARD_DATA;
  const liveJsonUrl = fallback.source.liveJsonUrl;

  if (liveJsonUrl) {
    const url = new URL(liveJsonUrl);

    SOURCE_URL_KEYS.forEach((key) => {
      const value = fallback.source[key];

      if (value && key !== "liveJsonUrl") {
        url.searchParams.set(key, value);
      }
    });

    const response = await fetch(url.toString(), { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Live JSON fetch failed: ${response.status}`);
    }

    const payload = await response.json();
    const mergedSource = {
      ...fallback.source,
      ...payload.source,
      snapshotDate: payload.source?.snapshotDate || fallback.source.snapshotDate,
    };
    let incomeSummary = payload.incomeSummary || fallback.incomeSummary || normalizeIncomeSummaryMatrix([]);

    if (!hasIncomeSummaryValues(incomeSummary)) {
      incomeSummary = await fetchIncomeSummaryForSource(mergedSource);
    }

    return {
      ...fallback,
      source: mergedSource,
      rows: payload.rows || fallback.rows,
      pendingRows: payload.pendingRows || fallback.pendingRows || [],
      incomeSummary,
      paidDetailHeader: payload.paidDetailHeader || fallback.paidDetailHeader || [],
      paidDetailRows: payload.paidDetailRows || fallback.paidDetailRows || [],
      grandTotal: payload.grandTotal || fallback.grandTotal,
      actualGrandTotal: payload.actualGrandTotal || fallback.actualGrandTotal,
    };
  }

  try {
    return await fetchLiveSummaryDataViaScript();
  } catch (scriptError) {
    const source = fallback.source;
    const summaryMatrix = await fetchSheetMatrix(source.sheetName, source.range).catch(() => {
      throw scriptError;
    });
    const pendingMatrix = await fetchSheetMatrix(
      source.pendingSheetName || "เตรียมจ่าย",
      source.pendingRange || "A:G"
    ).catch(() => []);
    const paidDetailMatrix = await fetchSheetMatrix(
      source.paidSheetName || "จ่ายแล้ว",
      source.paidDetailRange || "G:V"
    ).catch(() => []);
    const incomeMatrix = await fetchSheetMatrix(
      source.incomeSheetName || "Summary รายรับ",
      source.incomeRange || "A:Z"
    ).catch(() => []);
    const paidDetailPayload = normalizePaidDetailPayload(paidDetailMatrix);

    return {
      ...normalizeSheetMatrix(summaryMatrix, fallback),
      pendingRows: normalizePendingMatrix(pendingMatrix),
      incomeSummary: normalizeIncomeSummaryMatrix(incomeMatrix),
      ...paidDetailPayload,
    };
  }
}

function renderKpis(data) {
  const root = document.getElementById("kpi-grid");
  const hiddenExpenseCategory = data.categories.find((item) => item.category === "ค่าใช้จ่ายแฝง");
  const hiddenExpenseBudget = hiddenExpenseCategory ? hiddenExpenseCategory.budget : 0;
  const remainingBudget = data.grandTotal - data.actualGrandTotal - hiddenExpenseBudget;

  const kpis = [
    { label: "งบรวม", value: formatCurrency(data.grandTotal) },
    { label: "จ่ายจริงรวม", value: formatCurrency(data.actualGrandTotal) },
    { label: "เตรียมจ่าย", value: formatCurrency(data.pendingGrandTotal) },
    { label: "เหลือจ่าย", value: formatCurrency(remainingBudget) },
  ];

  root.innerHTML = kpis
    .map(
      (item) => `
        <article class="kpi-card">
          <div class="kpi-label">${item.label}</div>
          <div class="kpi-value">${item.value}</div>
        </article>
      `
    )
    .join("");
}

function renderSourceInfo(data) {
  const sourceLink = document.getElementById("source-link");
  const sourceSheetName = document.getElementById("source-sheet-name");
  const snapshotDate = document.getElementById("snapshot-date");
  const sourceLinkInput = document.getElementById("source-link-input");
  const sourceSheetNameInput = document.getElementById("source-sheet-name-input");
  const sourceShareUrl = document.getElementById("source-share-url");

  if (sourceLink) {
    sourceLink.href = data.source.spreadsheetUrl || "#";
    sourceLink.textContent = data.source.spreadsheetUrl || "-";
  }

  if (sourceSheetName) {
    sourceSheetName.textContent = data.source.sheetName || "-";
  }

  if (snapshotDate) {
    snapshotDate.textContent = data.source.snapshotDate || "-";
  }

  if (sourceLinkInput) {
    sourceLinkInput.value = data.source.spreadsheetUrl || "";
  }

  if (sourceSheetNameInput) {
    sourceSheetNameInput.value = data.source.sheetName || "";
  }

  if (sourceShareUrl) {
    sourceShareUrl.value = buildSourceShareUrl(data.source);
  }
}

function renderCategoryCompare(data) {
  const root = document.getElementById("category-compare");
  const maxValue = Math.max(
    ...data.categories.map((item) => Math.max(item.budget, item.actual + item.pending)),
    0
  );

  root.innerHTML = data.categories
    .map((item) => {
      const usage = item.budget > 0 ? (item.actual / item.budget) * 100 : 0;
      const projected = item.actual + item.pending;
      const projectedUsage = item.budget > 0 ? (projected / item.budget) * 100 : 0;
      const budgetWidth = maxValue > 0 ? (item.budget / maxValue) * 100 : 0;
      const actualWidth = maxValue > 0 ? (item.actual / maxValue) * 100 : 0;
      const pendingWidth = maxValue > 0 ? (item.pending / maxValue) * 100 : 0;
      const overBudgetAmount = Math.max(projected - item.budget, 0);
      const overBudgetClass = projected > item.budget && item.budget > 0 ? " projected-over" : "";
      return `
        <div class="category-row${overBudgetClass}">
          <div class="category-label">
            <span>${item.category}</span>
            <span>${formatPercent(projectedUsage || usage || 0)}</span>
          </div>
          <div class="category-meta">
            <span>งบ ${formatCurrency(item.budget)}</span>
            <span>จ่ายจริง ${formatCurrency(item.actual)}</span>
            <span>เตรียมจ่าย ${formatCurrency(item.pending)}</span>
            <span>เกินงบ ${formatCurrency(overBudgetAmount)}</span>
          </div>
          <div class="dual-bars">
            <div class="bar-set">
              <span class="bar-set-label">งบ</span>
              <div class="bar-track"><div class="bar-fill budget" style="width: ${Math.max(budgetWidth, item.budget > 0 ? 2 : 0)}%"></div></div>
              <span class="bar-set-value">${formatCurrency(item.budget)}</span>
            </div>
            <div class="bar-set">
              <span class="bar-set-label">จ่าย</span>
              <div class="bar-track stacked-bar">
                <div class="bar-fill actual" style="width: ${Math.max(actualWidth, item.actual > 0 ? 2 : 0)}%"></div>
                <div class="bar-fill pending" style="width: ${Math.max(pendingWidth, item.pending > 0 ? 2 : 0)}%"></div>
              </div>
              <span class="bar-set-value">${formatCurrency(projected)}</span>
            </div>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderTopCategory(data) {
  const root = document.getElementById("top-category-card");
  const topCategory = data.categories[0];

  if (!topCategory) {
    root.innerHTML = "<p>ไม่พบข้อมูล</p>";
    return;
  }

  const variance = topCategory.actual - topCategory.budget;
  const usage = topCategory.budget > 0 ? (topCategory.actual / topCategory.budget) * 100 : 0;
  const nonZeroEntries = topCategory.entries.filter((item) => item.budget > 0 || item.actual > 0).length;

  root.innerHTML = `
    <div class="top-category-card">
      <div class="chip">${nonZeroEntries} รายการย่อยที่มีมูลค่า</div>
      <p class="top-category-amount">${formatCurrency(topCategory.actual)}</p>
      <div class="top-category-name">${topCategory.category}</div>
      <p class="top-category-share">ใช้งบ ${formatPercent(usage || 0)} · ส่วนต่าง ${formatCurrency(variance)}</p>
      <p class="muted">อ้างอิงจาก snapshot วันที่ ${data.source.snapshotDate}</p>
    </div>
  `;
}

function renderInsights(data) {
  const root = document.getElementById("insight-list");
  const topThree = data.categories.filter((item) => item.actual > 0).slice(0, 3);
  const concentration = topThree.reduce((sum, item) => sum + item.actual, 0);
  const concentrationShare = data.actualGrandTotal > 0 ? (concentration / data.actualGrandTotal) * 100 : 0;
  const highestSubcategory = data.activeRows[0];
  const unpaidCategories = data.categories.filter(
    (item) => item.category !== "ค่าใช้จ่ายแฝง" && item.budget > 0 && item.actual < item.budget
  );
  const unpaidRemainingTotal = unpaidCategories.reduce((sum, item) => sum + Math.max(item.budget - item.actual, 0), 0);
  const overBudget = data.categories
    .filter((item) => item.budget > 0 && item.actual > item.budget)
    .sort((left, right) => right.actual - right.budget - (left.actual - left.budget))[0];

  const insights = [
    unpaidRemainingTotal > 0
      ? `ยอดคงเหลือที่ยังไม่จ่ายให้ครบคือ ${formatCurrency(unpaidRemainingTotal)} จาก ${unpaidCategories.length} หมวดหลักที่ยังจ่ายไม่ครบ`
      : "ตอนนี้ไม่มีหมวดหลักที่ยังจ่ายไม่ครบ",
    `3 หมวดหลักแรกกินสัดส่วน ${formatPercent(concentrationShare)} ของยอดจ่ายจริงทั้งหมด`,
    highestSubcategory
      ? `หมวดย่อยที่จ่ายจริงสูงสุดคือ ${highestSubcategory.subcategory} ในหมวด ${highestSubcategory.category} มูลค่า ${formatCurrency(highestSubcategory.actual)}`
      : "ยังไม่พบหมวดย่อยที่มียอดใช้จ่าย",
    overBudget
      ? `หมวดหลักที่เกินงบมากสุดตอนนี้คือ ${overBudget.category} เกิน ${formatCurrency(overBudget.actual - overBudget.budget)}`
      : "ตอนนี้ยังไม่พบหมวดหลักที่เกินงบ",
    data.pendingGrandTotal > 0
      ? `ถ้าจ่ายรายการเตรียมจ่ายทั้งหมด ยอดรวมจะเพิ่มอีก ${formatCurrency(data.pendingGrandTotal)} เป็น ${formatCurrency(data.actualGrandTotal + data.pendingGrandTotal)}`
      : "ยังไม่พบยอดเตรียมจ่ายจากชีตเตรียมจ่าย",
    `มี ${data.rows.length} รายการย่อย และ ${data.activeRows.length} รายการที่มีงบหรือจ่ายจริงมากกว่า 0`,
    `หน้าเว็บนี้ดึงมาจาก ${data.source.sheetName}!${data.source.range} และจับคู่คอลัมน์งบกับจ่ายจริงแล้ว`,
  ];

  root.innerHTML = insights.map((item) => `<li>${item}</li>`).join("");
}

function getDetailValue(row, key) {
  return row[key] == null ? "" : row[key];
}

function getFilteredSortedDetailRows(data) {
  const normalizedGlobalQuery = detailTableState.globalQuery.trim().toLowerCase();
  const columns = data.paidDetailColumns || DETAIL_COLUMNS;

  return data.paidDetails
    .filter((row) => {
      const values = columns.map((column) => String(getDetailValue(row, column.key)).toLowerCase());

      if (normalizedGlobalQuery && !values.some((value) => value.includes(normalizedGlobalQuery))) {
        return false;
      }

      return columns.every((column) => {
        const filterValue = (detailTableState.filters[column.key] || "").trim().toLowerCase();

        if (!filterValue) {
          return true;
        }

        return String(getDetailValue(row, column.key)).toLowerCase().includes(filterValue);
      });
    })
    .sort((left, right) => {
      const column = columns.find((item) => item.key === detailTableState.sortKey) || columns[0];
      const leftValue = getDetailValue(left, column.key);
      const rightValue = getDetailValue(right, column.key);
      const sortFactor = detailTableState.sortDirection === "asc" ? 1 : -1;

      if (column.type === "number") {
        return (Number(leftValue) - Number(rightValue)) * sortFactor;
      }

      return String(leftValue).localeCompare(String(rightValue), "th") * sortFactor;
    });
}

function renderDetailHead(data) {
  const root = document.getElementById("detail-head");

  if (!root) {
    return;
  }

  const columns = data.paidDetailColumns || DETAIL_COLUMNS;
  const filteredRows = getFilteredSortedDetailRows(data);
  const visibleTotal = filteredRows.reduce((sum, row) => sum + row.amount, 0);
  const selectedVisibleCount = filteredRows.filter((row) => detailTableState.selectedRows.has(row.id)).length;
  const allVisibleSelected = filteredRows.length > 0 && selectedVisibleCount === filteredRows.length;
  const partiallySelected = selectedVisibleCount > 0 && !allVisibleSelected;
  const selectedTotal = filteredRows.reduce(
    (sum, row) => (detailTableState.selectedRows.has(row.id) ? sum + row.amount : sum),
    0
  );

  root.innerHTML = `
    <tr class="detail-total-row">
      <th class="select-column"></th>
      ${columns.map(
        (column) => `
          <th>
            ${
              column.key === "description"
                ? `<span class="selected-total-label">ยอดรวมที่เลือก</span><br /><span class="selected-total-value">${formatCurrency(selectedTotal)}</span>`
                : column.key === "amount"
                  ? `<span>ยอดรวมทั้งหมด</span><br />${formatCurrency(visibleTotal)}`
                  : ""
            }
          </th>
        `
      ).join("")}
    </tr>
    <tr>
      <th class="select-column">
        <input
          class="detail-row-checkbox select-all-checkbox"
          type="checkbox"
          aria-label="เลือกทั้งหมด"
          data-detail-select-all
          ${allVisibleSelected ? "checked" : ""}
          ${filteredRows.length === 0 ? "disabled" : ""}
        />
      </th>
      ${columns.map((column) => {
        const isActive = detailTableState.sortKey === column.key;
        const directionLabel = isActive ? (detailTableState.sortDirection === "asc" ? " ↑" : " ↓") : "";

        return `
          <th>
            <button class="sortable-header" type="button" data-detail-sort="${column.key}">
              ${column.label}${directionLabel}
            </button>
          </th>
        `;
      }).join("")}
    </tr>
    <tr class="filter-row">
      <th class="select-column"></th>
      ${columns.map(
        (column) => `
          <th>
            <input
              class="column-filter"
              type="search"
              value="${escapeHtml(detailTableState.filters[column.key] || "")}"
              placeholder="กรอง"
              data-detail-filter="${column.key}"
            />
          </th>
        `
      ).join("")}
    </tr>
  `;

  const selectAllCheckbox = root.querySelector("[data-detail-select-all]");

  if (selectAllCheckbox) {
    selectAllCheckbox.indeterminate = partiallySelected;
    selectAllCheckbox.addEventListener("change", () => {
      filteredRows.forEach((row) => {
        if (selectAllCheckbox.checked) {
          detailTableState.selectedRows.add(row.id);
        } else {
          detailTableState.selectedRows.delete(row.id);
        }
      });

      renderTableRows(data);
      renderDetailHead(data);
    });
  }

  root.querySelectorAll("[data-detail-sort]").forEach((button) => {
    button.addEventListener("click", () => {
      const nextSortKey = button.dataset.detailSort;

      if (detailTableState.sortKey === nextSortKey) {
        detailTableState.sortDirection = detailTableState.sortDirection === "asc" ? "desc" : "asc";
      } else {
        detailTableState.sortKey = nextSortKey;
        detailTableState.sortDirection = "asc";
      }

      renderTableRows(data);
      renderDetailHead(data);
    });
  });

  root.querySelectorAll("[data-detail-filter]").forEach((input) => {
    input.addEventListener("input", () => {
      detailTableState.filters[input.dataset.detailFilter] = input.value;
      renderTableRows(data);
      renderDetailHead(data);
    });
  });
}

function renderTableRows(data) {
  const root = document.getElementById("subcategory-body");

  if (!root) {
    return;
  }

  const filteredRows = getFilteredSortedDetailRows(data);

  if (!filteredRows.length) {
    root.innerHTML = `
      <tr>
        <td colspan="${(data.paidDetailColumns || DETAIL_COLUMNS).length + 1}" class="pending-empty-cell">ไม่พบรายการจ่ายแล้ว</td>
      </tr>
    `;
    return;
  }

  root.innerHTML = filteredRows
    .map(
      (row) => `
        <tr>
          <td class="select-column">
            <input
              class="detail-row-checkbox"
              type="checkbox"
              data-detail-row-id="${escapeHtml(row.id)}"
              ${detailTableState.selectedRows.has(row.id) ? "checked" : ""}
            />
          </td>
          <td>${escapeHtml(row.category)}</td>
          <td>${escapeHtml(row.subcategory)}</td>
          <td>${escapeHtml(row.description)}</td>
          <td>${formatCurrency(row.amount)}</td>
          <td>${escapeHtml(row.paidDate)}</td>
        </tr>
      `
    )
    .join("");

  root.querySelectorAll("[data-detail-row-id]").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        detailTableState.selectedRows.add(checkbox.dataset.detailRowId);
      } else {
        detailTableState.selectedRows.delete(checkbox.dataset.detailRowId);
      }

      renderDetailHead(data);
    });
  });
}

function setupSearch(data) {
  const searchInput = document.getElementById("search-input");
  searchInput.addEventListener("input", (event) => {
    detailTableState.globalQuery = event.target.value;
    renderTableRows(data);
    renderDetailHead(data);
  });
}

function setupSourceControls() {
  const applyButton = document.getElementById("source-apply-button");
  const sourceLinkInput = document.getElementById("source-link-input");
  const sourceSheetNameInput = document.getElementById("source-sheet-name-input");
  const sourceShareUrl = document.getElementById("source-share-url");
  const copyUrlButton = document.getElementById("source-copy-url-button");
  const sourceModal = document.getElementById("source-modal");
  const openButton = document.getElementById("source-modal-open-button");
  const closeButton = document.getElementById("source-modal-close-button");

  function getPendingSourceConfig() {
    const spreadsheetUrl = sourceLinkInput.value.trim();
    const sheetName = sourceSheetNameInput.value.trim();
    const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);

    return {
      spreadsheetUrl,
      spreadsheetId,
      sheetName: sheetName || "Summary รายจ่าย",
      incomeSheetName: window.SUMMARY_DASHBOARD_DATA.source.incomeSheetName || "Summary รายรับ",
      pendingSheetName: window.SUMMARY_DASHBOARD_DATA.source.pendingSheetName || "เตรียมจ่าย",
      range: window.SUMMARY_DASHBOARD_DATA.source.range || "A:F",
      incomeRange: window.SUMMARY_DASHBOARD_DATA.source.incomeRange || "A:Z",
      pendingRange: window.SUMMARY_DASHBOARD_DATA.source.pendingRange || "A:G",
      paidSheetName: window.SUMMARY_DASHBOARD_DATA.source.paidSheetName || "จ่ายแล้ว",
      paidDetailRange: window.SUMMARY_DASHBOARD_DATA.source.paidDetailRange || "G:V",
      liveJsonUrl: window.SUMMARY_DASHBOARD_DATA.source.liveJsonUrl || "",
    };
  }

  function updateShareUrlPreview() {
    if (sourceShareUrl) {
      sourceShareUrl.value = buildSourceShareUrl({
        ...window.SUMMARY_DASHBOARD_DATA.source,
        ...getPendingSourceConfig(),
      });
    }
  }

  function openModal() {
    updateShareUrlPreview();
    sourceModal.hidden = false;
  }

  function closeModal() {
    sourceModal.hidden = true;
  }

  openButton.addEventListener("click", openModal);
  closeButton.addEventListener("click", closeModal);
  sourceModal.addEventListener("click", (event) => {
    if (event.target === sourceModal) {
      closeModal();
    }
  });

  sourceLinkInput.addEventListener("input", updateShareUrlPreview);
  sourceSheetNameInput.addEventListener("input", updateShareUrlPreview);

  copyUrlButton.addEventListener("click", async () => {
    updateShareUrlPreview();

    try {
      await navigator.clipboard.writeText(sourceShareUrl.value);
      setLiveStatus("คัดลอก URL พร้อม Data Source แล้ว");
    } catch (error) {
      sourceShareUrl.select();
      setLiveStatus("คัดลอกอัตโนมัติไม่ได้ เลือก URL ให้แล้ว");
    }
  });

  applyButton.addEventListener("click", async () => {
    const sourceConfig = getPendingSourceConfig();

    applyButton.disabled = true;
    setLiveStatus("กำลังบันทึก Data Source...");

    try {
      const savedSourceConfig = await saveSourceConfig(sourceConfig);

      updateGlobalSource(savedSourceConfig);
      await refreshDashboard();
      closeModal();
    } catch (error) {
      setLiveStatus("บันทึก Data Source ไม่สำเร็จ");
      console.warn("Source config save failed", error);
    } finally {
      applyButton.disabled = false;
    }
  });
}

function setupPanelTabs() {
  const tabs = Array.from(document.querySelectorAll("[data-tab-target]"));

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((item) => {
        const isActive = item === tab;
        const panel = document.getElementById(item.dataset.tabTarget);

        item.classList.toggle("is-active", isActive);
        item.setAttribute("aria-selected", String(isActive));

        if (panel) {
          panel.hidden = !isActive;
        }
      });
    });
  });
}

function setupMainViewTabs() {
  const tabs = Array.from(document.querySelectorAll("[data-view-target]"));

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((item) => {
        const isActive = item === tab;
        const panel = document.getElementById(item.dataset.viewTarget);

        item.classList.toggle("is-active", isActive);
        item.setAttribute("aria-selected", String(isActive));

        if (panel) {
          panel.hidden = !isActive;
        }
      });
    });
  });
}

function parsePendingPaste(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.split("\t").map((cell) => cell.trim()))
    .filter((row) => row.some((cell) => cell))
    .filter((row) => row[0] !== "หมวดหลัก")
    .map((row) => ({
      category: row[0] || "",
      subcategory: row[1] || "",
      subcategoryActual: parseAmount(row[2]),
      categoryActual: parseAmount(row[3]),
      description: row[4] || "",
    }))
    .filter((row) => row.category || row.subcategory || row.description);
}

function renderPendingPreview(rows) {
  const body = document.getElementById("pending-preview-body");
  const summary = document.getElementById("pending-preview-summary");

  if (!body || !summary) {
    return;
  }

  const total = rows.reduce((sum, row) => sum + (row.subcategoryActual || row.categoryActual), 0);
  summary.textContent = `${rows.length} รายการ · ${formatCurrency(total)}`;

  if (!rows.length) {
    body.innerHTML = `
      <tr>
        <td colspan="5" class="pending-empty-cell">ยังไม่มีข้อมูลเตรียมจ่าย</td>
      </tr>
    `;
    return;
  }

  body.innerHTML = rows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.category)}</td>
          <td>${escapeHtml(row.subcategory)}</td>
          <td>${formatCurrency(row.subcategoryActual)}</td>
          <td>${formatCurrency(row.categoryActual)}</td>
          <td>${escapeHtml(row.description)}</td>
        </tr>
      `
    )
    .join("");
}

function setupPendingPastePreview() {
  const input = document.getElementById("pending-paste-input");
  const previewButton = document.getElementById("pending-preview-button");
  const clearButton = document.getElementById("pending-clear-button");

  if (!input || !previewButton || !clearButton) {
    return;
  }

  function updatePreview() {
    renderPendingPreview(parsePendingPaste(input.value));
  }

  input.addEventListener("paste", () => {
    window.setTimeout(updatePreview, 0);
  });
  previewButton.addEventListener("click", updatePreview);
  clearButton.addEventListener("click", () => {
    input.value = "";
    renderPendingPreview([]);
  });
}

function parseIncomePaste(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => line.split("\t").map((cell) => cell.trim()))
    .filter((row) => row.some((cell) => cell))
    .filter((row) => !["วันที่", "date"].includes(String(row[0] || "").toLowerCase()))
    .map((row, index) => ({
      id: [row[0] || "", row[1] || "", row[2] || "", row[3] || "", row[4] || "", index].join("|"),
      date: row[0] || "",
      source: row[1] || "",
      description: row[2] || "",
      amount: parseAmount(row[3]),
      status: row[4] || "รับแล้ว",
    }))
    .filter((row) => row.source || row.description || row.amount > 0);
}

function loadIncomeRows() {
  try {
    return JSON.parse(window.localStorage.getItem(INCOME_STORAGE_KEY) || "[]").map((row, index) => ({
      id: row.id || [row.date || "", row.source || "", row.description || "", row.amount || "", row.status || "", index].join("|"),
      date: row.date || "",
      source: row.source || "",
      description: row.description || "",
      amount: parseAmount(row.amount),
      status: row.status || "รับแล้ว",
    }));
  } catch (error) {
    return [];
  }
}

function saveIncomeRows(rows) {
  window.localStorage.setItem(INCOME_STORAGE_KEY, JSON.stringify(rows));
}

function isIncomeReceived(row) {
  return !String(row.status || "").includes("ค้าง");
}

function renderIncomeRows(rows) {
  const previewBody = document.getElementById("income-preview-body");
  const detailBody = document.getElementById("income-detail-body");
  const summary = document.getElementById("income-preview-summary");
  const total = rows.reduce((sum, row) => sum + row.amount, 0);
  const rowsHtml = rows.length
    ? rows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.date)}</td>
              <td>${escapeHtml(row.source)}</td>
              <td>${escapeHtml(row.description)}</td>
              <td>${formatCurrency(row.amount)}</td>
              <td><span class="chip income-status">${escapeHtml(row.status)}</span></td>
            </tr>
          `
        )
        .join("")
    : `<tr><td colspan="5" class="pending-empty-cell">ยังไม่มีข้อมูลรายรับ</td></tr>`;

  if (summary) {
    summary.textContent = `${rows.length} รายการ · ${formatCurrency(total)}`;
  }

  if (previewBody) {
    previewBody.innerHTML = rowsHtml;
  }

  if (detailBody) {
    detailBody.innerHTML = rowsHtml;
  }
}

function getIncomeSummaryFromState() {
  return dashboardState?.incomeSummary || window.SUMMARY_DASHBOARD_DATA.incomeSummary || normalizeIncomeSummaryMatrix([]);
}

function renderIncomeCompare(incomeSummary) {
  const root = document.getElementById("income-compare");

  if (!root) {
    return;
  }

  const salesMongo = parseAmount(incomeSummary.salesMongo);
  const receivedWithFee = parseAmount(incomeSummary.receivedWithFee);
  const fee = parseAmount(incomeSummary.fee);
  const channels = incomeSummary.channels || [];
  const maxValue = Math.max(
    salesMongo,
    receivedWithFee,
    ...channels.map((item) => Math.max(parseAmount(item.salesMongo), parseAmount(item.receivedWithFee))),
    0
  );
  const gap = salesMongo - receivedWithFee;

  root.innerHTML = `
    <div class="income-compare-summary">
      <div>
        <span>ส่วนต่าง</span>
        <strong class="${gap > 0 ? "income-gap" : "income-good"}">${formatCurrency(gap)}</strong>
      </div>
      <div>
        <span>Fee</span>
        <strong>${formatCurrency(fee)}</strong>
      </div>
      <div>
        <span>% รับจริง</span>
        <strong>${formatPercent(salesMongo > 0 ? (receivedWithFee / salesMongo) * 100 : 0)}</strong>
      </div>
    </div>
    <div class="income-bars">
      ${
        channels.length
          ? channels
              .map((channel) => {
                const channelSales = parseAmount(channel.salesMongo);
                const channelReceived = parseAmount(channel.receivedWithFee);
                const salesWidth = maxValue > 0 ? (channelSales / maxValue) * 100 : 0;
                const receivedWidth = maxValue > 0 ? (channelReceived / maxValue) * 100 : 0;

                return `
                  <div class="income-channel-row">
                    <div class="category-label">
                      <span>${escapeHtml(channel.channel)}</span>
                      <span>${formatPercent(channelSales > 0 ? (channelReceived / channelSales) * 100 : 0)}</span>
                    </div>
                    <div class="income-channel-meta">
                      <span>ยอดขาย (Mongo) ${formatCurrency(channelSales)}</span>
                      <span>รับจริงรวม ${formatCurrency(channelReceived)}</span>
                      <span>Fee ${formatCurrency(channel.fee)}</span>
                    </div>
                    <div class="dual-bars">
                      <div class="bar-set">
                        <span class="bar-set-label">Mongo</span>
                        <div class="bar-track income-bar-track">
                          <div class="bar-fill income-bar-fill sales" style="width: ${Math.max(salesWidth, channelSales > 0 ? 2 : 0)}%"></div>
                        </div>
                        <span class="bar-set-value">${formatCurrency(channelSales)}</span>
                      </div>
                      <div class="bar-set">
                        <span class="bar-set-label">รับจริง</span>
                        <div class="bar-track income-bar-track">
                          <div class="bar-fill income-bar-fill received" style="width: ${Math.max(receivedWidth, channelReceived > 0 ? 2 : 0)}%"></div>
                        </div>
                        <span class="bar-set-value">${formatCurrency(channelReceived)}</span>
                      </div>
                    </div>
                  </div>
                `;
              })
              .join("")
          : `<div class="pending-empty-cell">ยังไม่พบข้อมูลรายช่องทางจาก Summary รายรับ</div>`
      }
    </div>
  `;
}

function renderBankSummary(incomeSummary) {
  const bankSummary = incomeSummary.bankSummary || { latestDate: "", accounts: [], total: 0 };
  const latestDate = document.getElementById("bank-latest-date");
  const totalBalance = document.getElementById("bank-total-balance");
  const body = document.getElementById("bank-balance-body");

  if (latestDate) {
    latestDate.textContent = bankSummary.latestDate || "-";
  }

  if (totalBalance) {
    totalBalance.textContent = formatCurrency(bankSummary.total || 0);
  }

  if (!body) {
    return;
  }

  if (!bankSummary.accounts?.length) {
    body.innerHTML = `<tr><td colspan="2" class="pending-empty-cell">ยังไม่มีข้อมูลธนาคาร</td></tr>`;
    return;
  }

  body.innerHTML = bankSummary.accounts
    .map(
      (item) => `
        <tr>
          <td>${escapeHtml(item.account)}</td>
          <td>${formatCurrency(item.balance)}</td>
        </tr>
      `
    )
    .join("");
}

function renderIncomeDashboard(rows) {
  const kpiRoot = document.getElementById("income-kpi-grid");
  const topCard = document.getElementById("income-top-card");
  const noteList = document.getElementById("income-note-list");
  const incomeSummary = getIncomeSummaryFromState();
  const salesMongo = parseAmount(incomeSummary.salesMongo);
  const receivedWithFee = parseAmount(incomeSummary.receivedWithFee);
  const fee = parseAmount(incomeSummary.fee);
  const channels = incomeSummary.channels || [];
  const manualTotal = rows.reduce((sum, row) => sum + row.amount, 0);
  const total = salesMongo || manualTotal;
  const receivedTotal = receivedWithFee || rows.filter(isIncomeReceived).reduce((sum, row) => sum + row.amount, 0);
  const pendingTotal = Math.max(total - receivedTotal, 0);
  const sourceMap = new Map();

  rows.forEach((row) => {
    const sourceName = row.source || "ไม่ระบุช่องทาง";
    sourceMap.set(sourceName, (sourceMap.get(sourceName) || 0) + row.amount);
  });
  channels.forEach((row) => {
    sourceMap.set(row.channel, parseAmount(row.salesMongo));
  });

  const topSource = Array.from(sourceMap.entries()).sort((left, right) => right[1] - left[1])[0];

  if (kpiRoot) {
    const kpis = [
      { label: "ยอดขาย (Mongo)", value: formatCurrency(salesMongo) },
      { label: "รับจริงรวม", value: formatCurrency(receivedTotal) },
      { label: "Fee", value: formatCurrency(fee) },
      { label: "คงเหลือรับ", value: formatCurrency(pendingTotal) },
    ];

    kpiRoot.innerHTML = kpis
      .map(
        (item) => `
          <article class="kpi-card income-kpi-card">
            <div class="kpi-label">${item.label}</div>
            <div class="kpi-value">${item.value}</div>
          </article>
        `
      )
      .join("");
  }

  if (topCard) {
    topCard.innerHTML = salesMongo || receivedTotal
      ? `
        <div class="top-category-card">
          <div class="chip">ยอดขาย (Mongo)</div>
          <p class="top-category-amount">${formatCurrency(salesMongo)}</p>
          <div class="top-category-name">ยอดขาย (Mongo)</div>
          <p class="top-category-share">รับจริงรวม ${formatCurrency(receivedTotal)}</p>
        </div>
      `
      : "<p class=\"muted\">ยังไม่มีข้อมูลรายรับ</p>";
  }

  if (noteList) {
    const notes = salesMongo || receivedWithFee
      ? [
          `ยอดขาย (Mongo) คือ ${formatCurrency(salesMongo)}`,
          `รับจริงรวมคือ ${formatCurrency(receivedTotal)} คิดเป็น ${formatPercent(salesMongo > 0 ? (receivedTotal / salesMongo) * 100 : 0)}`,
          `ยังต่างจากยอดขายอยู่ ${formatCurrency(salesMongo - receivedTotal)}`,
          topSource ? `ช่องทางรายรับสูงสุดคือ ${topSource[0]} มูลค่า ${formatCurrency(topSource[1])}` : "ยังไม่พบช่องทางรายรับ",
        ]
      : ["เริ่มจากวางข้อมูลรายรับทางซ้าย แล้วกดบันทึกบนเครื่องนี้", "โครงนี้พร้อมต่อเข้าชีทรายรับจริงในขั้นถัดไป"];

    noteList.innerHTML = notes.map((item) => `<li>${item}</li>`).join("");
  }

  renderIncomeCompare(incomeSummary);
  renderBankSummary(incomeSummary);
  renderIncomeRows(rows);
}

function setupIncomeWorkspace() {
  const input = document.getElementById("income-paste-input");
  const previewButton = document.getElementById("income-preview-button");
  const saveButton = document.getElementById("income-save-button");
  const clearButton = document.getElementById("income-clear-button");

  incomeRows = loadIncomeRows();
  renderIncomeDashboard(incomeRows);

  if (!input || !previewButton || !saveButton || !clearButton) {
    return;
  }

  function previewRows() {
    const rows = parseIncomePaste(input.value);
    renderIncomeDashboard(rows.length ? rows : incomeRows);
    return rows;
  }

  input.addEventListener("paste", () => {
    window.setTimeout(previewRows, 0);
  });
  previewButton.addEventListener("click", previewRows);
  saveButton.addEventListener("click", () => {
    incomeRows = previewRows();
    saveIncomeRows(incomeRows);
    renderIncomeDashboard(incomeRows);
    setLiveStatus("บันทึกข้อมูลรายรับบนเครื่องนี้แล้ว");
  });
  clearButton.addEventListener("click", () => {
    input.value = "";
    incomeRows = [];
    saveIncomeRows(incomeRows);
    renderIncomeDashboard(incomeRows);
  });
}

function renderDashboard(sourceData) {
  const data = prepareData(sourceData);
  dashboardState = data;
  renderKpis(data);
  renderSourceInfo(data);
  renderCategoryCompare(data);
  renderTopCategory(data);
  renderInsights(data);
  renderDetailHead(data);
  renderTableRows(data);
  renderIncomeDashboard(incomeRows);
}

async function refreshDashboard() {
  setRefreshDisabled(true);
  setLiveStatus("กำลังดึงข้อมูลล่าสุดจาก Google Sheet...");

  try {
    const liveData = await fetchLiveSummaryData();
    updateGlobalSource(liveData.source);
    renderDashboard(liveData);
    setLiveStatus(liveData.source.liveJsonUrl ? "เชื่อมต่อ Real Time ผ่าน Apps Script สำเร็จ" : "เชื่อมต่อ Real Time สำเร็จ");
  } catch (error) {
    renderDashboard(window.SUMMARY_DASHBOARD_DATA);
    setLiveStatus("ดึงสดไม่ได้ ใช้ snapshot ล่าสุดแทน");
    console.warn("Live Google Sheet fetch failed", error);
  } finally {
    setRefreshDisabled(false);
  }
}

function bootstrap() {
  const refreshButton = document.getElementById("refresh-button");
  const searchInput = document.getElementById("search-input");
  const sourceConfig = getSourceConfig();

  updateGlobalSource(sourceConfig);

  refreshButton.addEventListener("click", () => {
    refreshDashboard();
  });

  searchInput.addEventListener("input", (event) => {
    if (dashboardState) {
      detailTableState.globalQuery = event.target.value;
      renderTableRows(dashboardState);
      renderDetailHead(dashboardState);
    }
  });

  setupSourceControls();
  setupMainViewTabs();
  setupPanelTabs();
  setupPendingPastePreview();
  setupIncomeWorkspace();
  refreshDashboard();
}

bootstrap();
