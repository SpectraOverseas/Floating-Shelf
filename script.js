const DATA_URL = "data/Combined.xlsx";
const SHEET_NAME = "Sheet1";
const TOKEN_STORAGE_KEY = "floating-shelf-gh-token";
const REPO_STORAGE_KEY = "floating-shelf-gh-repo";
const FILTER_COLUMNS = [
  "ASIN",
  "Colour",
  "Material",
  "L X W X H",
  "Seller Country/Region",
  "Seller",
];

const KPI_CONFIG = [
  { key: "asins", column: "ASIN", type: "unique" },
  { key: "colours", column: "Colour", type: "unique" },
  { key: "asinRevenue", column: "ASIN Revenue", type: "sum" },
  { key: "sellerCountries", column: "Seller Country/Region", type: "unique" },
  { key: "sellers", column: "Seller", type: "unique" },
];

const PRICE_COLUMN_KEY = "Price  $";

const filtersContainer = document.getElementById("filters");
const resetButton = document.getElementById("resetFilters");
const showRecordViewButton = document.getElementById("showRecordView");
const showGraphViewButton = document.getElementById("showGraphView");
const showTableViewButton = document.getElementById("showTableView");
const showTableViewFromRecordButton = document.getElementById(
  "showTableViewFromRecord"
);
const tableBody = document.querySelector("[data-table-body]");
const recordView = document.querySelector('[data-view="record"]');
const graphView = document.querySelector('[data-view="graph"]');
const tableView = document.querySelector('[data-view="table"]');
const sharedSection = document.querySelector('[data-shared="dashboard"]');
const form = document.getElementById("recordForm");
const formFields = document.getElementById("formFields");
const formMessage = document.getElementById("formMessage");
const connectionStatus = document.getElementById("connectionStatus");
const manageTokenButton = document.getElementById("manageToken");
const trendChartCanvas = document.getElementById("trendChart");
const comparisonChartCanvas = document.getElementById("comparisonChart");
const distributionChartCanvas = document.getElementById("distributionChart");
const kpiElements = new Map(
  Array.from(document.querySelectorAll("[data-kpi]")).map((el) => [
    el.dataset.kpi,
    el,
  ])
);

let rawData = [];
let lastDataSignature = "";
const activeFilters = {};

const numberFormatter = new Intl.NumberFormat("en-US", {
  maximumFractionDigits: 2,
});
const currencyFormatter = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const integerFormatter = new Intl.NumberFormat("en-US", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 0,
});
const kpiNumberFormatter = new Intl.NumberFormat("en-US", {
  maximumFractionDigits: 0,
});

const hexToRgba = (hex, alpha) => {
  if (!hex) {
    return `rgba(0, 0, 0, ${alpha})`;
  }
  let sanitized = hex.replace("#", "");
  if (sanitized.length === 3) {
    sanitized = sanitized
      .split("")
      .map((char) => char + char)
      .join("");
  }
  const value = Number.parseInt(sanitized, 16);
  if (Number.isNaN(value)) {
    return `rgba(0, 0, 0, ${alpha})`;
  }
  const red = (value >> 16) & 255;
  const green = (value >> 8) & 255;
  const blue = value & 255;
  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
};

const hashString = (value) => {
  if (!value) {
    return 0;
  }
  let hash = 0;
  for (let i = 0; i < value.length; i += 1) {
    hash = (hash << 5) - hash + value.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash);
};

const applyHorizontalJitter = (value, seed, jitterPercent = 0.02) => {
  if (!value || !seed) {
    return value;
  }
  const normalized = (hashString(seed) % 2001) / 2000;
  const offset = (normalized - 0.5) * 2 * jitterPercent;
  return value * (1 + offset);
};

const cleanValue = (value) => {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
};

const DATE_KEYWORDS = ["date", "month", "year", "created", "updated"];
const FORCED_NUMBER_COLUMNS = new Set(["Parent Level Sales", "Review Count"]);

const inferColumnType = (header, values) => {
  if (FORCED_NUMBER_COLUMNS.has(header)) {
    return "number";
  }

  const headerLower = header.toLowerCase();
  if (DATE_KEYWORDS.some((keyword) => headerLower.includes(keyword))) {
    return "date";
  }

  const sample = values.filter((value) => value !== "").slice(0, 30);
  if (!sample.length) {
    return "text";
  }

  const numericValues = sample.filter((value) => !Number.isNaN(Number(value)));
  if (numericValues.length === sample.length) {
    return "number";
  }

  const uniqueValues = Array.from(
    new Set(sample.map((value) => cleanValue(value)))
  ).filter(Boolean);
  if (uniqueValues.length > 0 && uniqueValues.length <= 8) {
    return "select";
  }

  return "text";
};

const inferSelectOptions = (values) => {
  const uniqueValues = Array.from(
    new Set(values.map((value) => cleanValue(value)))
  ).filter(Boolean);
  return uniqueValues.sort((a, b) => a.localeCompare(b));
};

const buildRepoDetails = () => {
  const saved = localStorage.getItem(REPO_STORAGE_KEY);
  if (saved) {
    try {
      return JSON.parse(saved);
    } catch (error) {
      console.warn("Failed to parse repo details", error);
    }
  }

  const host = window.location.hostname;
  const pathParts = window.location.pathname.split("/").filter(Boolean);
  if (host.endsWith("github.io") && pathParts.length > 0) {
    const owner = host.replace(".github.io", "");
    const repo = pathParts[0];
    return { owner, repo, branch: "main", path: DATA_URL };
  }

  return { owner: "", repo: "", branch: "main", path: DATA_URL };
};

const saveRepoDetails = (details) => {
  localStorage.setItem(REPO_STORAGE_KEY, JSON.stringify(details));
};

const getToken = () => localStorage.getItem(TOKEN_STORAGE_KEY) || "";

const updateConnectionStatus = (details) => {
  if (!details.owner || !details.repo) {
    connectionStatus.textContent =
      "Repository details are missing. Click Manage GitHub Token to configure the owner and repo.";
    return;
  }

  connectionStatus.textContent = `Saving to ${details.owner}/${details.repo}/${details.path} on ${details.branch}.`;
};

const requestToken = (details) => {
  const owner = window.prompt("GitHub owner/user name", details.owner || "");
  if (owner === null) {
    return null;
  }
  const repo = window.prompt("GitHub repository name", details.repo || "");
  if (repo === null) {
    return null;
  }
  const branch = window.prompt("Branch name", details.branch || "main");
  if (branch === null) {
    return null;
  }
  const token = window.prompt(
    "GitHub Personal Access Token with repo contents permission (stored locally)",
    getToken() || ""
  );
  if (token === null) {
    return null;
  }
  const nextDetails = {
    owner: owner.trim(),
    repo: repo.trim(),
    branch: branch.trim() || "main",
    path: details.path,
  };
  saveRepoDetails(nextDetails);
  if (token.trim()) {
    localStorage.setItem(TOKEN_STORAGE_KEY, token.trim());
  }
  updateConnectionStatus(nextDetails);
  return nextDetails;
};

const fetchWorkbook = async () => {
  const response = await fetch(buildDataUrl(), {
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error("Unable to load workbook.");
  }
  const arrayBuffer = await response.arrayBuffer();
  return XLSX.read(arrayBuffer, { type: "array" });
};

const setValidationMessage = (input, fieldType) => {
  const value = input.value.trim();
  if (!value) {
    input.setCustomValidity("");
    return;
  }
  if (fieldType === "number" && Number.isNaN(Number(value))) {
    input.setCustomValidity("Please enter a valid number.");
    return;
  }
  input.setCustomValidity("");
};

const buildForm = (headers, rows) => {
  formFields.replaceChildren();
  headers.forEach((header, index) => {
    const fieldWrapper = document.createElement("div");
    fieldWrapper.className = "form-field";

    const label = document.createElement("label");
    label.textContent = header;
    label.setAttribute("for", `field-${index}`);

    const columnValues = rows.map((row) => row[index]);
    const fieldType = inferColumnType(header, columnValues);
    let input;
    let datalist;

    if (fieldType === "select") {
      input = document.createElement("input");
      datalist = document.createElement("datalist");
      const listId = `field-options-${index}`;
      datalist.id = listId;
      input.setAttribute("list", listId);
      input.placeholder = "Select or type";
      inferSelectOptions(columnValues).forEach((optionValue) => {
        const option = document.createElement("option");
        option.value = optionValue;
        datalist.appendChild(option);
      });
    } else {
      input = document.createElement("input");
      input.type = fieldType;
      if (fieldType === "date") {
        input.placeholder = "YYYY-MM-DD";
      } else if (fieldType === "number") {
        input.step = "any";
        input.inputMode = "decimal";
      }
    }

    input.id = `field-${index}`;
    input.name = header;
    input.required = true;
    if (fieldType === "number") {
      fieldWrapper.classList.add("form-field--number");
    }
    input.addEventListener("input", () => setValidationMessage(input, fieldType));

    fieldWrapper.appendChild(label);
    fieldWrapper.appendChild(input);
    if (datalist) {
      fieldWrapper.appendChild(datalist);
    }
    formFields.appendChild(fieldWrapper);
  });
};

const encodeBase64 = (arrayBuffer) => {
  let binary = "";
  const bytes = new Uint8Array(arrayBuffer);
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
};

const isRowEmpty = (row) =>
  !row || row.every((cell) => cleanValue(cell) === "");

const getReindexColumnIndex = (headers) => {
  const candidates = new Set([
    "index",
    "row",
    "row #",
    "row no",
    "row number",
    "s.no",
    "s. no",
    "serial no",
    "serial number",
  ]);
  return headers.findIndex((header) =>
    candidates.has(cleanValue(header).toLowerCase())
  );
};

const compactSheetData = (data, fallbackHeaders) => {
  const safeHeaders = data[0] || fallbackHeaders;
  const rows = data.slice(1).filter((row) => !isRowEmpty(row));
  const reindexColumn = getReindexColumnIndex(safeHeaders);
  const normalizedRows =
    reindexColumn >= 0
      ? rows.map((row, rowIndex) => {
          const nextRow = [...row];
          nextRow[reindexColumn] = rowIndex + 1;
          return nextRow;
        })
      : rows;
  return { headers: safeHeaders, rows: normalizedRows };
};

const appendRowToSheet = (workbook, headers, rowValues) => {
  const sheet = workbook.Sheets[SHEET_NAME];
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  }

  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const { headers: safeHeaders, rows: compactedRows } = compactSheetData(
    data,
    headers
  );
  const nextRow = safeHeaders.map((header, index) => {
    const value = rowValues[index];
    return value === undefined || value === null ? "" : value;
  });
  compactedRows.push(nextRow);
  const normalizedData = [safeHeaders, ...compactedRows];

  const updatedSheet = XLSX.utils.aoa_to_sheet(normalizedData);
  workbook.Sheets[SHEET_NAME] = updatedSheet;
  return workbook;
};

const fetchRepoFile = async (details, token) => {
  const apiUrl = `https://api.github.com/repos/${details.owner}/${details.repo}/contents/${details.path}`;
  const response = await fetch(apiUrl, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github+json",
    },
  });
  if (!response.ok) {
    throw new Error("Unable to fetch repository file metadata.");
  }
  return response.json();
};

const updateRepoFile = async (details, token, content, sha) => {
  const apiUrl = `https://api.github.com/repos/${details.owner}/${details.repo}/contents/${details.path}`;
  const response = await fetch(apiUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github+json",
    },
    body: JSON.stringify({
      message: "Append record to Combined.xlsx",
      content,
      sha,
      branch: details.branch,
    }),
  });
  if (!response.ok) {
    throw new Error("Unable to write to repository file.");
  }
  return response.json();
};

const formatMessage = (message, isError = false) => {
  formMessage.textContent = message;
  formMessage.classList.toggle("is-error", isError);
};

const findColumnByKeywords = (columns, keywords) => {
  const lowerColumns = columns.map((column) => ({
    key: column,
    lower: column.toLowerCase(),
  }));
  for (const keyword of keywords) {
    const match = lowerColumns.find((column) => column.lower.includes(keyword));
    if (match) {
      return match.key;
    }
  }
  return "";
};

const parseDateValue = (value) => {
  if (!value) {
    return null;
  }
  const raw = cleanValue(value);
  const parsed = Date.parse(raw);
  if (!Number.isNaN(parsed)) {
    return new Date(parsed);
  }
  const asNumber = Number(raw);
  if (!Number.isNaN(asNumber)) {
    return new Date(asNumber, 0, 1);
  }
  return null;
};

let trendChart;
let comparisonChart;
let distributionChart;
const parseNumericValue = (value) => {
  const normalized = cleanValue(value);
  if (!normalized) {
    return null;
  }
  const raw = normalized.replace(/[^0-9.-]+/g, "");
  if (!raw) {
    return null;
  }
  const parsed = Number.parseFloat(raw);
  return Number.isFinite(parsed) ? parsed : null;
};

const formatCurrencyValue = (value) => {
  const parsed = parseNumericValue(value);
  if (parsed === null) {
    return cleanValue(value);
  }
  return currencyFormatter.format(parsed);
};

const formatIntegerValue = (value) => {
  const parsed = parseNumericValue(value);
  if (parsed === null) {
    return cleanValue(value);
  }
  return integerFormatter.format(parsed);
};

const TABLE_COLUMNS = [
  { label: "Design", key: "Design" },
  { label: "Seller", key: "Seller" },
  { label: "ASIN", key: "ASIN" },
  { label: "Pack", key: "Pack", formatter: formatIntegerValue, isNumeric: true },
  { label: "L X W X H", key: "L X W X H" },
  { label: "Colour", key: "Colour" },
  { label: "Advantage", key: "Advantage" },
  {
    label: "Price $",
    key: PRICE_COLUMN_KEY,
    formatter: formatCurrencyValue,
    isNumeric: true,
  },
  {
    label: "ASIN Revenue",
    key: "ASIN Revenue",
    formatter: formatCurrencyValue,
    isNumeric: true,
  },
];

const applyNumericAlignment = () => {
  const headerCells = document.querySelectorAll(".data-table thead th");
  TABLE_COLUMNS.forEach((column, index) => {
    if (!column.isNumeric) {
      return;
    }
    const cell = headerCells[index];
    if (cell) {
      cell.classList.add("data-table__cell--numeric");
    }
  });
};

const uniqueValues = (rows, column) => {
  const values = new Set();
  rows.forEach((row) => {
    const value = cleanValue(row[column]);
    if (value) {
      values.add(value);
    }
  });
  return Array.from(values).sort((a, b) => a.localeCompare(b));
};

const buildFilter = (column, values) => {
  const wrapper = document.createElement("div");
  wrapper.className = "filter";

  const label = document.createElement("label");
  label.textContent = column;
  wrapper.appendChild(label);

  const select = document.createElement("div");
  select.className = "multi-select";
  select.dataset.filter = column;

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "multi-select__toggle";
  toggle.textContent = "All";

  const menu = document.createElement("div");
  menu.className = "multi-select__menu";

  const options = ["All", ...values];
  options.forEach((value) => {
    const option = document.createElement("label");
    option.className = "multi-select__option";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = value;
    checkbox.checked = value === "All";

    const text = document.createElement("span");
    text.textContent = value;

    option.appendChild(checkbox);
    option.appendChild(text);
    menu.appendChild(option);
  });

  toggle.addEventListener("click", () => {
    select.classList.toggle("open");
  });

  select.appendChild(toggle);
  select.appendChild(menu);
  wrapper.appendChild(select);
  filtersContainer.appendChild(wrapper);

  activeFilters[column] = new Set(["All"]);
};

const updateToggleLabel = (selectElement, selections) => {
  const toggle = selectElement.querySelector(".multi-select__toggle");
  if (selections.has("All") || selections.size === 0) {
    toggle.textContent = "All";
  } else if (selections.size === 1) {
    toggle.textContent = Array.from(selections)[0];
  } else {
    toggle.textContent = `${selections.size} selected`;
  }
};

const syncCheckboxes = (selectElement, selections) => {
  selectElement
    .querySelectorAll("input[type='checkbox']")
    .forEach((checkbox) => {
      checkbox.checked = selections.has(checkbox.value);
    });
};

const applyFilterRules = (column, value, checked) => {
  const selections = activeFilters[column];

  if (value === "All" && checked) {
    selections.clear();
    selections.add("All");
  } else if (value !== "All") {
    if (checked) {
      selections.delete("All");
      selections.add(value);
    } else {
      selections.delete(value);
      if (selections.size === 0) {
        selections.add("All");
      }
    }
  }

  return selections;
};

const filteredRows = () => {
  return rawData.filter((row) => {
    return FILTER_COLUMNS.every((column) => {
      const selections = activeFilters[column];
      if (!selections || selections.has("All")) {
        return true;
      }
      const value = cleanValue(row[column]);
      return selections.has(value);
    });
  });
};

const calculateKpis = (rows) => {
  const result = {};

  KPI_CONFIG.forEach((kpi) => {
    if (kpi.type === "unique") {
      const values = new Set();
      rows.forEach((row) => {
        const value = cleanValue(row[kpi.column]);
        if (value) {
          values.add(value);
        }
      });
      result[kpi.key] = values.size;
    }

    if (kpi.type === "sum") {
      const total = rows.reduce((sum, row) => {
        const rawValue = cleanValue(row[kpi.column]).replace(/[^0-9.-]+/g, "");
        const parsed = Number.parseFloat(rawValue);
        return sum + (Number.isFinite(parsed) ? parsed : 0);
      }, 0);
      result[kpi.key] = total;
    }
  });

  return result;
};

const renderKpis = (rows) => {
  const values = calculateKpis(rows);
  KPI_CONFIG.forEach((kpi) => {
    const element = kpiElements.get(kpi.key);
    if (!element) {
      return;
    }
    const value = values[kpi.key] ?? 0;
    element.textContent = kpiNumberFormatter.format(Math.round(value));
  });
};

const renderTable = (rows) => {
  if (!tableBody) {
    return;
  }

  const fragment = document.createDocumentFragment();

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    TABLE_COLUMNS.forEach((column) => {
      const td = document.createElement("td");
      const rawValue = row[column.key];
      td.textContent = column.formatter
        ? column.formatter(rawValue)
        : cleanValue(rawValue);
      if (column.isNumeric) {
        td.classList.add("data-table__cell--numeric");
      }
      tr.appendChild(td);
    });
    fragment.appendChild(tr);
  });

  tableBody.replaceChildren(fragment);
};

const updateDashboard = () => {
  const rows = filteredRows();
  renderKpis(rows);
  renderTable(rows);
  renderCharts(rows);
};

const attachFilterListeners = () => {
  filtersContainer.addEventListener("change", (event) => {
    const checkbox = event.target;
    if (checkbox.tagName !== "INPUT") {
      return;
    }
    const selectElement = checkbox.closest(".multi-select");
    const column = selectElement.dataset.filter;
    const selections = applyFilterRules(column, checkbox.value, checkbox.checked);
    syncCheckboxes(selectElement, selections);
    updateToggleLabel(selectElement, selections);
    updateDashboard();
  });

  document.addEventListener("click", (event) => {
    if (event.target.closest(".multi-select")) {
      return;
    }
    document.querySelectorAll(".multi-select.open").forEach((select) => {
      select.classList.remove("open");
    });
  });
};

const resetFilters = () => {
  Object.keys(activeFilters).forEach((column) => {
    activeFilters[column] = new Set(["All"]);
  });

  document.querySelectorAll(".multi-select").forEach((select) => {
    const column = select.dataset.filter;
    const selections = activeFilters[column];
    syncCheckboxes(select, selections);
    updateToggleLabel(select, selections);
  });

  updateDashboard();
};

const buildDataUrl = () => `${DATA_URL}?v=${Date.now()}`;

const updateFilters = (rows) => {
  const previousSelections = {};
  Object.keys(activeFilters).forEach((column) => {
    previousSelections[column] = new Set(activeFilters[column]);
  });

  filtersContainer.replaceChildren();
  FILTER_COLUMNS.forEach((column) => {
    const values = uniqueValues(rows, column);
    buildFilter(column, values);
    const selectElement = filtersContainer.querySelector(
      `.multi-select[data-filter="${column}"]`
    );
    const savedSelections = previousSelections[column];
    if (!selectElement || !savedSelections) {
      return;
    }
    const available = new Set(["All", ...values]);
    const nextSelections = new Set(
      Array.from(savedSelections).filter((value) => available.has(value))
    );
    if (nextSelections.size === 0) {
      nextSelections.add("All");
    }
    activeFilters[column] = nextSelections;
    syncCheckboxes(selectElement, nextSelections);
    updateToggleLabel(selectElement, nextSelections);
  });
};

const loadData = async () => {
  const response = await fetch(buildDataUrl(), {
    cache: "no-store",
  });
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[SHEET_NAME];
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" not found in workbook.`);
  }
  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
  });
  return rows;
};

const getDataSignature = (rows) => {
  if (!rows.length) {
    return "0";
  }
  const lastRow = rows[rows.length - 1];
  return `${rows.length}-${JSON.stringify(lastRow)}`;
};

const refreshData = async () => {
  try {
    const rows = await loadData();
    const signature = getDataSignature(rows);
    if (signature !== lastDataSignature) {
      rawData = rows;
      lastDataSignature = signature;
      updateFilters(rawData);
      updateDashboard();
    }
  } catch (error) {
    console.error("Failed to refresh data", error);
  }
};

const buildChartCard = (canvas, config) => {
  if (!canvas) {
    return null;
  }
  return new Chart(canvas, config);
};

const buildBubbleChartData = (rows) => {
  if (!rows.length) {
    return { datasets: [], xMax: 0, yMax: 0 };
  }
  const columns = Object.keys(rows[0]);
  const priceColumn =
    columns.find((column) => column === PRICE_COLUMN_KEY) ||
    findColumnByKeywords(columns, ["price"]);
  const revenueColumn = columns.find(
    (column) => column.toLowerCase() === "asin revenue"
  );
  const reviewCountColumn = columns.find(
    (column) => column.toLowerCase() === "review count"
  );
  const asinColumn = columns.find((column) => column.toLowerCase() === "asin");
  const sellerColumn = columns.find(
    (column) => column.toLowerCase() === "seller"
  );
  const ratingsColumn = columns.find(
    (column) => column.toLowerCase() === "ratings"
  );
  const fulfillmentColumn = columns.find(
    (column) => column.toLowerCase() === "fulfillment"
  );
  const groupColumn = columns.find(
    (column) => column.toLowerCase() === "seller country/region"
  );

  if (!priceColumn || !revenueColumn) {
    return { datasets: [], xMax: 0, yMax: 0 };
  }

  const palette = [
    "#2563eb",
    "#38bdf8",
    "#22c55e",
    "#f97316",
    "#a855f7",
    "#facc15",
  ];
  const MAX_BUBBLE_POINTS = 80;
  const preparedRows = rows
    .map((row) => {
      const priceValue = parseNumericValue(row[priceColumn]);
      const revenueValue = parseNumericValue(row[revenueColumn]);
      if (priceValue === null || revenueValue === null) {
        return null;
      }
      return {
        row,
        priceValue,
        revenueValue,
        reviewCount: reviewCountColumn
          ? parseNumericValue(row[reviewCountColumn]) ?? 0
          : 0,
      };
    })
    .filter(Boolean);

  if (!preparedRows.length) {
    return { datasets: [], xMax: 0, yMax: 0 };
  }

  const displayRows =
    preparedRows.length > MAX_BUBBLE_POINTS
      ? [...preparedRows]
          .sort((a, b) => b.revenueValue - a.revenueValue)
          .slice(0, MAX_BUBBLE_POINTS)
      : preparedRows;

  const maxPrice = Math.max(
    ...displayRows.map((entry) => entry.priceValue),
    0
  );
  const maxRevenue = Math.max(
    ...displayRows.map((entry) => entry.revenueValue),
    0
  );

  const groups = new Map();
  displayRows.forEach((entry) => {
    const groupLabel = cleanValue(entry.row[groupColumn]) || "Unknown";
    const jitterSeed =
      cleanValue(entry.row[asinColumn]) ||
      cleanValue(entry.row[sellerColumn]) ||
      String(entry.priceValue);
    const jitteredPrice = applyHorizontalJitter(entry.priceValue, jitterSeed);
    const point = {
      x: jitteredPrice,
      y: entry.revenueValue,
      asin: cleanValue(entry.row[asinColumn]),
      seller: cleanValue(entry.row[sellerColumn]),
      price: entry.priceValue,
      revenue: entry.revenueValue,
      ratings: cleanValue(entry.row[ratingsColumn]),
      reviewCount: entry.reviewCount,
      fulfillment: cleanValue(entry.row[fulfillmentColumn]),
    };
    if (!groups.has(groupLabel)) {
      groups.set(groupLabel, []);
    }
    groups.get(groupLabel).push(point);
  });

  const datasets = Array.from(groups.entries()).map(([label, data], index) => {
    const color = palette[index % palette.length];
    return {
      label,
      data,
      backgroundColor: hexToRgba(color, 0.65),
      borderColor: hexToRgba(color, 0.7),
      borderWidth: 1,
      pointRadius: 6,
    };
  });

  return {
    datasets,
    xMax: maxPrice + 10,
    yMax: maxRevenue * 1.50,
  };
};

const buildComparisonData = (rows) => {
  if (!rows.length) {
    return { labels: [], values: [], valueLabel: "" };
  }
  const columns = Object.keys(rows[0]);
  const valueColumn = columns.find(
    (column) => column.toLowerCase() === "asin revenue"
  );
  const categoryColumn = columns.find(
    (column) => column.toLowerCase() === "seller"
  );
  if (!valueColumn || !categoryColumn) {
    return { labels: [], values: [], valueLabel: "" };
  }

  const totals = new Map();
  rows.forEach((row) => {
    const category = cleanValue(row[categoryColumn]);
    if (!category) {
      return;
    }
    const value = parseNumericValue(row[valueColumn]) ?? 0;
    totals.set(category, (totals.get(category) || 0) + value);
  });

  const sorted = Array.from(totals.entries()).sort((a, b) => b[1] - a[1]);
  const topEntries = sorted.slice(0, 10);

  return {
    labels: topEntries.map(([label]) => label),
    values: topEntries.map(([, value]) => value),
    valueLabel: `${valueColumn} by ${categoryColumn}`,
  };
};

const buildComparisonTooltipData = (rows) => {
  if (!rows.length) {
    return new Map();
  }
  const columns = Object.keys(rows[0]);
  const valueColumn = columns.find(
    (column) => column.toLowerCase() === "asin revenue"
  );
  const categoryColumn = columns.find(
    (column) => column.toLowerCase() === "seller"
  );
  const advantageColumn = columns.find(
    (column) => column.toLowerCase() === "advantage"
  );
  if (!valueColumn || !categoryColumn) {
    return new Map();
  }

  const details = new Map();
  rows.forEach((row) => {
    const seller = cleanValue(row[categoryColumn]);
    if (!seller) {
      return;
    }
    const value = parseNumericValue(row[valueColumn]) ?? 0;
    const entry = details.get(seller) || {
      total: 0,
      advantageCounts: new Map(),
      firstAdvantage: "",
    };
    entry.total += value;
    if (advantageColumn) {
      const advantage = cleanValue(row[advantageColumn]);
      if (advantage) {
        if (!entry.firstAdvantage) {
          entry.firstAdvantage = advantage;
        }
        entry.advantageCounts.set(
          advantage,
          (entry.advantageCounts.get(advantage) || 0) + 1
        );
      }
    }
    details.set(seller, entry);
  });
  return details;
};

const resolveAdvantageValue = (entry) => {
  if (!entry) {
    return "";
  }
  let topAdvantage = "";
  let topCount = -1;
  entry.advantageCounts.forEach((count, advantage) => {
    if (count > topCount) {
      topAdvantage = advantage;
      topCount = count;
    }
  });
  return topAdvantage || entry.firstAdvantage || "";
};

const buildDistributionData = (rows) => {
  if (!rows.length) {
    return { labels: [], values: [], valueLabel: "" };
  }
  const columns = Object.keys(rows[0]);
  const categoryColumn = findColumnByKeywords(columns, [
    "seller",
    "country",
    "colour",
    "material",
  ]);
  if (!categoryColumn) {
    return { labels: [], values: [], valueLabel: "" };
  }
  const counts = new Map();
  rows.forEach((row) => {
    const value = cleanValue(row[categoryColumn]);
    if (!value) {
      return;
    }
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  const sorted = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);
  const topEntries = sorted.slice(0, 6);
  return {
    labels: topEntries.map(([label]) => label),
    values: topEntries.map(([, value]) => value),
    valueLabel: `${categoryColumn} share`,
  };
};

const renderCharts = (rows) => {
  if (!trendChartCanvas || !comparisonChartCanvas || !distributionChartCanvas) {
    return;
  }

  const bubbleData = buildBubbleChartData(rows);
  const comparisonData = buildComparisonData(rows);
  const distributionData = buildDistributionData(rows);
  const comparisonTooltipDetails = buildComparisonTooltipData(rows);

  if (!trendChart) {
    trendChart = buildChartCard(trendChartCanvas, {
      type: "scatter",
      data: {
        datasets: bubbleData.datasets,
      },
      options: {
        responsive: true,
        plugins: {
          legend: {
            position: "bottom",
            align: "center",
          },
          tooltip: {
            callbacks: {
              title: (items) => {
                const raw = items[0]?.raw || {};
                return raw.asin ? `ASIN: ${raw.asin}` : "ASIN: N/A";
              },
              label: (context) => {
                const raw = context.raw || {};
                const currencyFormatter = new Intl.NumberFormat("en-US", {
                  style: "currency",
                  currency: "USD",
                  maximumFractionDigits: 2,
                });
                return [
                  `Seller: ${raw.seller || "N/A"}`,
                  `Price $: ${currencyFormatter.format(raw.price ?? 0)}`,
                  `ASIN Revenue: ${currencyFormatter.format(raw.revenue ?? 0)}`,
                  `Ratings: ${raw.ratings || "N/A"}`,
                  `Review Count: ${numberFormatter.format(raw.reviewCount ?? 0)}`,
                  `Fulfillment: ${raw.fulfillment || "N/A"}`,
                ];
              },
            },
          },
        },
        scales: {
          x: {
            min: 0,
            max: bubbleData.xMax,
            title: {
              display: true,
              text: "Product Price ($)",
            },
            ticks: {
              callback: (value) => numberFormatter.format(value),
            },
          },
          y: {
            type: "logarithmic",
            min: 1,
            max: bubbleData.yMax,
            title: {
              display: true,
              text: "ASIN Revenue (Log Scale)",
            },
            ticks: {
              callback: (value) => numberFormatter.format(value),
            },
          },
        },
      },
    });
  } else {
    trendChart.data.datasets = bubbleData.datasets;
    trendChart.options.scales.x.min = 0;
    trendChart.options.scales.x.max = bubbleData.xMax;
    trendChart.options.scales.y.type = "logarithmic";
    trendChart.options.scales.y.min = 1;
    trendChart.options.scales.y.max = bubbleData.yMax;
    trendChart.update();
  }

  if (!comparisonChart) {
    comparisonChart = buildChartCard(comparisonChartCanvas, {
      type: "bar",
      data: {
        labels: comparisonData.labels,
        datasets: [
          {
            label: comparisonData.valueLabel || "Totals",
            data: comparisonData.values,
            backgroundColor: ["#1E3A8A", "#3B82F6", "#10B981"],
            borderRadius: 8,
          },
        ],
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              title: () => "",
              label: (context) => {
                const seller = context.label || "";
                const details = context.chart.$comparisonTooltipDetails;
                const entry = details ? details.get(seller) : null;
                const revenue = entry ? entry.total : context.parsed.y ?? 0;
                const advantage = resolveAdvantageValue(entry);
                const formattedRevenue = new Intl.NumberFormat("en-US", {
                  style: "currency",
                  currency: "USD",
                  maximumFractionDigits: 2,
                }).format(revenue ?? 0);
                return [
                  `Seller: ${seller || "N/A"}`,
                  `ASIN Revenue: ${formattedRevenue}`,
                  `Advantage: ${advantage || "N/A"}`,
                ];
              },
            },
          },
        },
        scales: {
          y: {
            ticks: {
              callback: (value) => numberFormatter.format(value),
            },
          },
        },
      },
    });
    comparisonChart.$comparisonTooltipDetails = comparisonTooltipDetails;
  } else {
    comparisonChart.data.labels = comparisonData.labels;
    comparisonChart.data.datasets[0].data = comparisonData.values;
    comparisonChart.$comparisonTooltipDetails = comparisonTooltipDetails;
    comparisonChart.update();
  }

  if (!distributionChart) {
    distributionChart = buildChartCard(distributionChartCanvas, {
      type: "pie",
      data: {
        labels: distributionData.labels,
        datasets: [
          {
            label: distributionData.valueLabel || "Distribution",
            data: distributionData.values,
            backgroundColor: [
              "#2563eb",
              "#38bdf8",
              "#22c55e",
              "#f97316",
              "#a855f7",
              "#facc15",
            ],
          },
        ],
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: "bottom" },
        },
      },
    });
  } else {
    distributionChart.data.labels = distributionData.labels;
    distributionChart.data.datasets[0].data = distributionData.values;
    distributionChart.update();
  }
};

const setActiveView = (viewName) => {
  const views = {
    record: recordView,
    graph: graphView,
    table: tableView,
  };

  Object.entries(views).forEach(([key, view]) => {
    if (!view) {
      return;
    }
    view.classList.toggle("is-active", key === viewName);
  });

  if (sharedSection) {
    sharedSection.classList.toggle("is-hidden", viewName === "record");
  }
};

const initRecordForm = async () => {
  if (!form) {
    return;
  }
  const repoDetails = buildRepoDetails();
  updateConnectionStatus(repoDetails);
  if (manageTokenButton) {
    manageTokenButton.addEventListener("click", () => {
      requestToken(repoDetails);
    });
  }

  try {
    const workbook = await fetchWorkbook();
    const sheet = workbook.Sheets[SHEET_NAME];
    if (!sheet) {
      throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    }
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const headers = data[0];
    if (!headers || headers.length === 0) {
      throw new Error(`${SHEET_NAME} does not contain headers.`);
    }
    buildForm(headers, data.slice(1));
  } catch (error) {
    formatMessage(error.message, true);
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    formatMessage("");

    const repoDetails = buildRepoDetails();
    const token = getToken();
    if (!repoDetails.owner || !repoDetails.repo || !token) {
      const updated = requestToken(repoDetails);
      if (!updated || !getToken()) {
        formatMessage("Repository details and token are required to save.", true);
        return;
      }
    }

    const inputs = Array.from(formFields.querySelectorAll("input"));
    inputs.forEach((input) => {
      const fieldType = input.type;
      setValidationMessage(input, fieldType);
    });
    if (!form.checkValidity()) {
      form.reportValidity();
      formatMessage("Please correct the highlighted fields.", true);
      return;
    }

    const rowValues = inputs.map((input) => input.value.trim());

    try {
      const workbook = await fetchWorkbook();
      const sheet = workbook.Sheets[SHEET_NAME];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const headers = data[0] || [];

      appendRowToSheet(workbook, headers, rowValues);

      const arrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const content = encodeBase64(arrayBuffer);
      const details = buildRepoDetails();
      const fileMeta = await fetchRepoFile(details, getToken());
      await updateRepoFile(details, getToken(), content, fileMeta.sha);

      formatMessage("Record added successfully");
      await refreshData();
      setActiveView("table");
      form.reset();
    } catch (error) {
      formatMessage(error.message || "Unable to save record.", true);
    }
  });
};

const init = async () => {
  rawData = await loadData();
  lastDataSignature = getDataSignature(rawData);
  updateFilters(rawData);
  attachFilterListeners();
  applyNumericAlignment();
  resetButton.addEventListener("click", resetFilters);
  if (showRecordViewButton) {
    showRecordViewButton.addEventListener("click", () => {
      setActiveView("record");
    });
  }
  if (showGraphViewButton) {
    showGraphViewButton.addEventListener("click", () => {
      setActiveView("graph");
    });
  }
  if (showTableViewButton) {
    showTableViewButton.addEventListener("click", () => {
      setActiveView("table");
    });
  }
  if (showTableViewFromRecordButton) {
    showTableViewFromRecordButton.addEventListener("click", () => {
      setActiveView("table");
    });
  }
  await initRecordForm();
  setActiveView("table");
  updateDashboard();
  window.setInterval(refreshData, 15000);
};

init();
