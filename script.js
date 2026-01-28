const FILTER_COLUMNS = [
  "ASIN",
  "L X W X H",
  "Colour",
  "Advantage",
  "Seller Country/Region",
  "Seller",
];

const state = {
  data: [],
  columns: [],
  filters: {},
  filteredData: [],
};

const elements = {
  loadingIndicator: document.getElementById("loadingIndicator"),
  errorMessage: document.getElementById("errorMessage"),
  filtersContainer: document.getElementById("filtersContainer"),
  resetFilters: document.getElementById("resetFilters"),
  kpiContainer: document.getElementById("kpiContainer"),
  tableHead: document.getElementById("tableHead"),
  tableBody: document.getElementById("tableBody"),
  tableCount: document.getElementById("tableCount"),
  filterSelects: {},
};

const numberFormatter = new Intl.NumberFormat("en-US");

const formatValue = (value) => {
  if (value instanceof Date) {
    return value.toLocaleDateString();
  }
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
};

const showError = (message) => {
  elements.errorMessage.textContent = message;
};

const clearError = () => {
  elements.errorMessage.textContent = "";
};

const normalizeHeaders = (headerRow) => {
  const seen = new Map();
  return headerRow.map((header, index) => {
    const base = header ? String(header).trim() : `Column ${index + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count ? `${base} (${count + 1})` : base;
  });
};

const loadExcel = async () => {
  try {
    clearError();
    elements.loadingIndicator.style.display = "inline-flex";

    const response = await fetch("data/Combined.xlsx");
    if (!response.ok) {
      throw new Error("Unable to load Combined.xlsx. Please confirm the file path.");
    }

    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true });

    const sheetName = workbook.SheetNames.find((name) => name === "Sheet1");
    if (!sheetName) {
      throw new Error("Sheet1 is missing. Please ensure the Excel file includes a Sheet1 tab.");
    }

    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: true,
      blankrows: false,
    });

    if (!rows.length) {
      throw new Error("Sheet1 does not contain any data.");
    }

    const headers = normalizeHeaders(rows[0]);
    const dataRows = rows.slice(1).filter((row) =>
      row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== "")
    );

    const parsedData = dataRows.map((row) => {
      const record = {};
      headers.forEach((header, index) => {
        record[header] = row[index] !== undefined ? row[index] : "";
      });
      return record;
    });

    state.data = parsedData;
    state.columns = headers;
    state.filters = FILTER_COLUMNS.reduce((acc, column) => {
      acc[column] = new Set();
      return acc;
    }, {});

    buildFilters();
    updateFilteredData();
  } catch (error) {
    showError(error.message || "An unexpected error occurred while loading the Excel file.");
  } finally {
    elements.loadingIndicator.style.display = "none";
  }
};

const buildFilters = () => {
  elements.filtersContainer.innerHTML = "";
  elements.filterSelects = {};

  FILTER_COLUMNS.forEach((column) => {
    const filterCard = document.createElement("div");
    filterCard.className = "filter-card";

    const label = document.createElement("label");
    label.textContent = column;
    label.setAttribute("for", `filter-${column}`);

    const select = document.createElement("select");
    select.id = `filter-${column}`;
    select.multiple = true;
    select.size = 6;
    select.dataset.column = column;

    const values = new Set(state.data.map((row) => formatValue(row[column])));

    [...values]
      .filter((value) => value !== "")
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
      .forEach((value) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        select.appendChild(option);
      });

    if (values.has("")) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "(Blank)";
      select.appendChild(option);
    }

    select.addEventListener("change", (event) => {
      const selected = new Set([...event.target.selectedOptions].map((opt) => opt.value));
      state.filters[column] = selected;
      updateFilteredData();
    });

    filterCard.appendChild(label);
    filterCard.appendChild(select);
    elements.filtersContainer.appendChild(filterCard);
    elements.filterSelects[column] = select;
  });
};

const applyFilters = () => {
  return state.data.filter((row) =>
    FILTER_COLUMNS.every((column) => {
      const selected = state.filters[column];
      if (!selected || selected.size === 0) {
        return true;
      }
      const value = formatValue(row[column]);
      return selected.has(value);
    })
  );
};

const getUniqueCount = (data, column) => {
  if (!state.columns.includes(column)) {
    return 0;
  }
  const unique = new Set(
    data
      .map((row) => formatValue(row[column]))
      .filter((value) => value !== "")
  );
  return unique.size;
};

const renderKpis = () => {
  const totalRows = state.filteredData.length;
  const asinCount = getUniqueCount(state.filteredData, "ASIN");
  const sellerCount = getUniqueCount(state.filteredData, "Seller");
  const colourCount = getUniqueCount(state.filteredData, "Colour");

  const kpis = [
    { label: "Total Records", value: numberFormatter.format(totalRows) },
    { label: "Unique ASINs", value: numberFormatter.format(asinCount) },
    { label: "Unique Sellers", value: numberFormatter.format(sellerCount) },
    { label: "Unique Colours", value: numberFormatter.format(colourCount) },
  ];

  elements.kpiContainer.innerHTML = "";
  kpis.forEach((kpi) => {
    const card = document.createElement("div");
    card.className = "kpi-card";

    const label = document.createElement("p");
    label.className = "kpi-label";
    label.textContent = kpi.label;

    const value = document.createElement("p");
    value.className = "kpi-value";
    value.textContent = kpi.value;

    card.appendChild(label);
    card.appendChild(value);
    elements.kpiContainer.appendChild(card);
  });
};

const renderTable = () => {
  elements.tableHead.innerHTML = "";
  elements.tableBody.innerHTML = "";

  if (!state.columns.length) {
    return;
  }

  const headerRow = document.createElement("tr");
  state.columns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column;
    headerRow.appendChild(th);
  });
  elements.tableHead.appendChild(headerRow);

  state.filteredData.forEach((row) => {
    const tr = document.createElement("tr");
    state.columns.forEach((column) => {
      const td = document.createElement("td");
      td.textContent = formatValue(row[column]);
      tr.appendChild(td);
    });
    elements.tableBody.appendChild(tr);
  });

  elements.tableCount.textContent = `${numberFormatter.format(state.filteredData.length)} rows`;
};

const updateFilteredData = () => {
  state.filteredData = applyFilters();
  renderKpis();
  renderTable();
};

const resetFilters = () => {
  FILTER_COLUMNS.forEach((column) => {
    state.filters[column].clear();
    const select = elements.filterSelects[column];
    if (select) {
      select.selectedIndex = -1;
    }
  });
  updateFilteredData();
};

elements.resetFilters.addEventListener("click", resetFilters);

loadExcel();
