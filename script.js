const FILTER_COLUMNS = [
  "ASIN",
  "L X W X H",
  "Colour",
  "Material",
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
};

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

  FILTER_COLUMNS.forEach((column) => {
    const filterCard = document.createElement("div");
    filterCard.className = "filter-card";

    const label = document.createElement("label");
    label.textContent = column;
    label.setAttribute("for", `filter-${column}`);

    const select = document.createElement("select");
    select.id = `filter-${column}`;
    select.multiple = true;
    select.dataset.column = column;

    const values = new Set(
      state.data.map((row) => formatValue(row[column]))
    );

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

    const clearButton = document.createElement("button");
    clearButton.type = "button";
    clearButton.className = "clear-button";
    clearButton.setAttribute("aria-label", `Clear ${column} filter`);
    clearButton.textContent = "✕";
    clearButton.addEventListener("click", () => {
      select.selectedIndex = -1;
      state.filters[column].clear();
      updateFilteredData();
    });

    const fieldRow = document.createElement("div");
    fieldRow.className = "filter-field";
    fieldRow.appendChild(select);
    fieldRow.appendChild(clearButton);

    filterCard.appendChild(label);
    filterCard.appendChild(fieldRow);
    elements.filtersContainer.appendChild(filterCard);
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

const updateFilteredData = () => {
  state.filteredData = applyFilters();
};

loadExcel();
