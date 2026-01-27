const state = {
  data: [],
  columns: [],
  filters: {},
  search: "",
  sort: { column: null, direction: "asc" },
  currentPage: 1,
  rowsPerPage: 10,
  visibleColumns: new Set(),
};

const elements = {
  loadingIndicator: document.getElementById("loadingIndicator"),
  errorMessage: document.getElementById("errorMessage"),
  totalRecords: document.getElementById("totalRecords"),
  totalColumns: document.getElementById("totalColumns"),
  filtersContainer: document.getElementById("filtersContainer"),
  columnToggleContainer: document.getElementById("columnToggleContainer"),
  tableHead: document.querySelector("#dataTable thead"),
  tableBody: document.querySelector("#dataTable tbody"),
  globalSearch: document.getElementById("globalSearch"),
  rowsPerPage: document.getElementById("rowsPerPage"),
  resetFilters: document.getElementById("resetFilters"),
  paginationControls: document.getElementById("paginationControls"),
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

const getComparable = (value) => {
  if (value instanceof Date) {
    return value.getTime();
  }
  if (typeof value === "number") {
    return value;
  }
  const numeric = Number(value);
  if (!Number.isNaN(numeric) && value !== "") {
    return numeric;
  }
  return String(value).toLowerCase();
};

const showError = (message) => {
  elements.errorMessage.textContent = message;
};

const clearError = () => {
  elements.errorMessage.textContent = "";
};

const updateSummary = (filteredData) => {
  elements.totalRecords.textContent = filteredData.length.toLocaleString();
  elements.totalColumns.textContent = state.columns.length.toLocaleString();
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
    elements.loadingIndicator.style.display = "inline-block";

    const response = await fetch("data/Combined.xlsx");
    if (!response.ok) {
      throw new Error("Unable to load Combined.xlsx. Please confirm the file path.");
    }

    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true });

    // Sheet1-only logic: explicitly select Sheet1 and ignore all other sheets.
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

    // Dynamic column detection based on the first row of Sheet1.
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
    state.visibleColumns = new Set(headers);
    state.filters = headers.reduce((acc, header) => {
      acc[header] = new Set();
      return acc;
    }, {});

    buildFilters();
    buildColumnToggles();
    updateTable();
  } catch (error) {
    showError(error.message || "An unexpected error occurred while loading the Excel file.");
  } finally {
    elements.loadingIndicator.style.display = "none";
  }
};

const buildFilters = () => {
  elements.filtersContainer.innerHTML = "";

  state.columns.forEach((column) => {
    const filterCard = document.createElement("div");
    filterCard.className = "filter-card";

    const label = document.createElement("label");
    label.textContent = column;

    const select = document.createElement("select");
    select.multiple = true;
    select.dataset.column = column;

    const values = new Set(
      state.data.map((row) => formatValue(row[column]))
    );

    [...values].sort().forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value === "" ? "(Blank)" : value;
      select.appendChild(option);
    });

    select.addEventListener("change", (event) => {
      const selected = new Set([...event.target.selectedOptions].map((opt) => opt.value));
      state.filters[column] = selected;
      state.currentPage = 1;
      updateTable();
    });

    filterCard.appendChild(label);
    filterCard.appendChild(select);
    elements.filtersContainer.appendChild(filterCard);
  });
};

const buildColumnToggles = () => {
  elements.columnToggleContainer.innerHTML = "";

  state.columns.forEach((column) => {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = true;
    checkbox.dataset.column = column;

    checkbox.addEventListener("change", (event) => {
      if (event.target.checked) {
        state.visibleColumns.add(column);
      } else {
        state.visibleColumns.delete(column);
      }
      updateTable();
    });

    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(column));
    elements.columnToggleContainer.appendChild(label);
  });
};

const applyFilters = () => {
  const searchTerm = state.search.toLowerCase();

  return state.data.filter((row) => {
    const matchesFilters = state.columns.every((column) => {
      const selected = state.filters[column];
      if (!selected || selected.size === 0) {
        return true;
      }
      const value = formatValue(row[column]);
      return selected.has(value);
    });

    const matchesSearch = !searchTerm
      ? true
      : state.columns.some((column) =>
          formatValue(row[column]).toLowerCase().includes(searchTerm)
        );

    return matchesFilters && matchesSearch;
  });
};

const applySorting = (rows) => {
  const { column, direction } = state.sort;
  if (!column) {
    return rows;
  }

  const sorted = [...rows].sort((a, b) => {
    const valueA = getComparable(a[column]);
    const valueB = getComparable(b[column]);

    if (valueA < valueB) return direction === "asc" ? -1 : 1;
    if (valueA > valueB) return direction === "asc" ? 1 : -1;
    return 0;
  });

  return sorted;
};

const updateTable = () => {
  const filtered = applyFilters();
  const sorted = applySorting(filtered);
  updateSummary(filtered);

  const totalPages = Math.max(1, Math.ceil(sorted.length / state.rowsPerPage));
  if (state.currentPage > totalPages) {
    state.currentPage = totalPages;
  }

  const startIndex = (state.currentPage - 1) * state.rowsPerPage;
  const paged = sorted.slice(startIndex, startIndex + state.rowsPerPage);

  renderTableHead();
  renderTableBody(paged);
  renderPagination(totalPages);
};

const renderTableHead = () => {
  elements.tableHead.innerHTML = "";
  const row = document.createElement("tr");

  state.columns.forEach((column) => {
    if (!state.visibleColumns.has(column)) {
      return;
    }

    const th = document.createElement("th");
    th.textContent = column;

    if (state.sort.column === column) {
      th.textContent += state.sort.direction === "asc" ? " ▲" : " ▼";
    }

    th.addEventListener("click", () => {
      if (state.sort.column === column) {
        state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
      } else {
        state.sort.column = column;
        state.sort.direction = "asc";
      }
      updateTable();
    });

    row.appendChild(th);
  });

  elements.tableHead.appendChild(row);
};

const renderTableBody = (rows) => {
  elements.tableBody.innerHTML = "";

  if (!rows.length) {
    const emptyRow = document.createElement("tr");
    const emptyCell = document.createElement("td");
    emptyCell.colSpan = Math.max(state.visibleColumns.size, 1);
    emptyCell.textContent = "No matching records.";
    emptyRow.appendChild(emptyCell);
    elements.tableBody.appendChild(emptyRow);
    return;
  }

  rows.forEach((rowData) => {
    const row = document.createElement("tr");
    state.columns.forEach((column) => {
      if (!state.visibleColumns.has(column)) {
        return;
      }
      const cell = document.createElement("td");
      cell.textContent = formatValue(rowData[column]);
      row.appendChild(cell);
    });
    elements.tableBody.appendChild(row);
  });
};

const renderPagination = (totalPages) => {
  elements.paginationControls.innerHTML = "";

  const info = document.createElement("span");
  info.textContent = `Page ${state.currentPage} of ${totalPages}`;

  const prev = document.createElement("button");
  prev.textContent = "Previous";
  prev.disabled = state.currentPage === 1;
  prev.addEventListener("click", () => {
    state.currentPage -= 1;
    updateTable();
  });

  const next = document.createElement("button");
  next.textContent = "Next";
  next.disabled = state.currentPage === totalPages;
  next.addEventListener("click", () => {
    state.currentPage += 1;
    updateTable();
  });

  elements.paginationControls.appendChild(prev);
  elements.paginationControls.appendChild(info);
  elements.paginationControls.appendChild(next);
};

const resetAllFilters = () => {
  state.search = "";
  elements.globalSearch.value = "";
  state.currentPage = 1;

  Object.keys(state.filters).forEach((column) => {
    state.filters[column].clear();
  });

  document.querySelectorAll("#filtersContainer select").forEach((select) => {
    select.selectedIndex = -1;
  });

  updateTable();
};

const attachEventListeners = () => {
  elements.globalSearch.addEventListener("input", (event) => {
    state.search = event.target.value;
    state.currentPage = 1;
    updateTable();
  });

  elements.rowsPerPage.addEventListener("change", (event) => {
    state.rowsPerPage = Number(event.target.value);
    state.currentPage = 1;
    updateTable();
  });

  elements.resetFilters.addEventListener("click", () => {
    resetAllFilters();
  });
};

attachEventListeners();
loadExcel();
