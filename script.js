const state = {
  rawData: [],
  columns: [],
  filterState: new Map(),
  visibleColumns: new Set(),
  searchQuery: "",
  sortState: { column: null, direction: "asc" },
  currentPage: 1,
  rowsPerPage: 10
};

const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });

const elements = {
  filters: document.getElementById("filters"),
  columnToggles: document.getElementById("columnToggles"),
  tableHead: document.getElementById("tableHead"),
  tableBody: document.getElementById("tableBody"),
  pagination: document.getElementById("pagination"),
  totalRecords: document.getElementById("totalRecords"),
  totalColumns: document.getElementById("totalColumns"),
  globalSearch: document.getElementById("globalSearch"),
  rowsPerPage: document.getElementById("rowsPerPage"),
  resetFilters: document.getElementById("resetFilters"),
  loading: document.getElementById("loadingIndicator"),
  error: document.getElementById("errorMessage")
};

function formatValue(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    return value.toLocaleDateString();
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return value.toString();
  }
  return String(value).trim();
}

function isRowEmpty(row) {
  return row.every((cell) => formatValue(cell) === "");
}

function showError(message) {
  elements.error.textContent = message;
  elements.error.hidden = false;
}

function setLoading(isLoading) {
  elements.loading.hidden = !isLoading;
}

function buildDataset(rows) {
  if (!rows.length) {
    throw new Error("Sheet1 is empty or missing headers.");
  }

  const headerRow = rows.shift();
  const maxLength = Math.max(headerRow.length, ...rows.map((row) => row.length));
  const columns = Array.from({ length: maxLength }, (_, index) => {
    const header = headerRow[index];
    const label = formatValue(header);
    return label !== "" ? label : `Column ${index + 1}`;
  });

  const data = rows
    .filter((row) => !isRowEmpty(row))
    .map((row) => {
      const record = {};
      columns.forEach((column, index) => {
        record[column] = formatValue(row[index]);
      });
      return record;
    });

  return { columns, data };
}

function buildFilters() {
  elements.filters.innerHTML = "";
  state.filterState.clear();

  state.columns.forEach((column) => {
    const group = document.createElement("div");
    group.className = "filter-group";

    const label = document.createElement("label");
    label.textContent = column;

    const select = document.createElement("select");
    select.multiple = true;
    select.dataset.column = column;

    const values = Array.from(
      new Set(state.rawData.map((row) => row[column]).filter((value) => value !== ""))
    ).sort((a, b) => collator.compare(a, b));

    values.forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });

    select.addEventListener("change", () => {
      const selectedValues = new Set(Array.from(select.selectedOptions).map((opt) => opt.value));
      state.filterState.set(column, selectedValues);
      state.currentPage = 1;
      render();
    });

    group.appendChild(label);
    group.appendChild(select);
    elements.filters.appendChild(group);
    state.filterState.set(column, new Set());
  });
}

function buildColumnToggles() {
  elements.columnToggles.innerHTML = "";
  state.visibleColumns = new Set(state.columns);

  state.columns.forEach((column) => {
    const wrapper = document.createElement("label");
    wrapper.className = "column-toggle";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = true;
    checkbox.dataset.column = column;

    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        state.visibleColumns.add(column);
      } else {
        state.visibleColumns.delete(column);
        if (state.visibleColumns.size === 0) {
          checkbox.checked = true;
          state.visibleColumns.add(column);
          return;
        }
      }
      renderTable();
    });

    const text = document.createElement("span");
    text.textContent = column;

    wrapper.appendChild(checkbox);
    wrapper.appendChild(text);
    elements.columnToggles.appendChild(wrapper);
  });
}

function applyFilters() {
  let data = [...state.rawData];

  // Filter updates are column-agnostic: apply any selected values per column.
  state.filterState.forEach((selectedValues, column) => {
    if (selectedValues.size === 0) return;
    data = data.filter((row) => selectedValues.has(row[column]));
  });

  if (state.searchQuery) {
    const query = state.searchQuery.toLowerCase();
    data = data.filter((row) =>
      state.columns.some((column) => String(row[column]).toLowerCase().includes(query))
    );
  }

  return data;
}

function sortData(data) {
  const { column, direction } = state.sortState;
  if (!column) return data;

  const sorted = [...data].sort((a, b) => collator.compare(a[column], b[column]));
  return direction === "asc" ? sorted : sorted.reverse();
}

function paginateData(data) {
  const totalPages = Math.max(1, Math.ceil(data.length / state.rowsPerPage));
  state.currentPage = Math.min(state.currentPage, totalPages);
  const start = (state.currentPage - 1) * state.rowsPerPage;
  return {
    totalPages,
    pageData: data.slice(start, start + state.rowsPerPage)
  };
}

function updateSummary(filteredCount) {
  elements.totalRecords.textContent = filteredCount.toLocaleString();
  elements.totalColumns.textContent = state.columns.length.toLocaleString();
}

function renderTable() {
  const visibleColumns = state.columns.filter((column) => state.visibleColumns.has(column));
  elements.tableHead.innerHTML = "";

  const headRow = document.createElement("tr");
  visibleColumns.forEach((column) => {
    const th = document.createElement("th");
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.column = column;
    button.textContent = column;

    const indicator = document.createElement("span");
    if (state.sortState.column === column) {
      indicator.textContent = state.sortState.direction === "asc" ? "▲" : "▼";
    } else {
      indicator.textContent = "↕";
    }

    button.appendChild(indicator);
    button.addEventListener("click", () => {
      if (state.sortState.column === column) {
        state.sortState.direction = state.sortState.direction === "asc" ? "desc" : "asc";
      } else {
        state.sortState.column = column;
        state.sortState.direction = "asc";
      }
      render();
    });

    th.appendChild(button);
    headRow.appendChild(th);
  });
  elements.tableHead.appendChild(headRow);

  const filteredData = applyFilters();
  const sortedData = sortData(filteredData);
  const { totalPages, pageData } = paginateData(sortedData);

  elements.tableBody.innerHTML = "";
  pageData.forEach((row) => {
    const tr = document.createElement("tr");
    visibleColumns.forEach((column) => {
      const td = document.createElement("td");
      td.textContent = row[column];
      tr.appendChild(td);
    });
    elements.tableBody.appendChild(tr);
  });

  renderPagination(totalPages, filteredData.length);
}

function renderPagination(totalPages, totalRecords) {
  elements.pagination.innerHTML = "";

  const info = document.createElement("span");
  info.textContent = `Page ${state.currentPage} of ${totalPages} • ${totalRecords.toLocaleString()} records`;
  elements.pagination.appendChild(info);

  const prevButton = document.createElement("button");
  prevButton.type = "button";
  prevButton.textContent = "Previous";
  prevButton.disabled = state.currentPage === 1;
  prevButton.addEventListener("click", () => {
    state.currentPage -= 1;
    renderTable();
  });

  const nextButton = document.createElement("button");
  nextButton.type = "button";
  nextButton.textContent = "Next";
  nextButton.disabled = state.currentPage === totalPages;
  nextButton.addEventListener("click", () => {
    state.currentPage += 1;
    renderTable();
  });

  elements.pagination.appendChild(prevButton);
  elements.pagination.appendChild(nextButton);
}

function render() {
  const filteredData = applyFilters();
  updateSummary(filteredData.length);
  renderTable();
}

function resetFilters() {
  state.filterState.forEach((_, column) => {
    state.filterState.set(column, new Set());
  });

  elements.filters.querySelectorAll("select").forEach((select) => {
    Array.from(select.options).forEach((option) => {
      option.selected = false;
    });
  });

  state.searchQuery = "";
  elements.globalSearch.value = "";
  state.sortState = { column: null, direction: "asc" };
  state.currentPage = 1;
  render();
}

async function loadWorkbook() {
  setLoading(true);
  try {
    const response = await fetch("data/Combined.xlsx");
    if (!response.ok) {
      throw new Error("Unable to fetch Combined.xlsx.");
    }

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });

    // Sheet1-only logic: we explicitly ignore every other sheet, including Sheet2.
    const sheet = workbook.Sheets["Sheet1"];
    if (!sheet) {
      throw new Error("Sheet1 is missing. Please ensure Combined.xlsx includes Sheet1.");
    }

    // Dynamic column detection: read Sheet1 into a 2D array to build headers safely.
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
    const dataset = buildDataset(rows);

    state.rawData = dataset.data;
    state.columns = dataset.columns;

    buildFilters();
    buildColumnToggles();
    updateSummary(state.rawData.length);
    renderTable();
  } catch (error) {
    console.error(error);
    showError(error.message || "Unable to load dataset.");
  } finally {
    setLoading(false);
  }
}

// Filter and table updates: keep the table, pagination, and summary in sync.
elements.globalSearch.addEventListener("input", (event) => {
  state.searchQuery = event.target.value.trim();
  state.currentPage = 1;
  render();
});

elements.rowsPerPage.addEventListener("change", (event) => {
  state.rowsPerPage = Number(event.target.value);
  state.currentPage = 1;
  render();
});

elements.resetFilters.addEventListener("click", resetFilters);

loadWorkbook();
