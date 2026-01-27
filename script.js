const STATUS = document.getElementById("status");
const FILTER_FIELDS = document.getElementById("filterFields");
const TABLES_CONTAINER = document.getElementById("tables");
const RESET_BUTTON = document.getElementById("resetFilters");

let pivotTables = [];
let filterConfig = [];
let filterState = {};

// Treat empty cells consistently when parsing Sheet 2.
const isEmptyValue = (value) =>
  value === null || value === undefined || String(value).trim() === "";

const isRowEmpty = (row) => row.every(isEmptyValue);

const trimRow = (row) => {
  let lastIndex = row.length - 1;
  while (lastIndex >= 0 && isEmptyValue(row[lastIndex])) {
    lastIndex -= 1;
  }
  return row.slice(0, lastIndex + 1).map((cell) => (cell ?? ""));
};

const updateStatus = (message, isError = false) => {
  STATUS.textContent = message;
  STATUS.style.color = isError ? "#b91c1c" : "";
};

// --- Sheet 2 parsing & pivot table separation logic ---
// We read Sheet 2 into a 2D array and split pivot tables using empty rows as
// natural separators. The first non-empty row in each block becomes the header.
const parsePivotTables = (rows) => {
  const tables = [];
  let currentTable = null;

  rows.forEach((row) => {
    if (isRowEmpty(row)) {
      if (currentTable) {
        tables.push(currentTable);
        currentTable = null;
      }
      return;
    }

    const trimmedRow = trimRow(row);

    if (!currentTable) {
      currentTable = {
        headers: trimmedRow.map((header) => String(header).trim()),
        rows: [],
      };
      return;
    }

    const record = {};
    currentTable.headers.forEach((header, index) => {
      record[header] = trimmedRow[index] ?? "";
    });
    currentTable.rows.push(record);
  });

  if (currentTable) {
    tables.push(currentTable);
  }

  return tables;
};

// --- Filter generation and filtering logic ---
// Filters are derived from the shared column headers and data values.
// Columns with smaller unique counts become dropdowns; others are free-text.
const buildFilters = (tables) => {
  const allRows = tables.flatMap((table) => table.rows);
  const headers = tables[0]?.headers ?? [];
  const config = headers.map((header) => {
    const values = new Set();
    allRows.forEach((row) => {
      const value = row[header];
      if (!isEmptyValue(value)) {
        values.add(String(value));
      }
    });

    const uniqueValues = Array.from(values).sort();
    const useDropdown = uniqueValues.length > 0 && uniqueValues.length <= 30;

    return {
      header,
      type: useDropdown ? "dropdown" : "text",
      options: uniqueValues,
    };
  });

  return config;
};

const renderFilters = () => {
  FILTER_FIELDS.innerHTML = "";

  filterConfig.forEach(({ header, type, options }) => {
    const wrapper = document.createElement("div");
    wrapper.className = "filter";

    const label = document.createElement("label");
    label.textContent = header;
    label.setAttribute("for", `filter-${header}`);

    let input;
    if (type === "dropdown") {
      input = document.createElement("select");
      input.innerHTML = `<option value="">All</option>`;
      options.forEach((option) => {
        const optionEl = document.createElement("option");
        optionEl.value = option;
        optionEl.textContent = option;
        input.appendChild(optionEl);
      });
    } else {
      input = document.createElement("input");
      input.type = "search";
      input.placeholder = `Search ${header}`;
    }

    input.id = `filter-${header}`;
    input.addEventListener("input", (event) => {
      filterState[header] = event.target.value;
      renderTables();
    });

    wrapper.appendChild(label);
    wrapper.appendChild(input);
    FILTER_FIELDS.appendChild(wrapper);
  });
};

const rowMatchesFilters = (row) => {
  return filterConfig.every(({ header, type }) => {
    const value = filterState[header];
    if (!value) {
      return true;
    }

    const cellValue = String(row[header] ?? "");
    if (type === "dropdown") {
      return cellValue === value;
    }

    return cellValue.toLowerCase().includes(String(value).toLowerCase());
  });
};

const renderTables = () => {
  TABLES_CONTAINER.innerHTML = "";

  pivotTables.forEach((table, index) => {
    const card = document.createElement("article");
    card.className = "table-card";

    const title = document.createElement("h3");
    title.textContent = `Pivot Table ${index + 1}`;

    const wrapper = document.createElement("div");
    wrapper.className = "table-wrapper";

    const tableEl = document.createElement("table");
    tableEl.className = "table";

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    table.headers.forEach((header) => {
      const th = document.createElement("th");
      th.textContent = header;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    const tbody = document.createElement("tbody");
    const filteredRows = table.rows.filter(rowMatchesFilters);

    if (filteredRows.length === 0) {
      const emptyRow = document.createElement("tr");
      const emptyCell = document.createElement("td");
      emptyCell.colSpan = table.headers.length;
      emptyCell.className = "empty-state";
      emptyCell.textContent = "No rows match the selected filters.";
      emptyRow.appendChild(emptyCell);
      tbody.appendChild(emptyRow);
    } else {
      filteredRows.forEach((row) => {
        const tr = document.createElement("tr");
        table.headers.forEach((header) => {
          const td = document.createElement("td");
          td.textContent = row[header] ?? "";
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
    }

    tableEl.appendChild(thead);
    tableEl.appendChild(tbody);
    wrapper.appendChild(tableEl);

    card.appendChild(title);
    card.appendChild(wrapper);
    TABLES_CONTAINER.appendChild(card);
  });
};

const resetFilters = () => {
  filterState = {};
  filterConfig.forEach(({ header }) => {
    const input = document.getElementById(`filter-${header}`);
    if (input) {
      input.value = "";
    }
  });
  renderTables();
};

const loadWorkbook = async () => {
  try {
    updateStatus("Loading Combined.xlsx…");
    const response = await fetch("Combined.xlsx");
    if (!response.ok) {
      throw new Error("Unable to load Combined.xlsx");
    }

    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });

    const sheet = workbook.Sheets["Sheet 2"];
    if (!sheet) {
      throw new Error("Sheet 2 not found in Combined.xlsx");
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: true });
    pivotTables = parsePivotTables(rows).filter((table) => table.rows.length > 0);

    if (pivotTables.length === 0) {
      updateStatus("No pivot tables detected in Sheet 2.", true);
      return;
    }

    filterConfig = buildFilters(pivotTables);
    filterState = {};
    renderFilters();
    renderTables();
    updateStatus(`Loaded ${pivotTables.length} pivot table(s).`);
  } catch (error) {
    updateStatus(error.message, true);
    console.error(error);
  }
};

RESET_BUTTON.addEventListener("click", resetFilters);

loadWorkbook();
