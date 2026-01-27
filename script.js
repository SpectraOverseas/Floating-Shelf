"use strict";

// Global state for pivot tables and active filters.
const state = {
  tables: [],
  headers: [],
  filters: {}
};

const controlsContainer = document.getElementById("filterControls");
const tablesContainer = document.getElementById("tablesContainer");
const resetButton = document.getElementById("resetFilters");

/**
 * Fetch the Excel file and kick off the dashboard rendering.
 */
async function initDashboard() {
  try {
    const response = await fetch("Combined.xlsx");
    const arrayBuffer = await response.arrayBuffer();

    // Parse the workbook with SheetJS and read ONLY "Sheet 2".
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets["Sheet 2"];

    if (!sheet) {
      renderStatus("Sheet 2 not found in Combined.xlsx.");
      return;
    }

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false
    });

    // Detect pivot tables dynamically from Sheet 2.
    const { headers, tables } = extractPivotTables(rows);

    if (!headers.length || !tables.length) {
      renderStatus("No pivot tables found in Sheet 2.");
      return;
    }

    state.headers = headers;
    state.tables = tables;

    buildFilters(headers, tables);
    renderTables(tables, headers);
  } catch (error) {
    console.error(error);
    renderStatus("Unable to load Combined.xlsx. Please ensure the file is present.");
  }
}

/**
 * Parse Sheet 2 into pivot tables.
 * - Uses the first header row as the template.
 * - Treats identical header rows or blank rows as table separators.
 */
function extractPivotTables(rows) {
  const normalizedRows = rows.map((row) => row.map((cell) => String(cell).trim()));
  const firstHeaderIndex = normalizedRows.findIndex((row) => row.some((cell) => cell));

  if (firstHeaderIndex === -1) {
    return { headers: [], tables: [] };
  }

  const headers = rows[firstHeaderIndex];
  const headerTemplate = normalizedRows[firstHeaderIndex];
  const tables = [];
  let currentTable = [];

  const pushTable = () => {
    if (currentTable.length) {
      tables.push(currentTable);
      currentTable = [];
    }
  };

  for (let i = firstHeaderIndex + 1; i < normalizedRows.length; i += 1) {
    const row = normalizedRows[i];
    const rawRow = rows[i];

    if (isRowEmpty(row)) {
      pushTable();
      continue;
    }

    if (isHeaderRow(row, headerTemplate)) {
      pushTable();
      continue;
    }

    currentTable.push(rawRow);
  }

  pushTable();

  return { headers, tables };
}

/**
 * Determine if a row is empty (all values blank).
 */
function isRowEmpty(row) {
  return row.every((cell) => !String(cell).trim());
}

/**
 * Determine if a row matches the header template.
 */
function isHeaderRow(row, headers) {
  const trimmedHeaders = headers.map((cell) => String(cell).trim());
  const trimmedRow = row.map((cell) => String(cell).trim());
  const headerLength = trimmedHeaders.length;

  for (let i = 0; i < headerLength; i += 1) {
    if (!trimmedHeaders[i] && !trimmedRow[i]) {
      continue;
    }
    if (trimmedHeaders[i] !== trimmedRow[i]) {
      return false;
    }
  }

  return true;
}

/**
 * Build filters dynamically from header columns and table data.
 * - Dropdowns for categorical columns (limited unique values).
 * - Text inputs for free-text columns (large unique values).
 */
function buildFilters(headers, tables) {
  controlsContainer.innerHTML = "";
  state.filters = {};

  const columnValues = headers.map(() => new Set());

  tables.forEach((table) => {
    table.forEach((row) => {
      headers.forEach((_, columnIndex) => {
        const value = row[columnIndex] ?? "";
        columnValues[columnIndex].add(String(value).trim());
      });
    });
  });

  headers.forEach((header, columnIndex) => {
    const values = Array.from(columnValues[columnIndex]).filter(Boolean).sort();
    const useDropdown = values.length <= 20;

    const wrapper = document.createElement("div");
    wrapper.className = "filter-control";

    const label = document.createElement("label");
    label.textContent = header || `Column ${columnIndex + 1}`;

    if (useDropdown) {
      const select = document.createElement("select");
      select.innerHTML = `<option value="">All</option>${values
        .map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
        .join("")}`;
      select.addEventListener("change", () => {
        state.filters[columnIndex] = { type: "select", value: select.value };
        applyFilters();
      });
      wrapper.append(label, select);
    } else {
      const input = document.createElement("input");
      input.type = "search";
      input.placeholder = "Type to search";
      input.addEventListener("input", () => {
        state.filters[columnIndex] = { type: "text", value: input.value };
        applyFilters();
      });
      wrapper.append(label, input);
    }

    controlsContainer.appendChild(wrapper);
  });
}

/**
 * Render all pivot tables with headings.
 */
function renderTables(tables, headers) {
  tablesContainer.innerHTML = "";

  tables.forEach((table, index) => {
    const card = document.createElement("div");
    card.className = "table-card";

    const heading = document.createElement("h3");
    heading.textContent = `Pivot Table ${index + 1}`;

    const wrapper = document.createElement("div");
    wrapper.className = "table-wrapper";

    const tableEl = document.createElement("table");
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    headers.forEach((header) => {
      const th = document.createElement("th");
      th.textContent = header || "";
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    tableEl.appendChild(thead);

    const tbody = document.createElement("tbody");
    table.forEach((row) => {
      const tr = document.createElement("tr");
      headers.forEach((_, columnIndex) => {
        const td = document.createElement("td");
        td.textContent = row[columnIndex] ?? "";
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    tableEl.appendChild(tbody);
    wrapper.appendChild(tableEl);
    card.append(heading, wrapper);
    tablesContainer.appendChild(card);
  });
}

/**
 * Apply all active filters to every table simultaneously.
 */
function applyFilters() {
  const cards = tablesContainer.querySelectorAll(".table-card");

  cards.forEach((card, tableIndex) => {
    const tbodyRows = card.querySelectorAll("tbody tr");
    const tableData = state.tables[tableIndex];

    tbodyRows.forEach((rowEl, rowIndex) => {
      const rowData = tableData[rowIndex] || [];
      const isVisible = Object.entries(state.filters).every(([columnIndex, filter]) => {
        const value = String(rowData[columnIndex] ?? "");

        if (!filter || !filter.value) {
          return true;
        }

        if (filter.type === "select") {
          return value === filter.value;
        }

        return value.toLowerCase().includes(filter.value.toLowerCase());
      });

      rowEl.style.display = isVisible ? "" : "none";
    });
  });
}

/**
 * Reset all filters to their default state.
 */
function resetFilters() {
  controlsContainer.querySelectorAll("select").forEach((select) => {
    select.value = "";
  });
  controlsContainer.querySelectorAll("input").forEach((input) => {
    input.value = "";
  });

  state.filters = {};
  applyFilters();
}

/**
 * Render a message when data is missing or fails to load.
 */
function renderStatus(message) {
  tablesContainer.innerHTML = `<p class="status-message">${message}</p>`;
}

/**
 * Escape HTML to keep dynamically generated options safe.
 */
function escapeHtml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

resetButton.addEventListener("click", resetFilters);

initDashboard();
