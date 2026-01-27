const ALLOWED_COLUMN_INDICES = {
  A: 0,
  B: 1,
  C: 2,
  I: 8,
  M: 12,
  O: 14,
  R: 17,
  W: 22,
  AJ: 35,
};

const state = {
  rawData: [],
  columns: [],
  filters: {},
  search: "",
  sort: { key: null, direction: "asc" },
  page: 1,
  pageSize: 12,
};

const elements = {
  dataStatus: document.getElementById("dataStatus"),
  filters: document.getElementById("filters"),
  resetFilters: document.getElementById("resetFilters"),
  globalSearch: document.getElementById("globalSearch"),
  summaryHead: document.getElementById("summaryHead"),
  summaryBody: document.getElementById("summaryBody"),
  prevPage: document.getElementById("prevPage"),
  nextPage: document.getElementById("nextPage"),
  pageStatus: document.getElementById("pageStatus"),
  kpiTitle1: document.getElementById("kpiTitle1"),
  kpiValue1: document.getElementById("kpiValue1"),
  kpiMeta1: document.getElementById("kpiMeta1"),
  kpiTitle2: document.getElementById("kpiTitle2"),
  kpiValue2: document.getElementById("kpiValue2"),
  kpiMeta2: document.getElementById("kpiMeta2"),
  kpiTitle3: document.getElementById("kpiTitle3"),
  kpiValue3: document.getElementById("kpiValue3"),
  kpiMeta3: document.getElementById("kpiMeta3"),
};

const formatNumber = (value, type = "number") => {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return "—";
  }
  if (type === "currency") {
    return new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
      maximumFractionDigits: 2,
    }).format(value);
  }
  if (type === "percent") {
    return new Intl.NumberFormat("en-US", {
      style: "percent",
      maximumFractionDigits: 2,
    }).format(value);
  }
  return new Intl.NumberFormat("en-US", {
    maximumFractionDigits: 2,
  }).format(value);
};

const toNumber = (value) => {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const numericValue =
    typeof value === "number" ? value : Number(String(value).replace(/,/g, ""));
  return Number.isNaN(numericValue) ? null : numericValue;
};

const loadData = async () => {
  elements.dataStatus.textContent = "Loading data…";
  const response = await fetch("Combined.xlsx");
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets["Sheet1"];
  if (!sheet) {
    elements.dataStatus.textContent = "Sheet1 not found in Combined.xlsx.";
    return;
  }
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  const headers = rows[0] || [];
  const allowedEntries = Object.entries(ALLOWED_COLUMN_INDICES).map(
    ([code, index]) => ({
      code,
      index,
      name: headers[index] || `Column ${code}`,
    })
  );

  const data = rows.slice(1).map((row) => {
    const record = {};
    allowedEntries.forEach(({ code, index, name }) => {
      record[code] = {
        label: name,
        value: row[index],
      };
    });
    return record;
  });

  state.rawData = data;
  state.columns = allowedEntries;
  state.filters = {};
  state.search = "";
  state.page = 1;
  elements.dataStatus.textContent = `Loaded ${data.length.toLocaleString()} rows`;

  buildFilters();
  render();
};

const buildFilters = () => {
  elements.filters.innerHTML = "";
  const filterColumns = getCategoricalColumns({ maxUnique: 50 });

  filterColumns.forEach((col) => {
    const values = Array.from(
      new Set(
        state.rawData
          .map((row) => row[col.code].value)
          .filter((value) => value !== null && value !== undefined && value !== "")
          .map((value) => String(value))
      )
    ).sort();
    const wrapper = document.createElement("div");
    wrapper.className = "filter-card";
    wrapper.innerHTML = `
      <label for="filter-${col.code}">${col.name}</label>
      <select id="filter-${col.code}" multiple></select>
    `;
    const select = wrapper.querySelector("select");
    values.forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });
    select.addEventListener("change", () => {
      const selected = Array.from(select.selectedOptions).map(
        (option) => option.value
      );
      state.filters[col.code] = selected;
      state.page = 1;
      render();
    });
    elements.filters.appendChild(wrapper);
  });
};

const getFilteredData = () => {
  return state.rawData.filter((row) => {
    return Object.entries(state.filters).every(([code, selected]) => {
      if (!selected || selected.length === 0) {
        return true;
      }
      const value = row[code]?.value;
      if (value === null || value === undefined) {
        return false;
      }
      return selected.includes(String(value));
    });
  });
};

const getCategoricalColumns = ({ maxUnique = 50 } = {}) => {
  return state.columns.filter((col) => {
    const values = state.rawData
      .map((row) => row[col.code].value)
      .filter((value) => value !== null && value !== undefined && value !== "");
    if (values.length === 0) {
      return false;
    }
    const numericShare =
      values.filter((value) => toNumber(value) !== null).length / values.length;
    const uniqueCount = new Set(values.map((value) => String(value))).size;
    return numericShare < 0.5 && uniqueCount <= maxUnique && uniqueCount > 1;
  });
};

const getNumericColumns = () => {
  return state.columns.filter((col) => {
    const values = state.rawData
      .map((row) => row[col.code].value)
      .filter((value) => value !== null && value !== undefined && value !== "");
    if (values.length === 0) {
      return false;
    }
    const numericShare =
      values.filter((value) => toNumber(value) !== null).length / values.length;
    return numericShare >= 0.6;
  });
};

const updateKpis = (filteredData) => {
  const numericColumns = getNumericColumns();
  const first = numericColumns[0];
  const second = numericColumns[1];

  const sumFirst = first
    ? filteredData.reduce((acc, row) => acc + (toNumber(row[first.code].value) || 0), 0)
    : null;
  const avgFirst =
    first && filteredData.length
      ? sumFirst / filteredData.length
      : null;
  const sumSecond = second
    ? filteredData.reduce(
        (acc, row) => acc + (toNumber(row[second.code].value) || 0),
        0
      )
    : null;

  elements.kpiTitle1.textContent = first
    ? `${first.name} Sum`
    : "No numeric columns";
  elements.kpiValue1.textContent = first
    ? formatNumber(sumFirst, "currency")
    : "—";
  elements.kpiMeta1.textContent = first ? "Total after filters" : "";

  elements.kpiTitle2.textContent = first
    ? `${first.name} Average`
    : "No numeric columns";
  elements.kpiValue2.textContent = first
    ? formatNumber(avgFirst, "number")
    : "—";
  elements.kpiMeta2.textContent = first ? "Average per row" : "";

  const ratio =
    first && second && sumSecond !== 0 ? sumFirst / sumSecond : null;
  elements.kpiTitle3.textContent = second
    ? `${first.name} ÷ ${second.name}`
    : "Ratio";
  elements.kpiValue3.textContent =
    ratio !== null ? formatNumber(ratio, ratio <= 1 ? "percent" : "number") : "—";
  elements.kpiMeta3.textContent = second ? "Based on filtered totals" : "";
};

const buildPivotRows = (filteredData) => {
  const numericColumns = getNumericColumns();
  const categoricalColumns = getCategoricalColumns();
  const groupColumn = categoricalColumns[0] || state.columns[0];

  const groups = new Map();
  filteredData.forEach((row) => {
    const key = String(row[groupColumn.code]?.value ?? "Unspecified");
    if (!groups.has(key)) {
      groups.set(key, { label: key, count: 0, sums: {} });
    }
    const entry = groups.get(key);
    entry.count += 1;
    numericColumns.forEach((col) => {
      const value = toNumber(row[col.code].value) || 0;
      entry.sums[col.code] = (entry.sums[col.code] || 0) + value;
    });
  });

  return {
    groupColumn,
    numericColumns,
    rows: Array.from(groups.values()),
  };
};

const renderTable = ({ groupColumn, numericColumns, rows }) => {
  const searchTerm = state.search.toLowerCase();
  let displayRows = rows.filter((row) =>
    row.label.toLowerCase().includes(searchTerm)
  );

  if (state.sort.key) {
    const { key, direction } = state.sort;
    const getSortValue = (row) => {
      if (key === "label") {
        return row.label;
      }
      if (key === "count") {
        return row.count;
      }
      if (key.startsWith("sum:")) {
        const code = key.split(":")[1];
        return row.sums[code] || 0;
      }
      if (key.startsWith("avg:")) {
        const code = key.split(":")[1];
        return row.count ? (row.sums[code] || 0) / row.count : 0;
      }
      return 0;
    };
    displayRows = displayRows.sort((a, b) => {
      const valueA = getSortValue(a);
      const valueB = getSortValue(b);
      if (typeof valueA === "string") {
        return direction === "asc"
          ? valueA.localeCompare(valueB)
          : valueB.localeCompare(valueA);
      }
      return direction === "asc" ? valueA - valueB : valueB - valueA;
    });
  }

  const totalPages = Math.max(1, Math.ceil(displayRows.length / state.pageSize));
  state.page = Math.min(state.page, totalPages);
  const start = (state.page - 1) * state.pageSize;
  const pagedRows = displayRows.slice(start, start + state.pageSize);

  elements.summaryHead.innerHTML = "";
  const headerRow = document.createElement("tr");
  const headerCells = [
    { key: "label", label: groupColumn.name },
    { key: "count", label: "Rows" },
    ...numericColumns.flatMap((col) => [
      { key: `sum:${col.code}`, label: `${col.name} Sum` },
      { key: `avg:${col.code}`, label: `${col.name} Avg` },
    ]),
  ];

  headerCells.forEach((cell) => {
    const th = document.createElement("th");
    th.textContent = cell.label;
    th.addEventListener("click", () => {
      if (state.sort.key === cell.key) {
        state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
      } else {
        state.sort.key = cell.key;
        state.sort.direction = "asc";
      }
      render();
    });
    headerRow.appendChild(th);
  });
  elements.summaryHead.appendChild(headerRow);

  elements.summaryBody.innerHTML = "";
  pagedRows.forEach((row) => {
    const tr = document.createElement("tr");
    const labelCell = document.createElement("td");
    labelCell.textContent = row.label;
    tr.appendChild(labelCell);

    const countCell = document.createElement("td");
    countCell.textContent = row.count.toLocaleString();
    tr.appendChild(countCell);

    numericColumns.forEach((col) => {
      const sumCell = document.createElement("td");
      sumCell.textContent = formatNumber(row.sums[col.code] || 0, "currency");
      tr.appendChild(sumCell);
      const avgCell = document.createElement("td");
      const avgValue = row.count ? (row.sums[col.code] || 0) / row.count : null;
      avgCell.textContent = formatNumber(avgValue, "currency");
      tr.appendChild(avgCell);
    });
    elements.summaryBody.appendChild(tr);
  });

  elements.pageStatus.textContent = `Page ${state.page} of ${totalPages}`;
  elements.prevPage.disabled = state.page <= 1;
  elements.nextPage.disabled = state.page >= totalPages;
};

const render = () => {
  const filteredData = getFilteredData();
  updateKpis(filteredData);
  const pivot = buildPivotRows(filteredData);
  renderTable(pivot);
};

elements.resetFilters.addEventListener("click", () => {
  state.filters = {};
  document.querySelectorAll("#filters select").forEach((select) => {
    select.selectedIndex = -1;
  });
  state.page = 1;
  render();
});

elements.globalSearch.addEventListener("input", (event) => {
  state.search = event.target.value;
  state.page = 1;
  render();
});

elements.prevPage.addEventListener("click", () => {
  state.page = Math.max(1, state.page - 1);
  render();
});

elements.nextPage.addEventListener("click", () => {
  state.page += 1;
  render();
});

document.addEventListener("DOMContentLoaded", loadData);
