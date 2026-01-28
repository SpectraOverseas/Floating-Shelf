const DATA_URL = "data/Combined.xlsx";
const FILTER_COLUMNS = [
  "ASIN",
  "Colour",
  "Material",
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

const filtersContainer = document.getElementById("filters");
const resetButton = document.getElementById("resetFilters");
const kpiElements = new Map(
  Array.from(document.querySelectorAll("[data-kpi]")).map((el) => [
    el.dataset.kpi,
    el,
  ])
);

let rawData = [];
const activeFilters = {};

const numberFormatter = new Intl.NumberFormat("en-US", {
  maximumFractionDigits: 2,
});

const cleanValue = (value) => {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
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
    element.textContent = numberFormatter.format(value);
  });
};

const updateDashboard = () => {
  const rows = filteredRows();
  renderKpis(rows);
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

const loadData = async () => {
  const response = await fetch(DATA_URL);
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    defval: "",
  });
  return rows;
};

const init = async () => {
  rawData = await loadData();
  FILTER_COLUMNS.forEach((column) => {
    buildFilter(column, uniqueValues(rawData, column));
  });
  attachFilterListeners();
  resetButton.addEventListener("click", resetFilters);
  updateDashboard();
};

init();
