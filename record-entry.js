const DATA_PATH = "data/Combined.xlsx";
const SHEET_NAME = "Sheet1";
const TOKEN_STORAGE_KEY = "floating-shelf-gh-token";
const REPO_STORAGE_KEY = "floating-shelf-gh-repo";

const form = document.getElementById("recordForm");
const formFields = document.getElementById("formFields");
const formMessage = document.getElementById("formMessage");
const connectionStatus = document.getElementById("connectionStatus");
const manageTokenButton = document.getElementById("manageToken");
const cancelButtons = [
  document.getElementById("cancelForm"),
  document.getElementById("backToDashboard"),
];

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

  const uniqueValues = Array.from(new Set(sample.map((value) => cleanValue(value)))).filter(
    Boolean
  );
  if (uniqueValues.length > 0 && uniqueValues.length <= 8) {
    return "select";
  }

  return "text";
};

const inferSelectOptions = (values) => {
  const uniqueValues = Array.from(new Set(values.map((value) => cleanValue(value)))).filter(
    Boolean
  );
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
    return { owner, repo, branch: "main", path: DATA_PATH };
  }

  return { owner: "", repo: "", branch: "main", path: DATA_PATH };
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
  const response = await fetch(`${DATA_PATH}?v=${Date.now()}`, {
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

const appendRowToSheet = (workbook, headers, rowValues) => {
  const sheet = workbook.Sheets[SHEET_NAME];
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  }

  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const safeHeaders = data[0] || headers;
  const nextRow = safeHeaders.map((header, index) => {
    const value = rowValues[index];
    return value === undefined || value === null ? "" : value;
  });
  data.push(nextRow);

  const updatedSheet = XLSX.utils.aoa_to_sheet(data);
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

const init = async () => {
  cancelButtons.forEach((button) => {
    if (button) {
      button.addEventListener("click", () => {
        window.location.href = "index.html";
      });
    }
  });

  const repoDetails = buildRepoDetails();
  updateConnectionStatus(repoDetails);

  manageTokenButton.addEventListener("click", () => {
    requestToken(repoDetails);
  });

  try {
    const workbook = await fetchWorkbook();
    const sheet = workbook.Sheets[SHEET_NAME];
    if (!sheet) {
      throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    }
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const headers = data[0];
    if (!headers || headers.length === 0) {
      throw new Error("Sheet1 does not contain headers.");
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
      window.setTimeout(() => {
        window.location.href = "index.html";
      }, 1200);
    } catch (error) {
      formatMessage(error.message || "Unable to save record.", true);
    }
  });
};

init();
