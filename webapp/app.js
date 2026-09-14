// ---- State ----
let parsedRows = []; // array of objects, one per data row
let parsedColumns = []; // array of column names
let chartInstance = null;

// ---- Elements ----
const fileInput = document.getElementById("fileInput");
const fileStatus = document.getElementById("fileStatus");
const previewSection = document.getElementById("previewSection");
const previewTable = document.getElementById("previewTable");
const previewHint = document.getElementById("previewHint");
const analyzeSection = document.getElementById("analyzeSection");
const analyzeBtn = document.getElementById("analyzeBtn");
const analyzeStatus = document.getElementById("analyzeStatus");
const resultsSection = document.getElementById("resultsSection");
const chartTypeSelect = document.getElementById("chartType");
const downloadBtn = document.getElementById("downloadBtn");
const resultCanvas = document.getElementById("resultChart");

const PREVIEW_ROW_LIMIT = 10;

// ---- File upload handling ----
fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (!file) return;

  resetAfterNewFile();
  setStatus(fileStatus, `Reading "${file.name}"...`, "");

  const extension = file.name.split(".").pop().toLowerCase();

  if (extension === "csv") {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        if (results.errors && results.errors.length > 0) {
          setStatus(fileStatus, `Error parsing CSV: ${results.errors[0].message}`, "error");
          return;
        }
        loadParsedData(results.data, results.meta.fields || []);
      },
      error: (err) => {
        setStatus(fileStatus, `Error parsing CSV: ${err.message}`, "error");
      },
    });
  } else if (extension === "xlsx" || extension === "xls") {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const columns = rows.length > 0 ? Object.keys(rows[0]) : [];
        loadParsedData(rows, columns);
      } catch (err) {
        setStatus(fileStatus, `Error parsing Excel file: ${err.message}`, "error");
      }
    };
    reader.onerror = () => setStatus(fileStatus, "Error reading file.", "error");
    reader.readAsArrayBuffer(file);
  } else {
    setStatus(fileStatus, "Unsupported file type. Please upload a .csv, .xlsx, or .xls file.", "error");
  }
});

function loadParsedData(rows, columns) {
  if (!rows || rows.length === 0) {
    setStatus(fileStatus, "The file has no data rows.", "error");
    return;
  }
  parsedRows = rows;
  parsedColumns = columns;
  setStatus(fileStatus, `Loaded ${rows.length} row(s), ${columns.length} column(s).`, "success");
  renderPreview();
  analyzeSection.hidden = false;
}

function resetAfterNewFile() {
  parsedRows = [];
  parsedColumns = [];
  previewSection.hidden = true;
  analyzeSection.hidden = true;
  resultsSection.hidden = true;
  setStatus(analyzeStatus, "", "");
  if (chartInstance) {
    chartInstance.destroy();
    chartInstance = null;
  }
}

// ---- Preview table ----
function renderPreview() {
  const rowsToShow = parsedRows.slice(0, PREVIEW_ROW_LIMIT);

  const thead = `<thead><tr>${parsedColumns.map((c) => `<th>${escapeHtml(c)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rowsToShow
    .map(
      (row) =>
        `<tr>${parsedColumns.map((c) => `<td>${escapeHtml(row[c])}</td>`).join("")}</tr>`
    )
    .join("")}</tbody>`;

  previewTable.innerHTML = thead + tbody;
  previewHint.textContent =
    parsedRows.length > PREVIEW_ROW_LIMIT
      ? `Showing first ${PREVIEW_ROW_LIMIT} of ${parsedRows.length} rows.`
      : `Showing all ${parsedRows.length} row(s).`;
  previewSection.hidden = false;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ---- Analysis ----
// Placeholder analysis logic. Replace the body of this function with the
// real analysis once it's provided. It must return { labels: string[], values: number[] }.
function analyzeData(rows, columns) {
  if (columns.length === 0) {
    return { labels: [], values: [] };
  }

  const groupColumn = columns[0];
  const numericColumn = columns.slice(1).find((col) =>
    rows.every((row) => row[col] === "" || !isNaN(parseFloat(row[col])))
  );

  const groups = new Map();

  for (const row of rows) {
    const key = String(row[groupColumn] ?? "Unknown");
    const increment = numericColumn ? parseFloat(row[numericColumn]) || 0 : 1;
    groups.set(key, (groups.get(key) || 0) + increment);
  }

  return {
    labels: Array.from(groups.keys()),
    values: Array.from(groups.values()),
  };
}

analyzeBtn.addEventListener("click", () => {
  if (parsedRows.length === 0) {
    setStatus(analyzeStatus, "Please upload a file first.", "error");
    return;
  }

  setStatus(analyzeStatus, "Analyzing...", "");

  try {
    const result = analyzeData(parsedRows, parsedColumns);
    if (!result.labels || result.labels.length === 0) {
      setStatus(analyzeStatus, "Analysis produced no results.", "error");
      return;
    }
    setStatus(analyzeStatus, "Analysis complete.", "success");
    renderChart(result);
    resultsSection.hidden = false;
  } catch (err) {
    setStatus(analyzeStatus, `Error during analysis: ${err.message}`, "error");
  }
});

// ---- Chart rendering ----
let latestResult = null;

function renderChart(result) {
  latestResult = result;
  const type = chartTypeSelect.value;

  if (chartInstance) {
    chartInstance.destroy();
  }

  const palette = generatePalette(result.labels.length);

  chartInstance = new Chart(resultCanvas, {
    type,
    data: {
      labels: result.labels,
      datasets: [
        {
          label: "Value",
          data: result.values,
          backgroundColor: palette,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        legend: {
          display: type === "pie",
        },
      },
      scales:
        type === "bar"
          ? {
              y: {
                beginAtZero: true,
              },
            }
          : undefined,
    },
  });
}

chartTypeSelect.addEventListener("change", () => {
  if (latestResult) {
    renderChart(latestResult);
  }
});

function generatePalette(count) {
  const baseColors = [
    "#2563eb",
    "#f97316",
    "#16a34a",
    "#dc2626",
    "#9333ea",
    "#0891b2",
    "#ca8a04",
    "#db2777",
    "#4338ca",
    "#65a30d",
  ];
  const colors = [];
  for (let i = 0; i < count; i++) {
    colors.push(baseColors[i % baseColors.length]);
  }
  return colors;
}

// ---- PNG download ----
downloadBtn.addEventListener("click", () => {
  if (!chartInstance) return;
  const url = chartInstance.toBase64Image("image/png", 1);
  const link = document.createElement("a");
  link.href = url;
  link.download = "chart.png";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
});

// ---- Helpers ----
function setStatus(element, message, kind) {
  element.textContent = message;
  element.className = "status" + (kind ? ` ${kind}` : "");
}
