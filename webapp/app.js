// ---- Configuration ----
const DEFAULT_BANDS = [
  { min: 1200, max: 1500 },
  { min: 1500, max: 1800 },
  { min: 1800, max: 2100 },
  { min: 2100, max: 2400 },
];
const BEDROOM_TYPES = [1, 2, 3];
const RENT_COLUMN_NAMES = ["askingrent", "rent", "askingprice"];
const BEDROOM_COLUMN_NAMES = ["bedrooms", "bedroom", "beds"];
const PREVIEW_ROW_LIMIT = 10;

// ---- State ----
let parsedRows = [];
let parsedColumns = [];
let bands = DEFAULT_BANDS.map((band) => ({ ...band }));
let latestResult = null;
const charts = new Map();

// ---- Elements ----
const fileInput = document.getElementById("fileInput");
const fileStatus = document.getElementById("fileStatus");
const previewSection = document.getElementById("previewSection");
const previewTable = document.getElementById("previewTable");
const previewHint = document.getElementById("previewHint");
const bandsSection = document.getElementById("bandsSection");
const bandList = document.getElementById("bandList");
const bandError = document.getElementById("bandError");
const addBandBtn = document.getElementById("addBandBtn");
const resetBandsBtn = document.getElementById("resetBandsBtn");
const analyzeSection = document.getElementById("analyzeSection");
const analyzeBtn = document.getElementById("analyzeBtn");
const analyzeStatus = document.getElementById("analyzeStatus");
const resultsSection = document.getElementById("resultsSection");
const chartTypeSelect = document.getElementById("chartType");
const chartCards = document.getElementById("chartCards");
const summary = document.getElementById("summary");

// ---- File upload ----
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
      error: (err) => setStatus(fileStatus, `Error parsing CSV: ${err.message}`, "error"),
    });
  } else if (extension === "xlsx" || extension === "xls") {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        loadParsedData(rows, rows.length > 0 ? Object.keys(rows[0]) : []);
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
  bandsSection.hidden = false;
  analyzeSection.hidden = false;
}

function resetAfterNewFile() {
  parsedRows = [];
  parsedColumns = [];
  latestResult = null;
  previewSection.hidden = true;
  bandsSection.hidden = true;
  analyzeSection.hidden = true;
  resultsSection.hidden = true;
  setStatus(analyzeStatus, "", "");
  destroyCharts();
}

// ---- Preview table ----
function renderPreview() {
  const rowsToShow = parsedRows.slice(0, PREVIEW_ROW_LIMIT);
  const thead = `<thead><tr>${parsedColumns.map((c) => `<th>${escapeHtml(c)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rowsToShow
    .map((row) => `<tr>${parsedColumns.map((c) => `<td>${escapeHtml(row[c])}</td>`).join("")}</tr>`)
    .join("")}</tbody>`;

  previewTable.innerHTML = thead + tbody;
  previewHint.textContent =
    parsedRows.length > PREVIEW_ROW_LIMIT
      ? `Showing first ${PREVIEW_ROW_LIMIT} of ${parsedRows.length} rows.`
      : `Showing all ${parsedRows.length} row(s).`;
  previewSection.hidden = false;
}

// ---- Price band controls ----
function renderBands() {
  bandList.innerHTML = bands
    .map(
      (band, index) => `
      <div class="band-row">
        <span class="band-name">Band ${index + 1}</span>
        <span class="currency">£</span>
        <input type="number" class="band-input" data-index="${index}" data-field="min" value="${band.min}" min="0" step="50" />
        <span class="band-to">to</span>
        <span class="currency">£</span>
        <input type="number" class="band-input" data-index="${index}" data-field="max" value="${band.max}" min="0" step="50" />
        <button class="secondary remove-band" data-index="${index}" ${bands.length <= 1 ? "disabled" : ""}>Remove</button>
      </div>`
    )
    .join("");
}

bandList.addEventListener("change", (event) => {
  const input = event.target.closest(".band-input");
  if (!input) return;
  const index = Number(input.dataset.index);
  bands[index][input.dataset.field] = parseFloat(input.value);
  reanalyzeIfShowingResults();
});

bandList.addEventListener("click", (event) => {
  const button = event.target.closest(".remove-band");
  if (!button) return;
  bands.splice(Number(button.dataset.index), 1);
  renderBands();
  reanalyzeIfShowingResults();
});

addBandBtn.addEventListener("click", () => {
  const last = bands[bands.length - 1];
  const width = last ? Math.max(last.max - last.min, 50) : 300;
  const min = last ? last.max : 0;
  bands.push({ min, max: min + width });
  renderBands();
  reanalyzeIfShowingResults();
});

resetBandsBtn.addEventListener("click", () => {
  bands = DEFAULT_BANDS.map((band) => ({ ...band }));
  renderBands();
  reanalyzeIfShowingResults();
});

function validateBands() {
  for (let i = 0; i < bands.length; i++) {
    const { min, max } = bands[i];
    if (!Number.isFinite(min) || !Number.isFinite(max)) {
      return `Band ${i + 1} needs two numbers.`;
    }
    if (min >= max) {
      return `Band ${i + 1}: the lower value must be less than the upper value.`;
    }
  }
  return null;
}

// ---- Analysis ----
// For each bedroom type, work out what share of that type's listings falls in
// each rent band. Percentages are relative to the bedroom type, so each chart
// sums to 100%.
function analyzeData(rows, columns, activeBands) {
  const rentColumn = findColumn(columns, RENT_COLUMN_NAMES);
  const bedroomColumn = findColumn(columns, BEDROOM_COLUMN_NAMES);

  if (!rentColumn) throw new Error('No "Asking Rent" column found in the file.');
  if (!bedroomColumn) throw new Error('No "Bedrooms" column found in the file.');

  const labels = activeBands.map(bandLabel).concat("Other");
  const otherIndex = labels.length - 1;
  const groups = BEDROOM_TYPES.map((bedrooms) => ({
    bedrooms,
    total: 0,
    counts: new Array(labels.length).fill(0),
  }));
  const skipped = { bedrooms: 0, rent: 0, otherBedrooms: 0 };

  for (const row of rows) {
    const bedrooms = parseNumber(row[bedroomColumn]);
    if (!Number.isFinite(bedrooms)) {
      skipped.bedrooms++;
      continue;
    }

    const group = groups.find((g) => g.bedrooms === Math.round(bedrooms));
    if (!group) {
      skipped.otherBedrooms++;
      continue;
    }

    const rent = parseNumber(row[rentColumn]);
    if (!Number.isFinite(rent)) {
      skipped.rent++;
      continue;
    }

    const bandIndex = activeBands.findIndex((band) => rent >= band.min && rent < band.max);
    group.counts[bandIndex === -1 ? otherIndex : bandIndex]++;
    group.total++;
  }

  for (const group of groups) {
    group.percentages = group.counts.map((count) => (group.total ? (count / group.total) * 100 : 0));
  }

  return { labels, groups, skipped, rentColumn, bedroomColumn };
}

analyzeBtn.addEventListener("click", runAnalysis);

function runAnalysis() {
  if (parsedRows.length === 0) {
    setStatus(analyzeStatus, "Please upload a file first.", "error");
    return;
  }

  const bandProblem = validateBands();
  setStatus(bandError, bandProblem || "", bandProblem ? "error" : "");
  if (bandProblem) {
    setStatus(analyzeStatus, "Fix the price bands before analysing.", "error");
    return;
  }

  try {
    latestResult = analyzeData(parsedRows, parsedColumns, bands);
    setStatus(analyzeStatus, "Analysis complete.", "success");
    renderSummary(latestResult);
    renderCharts(latestResult);
    resultsSection.hidden = false;
  } catch (err) {
    latestResult = null;
    resultsSection.hidden = true;
    setStatus(analyzeStatus, err.message, "error");
  }
}

function reanalyzeIfShowingResults() {
  if (latestResult) runAnalysis();
}

function renderSummary(result) {
  const analysed = result.groups.reduce((sum, group) => sum + group.total, 0);
  const notes = [];
  if (result.skipped.otherBedrooms > 0) {
    notes.push(`${result.skipped.otherBedrooms} row(s) outside 1–3 bedrooms`);
  }
  if (result.skipped.bedrooms > 0) notes.push(`${result.skipped.bedrooms} row(s) with unreadable bedrooms`);
  if (result.skipped.rent > 0) notes.push(`${result.skipped.rent} row(s) with unreadable rent`);

  summary.textContent =
    `Analysed ${analysed} listing(s) using "${result.rentColumn}" and "${result.bedroomColumn}".` +
    (notes.length > 0 ? ` Excluded: ${notes.join(", ")}.` : "");
}

// ---- Charts ----
const whiteBackground = {
  id: "whiteBackground",
  beforeDraw(chart) {
    const { ctx } = chart;
    ctx.save();
    ctx.globalCompositeOperation = "destination-over";
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, chart.width, chart.height);
    ctx.restore();
  },
};

// Percentages are the point of these charts, and a downloaded PNG has no
// tooltips, so bar values are drawn onto the canvas itself.
const barValueLabels = {
  id: "barValueLabels",
  afterDatasetsDraw(chart) {
    if (chart.config.type !== "bar") return;
    const { ctx } = chart;
    ctx.save();
    ctx.fillStyle = "#1a1a1a";
    ctx.font = '600 12px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    chart.getDatasetMeta(0).data.forEach((element, index) => {
      ctx.fillText(formatPercent(chart.data.datasets[0].data[index]), element.x, element.y - 4);
    });
    ctx.restore();
  },
};

Chart.register(whiteBackground, barValueLabels);

function renderCharts(result) {
  destroyCharts();
  chartCards.innerHTML = result.groups
    .map(
      (group) => `
      <div class="card chart-card">
        <div class="chart-head">
          <h3>${group.bedrooms} bedroom</h3>
          <button class="download-btn" data-bedrooms="${group.bedrooms}" ${group.total === 0 ? "disabled" : ""}>Download PNG</button>
        </div>
        ${
          group.total === 0
            ? `<p class="empty-msg">No ${group.bedrooms} bedroom listings found.</p>`
            : `<div class="chart-wrap"><canvas id="chart-${group.bedrooms}"></canvas></div>`
        }
      </div>`
    )
    .join("");

  const type = chartTypeSelect.value;
  const palette = generatePalette(result.labels.length);

  for (const group of result.groups) {
    if (group.total === 0) continue;
    const canvas = document.getElementById(`chart-${group.bedrooms}`);
    charts.set(group.bedrooms, new Chart(canvas, buildChartConfig(type, result.labels, group, palette)));
  }
}

function buildChartConfig(type, labels, group, palette) {
  return {
    type,
    data: {
      labels,
      datasets: [
        {
          label: "Share of listings",
          data: group.percentages,
          backgroundColor: palette,
          borderColor: "#ffffff",
          borderWidth: type === "pie" ? 1 : 0,
        },
      ],
    },
    options: {
      responsive: true,
      layout: { padding: { top: 16 } },
      plugins: {
        title: {
          display: true,
          text: `${group.bedrooms} bedroom — ${group.total} listing(s)`,
          font: { size: 15 },
          color: "#1a1a1a",
        },
        legend: {
          display: type === "pie",
          position: "right",
          labels: {
            color: "#1a1a1a",
            generateLabels(chart) {
              const dataset = chart.data.datasets[0];
              return chart.data.labels.map((label, index) => ({
                text: `${label} — ${formatPercent(dataset.data[index])}`,
                fillStyle: dataset.backgroundColor[index],
                strokeStyle: "#ffffff",
                lineWidth: 1,
                hidden: !chart.getDataVisibility(index),
                index,
              }));
            },
          },
        },
        tooltip: {
          callbacks: {
            label(context) {
              const value = typeof context.parsed === "number" ? context.parsed : context.parsed.y;
              const count = group.counts[context.dataIndex];
              return `${formatPercent(value)} (${count} listing${count === 1 ? "" : "s"})`;
            },
          },
        },
      },
      scales:
        type === "bar"
          ? {
              y: {
                beginAtZero: true,
                grace: "12%", // headroom so the value label above the tallest bar clears the title
                ticks: { callback: (value) => `${value}%`, color: "#1a1a1a" },
                grid: { color: "#e2e4e9" },
              },
              x: { ticks: { color: "#1a1a1a" }, grid: { display: false } },
            }
          : undefined,
    },
  };
}

chartTypeSelect.addEventListener("change", () => {
  if (latestResult) renderCharts(latestResult);
});

chartCards.addEventListener("click", (event) => {
  const button = event.target.closest(".download-btn");
  if (!button) return;

  const bedrooms = Number(button.dataset.bedrooms);
  const chart = charts.get(bedrooms);
  if (!chart) return;

  const link = document.createElement("a");
  link.href = chart.toBase64Image("image/png", 1);
  link.download = `${bedrooms}-bedroom-rent-distribution.png`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
});

function destroyCharts() {
  for (const chart of charts.values()) {
    chart.destroy();
  }
  charts.clear();
  chartCards.innerHTML = "";
}

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
  return Array.from({ length: count }, (_, index) => baseColors[index % baseColors.length]);
}

// ---- Helpers ----
function parseNumber(value) {
  if (typeof value === "number") return value;
  if (value === null || value === undefined) return NaN;
  const cleaned = String(value).replace(/[^0-9.]/g, "");
  return cleaned === "" ? NaN : parseFloat(cleaned);
}

function normalizeKey(name) {
  return String(name).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function findColumn(columns, candidates) {
  for (const candidate of candidates) {
    const match = columns.find((column) => normalizeKey(column) === candidate);
    if (match) return match;
  }
  return null;
}

function bandLabel(band) {
  return `£${band.min.toLocaleString("en-GB")}–£${band.max.toLocaleString("en-GB")}`;
}

function formatPercent(value) {
  return `${value.toFixed(1)}%`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function setStatus(element, message, kind) {
  element.textContent = message;
  element.className = "status" + (kind ? ` ${kind}` : "");
}

renderBands();
