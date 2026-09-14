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
const DEMAND_ROWS_ON_LOAD = 2;
const SUPPLY_COLOR = "#2563eb";
const DEMAND_COLOR = "#f97316";

// ---- State ----
let parsedRows = [];
let parsedColumns = [];
let bands = DEFAULT_BANDS.map((band) => ({ ...band }));
let demand = buildEmptyDemand();
let latestResult = null;
const charts = new Map();

// ---- Elements ----
const fileInput = document.getElementById("fileInput");
const fileStatus = document.getElementById("fileStatus");
const previewBlock = document.getElementById("previewBlock");
const previewTable = document.getElementById("previewTable");
const previewHint = document.getElementById("previewHint");
const demandSection = document.getElementById("demandSection");
const demandGroups = document.getElementById("demandGroups");
const bandsSection = document.getElementById("bandsSection");
const bandList = document.getElementById("bandList");
const bandError = document.getElementById("bandError");
const addBandBtn = document.getElementById("addBandBtn");
const resetBandsBtn = document.getElementById("resetBandsBtn");
const analyzeSection = document.getElementById("analyzeSection");
const analyzeBtn = document.getElementById("analyzeBtn");
const analyzeStatus = document.getElementById("analyzeStatus");
const resultsSection = document.getElementById("resultsSection");
const chartCards = document.getElementById("chartCards");
const summary = document.getElementById("summary");

// ---- Step 1: supply upload ----
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
  demandSection.hidden = false;
  bandsSection.hidden = false;
  analyzeSection.hidden = false;
}

function resetAfterNewFile() {
  parsedRows = [];
  parsedColumns = [];
  latestResult = null;
  previewBlock.hidden = true;
  demandSection.hidden = true;
  bandsSection.hidden = true;
  analyzeSection.hidden = true;
  resultsSection.hidden = true;
  setStatus(analyzeStatus, "", "");
  destroyCharts();
}

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
  previewBlock.hidden = false;
}

// ---- Step 2: demand entry ----
function buildEmptyDemand() {
  const entries = {};
  for (const bedrooms of BEDROOM_TYPES) {
    entries[bedrooms] = Array.from({ length: DEMAND_ROWS_ON_LOAD }, () => ({ price: "", quantity: "" }));
  }
  return entries;
}

function renderDemand() {
  demandGroups.innerHTML = BEDROOM_TYPES.map(
    (bedrooms) => `
      <div class="demand-group">
        <h3>${bedrooms} bedroom</h3>
        <div class="demand-rows">
          ${demand[bedrooms]
            .map(
              (row, index) => `
              <div class="demand-row">
                <span class="currency">£</span>
                <input type="number" class="demand-input" data-bedrooms="${bedrooms}" data-index="${index}"
                  data-field="price" value="${escapeHtml(row.price)}" placeholder="Price" min="0" step="50" />
                <span class="times">&times;</span>
                <input type="number" class="demand-input" data-bedrooms="${bedrooms}" data-index="${index}"
                  data-field="quantity" value="${escapeHtml(row.quantity)}" placeholder="Units" min="0" step="1" />
                <button class="secondary remove-demand" data-bedrooms="${bedrooms}" data-index="${index}"
                  ${demand[bedrooms].length <= 1 ? "disabled" : ""}>Remove</button>
              </div>`
            )
            .join("")}
        </div>
        <button class="secondary add-demand" data-bedrooms="${bedrooms}">Add price point</button>
      </div>`
  ).join("");
}

// Inputs update state without re-rendering, so typing never loses focus.
demandGroups.addEventListener("change", (event) => {
  const input = event.target.closest(".demand-input");
  if (!input) return;
  demand[input.dataset.bedrooms][Number(input.dataset.index)][input.dataset.field] = input.value;
  reanalyzeIfShowingResults();
});

demandGroups.addEventListener("click", (event) => {
  const addButton = event.target.closest(".add-demand");
  if (addButton) {
    demand[addButton.dataset.bedrooms].push({ price: "", quantity: "" });
    renderDemand();
    return;
  }

  const removeButton = event.target.closest(".remove-demand");
  if (removeButton) {
    demand[removeButton.dataset.bedrooms].splice(Number(removeButton.dataset.index), 1);
    renderDemand();
    reanalyzeIfShowingResults();
  }
});

// ---- Price bands ----
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
  bands[Number(input.dataset.index)][input.dataset.field] = parseFloat(input.value);
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
    if (!Number.isFinite(min) || !Number.isFinite(max)) return `Band ${i + 1} needs two numbers.`;
    if (min >= max) return `Band ${i + 1}: the lower value must be less than the upper value.`;
  }
  return null;
}

// ---- Step 3: analysis ----
// For each bedroom type, work out what share of that type's supply (listings)
// and demand (units) falls in each price band. Percentages are relative to the
// bedroom type and to the series, so supply and demand each sum to 100% and can
// be compared side by side.
function analyzeData(rows, columns, activeBands, demandInput) {
  const rentColumn = findColumn(columns, RENT_COLUMN_NAMES);
  const bedroomColumn = findColumn(columns, BEDROOM_COLUMN_NAMES);

  if (!rentColumn) throw new Error('No "Asking Rent" column found in the file.');
  if (!bedroomColumn) throw new Error('No "Bedrooms" column found in the file.');

  const labels = activeBands.map(bandLabel).concat("Other");
  const groups = BEDROOM_TYPES.map((bedrooms) => ({
    bedrooms,
    supply: emptySeries(labels.length),
    demand: emptySeries(labels.length),
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

    addToSeries(group.supply, bandIndexFor(rent, activeBands, labels.length), 1);
  }

  for (const group of groups) {
    for (const entry of demandInput[group.bedrooms]) {
      const price = parseNumber(entry.price);
      const quantity = parseNumber(entry.quantity);
      if (!Number.isFinite(price) || !Number.isFinite(quantity)) continue;
      addToSeries(group.demand, bandIndexFor(price, activeBands, labels.length), quantity);
    }
  }

  for (const group of groups) {
    finaliseSeries(group.supply);
    finaliseSeries(group.demand);
  }

  return { labels, groups, skipped, rentColumn, bedroomColumn };
}

function emptySeries(length) {
  return { total: 0, counts: new Array(length).fill(0), percentages: new Array(length).fill(0) };
}

function bandIndexFor(value, activeBands, labelCount) {
  const index = activeBands.findIndex((band) => value >= band.min && value < band.max);
  return index === -1 ? labelCount - 1 : index;
}

function addToSeries(series, index, amount) {
  series.counts[index] += amount;
  series.total += amount;
}

function finaliseSeries(series) {
  series.percentages = series.counts.map((count) => (series.total ? (count / series.total) * 100 : 0));
}

analyzeBtn.addEventListener("click", runAnalysis);

function runAnalysis() {
  if (parsedRows.length === 0) {
    setStatus(analyzeStatus, "Please upload a supply file first.", "error");
    return;
  }

  const bandProblem = validateBands();
  setStatus(bandError, bandProblem || "", bandProblem ? "error" : "");
  if (bandProblem) {
    setStatus(analyzeStatus, "Fix the price bands before analysing.", "error");
    return;
  }

  try {
    latestResult = analyzeData(parsedRows, parsedColumns, bands, demand);
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
  const supplyTotal = result.groups.reduce((sum, group) => sum + group.supply.total, 0);
  const demandTotal = result.groups.reduce((sum, group) => sum + group.demand.total, 0);
  const notes = [];
  if (result.skipped.otherBedrooms > 0) notes.push(`${result.skipped.otherBedrooms} row(s) outside 1–3 bedrooms`);
  if (result.skipped.bedrooms > 0) notes.push(`${result.skipped.bedrooms} row(s) with unreadable bedrooms`);
  if (result.skipped.rent > 0) notes.push(`${result.skipped.rent} row(s) with unreadable rent`);

  summary.textContent =
    `Supply: ${supplyTotal} listing(s) from "${result.rentColumn}" and "${result.bedroomColumn}". ` +
    `Demand: ${demandTotal} unit(s) entered.` +
    (notes.length > 0 ? ` Excluded from supply: ${notes.join(", ")}.` : "");
}

// ---- Step 4: charts ----
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
// tooltips, so the values are drawn onto the canvas itself.
const barValueLabels = {
  id: "barValueLabels",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    ctx.save();
    ctx.fillStyle = "#1a1a1a";
    ctx.font = '600 10px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      if (meta.hidden) return;
      meta.data.forEach((element, index) => {
        const value = dataset.data[index];
        if (value > 0) ctx.fillText(formatPercent(value), element.x, element.y - 3);
      });
    });
    ctx.restore();
  },
};

Chart.register(whiteBackground, barValueLabels);

function renderCharts(result) {
  destroyCharts();
  chartCards.innerHTML = result.groups
    .map((group) => {
      const hasData = group.supply.total > 0 || group.demand.total > 0;
      return `
      <div class="card chart-card">
        <div class="chart-head">
          <h3>${group.bedrooms} bedroom</h3>
          <button class="download-btn" data-bedrooms="${group.bedrooms}" ${hasData ? "" : "disabled"}>Download PNG</button>
        </div>
        ${
          hasData
            ? `<div class="chart-wrap"><canvas id="chart-${group.bedrooms}"></canvas></div>`
            : `<p class="empty-msg">No ${group.bedrooms} bedroom supply listings or demand entered.</p>`
        }
      </div>`;
    })
    .join("");

  for (const group of result.groups) {
    if (group.supply.total === 0 && group.demand.total === 0) continue;
    const canvas = document.getElementById(`chart-${group.bedrooms}`);
    charts.set(group.bedrooms, new Chart(canvas, buildChartConfig(result.labels, group)));
  }
}

function buildChartConfig(labels, group) {
  const datasets = [];
  if (group.supply.total > 0) {
    datasets.push({ label: "Supply", data: group.supply.percentages, backgroundColor: SUPPLY_COLOR });
  }
  if (group.demand.total > 0) {
    datasets.push({ label: "Demand", data: group.demand.percentages, backgroundColor: DEMAND_COLOR });
  }

  const titleParts = [];
  if (group.supply.total > 0) titleParts.push(`supply ${group.supply.total} listing(s)`);
  if (group.demand.total > 0) titleParts.push(`demand ${group.demand.total} unit(s)`);

  return {
    type: "bar",
    data: { labels, datasets },
    options: {
      responsive: true,
      layout: { padding: { top: 8 } },
      plugins: {
        title: {
          display: true,
          text: `${group.bedrooms} bedroom — ${titleParts.join(", ")}`,
          font: { size: 15 },
          color: "#1a1a1a",
        },
        legend: { display: true, position: "top", labels: { color: "#1a1a1a", boxWidth: 14 } },
        tooltip: {
          callbacks: {
            label(context) {
              const series = context.dataset.label === "Supply" ? group.supply : group.demand;
              const raw = series.counts[context.dataIndex];
              const unit = context.dataset.label === "Supply" ? "listing" : "unit";
              return `${context.dataset.label}: ${formatPercent(context.parsed.y)} (${raw} ${unit}${raw === 1 ? "" : "s"})`;
            },
          },
        },
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: "12%", // headroom so the value label above the tallest bar clears the title
          ticks: { callback: (value) => `${value}%`, color: "#1a1a1a" },
          grid: { color: "#e2e4e9" },
        },
        x: { ticks: { color: "#1a1a1a" }, grid: { display: false } },
      },
    },
  };
}

chartCards.addEventListener("click", (event) => {
  const button = event.target.closest(".download-btn");
  if (!button) return;

  const chart = charts.get(Number(button.dataset.bedrooms));
  if (!chart) return;

  const link = document.createElement("a");
  link.href = chart.toBase64Image("image/png", 1);
  link.download = `${button.dataset.bedrooms}-bedroom-supply-vs-demand.png`;
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
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function setStatus(element, message, kind) {
  element.textContent = message;
  element.className = "status" + (kind ? ` ${kind}` : "");
}

renderBands();
renderDemand();
