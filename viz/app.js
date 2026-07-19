const MONTH_NUMBERS = Array.from({ length: 12 }, (_, index) => index + 1);
const MIN_PACE_SECONDS = 3 * 60 + 45;
const MAX_PACE_SECONDS = 15 * 60;
const state = { year: null, startMonth: 1, endMonth: 12 };

const appEl = document.getElementById("app");
const statusEl = document.getElementById("status");
const chartEl = document.getElementById("chart");
const yearSelect = document.getElementById("year-select");
const startMonthSelect = document.getElementById("start-month-select");
const endMonthSelect = document.getElementById("end-month-select");

function formatPaceMMSS(seconds) {
  if (!Number.isFinite(seconds)) return "—";
  const rounded = Math.round(seconds);
  return `${Math.floor(rounded / 60)}:${String(rounded % 60).padStart(2, "0")}`;
}

function buildPaceData(activities) {
  const byYear = {};
  const cleanedPaces = [];

  for (const activity of activities) {
    const { year, month, pace_s_per_km: pace } = activity;
    if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12
        || !Number.isFinite(pace) || pace < MIN_PACE_SECONDS || pace > MAX_PACE_SECONDS) continue;

    const yearData = byYear[year] ||= {
      pace_sums: Array(12).fill(0),
      pace_s_per_km: Array(12).fill(null),
      activity_count: Array(12).fill(0),
    };
    yearData.pace_sums[month - 1] += pace;
    yearData.activity_count[month - 1] += 1;
    cleanedPaces.push(pace);
  }

  for (const yearData of Object.values(byYear)) {
    yearData.pace_s_per_km = yearData.pace_sums.map((sum, index) =>
      yearData.activity_count[index] ? sum / yearData.activity_count[index] : null);
    delete yearData.pace_sums;
  }

  return { byYear, cleanedPaces, originalCount: activities.length };
}

function summarizeMonths(yearData) {
  const selectedPaces = [];
  let activityCount = 0;

  for (const month of MONTH_NUMBERS) {
    if (month < state.startMonth || month > state.endMonth) continue;
    const pace = yearData.pace_s_per_km[month - 1];
    if (Number.isFinite(pace)) selectedPaces.push(pace);
    activityCount += yearData.activity_count[month - 1] || 0;
  }

  return {
    monthsPlotted: selectedPaces.length,
    averagePace: selectedPaces.length
      ? selectedPaces.reduce((sum, pace) => sum + pace, 0) / selectedPaces.length
      : null,
    activityCount,
  };
}

function buildPaceScale(paces) {
  const present = paces.filter(Number.isFinite);
  const top = present.length
    ? Math.ceil((Math.max(...present) + 30) / 60) * 60
    : 600;
  const step = top <= 600 ? 60 : top <= 1200 ? 120 : 300;
  const tickvals = [];
  for (let value = 0; value <= top; value += step) tickvals.push(value);
  return { range: [0, top], tickvals, ticktext: tickvals.map(formatPaceMMSS) };
}

function renderMonthTable(yearData) {
  const body = document.getElementById("month-table-body");
  body.replaceChildren();

  for (const month of MONTH_NUMBERS) {
    const inRange = month >= state.startMonth && month <= state.endMonth;
    const row = document.createElement("tr");
    if (!inRange) row.className = "outside-range";

    const values = [
      `Month ${month}`,
      inRange ? formatPaceMMSS(yearData.pace_s_per_km[month - 1]) : "—",
      inRange ? String(yearData.activity_count[month - 1] || 0) : "—",
      `${state.year}-${String(month).padStart(2, "0")}`,
    ];

    for (const value of values) {
      const cell = document.createElement("td");
      cell.textContent = value;
      row.appendChild(cell);
    }
    body.appendChild(row);
  }
}

function renderSelectionSummary(yearData) {
  const summary = summarizeMonths(yearData);
  document.getElementById("months-plotted").textContent = summary.monthsPlotted;
  document.getElementById("average-pace").textContent = formatPaceMMSS(summary.averagePace);
  document.getElementById("activities-included").textContent = summary.activityCount.toLocaleString();
}

function renderCleanedSummary(paceData) {
  const cleaned = paceData.cleanedPaces.length;
  document.getElementById("original-activities").textContent = paceData.originalCount.toLocaleString();
  document.getElementById("cleaned-activities").textContent = cleaned.toLocaleString();
  document.getElementById("excluded-activities").textContent =
    (paceData.originalCount - cleaned).toLocaleString();
  document.getElementById("min-cleaned-pace").textContent = paceData.cleanedPaces.length
    ? (Math.min(...paceData.cleanedPaces) / 60).toFixed(2)
    : "—";
  document.getElementById("max-cleaned-pace").textContent = paceData.cleanedPaces.length
    ? (Math.max(...paceData.cleanedPaces) / 60).toFixed(2)
    : "—";
}

function renderChart(yearData) {
  const paces = yearData.pace_s_per_km.map((pace, index) => {
    const month = index + 1;
    return month >= state.startMonth && month <= state.endMonth ? pace : null;
  });
  const paceLabels = paces.map((pace) => Number.isFinite(pace) ? formatPaceMMSS(pace) : "");
  const paceScale = buildPaceScale(paces);
  const isNarrow = window.innerWidth <= 720;

  const trace = {
    x: MONTH_NUMBERS,
    y: paces,
    type: "scatter",
    mode: "lines+markers+text",
    connectgaps: false,
    line: { color: "#4472c4", width: 3 },
    marker: { color: "#4472c4", size: 7 },
    text: paceLabels,
    textposition: "top center",
    textfont: { color: "#333333", size: 12 },
    cliponaxis: false,
    customdata: paceLabels,
    hovertemplate: "Month %{x}<br>Average pace: %{customdata} /km<extra></extra>",
  };

  const layout = {
    title: { text: "Average Pace by Month (mm:ss/km)", font: { size: isNarrow ? 17 : 21 } },
    autosize: true,
    paper_bgcolor: "#ffffff",
    plot_bgcolor: "#ffffff",
    font: { family: '"Segoe UI", Arial, sans-serif', color: "#4a4a4a" },
    margin: isNarrow ? { l: 58, r: 18, t: 64, b: 78 } : { l: 76, r: 26, t: 72, b: 72 },
    showlegend: false,
    hovermode: "closest",
    dragmode: false,
    xaxis: {
      title: { text: "Month 1 to Month 12" },
      range: [0.5, 12.5],
      tickmode: "array",
      tickvals: MONTH_NUMBERS,
      ticktext: MONTH_NUMBERS.map((month) => isNarrow ? String(month) : `Month ${month}`),
      fixedrange: true,
      showgrid: false,
      zeroline: false,
    },
    yaxis: {
      title: { text: "Pace (mm:ss/km)" },
      ...paceScale,
      fixedrange: true,
      gridcolor: "#d9d9d9",
      zerolinecolor: "#bfbfbf",
    },
    annotations: paceLabels.some(Boolean) ? [] : [{
      text: "No pace data for this selection",
      x: 0.5,
      y: 0.5,
      xref: "paper",
      yref: "paper",
      showarrow: false,
      font: { size: 16, color: "#666666" },
    }],
  };

  Plotly.react(chartEl, [trace], layout, {
    responsive: true,
    displayModeBar: false,
    displaylogo: false,
  });
}

function populateSelects(years) {
  for (const year of years) {
    const option = document.createElement("option");
    option.value = year;
    option.textContent = year;
    yearSelect.appendChild(option);
  }

  for (const month of MONTH_NUMBERS) {
    for (const select of [startMonthSelect, endMonthSelect]) {
      const option = document.createElement("option");
      option.value = month;
      option.textContent = month;
      select.appendChild(option);
    }
  }

  state.year = years.at(-1);
  yearSelect.value = state.year;
  startMonthSelect.value = state.startMonth;
  endMonthSelect.value = state.endMonth;
}

function render(paceData) {
  const yearData = paceData.byYear[state.year];
  renderMonthTable(yearData);
  renderSelectionSummary(yearData);
  renderChart(yearData);
  document.getElementById("chart-note").textContent =
    `Horizontal axis locked to Month 1 through Month 12; pace displayed as mm:ss/km for selected year ${state.year}`;
}

function addInteractions(paceData) {
  yearSelect.addEventListener("change", () => {
    state.year = yearSelect.value;
    render(paceData);
  });
  startMonthSelect.addEventListener("change", () => {
    state.startMonth = Number(startMonthSelect.value);
    if (state.startMonth > state.endMonth) {
      state.endMonth = state.startMonth;
      endMonthSelect.value = state.endMonth;
    }
    render(paceData);
  });
  endMonthSelect.addEventListener("change", () => {
    state.endMonth = Number(endMonthSelect.value);
    if (state.endMonth < state.startMonth) {
      state.startMonth = state.endMonth;
      startMonthSelect.value = state.startMonth;
    }
    render(paceData);
  });
  window.addEventListener("resize", () => renderChart(paceData.byYear[state.year]));
}

function showError(message) {
  statusEl.className = "status status-error";
  statusEl.textContent = message;
}

function selfCheck() {
  const previous = { ...state };
  state.startMonth = 1;
  state.endMonth = 3;
  const summary = summarizeMonths({
    pace_s_per_km: [300, null, 360, ...Array(9).fill(null)],
    activity_count: [1, 0, 2, ...Array(9).fill(0)],
  });
  console.assert(summary.monthsPlotted === 2 && summary.activityCount === 3
    && formatPaceMMSS(summary.averagePace) === "5:30", "Pace Explorer self-check failed");
  const paceData = buildPaceData([
    { year: 2024, month: 1, pace_s_per_km: 300 },
    { year: 2024, month: 1, pace_s_per_km: 920 },
  ]);
  console.assert(paceData.cleanedPaces.length === 1
    && paceData.byYear[2024].pace_s_per_km[0] === 300, "Pace cleaning self-check failed");
  Object.assign(state, previous);
}

async function init() {
  selfCheck();
  if (typeof Plotly === "undefined") {
    showError("Plotly could not be loaded. Check the network connection and reload.");
    return;
  }

  try {
    const response = await fetch("./data/garmin_activities.json", { cache: "no-store" });
    if (!response.ok) throw new Error("Run `uv run python get-garmin.py` to create or refresh the dashboard data.");
    const chartData = await response.json();
    if (!Array.isArray(chartData.activities)) {
      throw new Error("The Garmin data file has no chartable activities.");
    }
    const paceData = buildPaceData(chartData.activities);
    const years = Object.keys(paceData.byYear).sort((a, b) => Number(a) - Number(b));
    if (!years.length) throw new Error("The Garmin data file has no chartable activities.");

    populateSelects(years);
    renderCleanedSummary(paceData);
    addInteractions(paceData);
    statusEl.hidden = true;
    appEl.hidden = false;
    render(paceData);
  } catch (error) {
    showError(error.message);
  }
}

window.addEventListener("DOMContentLoaded", init);
