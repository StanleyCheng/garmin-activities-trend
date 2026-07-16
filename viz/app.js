// viz/app.js — phase 2 baseline; phase 3 adds features.

const appEl = document.getElementById("app");
const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const state = {
  fromYear: null,
  toYear: null,
  bucket: "all",
  showPace: true,
  showDistance: true,
  mode: "line",  // 'line' | 'bar'
};

function getYearlyValues(chartData, year) {
  if (state.bucket === "all") {
    return chartData.monthly.by_year[year];
  }
  return chartData.monthly.monthly_by_bucket[year][state.bucket];
}

async function init() {
  const res = await fetch("./data/garmin_activities.json");
  if (!res.ok) {
    appEl.innerHTML = `<div class="banner banner-error">
      Run <code>uv run python get-garmin.py</code> to refresh data.
    </div>`;
    return;
  }
  let chartData;
  try {
    chartData = await res.json();
  } catch (e) {
    appEl.innerHTML = `<div class="banner banner-error">Corrupt JSON: ${e.message}</div>`;
    return;
  }
  renderApp(chartData);
}

function renderApp(chartData) {
  // === Build page chrome (matches legacy garmin_activities_trend.html structure) ===
  const page = document.createElement("div");
  page.className = "page";
  appEl.appendChild(page);

  const topbar = document.createElement("header");
  topbar.className = "topbar";
  page.appendChild(topbar);

  const topbarInner = document.createElement("div");
  topbarInner.className = "topbar-inner";
  topbar.appendChild(topbarInner);

  const titleWrap = document.createElement("div");
  const title = document.createElement("h1");
  title.textContent = "Garmin Monthly Trends";
  titleWrap.appendChild(title);
  topbarInner.appendChild(titleWrap);

  const summary = document.createElement("div");
  summary.className = "summary";
  topbarInner.appendChild(summary);

  function makeMetric(label, id, { compact = false, extraClass = "" } = {}) {
    const m = document.createElement("div");
    m.className = extraClass ? `metric ${extraClass}` : "metric";
    const lbl = document.createElement("div");
    lbl.className = "metric-label";
    lbl.textContent = label;
    const val = document.createElement("div");
    val.className = compact ? "metric-value compact" : "metric-value";
    val.id = id;
    m.appendChild(lbl);
    m.appendChild(val);
    summary.appendChild(m);
    return val;
  }

  const garminUser = makeMetric("Garmin User", "garminUser", { compact: true, extraClass: "user-metric" });
  const monthlyMileage = makeMetric("Monthly Mileage", "monthlyMileage", { compact: true, extraClass: "mileage-metric" });
  const monthlyActivities = makeMetric("Monthly Count", "monthlyActivities");
  const activityCount = makeMetric("Total Activities", "activityCount");
  const yearCount = makeMetric("Total Years", "yearCount");

  const main = document.createElement("main");
  page.appendChild(main);

  const controls = document.createElement("div");
  controls.className = "controls";
  main.appendChild(controls);

  function controlField(label, id, extraClass = "") {
    const wrap = document.createElement("div");
    wrap.className = extraClass ? `control-field ${extraClass}` : "control-field";
    const lbl = document.createElement("label");
    lbl.setAttribute("for", id);
    lbl.textContent = label;
    wrap.appendChild(lbl);
    const sel = document.createElement("select");
    sel.id = id;
    wrap.appendChild(sel);
    controls.appendChild(wrap);
    return sel;
  }

  const yearSelect = controlField("Year", "yearSelect");
  const monthSelect = controlField("Month", "monthSelect");
  const parameterSelect = controlField("Parameter", "parameterSelect", "parameter-field");

  function renderBucketChips(chartData) {
    const wrap = document.createElement("div");
    wrap.className = "control-field";
    const label = document.createElement("label");
    label.textContent = "Distance";
    wrap.appendChild(label);
    const chips = document.createElement("div");
    chips.className = "chip-group";
    const buckets = ["all", ...chartData.meta.distance_buckets.filter(b => b !== "all")];
    for (const b of buckets) {
      const chip = document.createElement("button");
      chip.textContent = b === "all" ? "All" : `${b} km`;
      chip.className = "chip" + (state.bucket === b ? " chip-active" : "");
      chip.addEventListener("click", () => {
        state.bucket = b;
        renderBucketChips(chartData);  // re-render to update active class
        renderChart(chartData);
      });
      chips.appendChild(chip);
    }
    wrap.appendChild(chips);
    return wrap;
  }

  controls.appendChild(renderBucketChips(chartData));

  const chartShell = document.createElement("div");
  chartShell.className = "chart-shell";
  main.appendChild(chartShell);

  const chartDiv = document.createElement("div");
  chartDiv.id = "chart";
  chartShell.appendChild(chartDiv);

  // === Legacy chart logic (verbatim from garmin_activities_trend.html) ===

  function populateSelect(selectElem, values) {
    values.forEach((value) => {
      const option = document.createElement("option");
      if (Array.isArray(value)) {
        option.value = value[0];
        option.textContent = value[1];
      } else {
        option.value = value;
        option.textContent = value;
      }
      selectElem.appendChild(option);
    });
  }

  function getYLabel(parameter) {
    if (parameter === "Avg Distance") {
      return "Kilometers";
    }
    if (parameter === "Heart Rate" || parameter === "maxHR") {
      return "Beats per minute";
    }
    if (parameter === "Pace") {
      return "Seconds per km";
    }
    if (parameter === "Duration") {
      return "Seconds";
    }
    if (parameter === "Calories") {
      return "Calories";
    }
    if (parameter === "avgElevation" || parameter === "maxElevation") {
      return "Elevation";
    }
    if (parameter === "MaxSpeed") {
      return "Speed";
    }
    if (parameter === "Vo2maxvalue") {
      return "VO2 max";
    }
    return parameter;
  }

  function isSecondsParameter(parameter) {
    return parameter === "Pace" || parameter === "Duration";
  }

  function formatSeconds(totalSeconds) {
    if (totalSeconds === null || totalSeconds === undefined || Number.isNaN(totalSeconds)) {
      return "";
    }
    const roundedSeconds = Math.round(totalSeconds);
    const hours = Math.floor(roundedSeconds / 3600);
    const minutes = Math.floor((roundedSeconds % 3600) / 60);
    const seconds = roundedSeconds % 60;
    return [hours, minutes, seconds]
      .map((value) => String(value).padStart(2, "0"))
      .join(":");
  }

  function getSecondsTickConfig(yValues) {
    const numericValues = yValues.filter((value) => value !== null && value !== undefined);
    if (!numericValues.length) {
      return {};
    }

    const minValue = Math.min(...numericValues);
    const maxValue = Math.max(...numericValues);
    if (minValue === maxValue) {
      return {
        tickmode: "array",
        tickvals: [minValue],
        ticktext: [formatSeconds(minValue)]
      };
    }

    const tickCount = 5;
    const step = (maxValue - minValue) / (tickCount - 1);
    const tickvals = Array.from({ length: tickCount }, (_, index) => minValue + step * index);
    return {
      tickmode: "array",
      tickvals,
      ticktext: tickvals.map(formatSeconds)
    };
  }

  function formatMileage(value) {
    if (value === null || value === undefined || Number.isNaN(value)) {
      return "0 km/mo";
    }
    return `${value.toLocaleString(undefined, { maximumFractionDigits: 1 })} km`;
  }

  function updateSummary(year, monthIndex) {
    const mileageValues = chartData.monthlyMileage[year] || [];
    const activityCounts = chartData.monthlyActivityCounts[year] || [];
    const selectedMileage = mileageValues[monthIndex] || 0;
    const selectedActivityCount = activityCounts[monthIndex] || 0;

    garminUser.textContent = chartData.garminUsername;
    monthlyMileage.textContent = formatMileage(selectedMileage);
    monthlyActivities.textContent = selectedActivityCount.toLocaleString();
  }

  function getLatestActiveMonthIndex(year) {
    const activityCounts = chartData.monthlyActivityCounts[year] || [];
    for (let index = activityCounts.length - 1; index >= 0; index -= 1) {
      if (activityCounts[index] > 0) {
        return index;
      }
    }
    return new Date().getMonth();
  }

  function getChartMargin() {
    if (window.matchMedia("(max-width: 420px)").matches) {
      return { l: 40, r: 8, t: 34, b: 32 };
    }
    if (window.matchMedia("(max-width: 720px)").matches) {
      return { l: 44, r: 10, t: 38, b: 34 };
    }
    return { l: 68, r: 28, t: 68, b: 56 };
  }

  function getTitleSize() {
    return window.matchMedia("(max-width: 720px)").matches ? 14 : 20;
  }

  function getTickSize() {
    return window.matchMedia("(max-width: 720px)").matches ? 10 : 12;
  }

  function renderChart() {
    const year = yearSelect.value;
    const monthIndex = Number(monthSelect.value);
    const parameter = parameterSelect.value;
    const yValues = chartData.values[year][parameter];
    const secondsParameter = isSecondsParameter(parameter);
    const hoverValues = secondsParameter ? yValues.map(formatSeconds) : yValues;
    const markerSizes = months.map((_, index) => index === monthIndex ? 12 : 8);
    const markerColors = months.map((_, index) => index === monthIndex ? "#1aa7e8" : "#8a8a8a");

    const trace = {
      x: months,
      y: yValues,
      customdata: hoverValues,
      type: "scatter",
      mode: "lines+markers",
      line: { color: "#171717", width: 3, shape: "spline", smoothing: 0.35 },
      marker: {
        size: markerSizes,
        color: markerColors,
        line: { color: "#ffffff", width: 2 }
      },
      fill: "tozeroy",
      fillcolor: "rgba(26, 167, 232, 0.12)",
      hovertemplate: secondsParameter
        ? "%{x}<br>%{customdata}<extra></extra>"
        : "%{x}<br>%{y:.2f}<extra></extra>"
    };

    const yAxisConfig = {
      title: { text: getYLabel(parameter), font: { color: "#707070", family: "Helvetica Neue, Avenir Next, Segoe UI, sans-serif" } },
      tickfont: { color: "#707070", size: getTickSize(), family: "Helvetica Neue, Avenir Next, Segoe UI, sans-serif" },
      gridcolor: "rgba(0, 0, 0, 0.1)",
      zerolinecolor: "rgba(0, 0, 0, 0.16)",
      ...(secondsParameter ? getSecondsTickConfig(yValues) : {})
    };

    const layout = {
      title: {
        text: `${parameter} Trend in ${year}`,
        x: 0,
        xanchor: "left",
        font: { size: getTitleSize(), color: "#151515", family: "Helvetica Neue, Avenir Next, Segoe UI, sans-serif" }
      },
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#fbfbfb",
      xaxis: {
        title: "",
        tickfont: { color: "#707070", size: getTickSize(), family: "Helvetica Neue, Avenir Next, Segoe UI, sans-serif" },
        gridcolor: "rgba(0, 0, 0, 0.1)",
        zeroline: false
      },
      yaxis: yAxisConfig,
      shapes: [{
        type: "line",
        xref: "x",
        yref: "paper",
        x0: months[monthIndex],
        x1: months[monthIndex],
        y0: 0,
        y1: 1,
        line: { color: "rgba(26, 167, 232, 0.58)", width: 2, dash: "dot" }
      }],
      hovermode: "x unified",
      margin: getChartMargin()
    };

    const config = {
      responsive: true,
      displaylogo: false,
      modeBarButtonsToRemove: ["lasso2d", "select2d"]
    };

    Plotly.newPlot("chart", [trace], layout, config);
    updateSummary(year, monthIndex);
  }

  activityCount.textContent = chartData.activityCount.toLocaleString();
  yearCount.textContent = chartData.years.length.toLocaleString();
  garminUser.textContent = chartData.garminUsername;
  populateSelect(yearSelect, chartData.years);
  populateSelect(monthSelect, months.map((month, index) => [index, month]));
  populateSelect(parameterSelect, chartData.parameters);
  yearSelect.value = chartData.years[chartData.years.length - 1];
  monthSelect.value = String(getLatestActiveMonthIndex(yearSelect.value));
  parameterSelect.value = chartData.parameters[0];
  yearSelect.addEventListener("change", () => {
    monthSelect.value = String(getLatestActiveMonthIndex(yearSelect.value));
    renderChart();
  });
  monthSelect.addEventListener("change", renderChart);
  parameterSelect.addEventListener("change", renderChart);
  window.addEventListener("resize", () => Plotly.Plots.resize("chart"));
  window.addEventListener("orientationchange", () => {
    setTimeout(() => {
      renderChart();
      Plotly.Plots.resize("chart");
    }, 250);
  });
  renderChart();
}

window.addEventListener("DOMContentLoaded", init);
