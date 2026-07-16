// viz/app.js — phase 2 baseline; phase 3 adds features.

// Year-range control uses two side-by-side <select> boxes (choice b).
// Swap to noUiSlider library when slider decision is finalized.
const MAX_OVERLAY_YEARS = 6;

// Client-side mirror of transform.to_bucket (Python is source of truth).
const BUCKET_UPPER_BOUNDS_CLIENT = [3, 5, 10, 15, 25, 40, Infinity];
const BUCKET_NAMES_CLIENT = ["<3", "3-5", "5-10", "10-15", "15-25", "25-40", "40+"];

function toBucketClient(km) {
  if (km == null) return "all";
  for (let i = 0; i < BUCKET_UPPER_BOUNDS_CLIENT.length; i++) {
    if (km < BUCKET_UPPER_BOUNDS_CLIENT[i]) return BUCKET_NAMES_CLIENT[i];
  }
  return "40+";
}

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

function rangeInclusive(from, to) {
  const result = [];
  for (let y = from; y <= to; y++) result.push(y);
  return result;
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

function computeTightBounds(values, reversed) {
  const present = values.filter((value) => typeof value === "number");
  if (!present.length) return { range: [0, 1] };
  const min = Math.min(...present);
  const max = Math.max(...present);
  const span = Math.max(max - min, 1);
  const bottom = span * 0.02;
  const top = span * 0.04;
  const lo = min - bottom;
  const hi = max + top;
  return reversed
    ? { range: [hi, lo] }
    : { range: [lo, hi] };
}

function showToast(msg) {
  const existing = document.getElementById("year-range-toast");
  if (existing) existing.remove();
  const toast = document.createElement("div");
  toast.id = "year-range-toast";
  toast.className = "toast";
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 4000);
}

function warnOnLongRange() {
  if (state.fromYear === null || state.toYear === null) return;
  const span = state.toYear - state.fromYear + 1;
  if (span > 7) {
    showToast(`Showing ${span} years; colors repeat. Recommended max ${MAX_OVERLAY_YEARS}.`);
  }
}

function getYearlyValues(chartData, year) {
  if (state.bucket === "all") {
    return chartData.monthly.by_year[year];
  }
  const monthlyByBucket = chartData.monthly.monthly_by_bucket || chartData.monthly_by_bucket;
  return monthlyByBucket?.[year]?.[state.bucket];
}

function filterActivities(chartData) {
  return chartData.activities.filter((a) => {
    if (a.year < state.fromYear || a.year > state.toYear) return false;
    if (state.bucket !== "all") {
      const b = toBucketClient(a.distance_km);
      if (b !== state.bucket) return false;
    }
    return true;
  });
}

function computeStats(filtered) {
  const paces = filtered.map((a) => a.pace_s_per_km).filter((v) => v !== null);
  const distances = filtered.map((a) => a.distance_km).filter((v) => v !== null);
  return {
    avgPace: paces.length ? paces.reduce((s, v) => s + v, 0) / paces.length : null,
    totalMileage: distances.reduce((s, v) => s + v, 0),
    activityCount: filtered.length,
    longestRun: distances.length ? Math.max(...distances) : null,
    bestPace: paces.length ? Math.min(...paces) : null,
  };
}

function formatPaceMMSS(seconds) {
  if (seconds === null || seconds === undefined || Number.isNaN(seconds)) return "—";
  const rounded = Math.round(seconds);
  const minutes = Math.floor(rounded / 60);
  const secs = rounded % 60;
  return `${minutes}:${String(secs).padStart(2, "0")}`;
}

function renderStats(stats) {
  const main = document.querySelector(".page > main");
  let footer = document.getElementById("stats-footer");
  if (!footer) {
    footer = document.createElement("footer");
    footer.id = "stats-footer";
    main.appendChild(footer);
  }
  footer.replaceChildren();

  const items = [
    { label: "Avg pace", value: formatPaceMMSS(stats.avgPace) },
    {
      label: "Total mileage",
      value: `${stats.totalMileage.toLocaleString(undefined, { maximumFractionDigits: 1 })} km`,
    },
    { label: "Activity count", value: stats.activityCount.toLocaleString() },
    {
      label: "Longest run",
      value: stats.longestRun !== null
        ? `${stats.longestRun.toLocaleString(undefined, { maximumFractionDigits: 1 })} km`
        : "—",
    },
    { label: "Best pace", value: formatPaceMMSS(stats.bestPace) },
  ];

  for (const item of items) {
    const card = document.createElement("div");
    card.className = "stat-card";
    const lbl = document.createElement("div");
    lbl.className = "stat-label";
    lbl.textContent = item.label;
    const val = document.createElement("div");
    val.className = "stat-value";
    val.textContent = item.value;
    card.appendChild(lbl);
    card.appendChild(val);
    footer.appendChild(card);
  }
}

function buildTraces(chartData) {
  const traces = [];
  const years = rangeInclusive(state.fromYear, state.toYear);
  const canInterpolateViridis = Plotly.d3 && typeof Plotly.d3.interpolateViridis === "function";
  const palette = years.length <= 7 && canInterpolateViridis
    ? years.map((_, i) => Plotly.d3.interpolateViridis(i / Math.max(1, years.length - 1)))
    : years.map((_, i) => `hsl(${(i * 360) / Math.max(1, years.length)}, 60%, 50%)`);

  for (let i = 0; i < years.length; i++) {
    const year = years[i];
    const values = getYearlyValues(chartData, String(year));
    if (!values) continue;
    const color = palette[i];

    if (state.showPace) {
      traces.push({
        x: months,
        y: values.pace_s_per_km,
        name: `${year} — Pace`,
        type: state.mode === "bar" ? "bar" : "scatter",
        mode: state.mode === "bar" ? undefined : "lines+markers",
        yaxis: "y2",
        line: { color, width: 2.5 },
        marker: { size: 6, color },
        hovertemplate: "%{x}<br>%{customdata}<extra></extra>",
        customdata: values.pace_s_per_km.map(formatSeconds),
      });
    }

    if (state.showDistance) {
      traces.push({
        x: months,
        y: values.distance_km_total,
        name: `${year} — Distance`,
        type: state.mode === "bar" ? "bar" : "scatter",
        mode: state.mode === "bar" ? undefined : "lines+markers",
        yaxis: "y",
        line: { color, width: 2.5 },
        marker: { size: 6, color },
        hovertemplate: "%{x}<br>%{y:.2f} km<extra></extra>",
      });
    }
  }
  return traces;
}

function buildLayout(chartData, traces) {
  const paceValues = traces
    .filter((trace) => trace.yaxis === "y2")
    .flatMap((trace) => trace.y || [])
    .filter((value) => value !== null && value !== undefined);

  const distValues = traces
    .filter((trace) => trace.yaxis === "y")
    .flatMap((trace) => trace.y || [])
    .filter((value) => value !== null && value !== undefined);

  return {
    title: `Pace vs Distance · ${state.fromYear}–${state.toYear}`,
    paper_bgcolor: "#ffffff",
    plot_bgcolor: "#fbfbfb",
    xaxis: { title: "", tickfont: { size: 11 } },
    yaxis: {
      title: "Distance (km)",
      ...computeTightBounds(distValues, false),
      autorange: false,
    },
    yaxis2: {
      title: "Pace (mm:ss / km)",
      ...computeTightBounds(paceValues, true),
      autorange: false,
      overlaying: "y",
      side: "right",
    },
    barmode: state.mode === "bar" ? "overlay" : undefined,
    showlegend: true,
    margin: { l: 60, r: 60, t: 60, b: 40 },
    hovermode: "x unified",
  };
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

  function renderBucketChips(chartData) {
    const wrap = document.createElement("div");
    wrap.className = "control-field";
    const label = document.createElement("label");
    label.textContent = "Distance";
    wrap.appendChild(label);
    const chips = document.createElement("div");
    chips.className = "chip-group";
    const buckets = ["all", ...chartData.meta.distance_buckets.filter(b => b !== "all")];
    const chipButtons = [];
    for (const b of buckets) {
      const chip = document.createElement("button");
      chip.textContent = b === "all" ? "All" : `${b} km`;
      chip.className = "chip" + (state.bucket === b ? " chip-active" : "");
      chip.addEventListener("click", () => {
        state.bucket = b;
        for (const btn of chipButtons) {
          btn.classList.toggle("chip-active", btn === chip);
        }
        renderChart();
      });
      chips.appendChild(chip);
      chipButtons.push(chip);
    }
    wrap.appendChild(chips);
    return wrap;
  }

  function renderYearRange(chartData) {
    const years = chartData.meta.year_range;  // [min, max]
    if (!years || !years.length) return document.createDocumentFragment();

    // Initial state: full range
    if (state.fromYear === null) {
      state.fromYear = years[0];
      state.toYear = years[1];
    }

    const wrap = document.createElement("div");
    wrap.className = "control-field";
    const lbl = document.createElement("label");
    lbl.textContent = "Year range";
    wrap.appendChild(lbl);

    // Two-select implementation (choice b)
    const from = document.createElement("select");
    const to = document.createElement("select");
    for (let y = years[0]; y <= years[1]; y++) {
      for (const sel of [from, to]) {
        const opt = document.createElement("option");
        opt.value = String(y);
        opt.textContent = String(y);
        sel.appendChild(opt);
      }
    }
    from.value = String(state.fromYear);
    to.value = String(state.toYear);

    from.addEventListener("change", () => {
      state.fromYear = Number(from.value);
      if (state.fromYear > state.toYear) state.toYear = state.fromYear;
      renderChart();
      warnOnLongRange();
    });
    to.addEventListener("change", () => {
      state.toYear = Number(to.value);
      if (state.toYear < state.fromYear) state.fromYear = state.toYear;
      renderChart();
      warnOnLongRange();
    });

    const row = document.createElement("div");
    row.className = "year-range-row";
    row.appendChild(from);
    row.appendChild(document.createTextNode("→"));
    row.appendChild(to);
    wrap.appendChild(row);
    return wrap;
  }

  function renderModeToggle(chartData) {
    const wrap = document.createElement("div");
    wrap.className = "control-field";
    const lbl = document.createElement("label");
    lbl.textContent = "Chart";
    wrap.appendChild(lbl);

    const seg = document.createElement("div");
    seg.className = "segmented";
    for (const mode of ["line", "bar"]) {
      const btn = document.createElement("button");
      btn.textContent = mode === "line" ? "Line" : "Bar";
      btn.className = state.mode === mode ? "seg-active" : "";
      btn.addEventListener("click", () => {
        state.mode = mode;
        // Bar mode restricted to single year (PLAN §4.3)
        if (mode === "bar" && state.fromYear !== state.toYear) {
          const snapped = state.toYear;
          state.fromYear = state.toYear = snapped;
          showToast(`Bar mode shows one year — narrowed to ${snapped}.`);
        }
        renderChart();
      });
      seg.appendChild(btn);
    }
    wrap.appendChild(seg);
    return wrap;
  }

  controls.appendChild(renderBucketChips(chartData));
  controls.appendChild(renderYearRange(chartData));
  controls.appendChild(renderModeToggle(chartData));
  warnOnLongRange();

  const chartShell = document.createElement("div");
  chartShell.className = "chart-shell";
  main.appendChild(chartShell);

  const chartDiv = document.createElement("div");
  chartDiv.id = "chart";
  chartShell.appendChild(chartDiv);

  function updateSummary() {
    const values = getYearlyValues(chartData, String(state.toYear));
    const activityCounts = values?.activity_count || [];
    let monthIndex = activityCounts.length - 1;
    while (monthIndex >= 0 && !activityCounts[monthIndex]) monthIndex -= 1;

    const selectedMileage = monthIndex >= 0 ? values.distance_km_total[monthIndex] || 0 : 0;
    const selectedActivityCount = monthIndex >= 0 ? activityCounts[monthIndex] || 0 : 0;
    monthlyMileage.textContent = `${selectedMileage.toLocaleString(undefined, { maximumFractionDigits: 1 })} km`;
    monthlyActivities.textContent = selectedActivityCount.toLocaleString();
  }

  function renderChart() {
    const traces = buildTraces(chartData);
    const layout = buildLayout(chartData, traces);
    Plotly.react(chartDiv, traces, layout, { responsive: true, displaylogo: false });
    updateSummary();
    renderStats(computeStats(filterActivities(chartData)));
  }

  const years = chartData.monthly.by_year || {};
  activityCount.textContent = (chartData.meta.activity_count_after_clean || 0).toLocaleString();
  yearCount.textContent = Object.keys(years).length.toLocaleString();
  garminUser.textContent = chartData.meta.garmin_username_masked || "";
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
