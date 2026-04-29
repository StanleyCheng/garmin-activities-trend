import argparse
import json
import os
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path


DEFAULT_OUTPUT_FILE = "garmin_activities_formatted.xlsx"
DEFAULT_CHART_OUTPUT_FILE = "garmin_activities_trend.html"
DEFAULT_MAX_ACTIVITIES = 5000
DEFAULT_BATCH_SIZE = 100
MAX_BATCH_SIZE = 100
DURATION_FIELDS = {"duration", "elapsedDuration", "movingDuration"}
MIN_PACE_SECONDS_PER_KM = 3 * 60 + 45
MAX_PACE_SECONDS_PER_KM = 15 * 60
CHART_PARAMETERS = [
    {"label": "Avg Distance", "fields": ["Distance (km)", "distance"]},
    {"label": "Heart Rate", "fields": ["Heart Rate (bpm)", "averageHR", "avgHR"]},
    {"label": "Pace", "fields": ["Pace (min/km)"]},
    {"label": "aerobicTrainingEffect", "fields": ["aerobicTrainingEffect"]},
    {"label": "anaerobicTrainingEffect", "fields": ["anaerobicTrainingEffect"]},
    {"label": "avgElevation", "fields": ["avgElevation", "averageElevation"]},
    {"label": "Calories", "fields": ["Calories", "calories"]},
    {"label": "Duration", "fields": ["duration"]},
    {"label": "maxHR", "fields": ["maxHR", "maxHeartRate"]},
    {"label": "maxElevation", "fields": ["maxElevation"]},
    {"label": "MaxSpeed", "fields": ["MaxSpeed", "maxSpeed"]},
    {"label": "Vo2maxvalue", "fields": ["Vo2maxvalue", "vo2MaxValue", "vO2MaxValue", "VO2MaxValue"]},
]


missing_packages = []
try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    missing_packages.append("python-dotenv")
    load_dotenv = None

try:
    from openpyxl import Workbook
except ModuleNotFoundError:
    missing_packages.append("openpyxl")
    Workbook = None

try:
    from garminconnect import Garmin, GarminConnectConnectionError, GarminConnectTooManyRequestsError
except ModuleNotFoundError:
    missing_packages.append("garminconnect")
    Garmin = None

    class GarminConnectConnectionError(Exception):
        pass

    class GarminConnectTooManyRequestsError(Exception):
        pass


def parse_args():
    parser = argparse.ArgumentParser(
        description="Export Garmin Connect activities to Excel and an optional trend chart."
    )
    parser.add_argument(
        "--max-activities",
        type=positive_int,
        default=DEFAULT_MAX_ACTIVITIES,
        help=f"Maximum number of activities to fetch. Default: {DEFAULT_MAX_ACTIVITIES}",
    )
    parser.add_argument(
        "--batch-size",
        type=positive_int,
        default=DEFAULT_BATCH_SIZE,
        help=f"Activities per Garmin request, capped at {MAX_BATCH_SIZE}. Default: {DEFAULT_BATCH_SIZE}",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT_FILE,
        help=f"Excel output path. Default: {DEFAULT_OUTPUT_FILE}",
    )
    parser.add_argument(
        "--chart-output",
        default=DEFAULT_CHART_OUTPUT_FILE,
        help=f"HTML chart output path. Default: {DEFAULT_CHART_OUTPUT_FILE}",
    )
    parser.add_argument(
        "--no-chart",
        action="store_true",
        help="Skip HTML chart generation.",
    )
    return parser.parse_args()


def positive_int(value):
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"{value!r} is not an integer") from exc
    if parsed <= 0:
        raise argparse.ArgumentTypeError("value must be greater than zero")
    return parsed


def check_dependencies():
    if not missing_packages:
        return
    package_list = " ".join(sorted(set(missing_packages)))
    raise ModuleNotFoundError(
        "Missing Python package(s): "
        f"{package_list}\n"
        "Install with:\n"
        f"python3 -m pip install {package_list}"
    )


def load_credentials():
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent

    for env_path in [script_dir / ".env", repo_root / ".env"]:
        if env_path.exists():
            load_dotenv(dotenv_path=env_path)
            break

    username = os.getenv("GARMIN_USERNAME")
    password = os.getenv("GARMIN_PASSWORD")
    if username and password:
        return username, password

    raise ValueError(
        "Missing GARMIN_USERNAME or GARMIN_PASSWORD in environment variables.\n"
        "Create a .env file at either:\n"
        f"- {script_dir / '.env'}\n"
        f"- {repo_root / '.env'}\n"
        "with:\n"
        "GARMIN_USERNAME=your_email\n"
        "GARMIN_PASSWORD=your_password\n"
        "Or export them in your shell before running."
    )


def fetch_activities(client, max_activities=DEFAULT_MAX_ACTIVITIES, batch_size=DEFAULT_BATCH_SIZE, retries=3):
    activities = []
    start = 0
    limit = min(batch_size, MAX_BATCH_SIZE, max_activities)

    while start < max_activities:
        request_limit = min(limit, max_activities - start)
        print(f"Fetching activities {start} to {start + request_limit}...")

        for attempt in range(1, retries + 1):
            try:
                batch = client.get_activities(start, request_limit)
                break
            except GarminConnectTooManyRequestsError:
                raise
            except GarminConnectConnectionError:
                if attempt == retries:
                    raise
                delay = 2 ** (attempt - 1)
                print(f"Connection error. Retrying in {delay} second(s)...")
                time.sleep(delay)

        if not batch:
            break

        activities.extend(batch)
        fetched_count = len(batch)
        start += fetched_count

        if fetched_count < request_limit:
            break

    return activities


def safe_float(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_int(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return None


def seconds_to_excel_days(seconds):
    if seconds is None:
        return None
    return seconds / 86400


def parse_activity_datetime(value):
    if not value or not isinstance(value, str):
        return None

    normalized = value.split(".", 1)[0]
    for date_format in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(normalized, date_format)
        except ValueError:
            continue
    print(f"Error parsing datetime for value: {value}")
    return None


def speed_to_pace_seconds(speed_mps):
    speed = safe_float(speed_mps)
    if not speed or speed <= 0:
        return None
    return 1000 / speed


def normalize_duration_seconds(value):
    seconds = safe_float(value)
    if seconds is None or seconds < 0:
        return None
    return seconds


def normalize_activity(activity):
    normalized = {}

    for key, value in activity.items():
        if key == "startTimeLocal":
            dt = parse_activity_datetime(value)
            normalized["Date"] = dt.date() if dt else None
            normalized["Time"] = dt.time() if dt else None
        elif key == "averageSpeed":
            normalized["Pace (min/km)"] = speed_to_pace_seconds(value)
        elif key == "distance":
            distance = safe_float(value)
            normalized["Distance (km)"] = round(distance / 1000, 2) if distance is not None else None
        elif key == "averageHR":
            normalized["Heart Rate (bpm)"] = safe_int(value)
        elif key in DURATION_FIELDS:
            normalized[key] = normalize_duration_seconds(value)
        elif isinstance(value, (dict, list)):
            normalized[key] = json.dumps(value, ensure_ascii=False)
        else:
            normalized[key] = value

    return normalized


def pace_is_in_range(activity):
    pace_seconds = activity.get("Pace (min/km)")
    if pace_seconds is None:
        return True
    return MIN_PACE_SECONDS_PER_KM <= pace_seconds <= MAX_PACE_SECONDS_PER_KM


def filter_activities_by_pace(normalized_activities):
    filtered_activities = []
    excluded_count = 0

    for activity in normalized_activities:
        if pace_is_in_range(activity):
            filtered_activities.append(activity)
        else:
            excluded_count += 1

    return filtered_activities, excluded_count


def to_numeric_chart_value(field, value):
    if value is None or isinstance(value, bool):
        return None
    if field == "Time":
        return None
    if hasattr(value, "year") or hasattr(value, "hour"):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return safe_float(value)


def get_activity_value(activity, field_names):
    lower_key_lookup = {key.lower(): key for key in activity.keys()}
    for field_name in field_names:
        if field_name in activity:
            return activity[field_name]
        matching_key = lower_key_lookup.get(field_name.lower())
        if matching_key:
            return activity[matching_key]
    return None


def format_public_username(username):
    if not username:
        return "Garmin User"
    if "@" not in username:
        return username

    local_part, domain = username.split("@", 1)
    if len(local_part) <= 2:
        masked_local = local_part[0] + "*"
    else:
        masked_local = local_part[:2] + "*" * min(4, len(local_part) - 2)
    return f"{masked_local}@{domain}"


def build_monthly_trend_data(normalized_activities, garmin_username=None):
    aggregates = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"sum": 0.0, "count": 0})))
    monthly_mileage = defaultdict(lambda: defaultdict(float))
    monthly_activity_counts = defaultdict(lambda: defaultdict(int))
    years = set()

    for activity in normalized_activities:
        date_value = activity.get("Date")
        if not date_value:
            continue

        year = str(date_value.year)
        month = date_value.month
        years.add(year)
        monthly_activity_counts[year][month] += 1

        distance_value = to_numeric_chart_value(
            "Avg Distance",
            get_activity_value(activity, ["Distance (km)", "distance"]),
        )
        if distance_value is not None:
            monthly_mileage[year][month] += distance_value

        for parameter in CHART_PARAMETERS:
            raw_value = get_activity_value(activity, parameter["fields"])
            numeric_value = to_numeric_chart_value(parameter["label"], raw_value)
            if numeric_value is None:
                continue
            label = parameter["label"]
            aggregates[year][label][month]["sum"] += numeric_value
            aggregates[year][label][month]["count"] += 1

    years = sorted(years)
    parameters = [parameter["label"] for parameter in CHART_PARAMETERS]
    values = {}
    mileage_values = {}
    activity_count_values = {}
    for year in years:
        values[year] = {}
        mileage_values[year] = []
        activity_count_values[year] = []
        for parameter in parameters:
            monthly_values = []
            for month in range(1, 13):
                month_stats = aggregates[year][parameter].get(month)
                if month_stats and month_stats["count"] > 0:
                    monthly_values.append(round(month_stats["sum"] / month_stats["count"], 3))
                else:
                    monthly_values.append(None)
            values[year][parameter] = monthly_values
        for month in range(1, 13):
            mileage = monthly_mileage[year].get(month)
            mileage_values[year].append(round(mileage, 2) if mileage else None)
            activity_count_values[year].append(monthly_activity_counts[year].get(month, 0))

    return {
        "years": years,
        "parameters": parameters,
        "values": values,
        "monthlyMileage": mileage_values,
        "monthlyActivityCounts": activity_count_values,
        "activityCount": len(normalized_activities),
        "garminUsername": format_public_username(garmin_username),
    }


def create_interactive_trend_chart(normalized_activities, output_html, garmin_username=None):
    chart_data = build_monthly_trend_data(normalized_activities, garmin_username=garmin_username)
    if not chart_data["years"] or not chart_data["parameters"]:
        print("No numeric data available to build chart.")
        return

    month_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    json_data = json.dumps(chart_data)
    json_months = json.dumps(month_labels)

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Garmin Activity Trends</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>
    :root {{
      color-scheme: light;
      --black: #171717;
      --charcoal: #242424;
      --accent: #1aa7e8;
      --accent-dark: #0d78bd;
      --ink: #151515;
      --muted: #707070;
      --line: #d7d7d7;
      --panel: #ffffff;
      --page: #f4f4f4;
      --soft-accent: rgba(26, 167, 232, 0.14);
      --soft-gray: rgba(0, 0, 0, 0.06);
    }}
    * {{
      box-sizing: border-box;
    }}
    html {{
      -webkit-text-size-adjust: 100%;
    }}
    body {{
      margin: 0;
      background: var(--page);
      color: var(--ink);
      font-family: "Helvetica Neue", "Avenir Next", "Segoe UI", system-ui, sans-serif;
      line-height: 1.45;
    }}
    .page {{
      min-height: 100vh;
      padding-bottom: 24px;
    }}
    .topbar {{
      background: linear-gradient(135deg, var(--black), var(--charcoal));
      color: #ffffff;
      position: relative;
      overflow: hidden;
    }}
    .topbar::after {{
      content: "";
      position: absolute;
      inset: auto -8% -56px auto;
      width: min(420px, 56vw);
      height: 180px;
      background: linear-gradient(90deg, rgba(26, 167, 232, 0.92), rgba(255, 255, 255, 0.42));
      border-radius: 999px 0 0 0;
      opacity: 0.78;
    }}
    .topbar-inner {{
      width: min(1480px, calc(100% - 32px));
      margin: 0 auto;
      padding: 26px 0 24px;
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: flex-end;
      position: relative;
      z-index: 1;
    }}
    h1 {{
      margin: 0;
      flex: 1 1 auto;
      min-width: 0;
      max-width: none;
      font-size: clamp(30px, 4vw, 56px);
      line-height: 1.02;
      font-weight: 300;
      letter-spacing: 0;
      white-space: nowrap;
    }}
    .summary {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      justify-content: flex-end;
      flex: 0 1 780px;
    }}
    .metric {{
      min-width: 0;
      flex: 1 1 132px;
      border: 1px solid rgba(255, 255, 255, 0.18);
      border-radius: 14px;
      padding: 12px 13px 11px;
      background: rgba(255, 255, 255, 0.08);
      box-shadow: 0 16px 34px rgba(0, 0, 0, 0.22);
      backdrop-filter: blur(12px);
    }}
    .metric:nth-child(2) {{
      background: rgba(26, 167, 232, 0.16);
    }}
    .metric-label {{
      color: rgba(255, 255, 255, 0.76);
      font-size: clamp(9px, 0.72vw, 11px);
      font-weight: 500;
      text-transform: uppercase;
      white-space: nowrap;
    }}
    .metric-value {{
      margin-top: 3px;
      font-size: clamp(18px, 1.7vw, 26px);
      line-height: 1;
      font-weight: 300;
    }}
    .metric-value.compact {{
      max-width: 160px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      font-size: clamp(13px, 1.08vw, 16px);
      line-height: 1.2;
    }}
    main {{
      width: min(1180px, calc(100% - 32px));
      margin: 18px auto 0;
    }}
    .controls {{
      display: flex;
      gap: 14px;
      flex-wrap: wrap;
      align-items: flex-end;
      padding: 16px 18px;
      border: 1px solid var(--line);
      border-radius: 14px;
      background: var(--panel);
      margin-bottom: 16px;
      box-shadow: 0 14px 38px rgba(0, 0, 0, 0.06);
    }}
    .control-field {{
      min-width: 0;
    }}
    label {{
      display: block;
      margin-bottom: 6px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 500;
      text-transform: uppercase;
    }}
    select {{
      min-width: 240px;
      height: 42px;
      border: 1px solid #cfcfcf;
      border-radius: 10px;
      background: #ffffff;
      color: var(--ink);
      font-size: 15px;
      font-weight: 400;
      padding: 0 36px 0 12px;
    }}
    select:focus {{
      outline: 3px solid rgba(26, 167, 232, 0.24);
      border-color: var(--accent);
    }}
    .chart-shell {{
      border: 1px solid var(--line);
      border-radius: 18px;
      background: var(--panel);
      overflow: hidden;
      box-shadow: 0 18px 52px rgba(0, 0, 0, 0.08);
    }}
    #chart {{
      width: 100%;
      height: min(640px, calc(100vh - 220px));
      min-height: 430px;
    }}
    @media (max-width: 1240px) {{
      .topbar-inner {{
        align-items: stretch;
        flex-direction: column;
      }}
      .summary {{
        justify-content: flex-start;
        flex-basis: auto;
      }}
      .metric {{
        max-width: 220px;
      }}
    }}
    @media (max-width: 900px) {{
      .topbar-inner {{
        align-items: stretch;
        flex-direction: column;
      }}
      .summary {{
        justify-content: flex-start;
      }}
      .metric {{
        flex: 1 1 150px;
      }}
      #chart {{
        height: min(620px, calc(100vh - 250px));
      }}
    }}
    @media (max-width: 720px) {{
      .page {{
        padding-bottom: 10px;
      }}
      .topbar::after {{
        inset: auto -22% -64px auto;
        width: 320px;
        height: 150px;
        opacity: 0.62;
      }}
      .topbar-inner {{
        padding-top: 10px;
        padding-bottom: 10px;
        width: min(100% - 24px, 1180px);
        gap: 8px;
      }}
      h1 {{
        font-size: clamp(18px, 5.7vw, 28px);
        max-width: 100%;
      }}
      .summary {{
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 4px;
      }}
      .user-metric {{
        display: none;
      }}
      .metric {{
        min-width: 0;
        padding: 6px 5px;
        border-radius: 9px;
      }}
      .metric-label {{
        color: #ffffff;
        font-size: 7.5px;
        line-height: 1;
        font-weight: 500;
        letter-spacing: 0;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.58);
        white-space: nowrap;
      }}
      .metric-value {{
        font-size: 16px;
        font-weight: 300;
      }}
      .metric-value.compact {{
        max-width: none;
        font-size: 11px;
      }}
      main {{
        width: min(100% - 24px, 1180px);
        margin-top: 8px;
      }}
      .controls {{
        display: grid;
        grid-template-columns: 0.8fr 0.9fr 1.35fr;
        padding: 8px;
        gap: 6px;
        margin-bottom: 8px;
        border-radius: 12px;
      }}
      label {{
        margin-bottom: 3px;
        font-size: 9px;
      }}
      select {{
        min-width: 100%;
        width: 100%;
        height: 34px;
        border-radius: 8px;
        font-size: 12px;
        padding: 0 20px 0 7px;
      }}
      .controls > div {{
        width: auto;
      }}
      .chart-shell {{
        border-radius: 12px;
      }}
      #chart {{
        height: min(76vh, 620px);
        min-height: 470px;
      }}
    }}
    @media (max-width: 420px) {{
      .topbar-inner {{
        width: min(100% - 16px, 1180px);
      }}
      main {{
        width: min(100% - 16px, 1180px);
        margin-top: 6px;
      }}
      .controls {{
        grid-template-columns: 0.76fr 0.86fr 1.38fr;
        gap: 5px;
        padding: 7px;
      }}
      select {{
        height: 32px;
        font-size: 11px;
      }}
      #chart {{
        height: 78vh;
        min-height: 500px;
      }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <header class="topbar">
      <div class="topbar-inner">
        <div>
          <h1>Garmin Monthly Trends</h1>
        </div>
        <div class="summary">
          <div class="metric user-metric">
            <div class="metric-label">Garmin User</div>
            <div class="metric-value compact" id="garminUser">Garmin User</div>
          </div>
          <div class="metric mileage-metric">
            <div class="metric-label">Monthly Mileage</div>
            <div class="metric-value compact" id="monthlyMileage">0 km</div>
          </div>
          <div class="metric">
            <div class="metric-label">Monthly Count</div>
            <div class="metric-value" id="monthlyActivities">0</div>
          </div>
          <div class="metric">
            <div class="metric-label">Total Activities</div>
            <div class="metric-value" id="activityCount">0</div>
          </div>
          <div class="metric">
            <div class="metric-label">Total Years</div>
            <div class="metric-value" id="yearCount">0</div>
          </div>
        </div>
      </div>
    </header>

    <main>
      <div class="controls">
        <div class="control-field">
          <label for="yearSelect">Year</label>
          <select id="yearSelect"></select>
        </div>
        <div class="control-field">
          <label for="monthSelect">Month</label>
          <select id="monthSelect"></select>
        </div>
        <div class="control-field parameter-field">
          <label for="parameterSelect">Parameter</label>
          <select id="parameterSelect"></select>
        </div>
      </div>
      <div class="chart-shell">
        <div id="chart"></div>
      </div>
    </main>
  </div>

  <script>
    const chartData = {json_data};
    const months = {json_months};
    const yearSelect = document.getElementById("yearSelect");
    const monthSelect = document.getElementById("monthSelect");
    const parameterSelect = document.getElementById("parameterSelect");
    const activityCount = document.getElementById("activityCount");
    const yearCount = document.getElementById("yearCount");
    const garminUser = document.getElementById("garminUser");
    const monthlyMileage = document.getElementById("monthlyMileage");
    const monthlyActivities = document.getElementById("monthlyActivities");

    function populateSelect(selectElem, values) {{
      values.forEach((value) => {{
        const option = document.createElement("option");
        if (Array.isArray(value)) {{
          option.value = value[0];
          option.textContent = value[1];
        }} else {{
          option.value = value;
          option.textContent = value;
        }}
        selectElem.appendChild(option);
      }});
    }}

    function getYLabel(parameter) {{
      if (parameter === "Avg Distance") {{
        return "Kilometers";
      }}
      if (parameter === "Heart Rate" || parameter === "maxHR") {{
        return "Beats per minute";
      }}
      if (parameter === "Pace") {{
        return "Seconds per km";
      }}
      if (parameter === "Duration") {{
        return "Seconds";
      }}
      if (parameter === "Calories") {{
        return "Calories";
      }}
      if (parameter === "avgElevation" || parameter === "maxElevation") {{
        return "Elevation";
      }}
      if (parameter === "MaxSpeed") {{
        return "Speed";
      }}
      if (parameter === "Vo2maxvalue") {{
        return "VO2 max";
      }}
      return parameter;
    }}

    function isSecondsParameter(parameter) {{
      return parameter === "Pace" || parameter === "Duration";
    }}

    function formatSeconds(totalSeconds) {{
      if (totalSeconds === null || totalSeconds === undefined || Number.isNaN(totalSeconds)) {{
        return "";
      }}
      const roundedSeconds = Math.round(totalSeconds);
      const hours = Math.floor(roundedSeconds / 3600);
      const minutes = Math.floor((roundedSeconds % 3600) / 60);
      const seconds = roundedSeconds % 60;
      return [hours, minutes, seconds]
        .map((value) => String(value).padStart(2, "0"))
        .join(":");
    }}

    function getSecondsTickConfig(yValues) {{
      const numericValues = yValues.filter((value) => value !== null && value !== undefined);
      if (!numericValues.length) {{
        return {{}};
      }}

      const minValue = Math.min(...numericValues);
      const maxValue = Math.max(...numericValues);
      if (minValue === maxValue) {{
        return {{
          tickmode: "array",
          tickvals: [minValue],
          ticktext: [formatSeconds(minValue)]
        }};
      }}

      const tickCount = 5;
      const step = (maxValue - minValue) / (tickCount - 1);
      const tickvals = Array.from({{ length: tickCount }}, (_, index) => minValue + step * index);
      return {{
        tickmode: "array",
        tickvals,
        ticktext: tickvals.map(formatSeconds)
      }};
    }}

    function formatMileage(value) {{
      if (value === null || value === undefined || Number.isNaN(value)) {{
        return "0 km/mo";
      }}
      return `${{value.toLocaleString(undefined, {{ maximumFractionDigits: 1 }})}} km`;
    }}

    function updateSummary(year, monthIndex) {{
      const mileageValues = chartData.monthlyMileage[year] || [];
      const activityCounts = chartData.monthlyActivityCounts[year] || [];
      const selectedMileage = mileageValues[monthIndex] || 0;
      const selectedActivityCount = activityCounts[monthIndex] || 0;

      garminUser.textContent = chartData.garminUsername;
      monthlyMileage.textContent = formatMileage(selectedMileage);
      monthlyActivities.textContent = selectedActivityCount.toLocaleString();
    }}

    function getLatestActiveMonthIndex(year) {{
      const activityCounts = chartData.monthlyActivityCounts[year] || [];
      for (let index = activityCounts.length - 1; index >= 0; index -= 1) {{
        if (activityCounts[index] > 0) {{
          return index;
        }}
      }}
      return new Date().getMonth();
    }}

    function getChartMargin() {{
      if (window.matchMedia("(max-width: 420px)").matches) {{
        return {{ l: 40, r: 8, t: 34, b: 32 }};
      }}
      if (window.matchMedia("(max-width: 720px)").matches) {{
        return {{ l: 44, r: 10, t: 38, b: 34 }};
      }}
      return {{ l: 68, r: 28, t: 68, b: 56 }};
    }}

    function getTitleSize() {{
      return window.matchMedia("(max-width: 720px)").matches ? 14 : 20;
    }}

    function getTickSize() {{
      return window.matchMedia("(max-width: 720px)").matches ? 10 : 12;
    }}

    function renderChart() {{
      const year = yearSelect.value;
      const monthIndex = Number(monthSelect.value);
      const parameter = parameterSelect.value;
      const yValues = chartData.values[year][parameter];
      const secondsParameter = isSecondsParameter(parameter);
      const hoverValues = secondsParameter ? yValues.map(formatSeconds) : yValues;
      const markerSizes = months.map((_, index) => index === monthIndex ? 12 : 8);
      const markerColors = months.map((_, index) => index === monthIndex ? "#1aa7e8" : "#8a8a8a");

      const trace = {{
        x: months,
        y: yValues,
        customdata: hoverValues,
        type: "scatter",
        mode: "lines+markers",
        line: {{ color: "#171717", width: 3, shape: "spline", smoothing: 0.35 }},
        marker: {{
          size: markerSizes,
          color: markerColors,
          line: {{ color: "#ffffff", width: 2 }}
        }},
        fill: "tozeroy",
        fillcolor: "rgba(26, 167, 232, 0.12)",
        hovertemplate: secondsParameter
          ? "%{{x}}<br>%{{customdata}}<extra></extra>"
          : "%{{x}}<br>%{{y:.2f}}<extra></extra>"
      }};

      const yAxisConfig = {{
        title: {{ text: getYLabel(parameter), font: {{ color: "#707070", family: "Helvetica Neue, Avenir Next, Segoe UI, sans-serif" }} }},
        tickfont: {{ color: "#707070", size: getTickSize(), family: "Helvetica Neue, Avenir Next, Segoe UI, sans-serif" }},
        gridcolor: "rgba(0, 0, 0, 0.1)",
        zerolinecolor: "rgba(0, 0, 0, 0.16)",
        ...(secondsParameter ? getSecondsTickConfig(yValues) : {{}})
      }};

      const layout = {{
        title: {{
          text: `${{parameter}} Trend in ${{year}}`,
          x: 0,
          xanchor: "left",
          font: {{ size: getTitleSize(), color: "#151515", family: "Helvetica Neue, Avenir Next, Segoe UI, sans-serif" }}
        }},
        paper_bgcolor: "#ffffff",
        plot_bgcolor: "#fbfbfb",
        xaxis: {{
          title: "",
          tickfont: {{ color: "#707070", size: getTickSize(), family: "Helvetica Neue, Avenir Next, Segoe UI, sans-serif" }},
          gridcolor: "rgba(0, 0, 0, 0.1)",
          zeroline: false
        }},
        yaxis: yAxisConfig,
        shapes: [{{
          type: "line",
          xref: "x",
          yref: "paper",
          x0: months[monthIndex],
          x1: months[monthIndex],
          y0: 0,
          y1: 1,
          line: {{ color: "rgba(26, 167, 232, 0.58)", width: 2, dash: "dot" }}
        }}],
        hovermode: "x unified",
        margin: getChartMargin()
      }};

      const config = {{
        responsive: true,
        displaylogo: false,
        modeBarButtonsToRemove: ["lasso2d", "select2d"]
      }};

      Plotly.newPlot("chart", [trace], layout, config);
      updateSummary(year, monthIndex);
    }}

    activityCount.textContent = chartData.activityCount.toLocaleString();
    yearCount.textContent = chartData.years.length.toLocaleString();
    garminUser.textContent = chartData.garminUsername;
    populateSelect(yearSelect, chartData.years);
    populateSelect(monthSelect, months.map((month, index) => [index, month]));
    populateSelect(parameterSelect, chartData.parameters);
    yearSelect.value = chartData.years[chartData.years.length - 1];
    monthSelect.value = String(getLatestActiveMonthIndex(yearSelect.value));
    parameterSelect.value = chartData.parameters[0];
    yearSelect.addEventListener("change", () => {{
      monthSelect.value = String(getLatestActiveMonthIndex(yearSelect.value));
      renderChart();
    }});
    monthSelect.addEventListener("change", renderChart);
    parameterSelect.addEventListener("change", renderChart);
    window.addEventListener("resize", () => Plotly.Plots.resize("chart"));
    window.addEventListener("orientationchange", () => {{
      setTimeout(() => {{
        renderChart();
        Plotly.Plots.resize("chart");
      }}, 250);
    }});
    renderChart();
  </script>
</body>
</html>
"""

    with open(output_html, "w", encoding="utf-8") as file:
        file.write(html_content)
    print(f"Interactive trend chart saved to {output_html}")


def excel_value(field, value):
    if value is None:
        return None
    if field in DURATION_FIELDS or field == "Pace (min/km)":
        return seconds_to_excel_days(value)
    return value


def save_activities_to_excel(normalized_activities, output_file):
    all_fieldnames = set()
    for activity in normalized_activities:
        all_fieldnames.update(activity.keys())

    wb = Workbook()
    ws = wb.active
    ws.title = "Garmin Activities"

    headers = sorted(all_fieldnames)
    ws.append(headers)

    for activity in normalized_activities:
        ws.append([excel_value(field, activity.get(field)) for field in headers])

    apply_excel_formats(ws, headers)
    wb.save(output_file)
    print(f"Activities saved to {output_file}!")


def apply_excel_formats(ws, headers):
    for column_index, header in enumerate(headers, start=1):
        column_cells = next(ws.iter_cols(
            min_row=2,
            max_row=ws.max_row,
            min_col=column_index,
            max_col=column_index,
        ))
        for cell in column_cells:
            if cell.value is None:
                continue
            if header == "Date":
                cell.number_format = "yyyy-mm-dd"
            elif header == "Time":
                cell.number_format = "hh:mm:ss"
            elif header == "Pace (min/km)":
                cell.number_format = "[mm]:ss"
            elif header in DURATION_FIELDS:
                cell.number_format = "[hh]:mm:ss"
            elif header == "Distance (km)":
                cell.number_format = "0.00"
            elif header == "Heart Rate (bpm)":
                cell.number_format = "0"


def login(username, password):
    client = Garmin(username, password)
    client.login()
    print("Logged in successfully!")
    return client


def run(args):
    check_dependencies()
    username, password = load_credentials()
    client = login(username, password)

    print("Fetching activities...")
    activities = fetch_activities(
        client,
        max_activities=args.max_activities,
        batch_size=args.batch_size,
    )

    if not activities:
        print("No activities found.")
        return 0

    normalized_activities = [normalize_activity(activity) for activity in activities]
    normalized_activities, excluded_count = filter_activities_by_pace(normalized_activities)
    if excluded_count:
        print(
            f"Excluded {excluded_count} activit"
            f"{'y' if excluded_count == 1 else 'ies'} outside pace range "
            "3:45 to 15:00 per km."
        )
    if not normalized_activities:
        print("No activities left after pace filtering.")
        return 0

    save_activities_to_excel(normalized_activities, args.output)

    if not args.no_chart:
        create_interactive_trend_chart(normalized_activities, args.chart_output, garmin_username=username)

    return 0


def main():
    args = parse_args()
    try:
        return run(args)
    except GarminConnectTooManyRequestsError as exc:
        print(f"Too many requests from Garmin Connect. Try again later. Details: {exc}", file=sys.stderr)
        return 2
    except GarminConnectConnectionError as exc:
        print(f"Error connecting to Garmin Connect: {exc}", file=sys.stderr)
        return 2
    except (ModuleNotFoundError, ValueError) as exc:
        print(exc, file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Unexpected error: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
