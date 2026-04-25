import os
import json
from datetime import datetime
from collections import defaultdict
from pathlib import Path

missing_packages = []
try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    missing_packages.append("python-dotenv")
    def load_dotenv():
        return False

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
    GarminConnectConnectionError = Exception
    GarminConnectTooManyRequestsError = Exception

# Load environment variables from likely .env locations
script_dir = Path(__file__).resolve().parent
repo_root = script_dir.parent
dotenv_loaded = False
for env_path in [script_dir / ".env", repo_root / ".env"]:
    if env_path.exists():
        load_dotenv(dotenv_path=env_path)
        dotenv_loaded = True
        break

# Fetch credentials from environment variables
username = os.getenv("GARMIN_USERNAME")
password = os.getenv("GARMIN_PASSWORD")

# Ensure credentials are loaded
if not username or not password:
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

if missing_packages:
    package_list = " ".join(sorted(set(missing_packages)))
    raise ModuleNotFoundError(
        "Missing Python package(s): "
        f"{package_list}\n"
        "Install with:\n"
        f"python3 -m pip install {package_list}"
    )

# Output file name
output_file = "garmin_activities_formatted.xlsx"
chart_output_file = "garmin_activities_trend.html"

def fetch_activities(client, max_activities=5000):
    """
    Fetch activities from Garmin Connect.
    :param client: Garmin client instance
    :param max_activities: Maximum number of activities to fetch
    :return: List of activities
    """
    activities = []
    start = 0
    limit = 100  # Garmin API supports up to 100 activities per request

    while start < max_activities:
        print(f"Fetching activities {start} to {start + limit}...")
        batch = client.get_activities(start, limit)
        if not batch:
            break
        activities.extend(batch)
        start += limit
    return activities

def format_speed(speed_mps):
    """
    Convert speed from meters/second to min:sec/km.
    :param speed_mps: Speed in meters per second
    :return: Speed in min:sec/km as a string (e.g., "5:30")
    """
    if speed_mps > 0:
        # Calculate total minutes per km
        total_minutes = 1000 / (speed_mps * 60)
        
        # Split into minutes and seconds
        minutes = int(total_minutes)
        seconds = int((total_minutes - minutes) * 60)
        
        # Format as "MM:SS"
        return f"{minutes}:{seconds:02d}"
    return None

def format_duration(duration_seconds):
    """
    Convert duration from seconds to HH:MM:SS format
    :param duration_seconds: Duration in seconds
    :return: Duration as string in HH:MM:SS format
    """
    if duration_seconds:
        hours = int(duration_seconds // 3600)
        minutes = int((duration_seconds % 3600) // 60)
        seconds = int(duration_seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    return None

def duration_to_seconds(duration_value):
    """
    Convert HH:MM:SS formatted duration string to total seconds.
    :param duration_value: Duration string
    :return: Duration in seconds (float) or None
    """
    if not duration_value or not isinstance(duration_value, str):
        return None
    parts = duration_value.split(":")
    if len(parts) != 3:
        return None
    try:
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = int(parts[2])
        return float(hours * 3600 + minutes * 60 + seconds)
    except ValueError:
        return None

def pace_to_seconds_per_km(pace_value):
    """
    Convert min:sec pace string to total seconds per km.
    :param pace_value: Pace string (M:SS)
    :return: Pace in seconds per km (float) or None
    """
    if not pace_value or not isinstance(pace_value, str):
        return None
    parts = pace_value.split(":")
    if len(parts) != 2:
        return None
    try:
        minutes = int(parts[0])
        seconds = int(parts[1])
        return float(minutes * 60 + seconds)
    except ValueError:
        return None

def to_numeric_chart_value(field, value):
    """
    Normalize activity field values into numeric values for charting.
    :param field: Field name
    :param value: Field value
    :return: Numeric value (float) or None if not chartable
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        if field in ["duration", "elapsedDuration", "movingDuration"]:
            return duration_to_seconds(value)
        if field == "Pace (min/km)":
            return pace_to_seconds_per_km(value)
        try:
            return float(value)
        except ValueError:
            return None
    return None

def build_monthly_trend_data(formatted_activities):
    """
    Build monthly averages by year and parameter.
    :param formatted_activities: List of formatted activity dictionaries
    :return: Dict for chart rendering: {years, parameters, values}
    """
    aggregates = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"sum": 0.0, "count": 0})))
    parameter_set = set()

    for activity in formatted_activities:
        date_value = activity.get("Date")
        if not date_value:
            continue
        try:
            dt = datetime.strptime(date_value, "%Y-%m-%d")
        except (ValueError, TypeError):
            continue

        year = str(dt.year)
        month = dt.month

        for field, raw_value in activity.items():
            if field in ["Date", "Time"]:
                continue
            numeric_value = to_numeric_chart_value(field, raw_value)
            if numeric_value is None:
                continue
            aggregates[year][field][month]["sum"] += numeric_value
            aggregates[year][field][month]["count"] += 1
            parameter_set.add(field)

    years = sorted(aggregates.keys())
    parameters = sorted(parameter_set)
    values = {}
    for year in years:
        values[year] = {}
        for parameter in parameters:
            monthly_values = []
            for month in range(1, 13):
                month_stats = aggregates[year][parameter].get(month)
                if month_stats and month_stats["count"] > 0:
                    monthly_values.append(round(month_stats["sum"] / month_stats["count"], 3))
                else:
                    monthly_values.append(None)
            values[year][parameter] = monthly_values

    return {
        "years": years,
        "parameters": parameters,
        "values": values
    }

def create_interactive_trend_chart(formatted_activities, output_html):
    """
    Create an interactive monthly trend line chart with year and parameter selectors.
    :param formatted_activities: List of formatted activity dictionaries
    :param output_html: Output HTML path
    """
    chart_data = build_monthly_trend_data(formatted_activities)
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
    body {{
      font-family: Arial, sans-serif;
      margin: 24px;
    }}
    .controls {{
      display: flex;
      gap: 16px;
      flex-wrap: wrap;
      margin-bottom: 20px;
      align-items: center;
    }}
    label {{
      font-weight: 600;
    }}
    select {{
      padding: 6px 8px;
      min-width: 220px;
    }}
    #chart {{
      width: 100%;
      height: 560px;
    }}
  </style>
</head>
<body>
  <h2>Garmin Monthly Trends</h2>
  <div class="controls">
    <div>
      <label for="yearSelect">Year</label><br />
      <select id="yearSelect"></select>
    </div>
    <div>
      <label for="parameterSelect">Parameter</label><br />
      <select id="parameterSelect"></select>
    </div>
  </div>
  <div id="chart"></div>

  <script>
    const chartData = {json_data};
    const months = {json_months};
    const yearSelect = document.getElementById("yearSelect");
    const parameterSelect = document.getElementById("parameterSelect");

    function populateSelect(selectElem, values) {{
      values.forEach((value) => {{
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        selectElem.appendChild(option);
      }});
    }}

    function getYLabel(parameter) {{
      if (["duration", "elapsedDuration", "movingDuration"].includes(parameter)) {{
        return "Seconds";
      }}
      if (parameter === "Pace (min/km)") {{
        return "Seconds per km";
      }}
      return parameter;
    }}

    function renderChart() {{
      const year = yearSelect.value;
      const parameter = parameterSelect.value;
      const yValues = chartData.values[year][parameter];

      const trace = {{
        x: months,
        y: yValues,
        type: "scatter",
        mode: "lines+markers",
        line: {{ width: 3 }},
        marker: {{ size: 7 }}
      }};

      const layout = {{
        title: `${{parameter}} Trend in ${{year}}`,
        xaxis: {{ title: "Month" }},
        yaxis: {{ title: getYLabel(parameter) }},
        hovermode: "x unified",
        margin: {{ l: 60, r: 20, t: 60, b: 60 }}
      }};

      Plotly.newPlot("chart", [trace], layout, {{ responsive: true }});
    }}

    populateSelect(yearSelect, chartData.years);
    populateSelect(parameterSelect, chartData.parameters);
    yearSelect.value = chartData.years[0];
    parameterSelect.value = chartData.parameters[0];
    yearSelect.addEventListener("change", renderChart);
    parameterSelect.addEventListener("change", renderChart);
    renderChart();
  </script>
</body>
</html>
"""

    with open(output_html, "w", encoding="utf-8") as file:
        file.write(html_content)
    print(f"Interactive trend chart saved to {output_html}")

def flatten_and_format_activity(activity):
    """
    Flatten and format an activity for Excel-friendly output.
    :param activity: Dictionary representing an activity
    :return: Flattened and formatted dictionary
    """
    flattened = {}
    for key, value in activity.items():
        if key == "startTimeLocal":
            try:
                # Try different datetime formats
                try:
                    # Try the T format first
                    value = value.split('.')[0]  # Remove milliseconds if present
                    dt = datetime.strptime(value, "%Y-%m-%dT%H:%M:%S")
                except ValueError:
                    # Try the space format
                    dt = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
                
                flattened["Date"] = dt.strftime("%Y-%m-%d")
                flattened["Time"] = dt.strftime("%H:%M:%S")
            except ValueError as e:
                print(f"Error parsing datetime: {e} for value: {value}")
                flattened["Date"] = None
                flattened["Time"] = None
        elif key == "averageSpeed":
            # Convert speed to min/km format
            flattened["Pace (min/km)"] = format_speed(value)
        elif key == "distance":
            # Convert distance to kilometers
            flattened["Distance (km)"] = round(value / 1000, 2) if value else None
        elif key == "averageHR":
            # Convert heart rate to integer
            flattened["Heart Rate (bpm)"] = int(value) if value else None
        elif key in ["duration", "elapsedDuration", "movingDuration"]:
            # Convert duration from seconds to HH:MM:SS
            flattened[key] = format_duration(value)
        elif isinstance(value, (dict, list)):
            # Convert nested structures to a JSON string
            flattened[key] = json.dumps(value, ensure_ascii=False)
        else:
            flattened[key] = value
    return flattened

try:
    # Log in to Garmin Connect
    client = Garmin(username, password)
    client.login()
    print("Logged in successfully!")

    # Fetch all activities
    print("Fetching activities...")
    activities = fetch_activities(client)

    if activities:
        # Dynamically gather all unique fieldnames across all activities
        all_fieldnames = set()
        formatted_activities = []
        for activity in activities:
            formatted_activity = flatten_and_format_activity(activity)
            formatted_activities.append(formatted_activity)
            all_fieldnames.update(formatted_activity.keys())

        # Create an Excel workbook and sheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Garmin Activities"

        # Write headers
        headers = sorted(all_fieldnames)
        ws.append(headers)

        # Write activity data
        for activity in formatted_activities:
            row = []
            for field in headers:
                value = activity.get(field, "")
                if field == "Date" and value:
                    try:
                        value = datetime.strptime(value, "%Y-%m-%d").date()
                    except (ValueError, TypeError) as e:
                        print(f"Error converting date: {e} for value: {value}")
                        value = None
                elif field == "Time" and value:
                    try:
                        value = datetime.strptime(value, "%H:%M:%S").time()
                    except (ValueError, TypeError) as e:
                        print(f"Error converting time: {e} for value: {value}")
                        value = None
                row.append(value)
            ws.append(row)

        # Format specific columns for Excel
        for col in ws.iter_cols(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
            header = ws.cell(row=1, column=col[0].column).value
            if header == "Date":
                for cell in col:
                    if cell.value:
                        cell.number_format = "yyyy-mm-dd"
            elif header == "Time":
                for cell in col:
                    if cell.value:
                        cell.number_format = "hh:mm:ss"
            elif header == "Pace (min/km)":
                for cell in col:
                    cell.number_format = "[mm]:ss"
            elif header in ["duration", "elapsedDuration", "movingDuration"]:
                for cell in col:
                    cell.number_format = "[hh]:mm:ss"
            elif header == "Distance (km)":
                for cell in col:
                    cell.number_format = "0.00"
            elif header == "Heart Rate (bpm)":
                for cell in col:
                    cell.number_format = "0"

        # Save the Excel file
        wb.save(output_file)
        print(f"Activities saved to {output_file}!")
        create_interactive_trend_chart(formatted_activities, chart_output_file)
    else:
        print("No activities found.")

except GarminConnectConnectionError as e:
    print(f"Error connecting to Garmin Connect: {e}")
except GarminConnectTooManyRequestsError as e:
    print(f"Too many requests: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
