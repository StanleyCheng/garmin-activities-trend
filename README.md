# Garmin activity dashboard

This project fetches Garmin Connect activities, cleans run/walk/hike records,
writes an Excel export, and builds a static interactive Plotly dashboard.

## One-time setup

Requirements: Python 3.14 and [uv](https://docs.astral.sh/uv/).

```sh
uv sync --dev
cp .env.example .env
```

Edit `.env` with your Garmin Connect login. The file is ignored by Git.

The first successful login stores refreshable authentication tokens in
`~/.garminconnect/garmin_tokens.json`. Later refreshes reuse those tokens, so
they do not repeatedly submit your password to Garmin. Set `GARMINTOKENS` in
`.env` only if you want a different private location.

## Refresh to the latest Garmin data

```sh
uv run python get-garmin.py --max-activities 10000
```

If Garmin reports `429`, `CAPTCHA_REQUIRED`, or HTTP `403`, stop retrying.
Sign in at [Garmin Connect](https://connect.garmin.com) in a browser and
complete any challenge, wait for Garmin's login cooldown, then run the command
once. After that successful run, the saved token is reused automatically. If
Garmin asks for MFA during the command, enter the code at the terminal prompt.

The command fetches activities newest-first and stops when it reaches the
requested limit or Garmin returns no more records. It writes:

- `garmin_activities_formatted.xlsx` — cleaned activity rows.
- `viz/data/garmin_activities.json` — dashboard data and monthly aggregates.

The terminal also reports how many activities were dropped by each cleaning
rule. Increase `--max-activities` if the Garmin account contains more than
10,000 activities.

## Display the interactive chart

Serve `viz/` over HTTP; opening `index.html` directly does not allow its JSON
request in every browser.

```sh
uv run python -m http.server 8000 --directory viz
```

Open [http://localhost:8000](http://localhost:8000). The page supports year
range and distance-bucket filters, pace/distance visibility, line/bar mode,
hover details, zoom, and PNG export through Plotly. To show newer activities,
run `get-garmin.py` again and reload the page.

For a non-interactive image from the Excel file:

```sh
uv run python run-analysis.py garmin_activities_formatted.xlsx --chart-type line --output chart.png
```

## Privacy

The generated JSON contains activity dates, distances, pace, and heart-rate
data. It is ignored by Git and intended for local use unless you explicitly
decide to publish that history. Garmin credentials must never be placed in
`viz/` or committed.

## Checks

```sh
uv run python -m pytest
node --check viz/app.js
```
