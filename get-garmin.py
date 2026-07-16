"""CLI entry point — kept for backward compatibility.

Delegates to get_data, transform, and aggregate modules.
"""
import argparse
import os
import sys
from datetime import datetime, timezone

try:
    from dotenv import load_dotenv
except ModuleNotFoundError as exc:
    print(f"python-dotenv is required: {exc}", file=sys.stderr); sys.exit(1)

from get_data import fetch_activities, get_client, DEFAULT_MAX_ACTIVITIES, DEFAULT_BATCH_SIZE
from garminconnect import GarminConnectConnectionError, GarminConnectTooManyRequestsError
from transform import normalize_activity, clean, format_public_username
from aggregate import build_payload, save_payload, now_iso

DEFAULT_OUTPUT_FILE = "garmin_activities_formatted.xlsx"
DEFAULT_JSON_OUTPUT = "viz/data/garmin_activities.json"
DEFAULT_CHART_OUTPUT_FILE = None  # deprecated; legacy inline HTML generation removed in Phase 2


def _positive_int(value):
    try:
        n = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(exc)
    if n <= 0:
        raise argparse.ArgumentTypeError("must be > 0")
    return n


def _load_credentials():
    from pathlib import Path
    username = os.getenv("GARMIN_USERNAME")
    password = os.getenv("GARMIN_PASSWORD")
    if username and password:
        return username, password

    # Look for .env in cwd or script dir
    candidates = [Path.cwd() / ".env", Path(__file__).parent / ".env"]
    for env_path in candidates:
        if env_path.exists():
            load_dotenv(dotenv_path=env_path)
            username = os.getenv("GARMIN_USERNAME")
            password = os.getenv("GARMIN_PASSWORD")
            if username and password:
                return username, password

    raise ValueError("Missing GARMIN_USERNAME or GARMIN_PASSWORD — see README.")


def _save_excel(activities, path):
    # Kept mostly verbatim from the legacy script — refactored into
    # aggregate.py in a future phase.
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Garmin Activities"
    all_keys = set()
    for a in activities:
        all_keys.update(a.keys())
    headers = sorted(all_keys)
    ws.append(headers)
    for a in activities:
        ws.append([a.get(h) for h in headers])
    wb.save(path)


def _run(args):
    username, password = _load_credentials()
    client = get_client(username, password)
    print("Fetching activities...")
    raw = fetch_activities(client, max_activities=args.max_activities,
                            batch_size=args.batch_size)

    if not raw:
        print("No activities found.")
        return 0

    normalized = [normalize_activity(a) for a in raw]
    kept, drops = clean(normalized)

    if drops:
        print(f"Dropped activities by reason: {drops}")

    if not kept:
        print("No activities left after cleaning.")
        return 0

    _save_excel(kept, args.output)
    print(f"Excel saved to {args.output}")

    payload = build_payload(
        kept, drops,
        generated_at=now_iso(),
        garmin_username_masked=format_public_username(username),
    )
    save_payload(payload, args.json_output)
    print(f"JSON saved to {args.json_output}")

    return 0


def _parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Fetch Garmin activities → xlsx + JSON.")
    parser.add_argument("--max-activities", type=_positive_int, default=DEFAULT_MAX_ACTIVITIES)
    parser.add_argument("--batch-size", type=_positive_int, default=DEFAULT_BATCH_SIZE)
    parser.add_argument("--output", default=DEFAULT_OUTPUT_FILE)
    parser.add_argument("--json-output", default=DEFAULT_JSON_OUTPUT)
    parser.add_argument("--no-chart", action="store_true")
    return parser.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv)
    try:
        return _run(args)
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
    sys.exit(main())
