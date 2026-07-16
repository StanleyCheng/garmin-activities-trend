from datetime import datetime


def _safe_float(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _safe_int(value):
    v = _safe_float(value)
    if v is None:
        return None
    return int(round(v))


def _parse_datetime(value):
    """Parse Garmin's startTimeLocal. Returns datetime or None.

    Handles ISO with optional .SSSZ suffix and ±HH:MM offsets.
    """
    if not value or not isinstance(value, str):
        return None
    s = value.strip()
    # Garmin sometimes returns "Z" suffix; replace with +00:00
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


def _activity_type_key(activity):
    type_ = activity.get("activityType")
    if isinstance(type_, dict):
        return type_.get("typeKey")
    return type_


def normalize_activity(activity):
    """Normalize a raw Garmin activity into the cleaned-record schema.

    Returns a dict with: date, year, month, distance_km, pace_s_per_km,
    duration_s, avg_hr, max_hr (None if absent), elevation_m, vo2max,
    calories, activity_type.
    """
    out = {}
    dt = _parse_datetime(activity.get("startTimeLocal"))
    out["date"] = dt.date().isoformat() if dt else None
    out["year"] = dt.year if dt else None
    out["month"] = dt.month if dt else None

    distance_m = _safe_float(activity.get("distance"))
    out["distance_km"] = round(distance_m / 1000, 2) if distance_m is not None else None

    speed = _safe_float(activity.get("averageSpeed"))
    out["pace_s_per_km"] = round(1000 / speed, 2) if speed and speed > 0 else None

    out["duration_s"] = _safe_int(activity.get("duration"))
    out["avg_hr"] = _safe_int(activity.get("averageHR"))
    out["max_hr"] = _safe_int(activity.get("maxHeartRate"))
    out["calories"] = _safe_int(activity.get("calories"))
    out["elevation_m"] = _safe_float(activity.get("elevationGain"))
    out["vo2max"] = _safe_float(activity.get("vO2MaxValue"))
    out["activity_type"] = _activity_type_key(activity)
    return out


BUCKET_NAMES = ("<3", "3-5", "5-10", "10-15", "15-25", "25-40", "40+")
_BUCKET_UPPER_BOUNDS = (3, 5, 10, 15, 25, 40, float("inf"))


def to_bucket(distance_km):
    if distance_km is None:
        return "all"
    for upper, name in zip(_BUCKET_UPPER_BOUNDS, BUCKET_NAMES):
        if distance_km < upper:
            return name
    return "40+"
