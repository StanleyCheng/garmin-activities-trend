import json
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

from transform import BUCKET_NAMES, to_bucket


_PARAMS = ("pace_s_per_km", "distance_km_total", "avg_hr", "duration_s", "activity_count")
_AGG_FUNCS = {
    "pace_s_per_km": ("mean", "pace_s_per_km"),
    "distance_km_total": ("sum", "distance_km"),
    "avg_hr": ("mean", "avg_hr"),
    "duration_s": ("sum", "duration_s"),
    "activity_count": ("count", None),
}


def _safe_num(d, key):
    v = d.get(key)
    return v if isinstance(v, (int, float)) else None


def _empty_year():
    """12-month arrays for each parameter, all None / 0."""
    return {p: [None] * 12 if p != "activity_count" else [0] * 12 for p in _PARAMS}


def _acc(month_arrays, params_to_track):
    """Helper to accumulate values into a month's arrays."""
    return month_arrays


def build_monthly(activities):
    """Return {year: {param: [12 vals]}} aggregating across ALL activities."""
    by_year = defaultdict(_empty_year)
    for a in activities:
        if a.get("date") is None:
            continue
        year = a.get("year")
        month = a.get("month")
        if year is None or month is None:
            continue
        yr = str(year)
        buckets = by_year[yr]
        for param, (agg, field) in _AGG_FUNCS.items():
            slot = buckets[param][month - 1]
            v = _safe_num(a, field) if field else 1
            if agg == "count":
                buckets["activity_count"][month - 1] = slot + 1
            elif agg == "sum":
                if v is not None:
                    buckets[param][month - 1] = (slot or 0) + v
            elif agg == "mean":
                if v is not None:
                    slot_sum, slot_n = (slot if isinstance(slot, tuple) else (0, 0))
                    new_sum = slot_sum + v
                    new_n = slot_n + 1
                    buckets[param][month - 1] = (new_sum, new_n)

    # Resolve mean tuples to actual means
    out = {}
    for yr, buckets in by_year.items():
        out[yr] = {}
        for param, vals in buckets.items():
            resolved = []
            for v in vals:
                if isinstance(v, tuple):
                    s, n = v
                    resolved.append(round(s / n, 2) if n else None)
                else:
                    resolved.append(v)
            out[yr][param] = resolved
    return out


def build_monthly_by_bucket(activities):
    """Partition first by year, then by bucket, then aggregate."""
    out = {}
    for yr_data in build_monthly(activities).keys():
        out[yr_data] = {b: None for b in (*BUCKET_NAMES, "all")}
        by_bucket = defaultdict(list)
        for a in activities:
            if str(a.get("year")) != yr_data:
                continue
            b = to_bucket(a.get("distance_km"))
            by_bucket[b].append(a)
        for bucket_name, acts in by_bucket.items():
            sub = build_monthly(acts)
            # `build_monthly` returns a dict keyed by year; we only have one year here
            out[yr_data][bucket_name] = sub.get(yr_data, _empty_year())
    return out


def build_payload(activities, drops, *, generated_at, garmin_username_masked):
    years = sorted({a["year"] for a in activities if a.get("year")})
    meta = {
        "generated_at": generated_at,
        "garmin_username_masked": garmin_username_masked,
        "activity_count_after_clean": len(activities),
        "activity_count_dropped": dict(drops),
        "year_range": [years[0], years[-1]] if years else [],
        "distance_buckets": [*BUCKET_NAMES, "all"],
        "params": list(_PARAMS),
    }
    monthly = build_monthly(activities)
    monthly_by_bucket = build_monthly_by_bucket(activities)
    return {
        "meta": meta,
        "activities": sorted(activities, key=lambda a: a.get("date") or ""),
        "monthly": {"by_year": monthly},
        "monthly_by_bucket": monthly_by_bucket,
    }


def save_payload(payload, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def now_iso():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")
