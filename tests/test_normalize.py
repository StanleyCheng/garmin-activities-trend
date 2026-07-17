import pytest

from transform import normalize_activity


def test_normalize_basic_running(synthetic_activity):
    out = normalize_activity(synthetic_activity)
    assert out["date"] == "2024-03-15"
    assert out["year"] == 2024
    assert out["month"] == 3
    assert out["distance_km"] == 10.0  # 10000 m = 10.0 km
    assert out["pace_s_per_km"] == pytest.approx(300.3, abs=0.5)  # 1000/3.33 ≈ 300.3
    assert out["avg_hr"] == 150
    assert out["activity_type"] == "running"
    assert out["duration_s"] == 3000


def test_normalize_zero_speed_returns_none_pace(synthetic_activity):
    raw = {**synthetic_activity, "averageSpeed": 0}
    out = normalize_activity(raw)
    assert out["pace_s_per_km"] is None


def test_normalize_missing_speed_returns_none_pace(synthetic_activity):
    raw = {k: v for k, v in synthetic_activity.items() if k != "averageSpeed"}
    out = normalize_activity(raw)
    assert out["pace_s_per_km"] is None


def test_normalize_iso_offset_z_parses(synthetic_activity):
    raw = {**synthetic_activity, "startTimeLocal": "2024-03-15T23:00:00.000Z"}
    out = normalize_activity(raw)
    assert out["date"] == "2024-03-15"
    assert out["year"] == 2024
    assert out["month"] == 3


def test_normalize_iso_offset_tz_parses(synthetic_activity):
    raw = {**synthetic_activity, "startTimeLocal": "2024-03-15T06:30:00+10:00"}
    out = normalize_activity(raw)
    assert out["date"] == "2024-03-15"


def test_normalize_unparseable_date_returns_none_date(synthetic_activity):
    raw = {**synthetic_activity, "startTimeLocal": "not-a-date"}
    out = normalize_activity(raw)
    assert out["date"] is None
    assert out["year"] is None
    assert out["month"] is None
