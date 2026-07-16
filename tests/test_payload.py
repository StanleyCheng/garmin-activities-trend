import pytest
from aggregate import build_payload, save_payload


def test_build_payload_basic_structure():
    activities = [{"date": "2024-03-15", "year": 2024, "month": 3, "distance_km": 10.0,
                   "pace_s_per_km": 300.0, "duration_s": 3000, "avg_hr": 150,
                   "activity_type": "running"}]
    drops = {"pace": 2, "no_pace": 1}
    payload = build_payload(
        activities, drops,
        generated_at="2026-07-16T10:00:00+00:00",
        garmin_username_masked="s***",
    )
    assert payload["meta"]["activity_count_after_clean"] == 1
    assert payload["meta"]["activity_count_dropped"] == drops
    assert payload["meta"]["garmin_username_masked"] == "s***"
    assert payload["meta"]["year_range"] == [2024, 2024]
    assert "2024" in payload["monthly"]["by_year"]


def test_build_payload_total_matches_sum():
    activities = [{"year": 2024, "month": 3, "distance_km": 10.0, "pace_s_per_km": 300.0,
                   "duration_s": 3000, "avg_hr": 150, "date": "2024-03-15",
                   "activity_type": "running"}]
    drops = {"pace": 3, "no_pace": 2, "distance": 1}
    p = build_payload(activities, drops, generated_at="2026-01-01T00:00:00+00:00",
                      garmin_username_masked="s***")
    raw_total = len(activities) + sum(drops.values())
    assert raw_total == p["meta"]["activity_count_after_clean"] + sum(p["meta"]["activity_count_dropped"].values())


def test_save_payload_writes_file(tmp_path):
    p = build_payload([], {}, generated_at="2026-01-01T00:00:00+00:00",
                      garmin_username_masked="s***")
    out = tmp_path / "out.json"
    save_payload(p, out)
    assert out.exists()
    import json
    assert json.loads(out.read_text())["meta"]["activity_count_after_clean"] == 0
