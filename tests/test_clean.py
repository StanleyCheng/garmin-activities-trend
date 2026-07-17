import pytest
from transform import clean


def _act(**overrides):
    base = {
        "date": "2024-03-15", "year": 2024, "month": 3,
        "distance_km": 10.0, "pace_s_per_km": 300,
        "duration_s": 3000, "avg_hr": 150,
        "activity_type": "running",
    }
    base.update(overrides)
    return base


def test_clean_keeps_valid_running():
    kept, drops = clean([_act()])
    assert len(kept) == 1
    assert drops == {}


def test_clean_drops_unparseable_date():
    kept, drops = clean([_act(date=None)])
    assert kept == [] and drops == {"date": 1}


def test_clean_drops_no_pace():
    kept, drops = clean([_act(pace_s_per_km=None)])
    assert kept == [] and drops == {"no_pace": 1}


def test_clean_drops_too_fast_pace():
    kept, drops = clean([_act(pace_s_per_km=200)])  # 3:20 — under 3:45
    assert kept == [] and drops == {"pace": 1}


def test_clean_keeps_walk_at_20min_per_km():
    # 20 min/km = 1200 s/km — outside run band but inside walk band
    kept, drops = clean([_act(activity_type="walking", pace_s_per_km=1200)])
    assert len(kept) == 1


def test_clean_drops_walk_too_slow():
    # 26 min/km = 1560 s/km — past walk band (25:00 = 1500 s)
    kept, drops = clean([_act(activity_type="walking", pace_s_per_km=1560)])
    assert kept == [] and drops == {"pace": 1}


def test_clean_keeps_hike_at_25min_per_km():
    # 25 min/km = 1500 s/km — inside hike band (3:45–30:00)
    kept, drops = clean([_act(activity_type="hiking", pace_s_per_km=1500)])
    assert len(kept) == 1


def test_clean_drops_distance_below_500m():
    kept, drops = clean([_act(distance_km=0.49)])
    assert kept == [] and drops == {"distance": 1}


def test_clean_keeps_distance_at_500m():
    kept, drops = clean([_act(distance_km=0.5)])
    assert len(kept) == 1


def test_clean_keeps_distance_at_200km():
    kept, drops = clean([_act(distance_km=200.0)])
    assert len(kept) == 1


def test_clean_drops_distance_above_200km():
    kept, drops = clean([_act(distance_km=200.01)])
    assert kept == [] and drops == {"distance": 1}


def test_clean_drops_cycling():
    kept, drops = clean([_act(activity_type="cycling")])
    assert kept == [] and drops == {"type": 1}


def test_clean_keeps_track_running_subtype():
    kept, drops = clean([_act(activity_type="track_running")])
    assert len(kept) == 1


def test_clean_keeps_ultra_run():
    kept, drops = clean([_act(activity_type="ultra_run")])
    assert len(kept) == 1


def test_clean_drops_duration_too_short():
    kept, drops = clean([_act(duration_s=30)])
    assert kept == [] and drops == {"duration": 1}


def test_clean_drops_duration_too_long():
    kept, drops = clean([_act(duration_s=13 * 3600)])
    assert kept == [] and drops == {"duration": 1}


def test_clean_drops_hr_too_low():
    kept, drops = clean([_act(avg_hr=20)])
    assert kept == [] and drops == {"hr": 1}


def test_clean_drops_hr_too_high():
    kept, drops = clean([_act(avg_hr=240)])
    assert kept == [] and drops == {"hr": 1}


def test_clean_aggregates_drop_reasons():
    acts = [
        _act(date=None),
        _act(pace_s_per_km=None),
        _act(activity_type="cycling"),
        _act(avg_hr=20),
    ]
    kept, drops = clean(acts)
    assert drops == {"date": 1, "no_pace": 1, "type": 1, "hr": 1}