import pytest
from aggregate import build_monthly, build_monthly_by_bucket


def _act(year, month, distance_km=10.0, pace_s_per_km=300.0, duration_s=3000, avg_hr=150,
         activity_type="running"):
    return {
        "date": f"{year}-{month:02d}-15",
        "year": year, "month": month,
        "distance_km": distance_km, "pace_s_per_km": pace_s_per_km,
        "duration_s": duration_s, "avg_hr": avg_hr,
        "activity_type": activity_type,
    }


def test_build_monthly_basic():
    activities = [
        _act(2024, 3, distance_km=10.0, pace_s_per_km=300),
        _act(2024, 3, distance_km=12.0, pace_s_per_km=320),
        _act(2024, 5, distance_km=8.0, pace_s_per_km=290),
    ]
    monthly = build_monthly(activities)
    assert set(monthly.keys()) == {"2024"}
    pace = monthly["2024"]["pace_s_per_km"]
    assert pace[2] == pytest.approx(310.0)  # (300+320)/2
    assert pace[4] == 290.0
    assert all(v is None for v in pace if v not in (310.0, 290.0))


def test_build_monthly_spans_multiple_years():
    activities = [_act(2023, 12), _act(2024, 1)]
    monthly = build_monthly(activities)
    assert set(monthly.keys()) == {"2023", "2024"}
    # December 2023
    assert monthly["2023"]["activity_count"][11] == 1
    # January 2024
    assert monthly["2024"]["activity_count"][0] == 1


def test_build_monthly_skips_activities_without_date():
    activities = [_act(2024, 3), {**_act(2024, 4), "date": None}]
    monthly = build_monthly(activities)
    assert monthly["2024"]["activity_count"][2] == 1
    assert monthly["2024"]["activity_count"][3] == 0


def test_build_monthly_by_bucket_partitions_correctly():
    activities = [
        _act(2024, 3, distance_km=2.0),    # <3
        _act(2024, 3, distance_km=7.0),    # 5-10
        _act(2024, 3, distance_km=20.0),   # 15-25
    ]
    by_bucket = build_monthly_by_bucket(activities)
    three_five = by_bucket["2024"]["5-10"]
    assert three_five["activity_count"][2] == 1
    assert by_bucket["2024"]["<3"]["activity_count"][2] == 1
    assert by_bucket["2024"]["15-25"]["activity_count"][2] == 1
    # Bucket <3 should have zero count in 5-10 slot
    assert by_bucket["2024"]["<3"]["activity_count"][2] == 1
    assert by_bucket["2024"]["<3"]["activity_count"][3] == 0


def test_build_monthly_by_bucket_empty_buckets_have_zero_arrays():
    """A year with only one bucket populated should still return zero-filled arrays for the other buckets."""
    activities = [_act(2024, 3, distance_km=10.0)]  # only one running activity in <3 or 5-10 depending on distance
    out = build_monthly_by_bucket(activities)
    # All 7 distance buckets + 'all' should be present
    assert set(out["2024"].keys()) == {"<3", "3-5", "5-10", "10-15", "15-25", "25-40", "40+", "all"}
    # Pick a bucket that doesn't have this activity (10 km is in '5-10', so '<3' is empty)
    empty_bucket = out["2024"]["<3"]
    assert empty_bucket is not None, "empty buckets should be zero-filled arrays, not None"
    assert empty_bucket["activity_count"] == [0] * 12
    assert empty_bucket["pace_s_per_km"] == [None] * 12
